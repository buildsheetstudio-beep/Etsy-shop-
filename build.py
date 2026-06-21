"""
Builds the Net Worth Tracker & Wealth Dashboard in Google Sheets.
Run: python3 build.py
Requires token.json (OAuth token with spreadsheets + drive.file scope) in same dir.
"""
import json
import time
import requests

# ---------------------------------------------------------------- auth
tok = json.load(open("token.json"))
csec = json.load(open("client_secret.json"))["installed"]

API_ROOT = "https://sheets.googleapis.com/v4/spreadsheets"


def _refresh_token():
    r = requests.post(csec["token_uri"], data={
        "client_id": csec["client_id"],
        "client_secret": csec["client_secret"],
        "refresh_token": tok["refresh_token"],
        "grant_type": "refresh_token",
    })
    r.raise_for_status()
    tok["access_token"] = r.json()["access_token"]
    json.dump(tok, open("token.json", "w"), indent=2)


def _headers():
    return {"Authorization": f"Bearer {tok['access_token']}", "Content-Type": "application/json"}


def _request(method, url, **kwargs):
    resp = requests.request(method, url, headers=_headers(), **kwargs)
    if resp.status_code == 401:
        _refresh_token()
        resp = requests.request(method, url, headers=_headers(), **kwargs)
    if not resp.ok:
        raise RuntimeError(f"{method} {url} -> {resp.status_code}: {resp.text}")
    return resp.json() if resp.text else {}


class _Values:
    def __init__(self, spreadsheet_id):
        self.sid = spreadsheet_id

    def batch_update(self, body):
        return _request("POST", f"{API_ROOT}/{self.sid}/values:batchUpdate", json=body)


class _SS:
    def create(self, body):
        return _request("POST", API_ROOT, json=body)

    def batch_update(self, spreadsheet_id, body):
        return _request("POST", f"{API_ROOT}/{spreadsheet_id}:batchUpdate", json=body)

    def values(self, spreadsheet_id):
        return _Values(spreadsheet_id)


ss = _SS()

# ---------------------------------------------------------------- constants
SH_DASH, SH_MONTHLY, SH_YEARLY, SH_ACCOUNTS, SH_SETTINGS = 0, 1, 2, 3, 4

PALETTE = {
    "deep": "#13413B", "mid": "#2F9E7E", "gold": "#C8A951",
    "liability": "#C0594B", "liquid": "#3D9BC9", "illiquid": "#7E6BB0",
    "bg": "#F4F7F6", "alt": "#EAF1EE", "text": "#1E2A28",
    "muted": "#6B7C77", "border": "#D4DEDA",
}

CATEGORY_COLORS = {
    "Cash": "#3D9BC9", "Investments": "#2F9E7E", "Retirement": "#13413B",
    "Property": "#C8A951", "Vehicles": "#A8895B", "Crypto": "#7E6BB0",
    "Other Asset": "#9AAFA9", "Mortgage": "#C0594B", "Loan": "#D88A5A",
    "Credit Card": "#A84B5C", "Other Debt": "#B58392",
}
ASSET_CATEGORIES = ["Cash", "Investments", "Retirement", "Property", "Vehicles", "Crypto", "Other Asset"]
ALL_CATEGORIES = ASSET_CATEGORIES + ["Mortgage", "Loan", "Credit Card", "Other Debt"]
TYPES = ["Asset", "Liability"]
LIQUIDITY = ["Liquid", "Illiquid"]
YEARS = list(range(2023, 2033))

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def hex_to_rgb(h):
    h = h.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}


def linear(start, end, n=12):
    return [round(start + (end - start) * i / (n - 1)) for i in range(n)]


def col_letter(idx):
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- sample data (24 of 80 accounts)
SEED_ACCOUNTS = [
    ("Monzo Current", "Asset", "Cash", "Liquid", linear(3000, 3800)),
    ("Chase Savings", "Asset", "Cash", "Liquid", linear(7000, 8200)),
    ("Premium Bonds", "Asset", "Cash", "Liquid", [5000] * 12),
    ("Vanguard S&S ISA", "Asset", "Investments", "Liquid", linear(35000, 40000)),
    ("Trading 212", "Asset", "Investments", "Liquid", linear(16000, 18500)),
    ("Workplace Pension", "Asset", "Retirement", "Illiquid", linear(50000, 53500)),
    ("SIPP", "Asset", "Retirement", "Illiquid", linear(29000, 31500)),
    ("Main Residence", "Asset", "Property", "Illiquid", [155000] * 12),
    ("Coinbase", "Asset", "Crypto", "Liquid", [6000, 6100, 6300, 6600, 6900, 4800, 5400, 5900, 6200, 6400, 6600, 6800]),
    ("Car – VW Golf", "Asset", "Vehicles", "Illiquid", linear(13000, 11500)),
    ("Nationwide ISA", "Asset", "Cash", "Liquid", linear(8000, 9200)),
    ("Marcus Savings", "Asset", "Cash", "Liquid", linear(5500, 6300)),
    ("Fidelity SIPP", "Asset", "Retirement", "Illiquid", linear(18000, 19800)),
    ("iShares Global Tracker", "Asset", "Investments", "Liquid", linear(9500, 11000)),
    ("Binance", "Asset", "Crypto", "Liquid", linear(2500, 3000)),
    ("Motorbike", "Asset", "Vehicles", "Illiquid", linear(3500, 3000)),
    ("Antique Watch Collection", "Asset", "Other Asset", "Illiquid", linear(2000, 2100)),
    ("Emergency Fund", "Asset", "Cash", "Liquid", linear(6000, 6500)),
    ("NS&I Bonds", "Asset", "Cash", "Liquid", [3500] * 12),
    ("Help to Buy ISA", "Asset", "Cash", "Liquid", linear(2500, 2800)),
    ("Mortgage", "Liability", "Mortgage", "Illiquid", linear(160000, 152500)),
    ("Student Loan", "Liability", "Loan", "Illiquid", linear(22000, 20800)),
    ("Barclaycard", "Liability", "Credit Card", "Liquid", [2000, 2200, 2600, 3000, 2700, 2500, 3200, 2900, 2400, 2100, 1900, 1700]),
    ("Car Finance", "Liability", "Loan", "Illiquid", linear(16000, 13000)),
]
N_ACCOUNTS = 80

# =================================================================
# 1. CREATE SPREADSHEET WITH 5 SHEETS
# =================================================================
def create_spreadsheet():
    body = {
        "properties": {"title": "Net Worth Tracker & Wealth Dashboard"},
        "sheets": [
            {"properties": {"sheetId": SH_DASH, "title": "Dashboard", "gridProperties": {"rowCount": 60, "columnCount": 14}}},
            {"properties": {"sheetId": SH_MONTHLY, "title": "Monthly Tracker", "gridProperties": {"rowCount": 90, "columnCount": 19}}},
            {"properties": {"sheetId": SH_YEARLY, "title": "Yearly Summary", "gridProperties": {"rowCount": 20, "columnCount": 8}}},
            {"properties": {"sheetId": SH_ACCOUNTS, "title": "Accounts Setup", "gridProperties": {"rowCount": 84, "columnCount": 7}}},
            {"properties": {"sheetId": SH_SETTINGS, "title": "Settings & Reference", "hidden": True, "gridProperties": {"rowCount": 40, "columnCount": 10}}},
        ],
    }
    resp = ss.create(body=body)
    return resp["spreadsheetId"]


# =================================================================
# 2. VALUES (data + formulas) per sheet
# =================================================================
def accounts_setup_values():
    rows = [["ACCOUNTS SETUP", "", "", "", "", ""],
            ["#", "Account Name", "Type", "Category", "Liquidity", "Include in Net Worth"]]
    for i in range(N_ACCOUNTS):
        if i < len(SEED_ACCOUNTS):
            name, typ, cat, liq, _ = SEED_ACCOUNTS[i]
            rows.append([i + 1, name, typ, cat, liq, True])
        else:
            rows.append([i + 1, "", "", "", "", True])
    return rows


def monthly_tracker_values():
    rows = []
    rows.append(["NET WORTH TRACKER"] + [""] * 17)
    rows.append(["Year", 2026, "", "=\"Balances shown in \"&'Settings & Reference'!$B$3"] + [""] * 14)
    rows.append([""] * 18)
    rows.append(["Account", "Type", "Category"] + MONTHS + ["Latest", "Include", "Liquidity"])
    for i in range(N_ACCOUNTS):
        setup_row = i + 3  # Accounts Setup data row for this account
        r = [
            f"=IF('Accounts Setup'!B{setup_row}=\"\",\"\",'Accounts Setup'!B{setup_row})",
            f"=IF('Accounts Setup'!C{setup_row}=\"\",\"\",'Accounts Setup'!C{setup_row})",
            f"=IF('Accounts Setup'!D{setup_row}=\"\",\"\",'Accounts Setup'!D{setup_row})",
        ]
        if i < len(SEED_ACCOUNTS):
            r += SEED_ACCOUNTS[i][4]
        else:
            r += [""] * 12
        mt_row = i + 5
        r.append(f"=IFERROR(LOOKUP(9E307,$D{mt_row}:$O{mt_row}),0)")
        r.append(f"=IF('Accounts Setup'!F{setup_row}=\"\",FALSE,'Accounts Setup'!F{setup_row})")
        r.append(f"=IF('Accounts Setup'!E{setup_row}=\"\",\"\",'Accounts Setup'!E{setup_row})")
        rows.append(r)
    # totals rows 85-87
    total_assets = ["TOTAL ASSETS", "", ""]
    total_liab = ["TOTAL LIABILITIES", "", ""]
    net_worth = ["NET WORTH", "", ""]
    for col in MONTHS + ["Latest"]:
        cl = "P" if col == "Latest" else col_letter(3 + MONTHS.index(col))
        total_assets.append(f"=SUMIFS({cl}$5:{cl}$84,$B$5:$B$84,\"Asset\",$Q$5:$Q$84,TRUE)")
        total_liab.append(f"=SUMIFS({cl}$5:{cl}$84,$B$5:$B$84,\"Liability\",$Q$5:$Q$84,TRUE)")
    for col in MONTHS + ["Latest"]:
        cl = "P" if col == "Latest" else col_letter(3 + MONTHS.index(col))
        net_worth.append(f"={cl}85-{cl}86")
    total_assets += ["", ""]
    total_liab += ["", ""]
    net_worth += ["", ""]
    rows.append(total_assets)
    rows.append(total_liab)
    rows.append(net_worth)
    return rows


def dashboard_values():
    rows = [[""] * 13 for _ in range(9)]
    rows[0][0] = "NET WORTH DASHBOARD"
    rows[1][0] = "Year"; rows[1][1] = 2026; rows[1][3] = "Currency"; rows[1][4] = "£"
    # row4 (idx3): card labels
    rows[3][0] = "CURRENT NET WORTH"
    rows[3][4] = "NET WORTH GOAL"
    # row5 (idx4): card values
    rows[4][0] = "='Settings & Reference'!$B$3&TEXT(N5,\"#,##0\")"  # N5 = 'Monthly Tracker'!$P$87
    rows[4][4] = 250000
    # row6 (idx5): card labels
    rows[5][0] = "GOAL PROGRESS"
    rows[5][4] = "YTD CHANGE"
    # row7 (idx6): card values
    rows[6][0] = "=N6"
    rows[6][4] = "='Settings & Reference'!$B$3&TEXT(N5-'Monthly Tracker'!$D$87,\"#,##0\")"
    # row8 (idx7): progress bar
    rows[7][0] = "=REPT(\"█\",ROUND(MIN(N6,1)*25,0))&REPT(\"░\",25-ROUND(MIN(N6,1)*25,0))"
    # row9 spacer
    return rows


def dashboard_helper_values():
    # hidden helper column N, rows 5 and 6
    return [
        ["='Monthly Tracker'!$P$87"],
        ["=IFERROR(N5/E5,0)"],
    ]


def yearly_summary_values():
    rows = [["YEARLY SUMMARY", "", "", "", "", ""],
            ["", "", "", "", "", ""],
            ["Year", "Total Assets", "Total Liabilities", "Net Worth", "YoY Change", "YoY %"]]
    data = {
        2024: (200000, 55000),
        2025: (235000, 63000),
    }
    for offset, year in enumerate(range(2024, 2034)):
        row_num = 4 + offset
        if year in data:
            assets, liab = data[year]
            a, l = str(assets), str(liab)
        elif year == 2026:
            a = "='Monthly Tracker'!$P$85"
            l = "='Monthly Tracker'!$P$86"
        else:
            a, l = "", ""
        d_formula = f"=IF(B{row_num}=\"\",\"\",B{row_num}-C{row_num})"
        if offset == 0:
            e_formula = ""
            f_formula = ""
        else:
            prev = row_num - 1
            e_formula = f"=IF(OR(D{row_num}=\"\",D{prev}=\"\"),\"\",D{row_num}-D{prev})"
            f_formula = f"=IF(OR(D{row_num}=\"\",D{prev}=\"\",D{prev}=0),\"\",(D{row_num}-D{prev})/D{prev})"
        rows.append([year if (year in data or year == 2026) else "", a, l, d_formula, e_formula, f_formula])
    return rows


def settings_values():
    rows = [[""] * 8 for _ in range(36)]
    rows[1][1] = "£"          # B2
    rows[2][1] = "=B2"             # B3
    rows[4][0] = "Types"
    rows[4][2] = "Categories"
    rows[4][4] = "Liquidity"
    rows[4][6] = "Years"
    for i, t in enumerate(TYPES):
        rows[5 + i][0] = t
    for i, c in enumerate(ALL_CATEGORIES):
        rows[5 + i][2] = c
    for i, lq in enumerate(LIQUIDITY):
        rows[5 + i][4] = lq
    for i, y in enumerate(YEARS):
        rows[5 + i][6] = y
    # Helper Table A: Asset Allocation by Category (rows 18-26, 0-idx 17-25)
    rows[17][0] = "Asset Allocation by Category"
    rows[17][1] = "Value"
    for i, cat in enumerate(ASSET_CATEGORIES):
        r = 19 + i
        rows[18 + i][0] = cat
        rows[18 + i][1] = (f"=SUMIFS('Monthly Tracker'!$P$5:$P$84,'Monthly Tracker'!$C$5:$C$84,\"{cat}\","
                            f"'Monthly Tracker'!$B$5:$B$84,\"Asset\")")
    # Helper Table B: Liquid vs Illiquid (rows 28-30, 0-idx 27-29)
    rows[27][0] = "Liquid vs Illiquid"
    rows[27][1] = "Value"
    rows[28][0] = "Liquid"
    rows[28][1] = ("=SUMIFS('Monthly Tracker'!$P$5:$P$84,'Monthly Tracker'!$R$5:$R$84,\"Liquid\","
                   "'Monthly Tracker'!$B$5:$B$84,\"Asset\")")
    rows[29][0] = "Illiquid"
    rows[29][1] = ("=SUMIFS('Monthly Tracker'!$P$5:$P$84,'Monthly Tracker'!$R$5:$R$84,\"Illiquid\","
                   "'Monthly Tracker'!$B$5:$B$84,\"Asset\")")
    # Helper Table C: Assets vs Liabilities latest (rows 33-35, 0-idx 32-34)
    rows[32][0] = "Assets vs Liabilities (Latest)"
    rows[32][1] = "Value"
    rows[33][0] = "Assets"
    rows[33][1] = "='Monthly Tracker'!$P$85"
    rows[34][0] = "Liabilities"
    rows[34][1] = "='Monthly Tracker'!$P$86"
    return rows


def write_values(spreadsheet_id):
    data = [
        {"range": "'Accounts Setup'!A1", "values": accounts_setup_values()},
        {"range": "'Monthly Tracker'!A1", "values": monthly_tracker_values()},
        {"range": "Dashboard!A1", "values": dashboard_values()},
        {"range": "Dashboard!N5", "values": dashboard_helper_values()},
        {"range": "'Yearly Summary'!A1", "values": yearly_summary_values()},
        {"range": "'Settings & Reference'!A1", "values": settings_values()},
    ]
    ss.values(spreadsheet_id).batch_update(
        {"valueInputOption": "USER_ENTERED", "data": data})


if __name__ == "__main__":
    sid = create_spreadsheet()
    print("SPREADSHEET_ID", sid)
    write_values(sid)
    print("values written")
    with open("spreadsheet_id.txt", "w") as f:
        f.write(sid)
