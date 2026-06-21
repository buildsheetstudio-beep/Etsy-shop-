"""
Builds the Airbnb / Short-Term Rental Manager workbook (phase 1: create + values/formulas).
Reuses auth/transport helpers from build.py (same Google account as the Net Worth Tracker).
Run: python3 airbnb_build.py
"""
import random
from build import ss, hex_to_rgb, PALETTE  # noqa: F401 (PALETTE/hex_to_rgb reused by formatting phase)

# ---------------------------------------------------------------- sheet ids
SH_DASH, SH_PROP_DASH, SH_BOOKINGS, SH_CALENDAR, SH_HEATMAP, SH_EXPENSES, \
    SH_RECURRING, SH_PNL, SH_PROPSETUP, SH_YEARLY, SH_SETTINGS = range(11)

MONTHS_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTHS_FULL = ["January", "February", "March", "April", "May", "June", "July", "August",
                "September", "October", "November", "December"]
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
STATUSES = ["Confirmed", "Pending", "Cancelled"]
SOURCES = ["Airbnb", "VRBO", "Booking.com", "Direct", "Other"]
CATEGORIES = ["Cleaning", "Maintenance", "Utilities", "Mortgage/Rent", "Insurance",
              "Supplies", "Platform Fees", "Management", "Marketing", "Other"]
YEARS = list(range(2023, 2033))
N_PROPERTIES = 50
N_BOOKINGS = 500
N_EXPENSES = 500
N_RECURRING = 50
SAMPLE_YEAR = 2026

SEED_PROPERTIES = [
    ("Maple Cottage", "12 Mill Lane, York", 2, 95),
    ("Harbour Loft", "4 Quay St, Bristol", 1, 110),
    ("Elm Studio", "7 Elm Rd, Manchester", 1, 75),
    ("Oak Barn", "Oak Farm, Cotswolds", 3, 160),
    ("City Flat 2B", "2B Park Ave, Leeds", 2, 90),
    ("Seaview Annexe", "1 Cliff Rd, Brighton", 1, 120),
    ("Garden Mews", "9 Mews Court, Bath", 2, 105),
]


def col_letter(idx):
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def date_serial(y, m, d):
    # store as ISO string; USER_ENTERED parses real dates from "YYYY-MM-DD"
    return f"{y:04d}-{m:02d}-{d:02d}"


# =================================================================
# sample bookings/expenses/recurring generator (deterministic)
# =================================================================
def gen_bookings():
    random.seed(42)
    rows = []
    n_props = len(SEED_PROPERTIES)
    cancelled_left, pending_left = 4, 2
    target = 50
    day_cursor = {name: 1 for name, *_ in SEED_PROPERTIES}
    month_cursor = {name: 1 for name, *_ in SEED_PROPERTIES}
    for i in range(target):
        name, addr, beds, rate = SEED_PROPERTIES[i % n_props]
        month = month_cursor[name]
        day = day_cursor[name]
        if day > 24:
            month += 1
            day = random.randint(1, 5)
        if month > 12:
            month = 12
            day = min(day, 20)
        nights = random.randint(2, 10)
        checkin = (SAMPLE_YEAR, month, day)
        # checkout date (handle month rollover simply by capping to 28)
        co_day = day + nights
        co_month = month
        if co_day > 28:
            co_day -= 28
            co_month = min(month + 1, 12)
        checkout = (SAMPLE_YEAR, co_month, co_day)
        # 30% chance: chain next booking back-to-back from this checkout (turnover demo)
        back_to_back = random.random() < 0.3
        if back_to_back:
            day_cursor[name] = co_day
            month_cursor[name] = co_month
        else:
            day_cursor[name] = co_day + random.randint(1, 4)
            month_cursor[name] = co_month
        guest = f"Guest {i + 1:03d}"
        revenue = round(rate * nights * random.uniform(0.95, 1.3))
        revenue = max(180, min(2400, revenue))
        if cancelled_left > 0 and random.random() < 0.12:
            status = "Cancelled"
            cancelled_left -= 1
        elif pending_left > 0 and random.random() < 0.08:
            status = "Pending"
            pending_left -= 1
        else:
            status = "Confirmed"
        source = random.choice(SOURCES)
        cleaning_fee = random.choice([0, 35, 45, 60, 75])
        rows.append((name, guest, checkin, checkout, revenue, status, source, cleaning_fee))
    return rows


def gen_expenses():
    random.seed(7)
    cats = ["Cleaning", "Maintenance", "Supplies", "Marketing", "Utilities", "Other"]
    rows = []
    for i in range(15):
        name = SEED_PROPERTIES[i % len(SEED_PROPERTIES)][0]
        month = random.randint(1, 12)
        day = random.randint(1, 28)
        cat = random.choice(cats)
        amount = random.choice([25, 40, 60, 85, 120, 150, 220])
        desc = f"{cat} - {name}"
        rows.append(((SAMPLE_YEAR, month, day), name, cat, desc, amount))
    return rows


SEED_RECURRING = [
    ("All", "Insurance", "Portfolio insurance", 45, (SAMPLE_YEAR, 1, 1), None),
    ("All", "Platform Fees", "Channel Manager subscription", 29, (SAMPLE_YEAR, 1, 1), None),
    ("Oak Barn", "Mortgage/Rent", "Mortgage - Oak Barn", 620, (SAMPLE_YEAR, 1, 1), None),
    ("Harbour Loft", "Cleaning", "Cleaning retainer - Harbour Loft", 80, (SAMPLE_YEAR, 1, 1), None),
    ("Elm Studio", "Management", "Property management fee", 60, (SAMPLE_YEAR, 1, 1), (SAMPLE_YEAR, 6, 30)),
]


# =================================================================
# PHASE 1: CREATE SPREADSHEET
# =================================================================
def create_spreadsheet():
    sheets = [
        {"properties": {"sheetId": SH_DASH, "title": "Dashboard", "gridProperties": {"rowCount": 50, "columnCount": 14}}},
        {"properties": {"sheetId": SH_PROP_DASH, "title": "Property Dashboard", "gridProperties": {"rowCount": 50, "columnCount": 14}}},
        {"properties": {"sheetId": SH_BOOKINGS, "title": "Bookings Log", "gridProperties": {"rowCount": 510, "columnCount": 11}}},
        {"properties": {"sheetId": SH_CALENDAR, "title": "Booking Calendar", "gridProperties": {"rowCount": 30, "columnCount": 14}}},
        {"properties": {"sheetId": SH_HEATMAP, "title": "Occupancy Heatmap", "gridProperties": {"rowCount": 60, "columnCount": 15}}},
        {"properties": {"sheetId": SH_EXPENSES, "title": "Expenses Log", "gridProperties": {"rowCount": 510, "columnCount": 6}}},
        {"properties": {"sheetId": SH_RECURRING, "title": "Recurring Expenses", "gridProperties": {"rowCount": 60, "columnCount": 8}}},
        {"properties": {"sheetId": SH_PNL, "title": "P&L Summary", "gridProperties": {"rowCount": 25, "columnCount": 18}}},
        {"properties": {"sheetId": SH_PROPSETUP, "title": "Property Setup", "gridProperties": {"rowCount": 60, "columnCount": 7}}},
        {"properties": {"sheetId": SH_YEARLY, "title": "Yearly Summary", "gridProperties": {"rowCount": 20, "columnCount": 8}}},
        {"properties": {"sheetId": SH_SETTINGS, "title": "Settings & Reference", "hidden": True, "gridProperties": {"rowCount": 140, "columnCount": 14}}},
    ]
    resp = ss.create(body={"properties": {"title": "Airbnb / Short-Term Rental Manager"}, "sheets": sheets})
    return resp["spreadsheetId"]


# =================================================================
# VALUE BUILDERS
# =================================================================
def property_setup_values():
    rows = [["PROPERTY SETUP", "", "", "", "", ""],
            ["", "", "", "", "", ""],
            ["#", "Property Name", "Address", "Bedrooms", "Base Nightly Rate", "Active"]]
    for i in range(N_PROPERTIES):
        if i < len(SEED_PROPERTIES):
            name, addr, beds, rate = SEED_PROPERTIES[i]
            rows.append([i + 1, name, addr, beds, rate, True])
        else:
            rows.append([i + 1, "", "", "", "", True])
    return rows


def bookings_values():
    rows = [["BOOKINGS LOG"] + [""] * 10,
            [""] * 11,
            [""] * 11,
            ["Booking ID", "Property", "Guest", "Check-in", "Check-out", "Nights", "Revenue",
             "Status", "Source", "Cleaning Fee", "Notes"]]
    seed = gen_bookings()
    for i in range(N_BOOKINGS):
        r5 = i + 5
        a = f"=IF(B{r5}=\"\",\"\",ROW()-4)"
        f = f"=IF(OR(D{r5}=\"\",E{r5}=\"\"),\"\",E{r5}-D{r5})"
        if i < len(seed):
            name, guest, ci, co, revenue, status, source, cleaning = seed[i]
            row = [a, name, guest, date_serial(*ci), date_serial(*co), f, revenue, status, source, cleaning, ""]
        else:
            row = [a, "", "", "", "", f, "", "", "", "", ""]
        rows.append(row)
    return rows


def expenses_values():
    rows = [["EXPENSES LOG"] + [""] * 5,
            [""] * 6,
            [""] * 6,
            ["Date", "Property", "Category", "Description", "Amount", "Month"]]
    seed = gen_expenses()
    for i in range(N_EXPENSES):
        r5 = i + 5
        f = f"=IF(A{r5}=\"\",\"\",DATE(YEAR(A{r5}),MONTH(A{r5}),1))"
        if i < len(seed):
            d, name, cat, desc, amount = seed[i]
            row = [date_serial(*d), name, cat, desc, amount, f]
        else:
            row = ["", "", "", "", "", f]
        rows.append(row)
    return rows


def recurring_values():
    rows = [["RECURRING EXPENSES", "", "", "", "", "", "", ""],
            [""] * 8,
            ["#", "Property", "Category", "Description", "Monthly Amount", "Start Month", "End Month", "Active"]]
    for i in range(N_RECURRING):
        if i < len(SEED_RECURRING):
            prop, cat, desc, amt, start, end = SEED_RECURRING[i]
            row = [i + 1, prop, cat, desc, amt, date_serial(*start), date_serial(*end) if end else "", True]
        else:
            row = [i + 1, "", "", "", "", "", "", True]
        rows.append(row)
    return rows


def settings_values():
    rows = [[""] * 12 for _ in range(140)]
    rows[1][1] = "£"               # B2
    rows[2][1] = "=B2"                  # B3
    rows[4][0] = "Properties"
    rows[4][2] = "Statuses"
    rows[4][4] = "Sources"
    rows[4][6] = "Categories"
    rows[4][8] = "Years"
    rows[4][10] = "Months"
    for i in range(N_PROPERTIES):
        rows[5 + i][0] = f"=IF('Property Setup'!B{4 + i}=\"\",\"\",'Property Setup'!B{4 + i})"
    for i, s in enumerate(STATUSES):
        rows[5 + i][2] = s
    for i, s in enumerate(SOURCES):
        rows[5 + i][4] = s
    for i, c in enumerate(CATEGORIES):
        rows[5 + i][6] = c
    for i, y in enumerate(YEARS):
        rows[5 + i][8] = y
    for i, m in enumerate(MONTHS_FULL):
        rows[5 + i][10] = m

    def recurring_term(month_expr, cat_expr=None, prop_clause=None):
        clauses = [
            "('Recurring Expenses'!$H$4:$H$53=TRUE)",
            f"('Recurring Expenses'!$F$4:$F$53<={month_expr})",
            f"(('Recurring Expenses'!$G$4:$G$53>={month_expr})+('Recurring Expenses'!$G$4:$G$53=\"\"))",
        ]
        if cat_expr:
            clauses.append(f"('Recurring Expenses'!$C$4:$C$53={cat_expr})")
        if prop_clause:
            clauses.append(prop_clause)
        return "SUMPRODUCT(" + "*".join(clauses) + "*'Recurring Expenses'!$E$4:$E$53)"

    def recurring_annual(year_expr, cat_expr=None, prop_clause=None):
        terms = [recurring_term(f"DATE({year_expr},{m},1)", cat_expr, prop_clause) for m in range(1, 13)]
        return "(" + "+".join(terms) + ")"

    # Helper Table: Booking Source Split (rows 58-63, 0-idx 57-62)
    rows[57][0] = "Booking Source Split"
    rows[57][1] = "Count"
    for i, s in enumerate(SOURCES):
        rows[58 + i][0] = s
        rows[58 + i][1] = (f"=IFERROR(COUNTIFS('Bookings Log'!$I$5:$I$504,\"{s}\","
                            f"'Bookings Log'!$H$5:$H$504,\"Confirmed\"),0)")

    # Helper Table: Expense Breakdown by Category (year = Dashboard!$B$2) rows 66-76 (0-idx 65-75)
    rows[65][0] = "Expense Breakdown by Category (Year)"
    rows[65][1] = "Value"
    for i, c in enumerate(CATEGORIES):
        r = 66 + i
        oneoff = (f"SUMIFS('Expenses Log'!$E$5:$E$504,'Expenses Log'!$C$5:$C$504,\"{c}\","
                  f"'Expenses Log'!$A$5:$A$504,\">=\"&DATE(Dashboard!$B$2,1,1),"
                  f"'Expenses Log'!$A$5:$A$504,\"<=\"&DATE(Dashboard!$B$2,12,31))")
        recur = recurring_annual("Dashboard!$B$2", cat_expr=f'"{c}"')
        rows[r][0] = c
        rows[r][1] = f"=IFERROR({oneoff}+{recur},0)"

    # Helper Table: Top Properties by Profit (year = Dashboard!$B$2) rows 80-129 (0-idx 79-128)
    rows[79][0] = "Top Properties by Profit (Year)"
    rows[79][1] = "Profit"
    for i in range(N_PROPERTIES):
        r = 80 + i
        prop_ref = f"$A${6 + i}"
        revenue = (f"SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
                   f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE(Dashboard!$B$2,1,1),"
                   f"'Bookings Log'!$D$5:$D$504,\"<=\"&DATE(Dashboard!$B$2,12,31),"
                   f"'Bookings Log'!$B$5:$B$504,{prop_ref})")
        oneoff = (f"SUMIFS('Expenses Log'!$E$5:$E$504,'Expenses Log'!$B$5:$B$504,{prop_ref},"
                  f"'Expenses Log'!$A$5:$A$504,\">=\"&DATE(Dashboard!$B$2,1,1),"
                  f"'Expenses Log'!$A$5:$A$504,\"<=\"&DATE(Dashboard!$B$2,12,31))")
        prop_clause = f"((('Recurring Expenses'!$B$4:$B$53={prop_ref})+('Recurring Expenses'!$B$4:$B$53=\"All\"))>0)"
        recur = recurring_annual("Dashboard!$B$2", prop_clause=prop_clause)
        rows[r][0] = f"=IF({prop_ref}=\"\",\"\",{prop_ref})"
        rows[r][1] = f"=IF({prop_ref}=\"\",\"\",IFERROR({revenue}-{oneoff}-{recur},0))"
    return rows, recurring_annual, recurring_term


def write_phase1(spreadsheet_id):
    settings_rows, recurring_annual, recurring_term = settings_values()
    data = [
        {"range": "'Property Setup'!A1", "values": property_setup_values()},
        {"range": "'Bookings Log'!A1", "values": bookings_values()},
        {"range": "'Expenses Log'!A1", "values": expenses_values()},
        {"range": "'Recurring Expenses'!A1", "values": recurring_values()},
        {"range": "'Settings & Reference'!A1", "values": settings_rows},
    ]
    ss.values(spreadsheet_id).batch_update({"valueInputOption": "USER_ENTERED", "data": data})
    return recurring_annual, recurring_term


if __name__ == "__main__":
    sid = create_spreadsheet()
    print("SPREADSHEET_ID", sid)
    write_phase1(sid)
    print("phase1 values written")
    with open("airbnb_spreadsheet_id.txt", "w") as f:
        f.write(sid)
