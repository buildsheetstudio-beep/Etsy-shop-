"""Phase 1b: write formulas for Heatmap, P&L, Calendar, Yearly, Dashboard, Property Dashboard."""
from build import ss
from airbnb_build import (SH_DASH, SH_PROP_DASH, SH_BOOKINGS, SH_CALENDAR, SH_HEATMAP,
                           SH_EXPENSES, SH_RECURRING, SH_PNL, SH_PROPSETUP, SH_YEARLY, SH_SETTINGS,
                           N_PROPERTIES, CATEGORIES, MONTHS_ABBR, SAMPLE_YEAR, date_serial)

sid = open("airbnb_spreadsheet_id.txt").read().strip()


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


# =================================================================
# Occupancy Heatmap (sheetId 4)
# Row1 banner, row2 spacer, row3 header (A "Property", B:M months, N annual occ, O annual revenue)
# Rows4-53 properties, row54 portfolio average (active only)
# =================================================================
def heatmap_values():
    rows = [["OCCUPANCY HEATMAP"] + [""] * 14,
            [""] * 15,
            ["Property"] + MONTHS_ABBR + ["Annual Occ %", "Annual Revenue"]]
    for i in range(N_PROPERTIES):
        r = i + 4
        setup_row = i + 4
        prop_ref = f"$A{r}"
        row = [f"=IF('Property Setup'!B{setup_row}=\"\",\"\",'Property Setup'!B{setup_row})"]
        for m in range(1, 13):
            month_expr = f"DATE(Dashboard!$B$2,{m},1)"
            nights = (f"SUMPRODUCT(('Bookings Log'!$B$5:$B$504={prop_ref})*('Bookings Log'!$H$5:$H$504=\"Confirmed\")"
                      f"*('Bookings Log'!$D$5:$D$504<EOMONTH({month_expr},0)+1)*('Bookings Log'!$E$5:$E$504>{month_expr})"
                      f"*(IF('Bookings Log'!$E$5:$E$504<EOMONTH({month_expr},0)+1,'Bookings Log'!$E$5:$E$504,EOMONTH({month_expr},0)+1)"
                      f"-IF('Bookings Log'!$D$5:$D$504>{month_expr},'Bookings Log'!$D$5:$D$504,{month_expr})))")
            cell = f"=IFERROR({nights}/DAY(EOMONTH({month_expr},0)),0)"
            row.append(cell)
        col_letters = [chr(ord('B') + m) for m in range(12)]
        annual_nights = "+".join(f"{cl}{r}*DAY(EOMONTH(DATE(Dashboard!$B$2,{m+1},1),0))" for m, cl in enumerate(col_letters))
        row.append(f"=IFERROR(({annual_nights})/365,0)")
        row.append(f"=IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
                    f"'Bookings Log'!$B$5:$B$504,{prop_ref},'Bookings Log'!$D$5:$D$504,\">=\"&DATE(Dashboard!$B$2,1,1),"
                    f"'Bookings Log'!$D$5:$D$504,\"<=\"&DATE(Dashboard!$B$2,12,31)),0)")
        rows.append(row)
    avg_row = ["Portfolio Avg (Active)"]
    for m in range(12):
        cl = chr(ord('B') + m)
        avg_row.append(f"=IFERROR(SUMPRODUCT(('Property Setup'!$F$4:$F$53=TRUE)*{cl}4:{cl}53)"
                        f"/SUM(IF('Property Setup'!$F$4:$F$53=TRUE,1,0)),0)")
    avg_row += ["", ""]
    rows.append(avg_row)
    return rows


# =================================================================
# P&L Summary (sheetId 7)
# Row1 banner(filled,not merged - col A frozen). Row2 selectors A2 Year/B2, C2 Property/D2.
# Row3 month-starts D3:O3, off-screen Q1 selAll, Q2 propCrit.
# Row4 headers. Row5 Income. Rows6-15 categories. Row16 Total Expenses. Row17 Net Profit. Row18 Margin%.
# =================================================================
def pnl_values():
    month_cols = [chr(ord('B') + m) for m in range(12)]  # B..M
    rows = [["P&L SUMMARY"] + [""] * 17,
            ["Year", SAMPLE_YEAR, "Property", "All"] + [""] * 14,
            [""] + [f"=DATE($B$2,{m},1)" for m in range(1, 13)] + ["", "", "=($D$2=\"All\")", "=IF($D$2=\"All\",\"*\",$D$2)"],
            ["P&L"] + MONTHS_ABBR + ["Total"] + [""] * 3,
    ]
    # selAll lands at P3, propCrit at Q3
    # row5 Income
    income_row = ["Income"]
    for cl in month_cols:
        income_row.append(
            f"=IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
            f"'Bookings Log'!$D$5:$D$504,\">=\"&{cl}$3,'Bookings Log'!$D$5:$D$504,\"<=\"&EOMONTH({cl}$3,0),"
            f"'Bookings Log'!$B$5:$B$504,$Q$3),0)")
    income_row.append("=SUM(B5:M5)")
    rows.append(income_row)
    # rows6-15 categories
    for i, cat in enumerate(CATEGORIES):
        r = 6 + i
        row = [cat]
        for cl in month_cols:
            oneoff = (f"SUMIFS('Expenses Log'!$E$5:$E$504,'Expenses Log'!$C$5:$C$504,$A{r},"
                      f"'Expenses Log'!$F$5:$F$504,{cl}$3,'Expenses Log'!$B$5:$B$504,$Q$3)")
            prop_clause = "((('Recurring Expenses'!$B$4:$B$53=$D$2)+('Recurring Expenses'!$B$4:$B$53=\"All\")+$P$3)>0)"
            recur = recurring_term(f"{cl}$3", cat_expr=f"$A{r}", prop_clause=prop_clause)
            row.append(f"=IFERROR({oneoff}+{recur},0)")
        row.append(f"=SUM(B{r}:M{r})")
        rows.append(row)
    # row16 total expenses
    total_row = ["Total Expenses"]
    for cl in month_cols:
        total_row.append(f"=SUM({cl}6:{cl}15)")
    total_row.append("=SUM(B16:M16)")
    rows.append(total_row)
    # row17 net profit
    profit_row = ["Net Profit"]
    for cl in month_cols:
        profit_row.append(f"={cl}5-{cl}16")
    profit_row.append("=N5-N16")
    rows.append(profit_row)
    # row18 margin
    margin_row = ["Margin %"]
    for cl in month_cols:
        margin_row.append(f"=IFERROR({cl}17/{cl}5,0)")
    margin_row.append("=IFERROR(N17/N5,0)")
    rows.append(margin_row)
    return rows


# =================================================================
# Yearly Summary (sheetId 9)
# =================================================================
def yearly_values():
    rows = [["YEARLY SUMMARY", "", "", "", "", "", ""],
            ["", "", "", "", "", "", ""],
            ["Year", "Revenue", "Expenses", "Profit", "Margin %", "Occupancy %", "Bookings"]]
    seed = {2024: (38000, 14500, 52, "61%"), 2025: (46500, 17200, 64, "68%")}
    for offset, year in enumerate(range(2024, 2034)):
        r = 4 + offset
        if year in seed:
            rev, exp, bookings, _ = seed[year]
            occ = [0.58, 0.66][offset]
            row = [year, rev, exp, f"=B{r}-C{r}", f"=IFERROR(D{r}/B{r},0)", occ, bookings]
        elif year == SAMPLE_YEAR:
            rev = (f"=IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
                   f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({year},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({year},12,31)),0)")
            oneoff = (f"SUMIFS('Expenses Log'!$E$5:$E$504,'Expenses Log'!$A$5:$A$504,\">=\"&DATE({year},1,1),"
                      f"'Expenses Log'!$A$5:$A$504,\"<=\"&DATE({year},12,31))")
            recur = recurring_annual(str(year))
            exp = f"=IFERROR({oneoff}+{recur},0)"
            bookings = (f"=IFERROR(COUNTIFS('Bookings Log'!$H$5:$H$504,\"Confirmed\","
                        f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({year},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({year},12,31)),0)")
            occ = (f"=IFERROR(SUMPRODUCT(('Property Setup'!$F$4:$F$53=TRUE)*'Occupancy Heatmap'!$N$4:$N$53)"
                   f"/SUM(IF('Property Setup'!$F$4:$F$53=TRUE,1,0)),0)")
            row = [year, rev, exp, f"=B{r}-C{r}", f"=IFERROR(D{r}/B{r},0)", occ, bookings]
        else:
            row = ["", "", "", f"=IF(B{r}=\"\",\"\",B{r}-C{r})", f"=IFERROR(D{r}/B{r},0)", "", ""]
        rows.append(row)
    return rows


# =================================================================
# Dashboard (sheetId 0) + hidden helper row 40
# =================================================================
def dashboard_values():
    rows = [[""] * 13 for _ in range(8)]
    rows[0][0] = "PORTFOLIO DASHBOARD"
    rows[1][0] = "Year"; rows[1][1] = SAMPLE_YEAR; rows[1][3] = "Currency"; rows[1][4] = "£"
    rows[3][0] = "TOTAL REVENUE"; rows[3][3] = "TOTAL EXPENSES"; rows[3][6] = "NET PROFIT"; rows[3][9] = "AVG MARGIN"
    rows[4][0] = "='Settings & Reference'!$B$3&TEXT(A40,\"#,##0\")"
    rows[4][3] = "='Settings & Reference'!$B$3&TEXT(B40,\"#,##0\")"
    rows[4][6] = "='Settings & Reference'!$B$3&TEXT(C40,\"#,##0\")"
    rows[4][9] = "=D40"
    rows[5][0] = "TOTAL BOOKINGS"; rows[5][3] = "NIGHTS BOOKED"; rows[5][6] = "AVG NIGHTLY (RevPAN)"; rows[5][9] = "CANCELLATION RATE"
    rows[6][0] = "=E40"
    rows[6][3] = "=F40"
    rows[6][6] = "='Settings & Reference'!$B$3&TEXT(G40,\"#,##0\")"
    rows[6][9] = "=H40"
    return rows


def dashboard_helpers():
    y = "$B$2"
    revenue = (f"IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
               f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31)),0)")
    oneoff = (f"SUMIFS('Expenses Log'!$E$5:$E$504,'Expenses Log'!$A$5:$A$504,\">=\"&DATE({y},1,1),"
              f"'Expenses Log'!$A$5:$A$504,\"<=\"&DATE({y},12,31))")
    recur = recurring_annual(y.replace("$", ""))
    expenses = f"IFERROR({oneoff}+{recur},0)"
    bookings = (f"IFERROR(COUNTIFS('Bookings Log'!$H$5:$H$504,\"Confirmed\","
                f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31)),0)")
    nights = (f"IFERROR(SUMIFS('Bookings Log'!$F$5:$F$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
              f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31)),0)")
    cancelled = (f"COUNTIFS('Bookings Log'!$H$5:$H$504,\"Cancelled\","
                 f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31))")
    allbk = (f"COUNTIFS('Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),"
             f"'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31))")
    return [[
        f"=A40_PLACEHOLDER",  # replaced below
    ]]


def dashboard_helper_row():
    y = "$B$2"
    revenue = (f"IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
               f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31)),0)")
    oneoff = (f"SUMIFS('Expenses Log'!$E$5:$E$504,'Expenses Log'!$A$5:$A$504,\">=\"&DATE({y},1,1),"
              f"'Expenses Log'!$A$5:$A$504,\"<=\"&DATE({y},12,31))")
    recur = recurring_annual("$B$2")
    expenses = f"IFERROR({oneoff}+{recur},0)"
    bookings = (f"IFERROR(COUNTIFS('Bookings Log'!$H$5:$H$504,\"Confirmed\","
                f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31)),0)")
    nights = (f"IFERROR(SUMIFS('Bookings Log'!$F$5:$F$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
              f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31)),0)")
    cancelled = (f"COUNTIFS('Bookings Log'!$H$5:$H$504,\"Cancelled\","
                 f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31))")
    allbk = (f"COUNTIFS('Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),"
             f"'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31))")
    a40 = f"={revenue}"
    b40 = f"={expenses}"
    c40 = "=A40-B40"
    d40 = "=IFERROR(C40/A40,0)"
    e40 = f"={bookings}"
    f40 = f"={nights}"
    g40 = "=IFERROR(A40/F40,0)"
    h40 = f"=IFERROR({cancelled}/{allbk},0)"
    return [[a40, b40, c40, d40, e40, f40, g40, h40]]


# =================================================================
# Property Dashboard (sheetId 1) + hidden helper rows 40 (monthly income), 41 (monthly occupancy)
# =================================================================
def property_dashboard_values():
    rows = [[""] * 13 for _ in range(8)]
    rows[0][0] = "PROPERTY DASHBOARD"
    rows[1][0] = "Property"; rows[1][1] = "Maple Cottage"; rows[1][3] = "Year"; rows[1][4] = SAMPLE_YEAR
    rows[3][0] = "REVENUE"; rows[3][3] = "EXPENSES"; rows[3][6] = "NET PROFIT"; rows[3][9] = "OCCUPANCY %"
    rows[4][0] = "='Settings & Reference'!$B$3&TEXT(A40,\"#,##0\")"
    rows[4][3] = "='Settings & Reference'!$B$3&TEXT(B40,\"#,##0\")"
    rows[4][6] = "='Settings & Reference'!$B$3&TEXT(C40,\"#,##0\")"
    rows[4][9] = "=D40"
    return rows


def property_dashboard_helper_rows():
    y, p = "$E$2", "$B$2"
    prop_crit = f"IF({p}=\"All\",\"*\",{p})"
    revenue = (f"IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
               f"'Bookings Log'!$D$5:$D$504,\">=\"&DATE({y},1,1),'Bookings Log'!$D$5:$D$504,\"<=\"&DATE({y},12,31),"
               f"'Bookings Log'!$B$5:$B$504,{prop_crit}),0)")
    oneoff = (f"SUMIFS('Expenses Log'!$E$5:$E$504,'Expenses Log'!$B$5:$B$504,{prop_crit},"
              f"'Expenses Log'!$A$5:$A$504,\">=\"&DATE({y},1,1),'Expenses Log'!$A$5:$A$504,\"<=\"&DATE({y},12,31))")
    prop_clause = f"((('Recurring Expenses'!$B$4:$B$53={p})+('Recurring Expenses'!$B$4:$B$53=\"All\")+({p}=\"All\"))>0)"
    recur = recurring_annual(y.replace("$", ""), prop_clause=prop_clause)
    expenses = f"IFERROR({oneoff}+{recur},0)"
    prop_match = f"((({p}=\"All\"))+('Property Setup'!$B$4:$B$53={p})>0)"
    occ = (f"IFERROR(SUMPRODUCT({prop_match}*'Occupancy Heatmap'!$N$4:$N$53)"
           f"/MAX(1,SUMPRODUCT({prop_match}*1)),0)")
    a40 = f"={revenue}"
    b40 = f"={expenses}"
    c40 = "=A40-B40"
    d40 = f"={occ}"
    # row41: monthly income for chart
    row41 = []
    for m in range(1, 13):
        ms = f"DATE({y},{m},1)"
        row41.append(f"=IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
                      f"'Bookings Log'!$D$5:$D$504,\">=\"&{ms},'Bookings Log'!$D$5:$D$504,\"<=\"&EOMONTH({ms},0),"
                      f"'Bookings Log'!$B$5:$B$504,{prop_crit}),0)")
    # row42: monthly occupancy % for chart
    row42 = []
    for m in range(1, 13):
        ms = f"DATE({y},{m},1)"
        nights = (f"SUMPRODUCT((({p}=\"All\")+('Bookings Log'!$B$5:$B$504={p})>0)*('Bookings Log'!$H$5:$H$504=\"Confirmed\")"
                  f"*('Bookings Log'!$D$5:$D$504<EOMONTH({ms},0)+1)*('Bookings Log'!$E$5:$E$504>{ms})"
                  f"*(IF('Bookings Log'!$E$5:$E$504<EOMONTH({ms},0)+1,'Bookings Log'!$E$5:$E$504,EOMONTH({ms},0)+1)"
                  f"-IF('Bookings Log'!$D$5:$D$504>{ms},'Bookings Log'!$D$5:$D$504,{ms})))")
        row42.append(f"=IFERROR({nights}/DAY(EOMONTH({ms},0)),0)")
    return [a40, b40, c40, d40], row41, row42


# =================================================================
# Booking Calendar (sheetId 3)
# =================================================================
def calendar_values():
    rows = [[""] * 9 for _ in range(13)]
    rows[0][0] = "BOOKING CALENDAR"
    rows[1][0] = "Property"; rows[1][1] = "All"; rows[1][2] = "Month"; rows[1][3] = "January"
    rows[1][4] = "Year"; rows[1][5] = SAMPLE_YEAR
    # off-screen helpers at K2,L2,M2 (idx 10,11,12)
    rows[1] += [""] * 5
    rows[1][10] = "=MATCH(D2,'Settings & Reference'!$K$6:$K$17,0)"
    rows[1][11] = "=IF(B2=\"All\",\"*\",B2)"
    rows[1][12] = "=DATE(F2,K2,1)-(WEEKDAY(DATE(F2,K2,1),2)-1)"
    rows[1][13] = "=(B2=\"All\")"
    weekday_row = [""] * 2 + ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    rows[4] = weekday_row + [""] * (len(rows[4]) - len(weekday_row)) if len(rows[4]) >= len(weekday_row) else weekday_row
    return rows


def calendar_grid_values():
    # 6x7 grid C8:I13
    grid = []
    grid.append([f"=M2"] + [f"={chr(ord('C')+c-1)}8+1" for c in range(1, 7)])
    for r in range(1, 6):
        prev_row_last_col = chr(ord('C') + 6)
        first = f"={prev_row_last_col}{8+r-1}+1"
        row = [first] + [f"={chr(ord('C')+c-1)}{8+r}+1" for c in range(1, 7)]
        grid.append(row)
    return grid


def calendar_stats_values():
    rows = [
        ["This Month Stats (selected property)"],
        ["Nights Booked", "=IFERROR(SUMPRODUCT((($N$2)+('Bookings Log'!$B$5:$B$504=$B$2)>0)*('Bookings Log'!$H$5:$H$504=\"Confirmed\")"
         "*('Bookings Log'!$D$5:$D$504<EOMONTH(DATE($F$2,$K$2,1),0)+1)*('Bookings Log'!$E$5:$E$504>DATE($F$2,$K$2,1))"
         "*(IF('Bookings Log'!$E$5:$E$504<EOMONTH(DATE($F$2,$K$2,1),0)+1,'Bookings Log'!$E$5:$E$504,EOMONTH(DATE($F$2,$K$2,1),0)+1)"
         "-IF('Bookings Log'!$D$5:$D$504>DATE($F$2,$K$2,1),'Bookings Log'!$D$5:$D$504,DATE($F$2,$K$2,1)))),0)"],
        ["Occupancy %", "=IFERROR(B16/DAY(EOMONTH(DATE($F$2,$K$2,1),0)),0)"],
        ["Revenue", "=IFERROR(SUMIFS('Bookings Log'!$G$5:$G$504,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
         "'Bookings Log'!$B$5:$B$504,$L$2,'Bookings Log'!$D$5:$D$504,\">=\"&DATE($F$2,$K$2,1),"
         "'Bookings Log'!$D$5:$D$504,\"<=\"&EOMONTH(DATE($F$2,$K$2,1),0)),0)"],
        ["Bookings Count", "=IFERROR(COUNTIFS('Bookings Log'!$B$5:$B$504,$L$2,'Bookings Log'!$H$5:$H$504,\"Confirmed\","
         "'Bookings Log'!$D$5:$D$504,\">=\"&DATE($F$2,$K$2,1),'Bookings Log'!$D$5:$D$504,\"<=\"&EOMONTH(DATE($F$2,$K$2,1),0)),0)"],
    ]
    return rows


def write_phase2():
    data = [
        {"range": "'Occupancy Heatmap'!A1", "values": heatmap_values()},
        {"range": "'P&L Summary'!A1", "values": pnl_values()},
        {"range": "'Yearly Summary'!A1", "values": yearly_values()},
        {"range": "Dashboard!A1", "values": dashboard_values()},
        {"range": "Dashboard!A40", "values": dashboard_helper_row()},
        {"range": "'Property Dashboard'!A1", "values": property_dashboard_values()},
    ]
    pd_kpi, pd_row41, pd_row42 = property_dashboard_helper_rows()
    data.append({"range": "'Property Dashboard'!A40", "values": [pd_kpi]})
    data.append({"range": "'Property Dashboard'!A41", "values": [pd_row41]})
    data.append({"range": "'Property Dashboard'!A42", "values": [pd_row42]})
    data.append({"range": "'Booking Calendar'!A1", "values": calendar_values()})
    data.append({"range": "'Booking Calendar'!C8", "values": calendar_grid_values()})
    data.append({"range": "'Booking Calendar'!A15", "values": calendar_stats_values()})
    ss.values(sid).batch_update({"valueInputOption": "USER_ENTERED", "data": data})


if __name__ == "__main__":
    write_phase2()
    print("phase2 values written")
