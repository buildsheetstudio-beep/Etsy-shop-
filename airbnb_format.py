"""Phase 2-6 for the Airbnb / Short-Term Rental Manager workbook:
formatting, validation, conditional formatting, charts, freeze/hide panes.
Run: python3 airbnb_format.py [spreadsheet_id] [phase]
phase in {prep, format, validate, cf, charts, freeze, all} (default: all)
"""
import sys
from build import ss, hex_to_rgb, PALETTE
from airbnb_build import (
    SH_DASH, SH_PROP_DASH, SH_BOOKINGS, SH_CALENDAR, SH_HEATMAP, SH_EXPENSES,
    SH_RECURRING, SH_PNL, SH_PROPSETUP, SH_YEARLY, SH_SETTINGS,
    STATUSES, SOURCES, CATEGORIES, YEARS, MONTHS_FULL, MONTHS_ABBR, N_PROPERTIES,
)

sid = sys.argv[1] if len(sys.argv) > 1 else open("airbnb_spreadsheet_id.txt").read().strip()

PROP_RANGE = "'Settings & Reference'!$A$6:$A$55"        # exact properties only
PROP_ALL_RANGE = "'Settings & Reference'!$A$6:$A$56"     # properties + "All" (row56)


# =================================================================
# request helpers
# =================================================================
def grid(sheet, r0, r1, c0, c1):
    return {"sheetId": sheet, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1}


def banner_request(sheet, end_col, text_size=16):
    return {
        "repeatCell": {
            "range": grid(sheet, 0, 1, 0, end_col),
            "cell": {"userEnteredFormat": {
                "backgroundColor": hex_to_rgb(PALETTE["deep"]),
                "textFormat": {"foregroundColor": hex_to_rgb("#FFFFFF"), "bold": True, "fontSize": text_size},
                "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
        }
    }


def header_row_request(sheet, row, end_col):
    return {
        "repeatCell": {
            "range": grid(sheet, row, row + 1, 0, end_col),
            "cell": {"userEnteredFormat": {
                "backgroundColor": hex_to_rgb(PALETTE["deep"]),
                "textFormat": {"foregroundColor": hex_to_rgb("#FFFFFF"), "bold": True, "fontSize": 11},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }


def alt_rows_requests(sheet, r0, r1, c0, c1):
    reqs = []
    for i, r in enumerate(range(r0, r1)):
        if i % 2 == 1:
            reqs.append({
                "repeatCell": {
                    "range": grid(sheet, r, r + 1, c0, c1),
                    "cell": {"userEnteredFormat": {"backgroundColor": hex_to_rgb(PALETTE["alt"])}},
                    "fields": "userEnteredFormat.backgroundColor",
                }
            })
    return reqs


def number_format_request(sheet, r0, r1, c0, c1, pattern, ftype="NUMBER"):
    return {
        "repeatCell": {
            "range": grid(sheet, r0, r1, c0, c1),
            "cell": {"userEnteredFormat": {"numberFormat": {"type": ftype, "pattern": pattern}}},
            "fields": "userEnteredFormat.numberFormat",
        }
    }


def merge_request(sheet, r0, r1, c0, c1):
    return {"mergeCells": {"range": grid(sheet, r0, r1, c0, c1), "mergeType": "MERGE_ALL"}}


def fill_request(sheet, r0, r1, c0, c1, color):
    return {
        "repeatCell": {
            "range": grid(sheet, r0, r1, c0, c1),
            "cell": {"userEnteredFormat": {"backgroundColor": hex_to_rgb(color)}},
            "fields": "userEnteredFormat.backgroundColor",
        }
    }


def bold_text_request(sheet, r0, r1, c0, c1, size=10, color=None):
    fmt = {"bold": True, "fontSize": size}
    if color:
        fmt["foregroundColor"] = hex_to_rgb(color)
    return {
        "repeatCell": {
            "range": grid(sheet, r0, r1, c0, c1),
            "cell": {"userEnteredFormat": {"textFormat": fmt}},
            "fields": "userEnteredFormat.textFormat",
        }
    }


def column_width_request(sheet, c0, c1, px):
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet, "dimension": "COLUMNS", "startIndex": c0, "endIndex": c1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def hide_columns_request(sheet, c0, c1):
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet, "dimension": "COLUMNS", "startIndex": c0, "endIndex": c1},
            "properties": {"hiddenByUser": True},
            "fields": "hiddenByUser",
        }
    }


def hide_rows_request(sheet, r0, r1):
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet, "dimension": "ROWS", "startIndex": r0, "endIndex": r1},
            "properties": {"hiddenByUser": True},
            "fields": "hiddenByUser",
        }
    }


def freeze_request(sheet, rows=None, cols=None):
    props = {}
    if rows is not None:
        props["frozenRowCount"] = rows
    if cols is not None:
        props["frozenColumnCount"] = cols
    return {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet, "gridProperties": props},
            "fields": ",".join(f"gridProperties.{k}" for k in props),
        }
    }


def checkbox_request(sheet, r0, r1, c0, c1):
    return {
        "setDataValidation": {
            "range": grid(sheet, r0, r1, c0, c1),
            "rule": {"condition": {"type": "BOOLEAN"}, "strict": True},
        }
    }


def list_validation_request(sheet, r0, r1, c0, c1, values):
    return {
        "setDataValidation": {
            "range": grid(sheet, r0, r1, c0, c1),
            "rule": {
                "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": str(v)} for v in values]},
                "strict": True,
                "showCustomUi": True,
            },
        }
    }


def range_validation_request(sheet, r0, r1, c0, c1, source_range_a1):
    return {
        "setDataValidation": {
            "range": grid(sheet, r0, r1, c0, c1),
            "rule": {
                "condition": {"type": "ONE_OF_RANGE", "values": [{"userEnteredValue": f"={source_range_a1}"}]},
                "strict": True,
                "showCustomUi": True,
            },
        }
    }


# =================================================================
# PHASE 1c (prep): small values needed before formatting/charts
# =================================================================
CAL_CONFIRMED_ROW = 20  # hidden helper grid (6x7), mirrors day grid C8:I13
CAL_PENDING_ROW = 27    # hidden helper grid (6x7)


def _calendar_status_grid(status):
    # CUSTOM_FORMULA conditional-format rules can't reference other sheets,
    # so this precomputes the cross-sheet lookup per day-cell on the calendar sheet itself.
    day_row0 = 8
    rows = []
    for r in range(6):
        row = []
        for c in range(7):
            col = chr(ord("C") + c)
            day_cell = f"{col}{day_row0 + r}"
            row.append(
                f"=SUMPRODUCT((($N$2)+('Bookings Log'!$B$5:$B$504=$B$2)>0)*"
                f"('Bookings Log'!$H$5:$H$504=\"{status}\")*"
                f"('Bookings Log'!$D$5:$D$504<={day_cell})*('Bookings Log'!$E$5:$E$504>{day_cell}))>0"
            )
        rows.append(row)
    return rows


def prep_values():
    data = [
        # "All" option appended to the property mirror list, used by dropdowns that need a wildcard choice
        {"range": "'Settings & Reference'!A56", "values": [["All"]]},
        # hidden month-label row for Property Dashboard monthly charts (row39)
        {"range": "'Property Dashboard'!A39", "values": [MONTHS_ABBR]},
        # hidden booking-status helper grids for Booking Calendar conditional formatting
        {"range": f"'Booking Calendar'!C{CAL_CONFIRMED_ROW}", "values": _calendar_status_grid("Confirmed")},
        {"range": f"'Booking Calendar'!C{CAL_PENDING_ROW}", "values": _calendar_status_grid("Pending")},
    ]
    return data


def write_prep():
    # grow Booking Calendar so the hidden status-helper grids (rows 20-25, 27-32) fit
    ss.batch_update(sid, {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": SH_CALENDAR, "gridProperties": {"rowCount": 34, "columnCount": 14}},
            "fields": "gridProperties(rowCount,columnCount)",
        }
    }]})
    ss.values(sid).batch_update({"valueInputOption": "USER_ENTERED", "data": prep_values()})


# =================================================================
# PHASE 2: FORMATTING
# =================================================================
def formatting_requests():
    reqs = []

    # ---- Property Setup ----
    reqs += [
        banner_request(SH_PROPSETUP, 6),
        header_row_request(SH_PROPSETUP, 2, 6),
        *alt_rows_requests(SH_PROPSETUP, 3, 53, 0, 6),
        number_format_request(SH_PROPSETUP, 3, 53, 4, 5, "#,##0"),
        column_width_request(SH_PROPSETUP, 0, 1, 40),
        column_width_request(SH_PROPSETUP, 1, 2, 170),
        column_width_request(SH_PROPSETUP, 2, 3, 220),
        column_width_request(SH_PROPSETUP, 3, 4, 90),
        column_width_request(SH_PROPSETUP, 4, 5, 130),
        column_width_request(SH_PROPSETUP, 5, 6, 80),
    ]

    # ---- Bookings Log ----
    reqs += [
        banner_request(SH_BOOKINGS, 11),
        header_row_request(SH_BOOKINGS, 3, 11),
        *alt_rows_requests(SH_BOOKINGS, 4, 504, 0, 11),
        number_format_request(SH_BOOKINGS, 4, 504, 3, 5, "dd mmm yyyy", "DATE"),
        number_format_request(SH_BOOKINGS, 4, 504, 5, 6, "0"),
        number_format_request(SH_BOOKINGS, 4, 504, 6, 7, "#,##0"),
        number_format_request(SH_BOOKINGS, 4, 504, 9, 10, "#,##0"),
        column_width_request(SH_BOOKINGS, 0, 1, 80),
        column_width_request(SH_BOOKINGS, 1, 2, 140),
        column_width_request(SH_BOOKINGS, 2, 3, 110),
        column_width_request(SH_BOOKINGS, 3, 5, 110),
        column_width_request(SH_BOOKINGS, 6, 7, 90),
        column_width_request(SH_BOOKINGS, 7, 9, 100),
    ]

    # ---- Expenses Log ----
    reqs += [
        banner_request(SH_EXPENSES, 6),
        header_row_request(SH_EXPENSES, 3, 6),
        *alt_rows_requests(SH_EXPENSES, 4, 504, 0, 6),
        number_format_request(SH_EXPENSES, 4, 504, 0, 1, "dd mmm yyyy", "DATE"),
        number_format_request(SH_EXPENSES, 4, 504, 4, 5, "#,##0"),
        number_format_request(SH_EXPENSES, 4, 504, 5, 6, "mmm yyyy", "DATE"),
        column_width_request(SH_EXPENSES, 0, 1, 110),
        column_width_request(SH_EXPENSES, 1, 2, 140),
        column_width_request(SH_EXPENSES, 2, 3, 120),
        column_width_request(SH_EXPENSES, 3, 4, 220),
        column_width_request(SH_EXPENSES, 5, 6, 110),
        hide_columns_request(SH_EXPENSES, 5, 6),
    ]

    # ---- Recurring Expenses ----
    reqs += [
        banner_request(SH_RECURRING, 8),
        header_row_request(SH_RECURRING, 2, 8),
        *alt_rows_requests(SH_RECURRING, 3, 53, 0, 8),
        number_format_request(SH_RECURRING, 3, 53, 4, 5, "#,##0"),
        number_format_request(SH_RECURRING, 3, 53, 5, 7, "dd mmm yyyy", "DATE"),
        column_width_request(SH_RECURRING, 0, 1, 40),
        column_width_request(SH_RECURRING, 1, 2, 140),
        column_width_request(SH_RECURRING, 2, 3, 120),
        column_width_request(SH_RECURRING, 3, 4, 220),
        column_width_request(SH_RECURRING, 4, 5, 130),
        column_width_request(SH_RECURRING, 5, 7, 120),
    ]

    # ---- Occupancy Heatmap ----
    reqs += [
        banner_request(SH_HEATMAP, 15),
        header_row_request(SH_HEATMAP, 2, 15),
        number_format_request(SH_HEATMAP, 3, 54, 1, 13, "0%"),
        number_format_request(SH_HEATMAP, 3, 54, 13, 14, "0%"),
        number_format_request(SH_HEATMAP, 3, 54, 14, 15, "#,##0"),
        bold_text_request(SH_HEATMAP, 53, 54, 0, 15, 10, PALETTE["deep"]),
        fill_request(SH_HEATMAP, 53, 54, 0, 15, PALETTE["alt"]),
        column_width_request(SH_HEATMAP, 0, 1, 150),
        column_width_request(SH_HEATMAP, 1, 13, 55),
        column_width_request(SH_HEATMAP, 13, 15, 110),
    ]

    # ---- P&L Summary ----
    reqs += [
        fill_request(SH_PNL, 0, 1, 0, 18, PALETTE["deep"]),
        bold_text_request(SH_PNL, 0, 1, 0, 1, 16, "#FFFFFF"),
        bold_text_request(SH_PNL, 1, 2, 0, 4, 11, PALETTE["muted"]),
        header_row_request(SH_PNL, 3, 14),
        number_format_request(SH_PNL, 4, 18, 1, 14, "#,##0"),
        number_format_request(SH_PNL, 17, 18, 1, 14, "0.0%"),
        bold_text_request(SH_PNL, 4, 5, 0, 14, 10, PALETTE["deep"]),
        bold_text_request(SH_PNL, 15, 18, 0, 14, 10, PALETTE["deep"]),
        fill_request(SH_PNL, 15, 16, 0, 14, PALETTE["alt"]),
        fill_request(SH_PNL, 16, 18, 0, 14, PALETTE["alt"]),
        column_width_request(SH_PNL, 0, 1, 130),
        column_width_request(SH_PNL, 1, 13, 75),
        column_width_request(SH_PNL, 13, 14, 90),
    ]

    # ---- Yearly Summary ----
    reqs += [
        banner_request(SH_YEARLY, 7),
        header_row_request(SH_YEARLY, 2, 7),
        *alt_rows_requests(SH_YEARLY, 3, 13, 0, 7),
        number_format_request(SH_YEARLY, 3, 13, 1, 4, "#,##0"),
        number_format_request(SH_YEARLY, 3, 13, 4, 6, "0%"),
        number_format_request(SH_YEARLY, 3, 13, 6, 7, "#,##0"),
        column_width_request(SH_YEARLY, 0, 1, 80),
        column_width_request(SH_YEARLY, 1, 7, 110),
    ]

    # ---- Dashboard ----
    reqs += [
        merge_request(SH_DASH, 0, 1, 0, 13),
        fill_request(SH_DASH, 0, 1, 0, 13, PALETTE["deep"]),
        bold_text_request(SH_DASH, 0, 1, 0, 13, 18, "#FFFFFF"),
        bold_text_request(SH_DASH, 1, 2, 0, 5, 11, PALETTE["muted"]),
        # KPI labels row4
        bold_text_request(SH_DASH, 3, 4, 0, 13, 10, PALETTE["muted"]),
        bold_text_request(SH_DASH, 5, 6, 0, 13, 10, PALETTE["muted"]),
        merge_request(SH_DASH, 3, 4, 0, 3), merge_request(SH_DASH, 3, 4, 3, 6),
        merge_request(SH_DASH, 3, 4, 6, 9), merge_request(SH_DASH, 3, 4, 9, 13),
        merge_request(SH_DASH, 5, 6, 0, 3), merge_request(SH_DASH, 5, 6, 3, 6),
        merge_request(SH_DASH, 5, 6, 6, 9), merge_request(SH_DASH, 5, 6, 9, 13),
        # KPI values row5/7
        merge_request(SH_DASH, 4, 5, 0, 3), merge_request(SH_DASH, 4, 5, 3, 6),
        merge_request(SH_DASH, 4, 5, 6, 9), merge_request(SH_DASH, 4, 5, 9, 13),
        merge_request(SH_DASH, 6, 7, 0, 3), merge_request(SH_DASH, 6, 7, 3, 6),
        merge_request(SH_DASH, 6, 7, 6, 9), merge_request(SH_DASH, 6, 7, 9, 13),
        bold_text_request(SH_DASH, 4, 5, 0, 13, 18, PALETTE["mid"]),
        bold_text_request(SH_DASH, 6, 7, 0, 13, 16, PALETTE["deep"]),
        number_format_request(SH_DASH, 4, 5, 9, 13, "0%"),
        number_format_request(SH_DASH, 6, 7, 9, 13, "0%"),
        column_width_request(SH_DASH, 0, 13, 95),
        hide_rows_request(SH_DASH, 39, 40),
    ]

    # ---- Property Dashboard ----
    reqs += [
        merge_request(SH_PROP_DASH, 0, 1, 0, 13),
        fill_request(SH_PROP_DASH, 0, 1, 0, 13, PALETTE["deep"]),
        bold_text_request(SH_PROP_DASH, 0, 1, 0, 13, 18, "#FFFFFF"),
        bold_text_request(SH_PROP_DASH, 1, 2, 0, 5, 11, PALETTE["muted"]),
        bold_text_request(SH_PROP_DASH, 3, 4, 0, 13, 10, PALETTE["muted"]),
        merge_request(SH_PROP_DASH, 3, 4, 0, 3), merge_request(SH_PROP_DASH, 3, 4, 3, 6),
        merge_request(SH_PROP_DASH, 3, 4, 6, 9), merge_request(SH_PROP_DASH, 3, 4, 9, 13),
        merge_request(SH_PROP_DASH, 4, 5, 0, 3), merge_request(SH_PROP_DASH, 4, 5, 3, 6),
        merge_request(SH_PROP_DASH, 4, 5, 6, 9), merge_request(SH_PROP_DASH, 4, 5, 9, 13),
        bold_text_request(SH_PROP_DASH, 4, 5, 0, 13, 18, PALETTE["mid"]),
        number_format_request(SH_PROP_DASH, 4, 5, 9, 13, "0%"),
        column_width_request(SH_PROP_DASH, 0, 13, 95),
        hide_rows_request(SH_PROP_DASH, 38, 42),
    ]

    # ---- Booking Calendar ----
    reqs += [
        merge_request(SH_CALENDAR, 0, 1, 0, 9),
        fill_request(SH_CALENDAR, 0, 1, 0, 9, PALETTE["deep"]),
        bold_text_request(SH_CALENDAR, 0, 1, 0, 9, 16, "#FFFFFF"),
        bold_text_request(SH_CALENDAR, 1, 2, 0, 6, 11, PALETTE["muted"]),
        header_row_request(SH_CALENDAR, 4, 9),
        number_format_request(SH_CALENDAR, 7, 13, 2, 9, "d mmm", "DATE"),
        bold_text_request(SH_CALENDAR, 7, 13, 2, 9, 10),
        column_width_request(SH_CALENDAR, 0, 2, 90),
        column_width_request(SH_CALENDAR, 2, 9, 100),
        bold_text_request(SH_CALENDAR, 14, 15, 0, 1, 11, PALETTE["deep"]),
        bold_text_request(SH_CALENDAR, 15, 19, 0, 1, 10),
        number_format_request(SH_CALENDAR, 16, 17, 1, 2, "0%"),
        number_format_request(SH_CALENDAR, 17, 18, 1, 2, "#,##0"),
        hide_columns_request(SH_CALENDAR, 10, 14),
        hide_rows_request(SH_CALENDAR, 19, 32),
    ]

    # ---- Settings & Reference ----
    reqs += [
        bold_text_request(SH_SETTINGS, 4, 5, 0, 11, 10, PALETTE["deep"]),
        bold_text_request(SH_SETTINGS, 57, 58, 0, 2, 10, PALETTE["deep"]),
        bold_text_request(SH_SETTINGS, 65, 66, 0, 2, 10, PALETTE["deep"]),
        bold_text_request(SH_SETTINGS, 79, 80, 0, 2, 10, PALETTE["deep"]),
        number_format_request(SH_SETTINGS, 58, 64, 1, 2, "#,##0"),
        number_format_request(SH_SETTINGS, 66, 76, 1, 2, "#,##0"),
        number_format_request(SH_SETTINGS, 80, 130, 1, 2, "#,##0"),
    ]
    return reqs


# =================================================================
# PHASE 3: DATA VALIDATION
# =================================================================
def validation_requests():
    reqs = [
        # Property Setup
        checkbox_request(SH_PROPSETUP, 3, 53, 5, 6),
        # Bookings Log
        range_validation_request(SH_BOOKINGS, 4, 504, 1, 2, PROP_RANGE),
        list_validation_request(SH_BOOKINGS, 4, 504, 7, 8, STATUSES),
        list_validation_request(SH_BOOKINGS, 4, 504, 8, 9, SOURCES),
        # Expenses Log
        range_validation_request(SH_EXPENSES, 4, 504, 1, 2, PROP_RANGE),
        list_validation_request(SH_EXPENSES, 4, 504, 2, 3, CATEGORIES),
        # Recurring Expenses
        range_validation_request(SH_RECURRING, 3, 53, 1, 2, PROP_ALL_RANGE),
        list_validation_request(SH_RECURRING, 3, 53, 2, 3, CATEGORIES),
        checkbox_request(SH_RECURRING, 3, 53, 7, 8),
        # P&L Summary selectors
        list_validation_request(SH_PNL, 1, 2, 1, 2, YEARS),
        range_validation_request(SH_PNL, 1, 2, 3, 4, PROP_ALL_RANGE),
        # Dashboard selectors
        list_validation_request(SH_DASH, 1, 2, 1, 2, YEARS),
        list_validation_request(SH_DASH, 1, 2, 4, 5, ["£", "$"]),
        # Property Dashboard selectors
        range_validation_request(SH_PROP_DASH, 1, 2, 1, 2, PROP_ALL_RANGE),
        list_validation_request(SH_PROP_DASH, 1, 2, 4, 5, YEARS),
        # Booking Calendar selectors
        range_validation_request(SH_CALENDAR, 1, 2, 1, 2, PROP_ALL_RANGE),
        list_validation_request(SH_CALENDAR, 1, 2, 3, 4, MONTHS_FULL),
        list_validation_request(SH_CALENDAR, 1, 2, 5, 6, YEARS),
        # Settings currency selector
        list_validation_request(SH_SETTINGS, 1, 2, 1, 2, ["£", "$"]),
    ]
    return reqs


# =================================================================
# PHASE 4: CONDITIONAL FORMATTING
# =================================================================
def conditional_formatting_requests():
    reqs = []
    idx = 0

    def add(sheet_range, rule_body):
        nonlocal idx
        reqs.append({"addConditionalFormatRule": {
            "rule": {"ranges": [sheet_range], **rule_body}, "index": idx}})
        idx += 1

    # 1. Bookings Log status colors
    add(grid(SH_BOOKINGS, 4, 504, 7, 8), {"booleanRule": {
        "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Confirmed"}]},
        "format": {"backgroundColor": hex_to_rgb("#E4F3EC"), "textFormat": {"foregroundColor": hex_to_rgb(PALETTE["mid"])}},
    }})
    add(grid(SH_BOOKINGS, 4, 504, 7, 8), {"booleanRule": {
        "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Pending"}]},
        "format": {"backgroundColor": hex_to_rgb("#FBF3E1"), "textFormat": {"foregroundColor": hex_to_rgb(PALETTE["gold"])}},
    }})
    add(grid(SH_BOOKINGS, 4, 504, 7, 8), {"booleanRule": {
        "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Cancelled"}]},
        "format": {"backgroundColor": hex_to_rgb("#FBEEEC"), "textFormat": {"foregroundColor": hex_to_rgb(PALETTE["liability"])}},
    }})

    # 2. Occupancy Heatmap gradient color scale (monthly occ % + annual occ %)
    add(grid(SH_HEATMAP, 3, 53, 1, 14), {"gradientRule": {
        "minpoint": {"color": hex_to_rgb(PALETTE["liability"]), "type": "NUMBER", "value": "0"},
        "midpoint": {"color": hex_to_rgb(PALETTE["gold"]), "type": "NUMBER", "value": "0.5"},
        "maxpoint": {"color": hex_to_rgb(PALETTE["mid"]), "type": "NUMBER", "value": "1"},
    }})

    # 3. P&L Summary Net Profit row (17) and Margin row (18) - green/terracotta text
    for row in (16, 17):
        add(grid(SH_PNL, row, row + 1, 1, 14), {"booleanRule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["mid"])}},
        }})
        add(grid(SH_PNL, row, row + 1, 1, 14), {"booleanRule": {
            "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
            "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["liability"])}},
        }})

    # 4. Yearly Summary Profit (col D) and Margin % (col E)
    for col in (3, 4):
        add(grid(SH_YEARLY, 3, 13, col, col + 1), {"booleanRule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["mid"])}},
        }})
        add(grid(SH_YEARLY, 3, 13, col, col + 1), {"booleanRule": {
            "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
            "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["liability"])}},
        }})

    # 5. Booking Calendar grid - highlight confirmed/pending occupied days.
    # CUSTOM_FORMULA rules can't reference other sheets, so these reference the local
    # hidden helper grids (CAL_CONFIRMED_ROW/CAL_PENDING_ROW) which precompute the lookup.
    confirmed_formula = f"=C{CAL_CONFIRMED_ROW}"
    pending_formula = f"=C{CAL_PENDING_ROW}"
    add(grid(SH_CALENDAR, 7, 13, 2, 9), {"booleanRule": {
        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": confirmed_formula}]},
        "format": {"backgroundColor": hex_to_rgb("#E4F3EC")},
    }})
    add(grid(SH_CALENDAR, 7, 13, 2, 9), {"booleanRule": {
        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": pending_formula}]},
        "format": {"backgroundColor": hex_to_rgb("#FBF3E1")},
    }})

    # 6. Property Setup - gray out inactive rows
    add(grid(SH_PROPSETUP, 3, 53, 0, 6), {"booleanRule": {
        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": "=$F4=FALSE"}]},
        "format": {"backgroundColor": hex_to_rgb(PALETTE["bg"]), "textFormat": {"foregroundColor": hex_to_rgb(PALETTE["muted"])}},
    }})

    # 7. Recurring Expenses - gray out inactive rows
    add(grid(SH_RECURRING, 3, 53, 0, 8), {"booleanRule": {
        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": "=$H4=FALSE"}]},
        "format": {"backgroundColor": hex_to_rgb(PALETTE["bg"]), "textFormat": {"foregroundColor": hex_to_rgb(PALETTE["muted"])}},
    }})

    return reqs


# =================================================================
# PHASE 5: CHARTS
# =================================================================
def chart_requests():
    reqs = []

    def add_chart(spec, sheet, row, col, w, h):
        reqs.append({"addChart": {"chart": {
            "spec": spec,
            "position": {"overlayPosition": {
                "anchorCell": {"sheetId": sheet, "rowIndex": row, "columnIndex": col},
                "widthPixels": w, "heightPixels": h,
            }},
        }}})

    # 1. Dashboard: Revenue vs Expenses by Month (from P&L Summary, "All" view)
    add_chart({
        "title": "Revenue vs Expenses by Month",
        "basicChart": {
            "chartType": "COLUMN", "legendPosition": "TOP_LEGEND",
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_PNL, 3, 4, 1, 13)]}}}],
            "series": [
                {"series": {"sourceRange": {"sources": [grid(SH_PNL, 4, 5, 1, 13)]}}, "targetAxis": "LEFT_AXIS", "color": hex_to_rgb(PALETTE["mid"])},
                {"series": {"sourceRange": {"sources": [grid(SH_PNL, 15, 16, 1, 13)]}}, "targetAxis": "LEFT_AXIS", "color": hex_to_rgb(PALETTE["liability"])},
            ],
        },
    }, SH_DASH, 9, 0, 480, 260)

    # 2. Dashboard: Booking Source Split donut
    add_chart({
        "title": "Booking Source Split",
        "pieChart": {
            "legendPosition": "RIGHT_LEGEND",
            "domain": {"sourceRange": {"sources": [grid(SH_SETTINGS, 58, 63, 0, 1)]}},
            "series": {"sourceRange": {"sources": [grid(SH_SETTINGS, 58, 63, 1, 2)]}},
            "pieHole": 0.5,
        },
    }, SH_DASH, 9, 7, 360, 260)

    # 3. Dashboard: Expense Breakdown by Category donut
    add_chart({
        "title": "Expense Breakdown by Category",
        "pieChart": {
            "legendPosition": "RIGHT_LEGEND",
            "domain": {"sourceRange": {"sources": [grid(SH_SETTINGS, 66, 76, 0, 1)]}},
            "series": {"sourceRange": {"sources": [grid(SH_SETTINGS, 66, 76, 1, 2)]}},
            "pieHole": 0.5,
        },
    }, SH_DASH, 23, 0, 400, 260)

    # 4. Dashboard: Top Properties by Profit (seeded properties only)
    add_chart({
        "title": "Top Properties by Profit",
        "basicChart": {
            "chartType": "BAR", "legendPosition": "NO_LEGEND",
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_SETTINGS, 80, 87, 0, 1)]}}}],
            "series": [{"series": {"sourceRange": {"sources": [grid(SH_SETTINGS, 80, 87, 1, 2)]}}, "targetAxis": "BOTTOM_AXIS", "color": hex_to_rgb(PALETTE["gold"])}],
        },
    }, SH_DASH, 23, 7, 400, 260)

    # 5. Property Dashboard: Monthly Income
    add_chart({
        "title": "Monthly Income",
        "basicChart": {
            "chartType": "COLUMN", "legendPosition": "NO_LEGEND",
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_PROP_DASH, 38, 39, 0, 12)]}}}],
            "series": [{"series": {"sourceRange": {"sources": [grid(SH_PROP_DASH, 40, 41, 0, 12)]}}, "targetAxis": "LEFT_AXIS", "color": hex_to_rgb(PALETTE["mid"])}],
        },
    }, SH_PROP_DASH, 9, 0, 480, 260)

    # 6. Property Dashboard: Monthly Occupancy %
    add_chart({
        "title": "Monthly Occupancy %",
        "basicChart": {
            "chartType": "LINE", "legendPosition": "NO_LEGEND",
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_PROP_DASH, 38, 39, 0, 12)]}}}],
            "series": [{"series": {"sourceRange": {"sources": [grid(SH_PROP_DASH, 41, 42, 0, 12)]}}, "targetAxis": "LEFT_AXIS", "color": hex_to_rgb(PALETTE["liquid"]), "lineStyle": {"width": 3}}],
        },
    }, SH_PROP_DASH, 9, 7, 400, 260)

    # 7. Yearly Summary: Revenue vs Expenses by Year
    add_chart({
        "title": "Revenue vs Expenses by Year",
        "basicChart": {
            "chartType": "COLUMN", "legendPosition": "TOP_LEGEND", "headerCount": 0,
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_YEARLY, 3, 13, 0, 1)]}}}],
            "series": [
                {"series": {"sourceRange": {"sources": [grid(SH_YEARLY, 3, 13, 1, 2)]}}, "targetAxis": "LEFT_AXIS", "color": hex_to_rgb(PALETTE["mid"])},
                {"series": {"sourceRange": {"sources": [grid(SH_YEARLY, 3, 13, 2, 3)]}}, "targetAxis": "LEFT_AXIS", "color": hex_to_rgb(PALETTE["liability"])},
            ],
        },
    }, SH_YEARLY, 14, 0, 520, 280)

    # 8. Yearly Summary: Net Profit by Year
    add_chart({
        "title": "Net Profit by Year",
        "basicChart": {
            "chartType": "COLUMN", "legendPosition": "NO_LEGEND", "headerCount": 0,
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_YEARLY, 3, 13, 0, 1)]}}}],
            "series": [{"series": {"sourceRange": {"sources": [grid(SH_YEARLY, 3, 13, 3, 4)]}}, "targetAxis": "LEFT_AXIS", "color": hex_to_rgb(PALETTE["gold"])}],
        },
    }, SH_YEARLY, 14, 7, 400, 280)

    return reqs


# =================================================================
# PHASE 6: FREEZE PANES + HIDE
# =================================================================
def freeze_requests():
    return [
        freeze_request(SH_DASH, rows=2),
        freeze_request(SH_PROP_DASH, rows=2),
        freeze_request(SH_BOOKINGS, rows=4, cols=2),
        freeze_request(SH_CALENDAR, rows=2),
        freeze_request(SH_HEATMAP, rows=3, cols=1),
        freeze_request(SH_EXPENSES, rows=4, cols=1),
        freeze_request(SH_RECURRING, rows=3, cols=1),
        freeze_request(SH_PNL, rows=4, cols=1),
        freeze_request(SH_PROPSETUP, rows=3, cols=1),
        freeze_request(SH_YEARLY, rows=3),
    ]


def chunked(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


if __name__ == "__main__":
    phase = sys.argv[2] if len(sys.argv) > 2 else "all"
    phases = {
        "format": formatting_requests,
        "validate": validation_requests,
        "cf": conditional_formatting_requests,
        "charts": chart_requests,
        "freeze": freeze_requests,
    }
    if phase in ("prep", "all"):
        write_prep()
        print("phase 'prep' done")
    to_run = phases.items() if phase == "all" else ([(phase, phases[phase])] if phase in phases else [])
    for name, fn in to_run:
        reqs = fn()
        for batch in chunked(reqs, 40):
            ss.batch_update(sid, {"requests": batch})
        print(f"phase '{name}' done: {len(reqs)} requests")
