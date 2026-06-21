"""Phase 2-6: formatting, validation, conditional formatting, charts, freeze/hide."""
import sys
from build import (ss, hex_to_rgb, PALETTE, CATEGORY_COLORS, ASSET_CATEGORIES, ALL_CATEGORIES,
                    TYPES, LIQUIDITY, SH_DASH, SH_MONTHLY, SH_YEARLY, SH_ACCOUNTS, SH_SETTINGS, MONTHS)

sid = sys.argv[1] if len(sys.argv) > 1 else open("spreadsheet_id.txt").read().strip()


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
                "textFormat": {"foregroundColor": hex_to_rgb("#FFFFFF"), "bold": True, "fontSize": 12},
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


# =================================================================
# PHASE 2: FORMATTING
# =================================================================
def formatting_requests():
    reqs = []

    # ---- Accounts Setup ----
    reqs += [
        banner_request(SH_ACCOUNTS, 6),
        merge_request(SH_ACCOUNTS, 0, 1, 0, 6),
        header_row_request(SH_ACCOUNTS, 1, 6),
        *alt_rows_requests(SH_ACCOUNTS, 2, 82, 0, 6),
        column_width_request(SH_ACCOUNTS, 1, 2, 200),
        column_width_request(SH_ACCOUNTS, 2, 3, 110),
        column_width_request(SH_ACCOUNTS, 3, 4, 130),
        column_width_request(SH_ACCOUNTS, 4, 5, 100),
        column_width_request(SH_ACCOUNTS, 5, 6, 160),
    ]

    # ---- Monthly Tracker ----
    reqs += [
        banner_request(SH_MONTHLY, 18),
        header_row_request(SH_MONTHLY, 3, 18),
        *alt_rows_requests(SH_MONTHLY, 4, 84, 0, 18),
        number_format_request(SH_MONTHLY, 4, 84, 3, 16, "#,##0.00"),
        number_format_request(SH_MONTHLY, 84, 87, 3, 16, "#,##0.00"),
        bold_text_request(SH_MONTHLY, 84, 87, 0, 18, 11),
        fill_request(SH_MONTHLY, 84, 87, 0, 18, PALETTE["mid"]),
        column_width_request(SH_MONTHLY, 0, 1, 180),
        column_width_request(SH_MONTHLY, 1, 2, 90),
        column_width_request(SH_MONTHLY, 2, 3, 120),
        column_width_request(SH_MONTHLY, 3, 15, 80),
        column_width_request(SH_MONTHLY, 15, 16, 100),
        hide_columns_request(SH_MONTHLY, 16, 18),
    ]

    # ---- Dashboard ----
    reqs += [
        merge_request(SH_DASH, 0, 1, 0, 13),
        fill_request(SH_DASH, 0, 1, 0, 13, PALETTE["deep"]),
        bold_text_request(SH_DASH, 0, 1, 0, 13, 18, "#FFFFFF"),
        # KPI labels
        merge_request(SH_DASH, 3, 4, 0, 3), merge_request(SH_DASH, 3, 4, 4, 7),
        merge_request(SH_DASH, 5, 6, 0, 3), merge_request(SH_DASH, 5, 6, 4, 7),
        bold_text_request(SH_DASH, 3, 4, 0, 7, 11, PALETTE["muted"]),
        bold_text_request(SH_DASH, 5, 6, 0, 7, 11, PALETTE["muted"]),
        # KPI values
        merge_request(SH_DASH, 4, 5, 0, 3), merge_request(SH_DASH, 4, 5, 4, 7),
        merge_request(SH_DASH, 6, 7, 0, 3), merge_request(SH_DASH, 6, 7, 4, 7),
        bold_text_request(SH_DASH, 4, 5, 0, 3, 20, PALETTE["mid"]),
        bold_text_request(SH_DASH, 6, 7, 4, 7, 20, PALETTE["mid"]),
        # Goal input cell - gold
        bold_text_request(SH_DASH, 4, 5, 4, 7, 20, PALETTE["deep"]),
        fill_request(SH_DASH, 4, 5, 4, 7, PALETTE["gold"]),
        number_format_request(SH_DASH, 4, 5, 4, 7, "#,##0"),
        # Goal progress percent
        number_format_request(SH_DASH, 6, 7, 0, 3, "0%"),
        # progress bar row
        merge_request(SH_DASH, 7, 8, 0, 7),
        bold_text_request(SH_DASH, 7, 8, 0, 7, 14, PALETTE["gold"]),
        column_width_request(SH_DASH, 0, 1, 160),
        # hide helper column N
        hide_columns_request(SH_DASH, 13, 14),
    ]

    # ---- Yearly Summary ----
    reqs += [
        banner_request(SH_YEARLY, 6),
        merge_request(SH_YEARLY, 0, 1, 0, 6),
        header_row_request(SH_YEARLY, 2, 6),
        *alt_rows_requests(SH_YEARLY, 3, 13, 0, 6),
        number_format_request(SH_YEARLY, 3, 13, 1, 4, "#,##0.00"),
        number_format_request(SH_YEARLY, 3, 13, 4, 5, "#,##0.00"),
        number_format_request(SH_YEARLY, 3, 13, 5, 6, "0%"),
        column_width_request(SH_YEARLY, 0, 1, 90),
        column_width_request(SH_YEARLY, 1, 5, 130),
    ]

    # ---- Settings & Reference ----
    reqs += [
        bold_text_request(SH_SETTINGS, 4, 5, 0, 8, 10, PALETTE["deep"]),
        bold_text_request(SH_SETTINGS, 17, 18, 0, 2, 10, PALETTE["deep"]),
        bold_text_request(SH_SETTINGS, 27, 28, 0, 2, 10, PALETTE["deep"]),
        bold_text_request(SH_SETTINGS, 32, 33, 0, 2, 10, PALETTE["deep"]),
        number_format_request(SH_SETTINGS, 18, 35, 1, 2, "#,##0.00"),
    ]
    return reqs


# =================================================================
# PHASE 3: DATA VALIDATION
# =================================================================
def validation_requests():
    return [
        list_validation_request(SH_ACCOUNTS, 2, 82, 2, 3, TYPES),
        list_validation_request(SH_ACCOUNTS, 2, 82, 3, 4, ALL_CATEGORIES),
        list_validation_request(SH_ACCOUNTS, 2, 82, 4, 5, LIQUIDITY),
        checkbox_request(SH_ACCOUNTS, 2, 82, 5, 6),
        list_validation_request(SH_MONTHLY, 1, 2, 1, 2, list(range(2023, 2033))),
        list_validation_request(SH_DASH, 1, 2, 1, 2, list(range(2023, 2033))),
        list_validation_request(SH_DASH, 1, 2, 4, 5, ["£", "$"]),
        list_validation_request(SH_SETTINGS, 1, 2, 1, 2, ["£", "$"]),
    ]


# =================================================================
# PHASE 4: CONDITIONAL FORMATTING
# =================================================================
def conditional_formatting_requests():
    reqs = []
    # 1. Monthly Tracker D87:O87 net worth row
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(SH_MONTHLY, 86, 87, 3, 15)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                    "format": {"textFormat": {"foregroundColor": hex_to_rgb("#1B6B53")}},
                },
            }, "index": 0,
        }
    })
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(SH_MONTHLY, 86, 87, 3, 15)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                    "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["liability"])}},
                },
            }, "index": 1,
        }
    })
    # 2. Monthly Tracker liability rows fill
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(SH_MONTHLY, 4, 84, 3, 15)],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": "=$B5=\"Liability\""}]},
                    "format": {"backgroundColor": hex_to_rgb("#FBEEEC")},
                },
            }, "index": 2,
        }
    })
    # 3. Yearly Summary YoY %
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(SH_YEARLY, 3, 13, 5, 6)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
                    "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["mid"])}},
                },
            }, "index": 0,
        }
    })
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(SH_YEARLY, 3, 13, 5, 6)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                    "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["liability"])}},
                },
            }, "index": 1,
        }
    })
    # 4. Dashboard Goal Progress color scale (A7:C7 merged - apply to A7)
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(SH_DASH, 6, 7, 0, 1)],
                "gradientRule": {
                    "minpoint": {"color": hex_to_rgb(PALETTE["liability"]), "type": "NUMBER", "value": "0"},
                    "midpoint": {"color": hex_to_rgb(PALETTE["gold"]), "type": "NUMBER", "value": "0.5"},
                    "maxpoint": {"color": hex_to_rgb(PALETTE["mid"]), "type": "NUMBER", "value": "1"},
                },
            }, "index": 0,
        }
    })
    # 5. Accounts Setup category debt text color
    debt_cats = ["Mortgage", "Loan", "Credit Card", "Other Debt"]
    formula = "=OR(" + ",".join(f'$D3="{c}"' for c in debt_cats) + ")"
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(SH_ACCOUNTS, 2, 82, 3, 4)],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                    "format": {"textFormat": {"foregroundColor": hex_to_rgb(PALETTE["liability"])}},
                },
            }, "index": 0,
        }
    })
    return reqs


# =================================================================
# PHASE 5: CHARTS
# =================================================================
def chart_requests():
    reqs = []
    # 1. Net Worth Growth line chart (Dashboard)
    reqs.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Net Worth Growth",
                    "basicChart": {
                        "chartType": "LINE",
                        "legendPosition": "NO_LEGEND",
                        "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
                        "domains": [{"domain": {"sourceRange": {"sources": [
                            grid(SH_MONTHLY, 3, 4, 3, 15)]}}}],
                        "series": [{
                            "series": {"sourceRange": {"sources": [grid(SH_MONTHLY, 86, 87, 3, 15)]}},
                            "targetAxis": "LEFT_AXIS",
                            "color": hex_to_rgb(PALETTE["mid"]),
                            "lineStyle": {"width": 3},
                        }],
                    },
                },
                "position": {"overlayPosition": {
                    "anchorCell": {"sheetId": SH_DASH, "rowIndex": 9, "columnIndex": 0},
                    "widthPixels": 480, "heightPixels": 260,
                }},
            }
        }
    })
    # 2. Asset Allocation donut
    reqs.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Asset Allocation",
                    "pieChart": {
                        "legendPosition": "RIGHT_LEGEND",
                        "domain": {"sourceRange": {"sources": [grid(SH_SETTINGS, 18, 25, 0, 1)]}},
                        "series": {"sourceRange": {"sources": [grid(SH_SETTINGS, 18, 25, 1, 2)]}},
                        "pieHole": 0.5,
                    },
                },
                "position": {"overlayPosition": {
                    "anchorCell": {"sheetId": SH_DASH, "rowIndex": 9, "columnIndex": 7},
                    "widthPixels": 360, "heightPixels": 260,
                }},
            }
        }
    })
    # 3. Liquid vs Illiquid donut
    reqs.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Liquid vs Illiquid",
                    "pieChart": {
                        "legendPosition": "BOTTOM_LEGEND",
                        "domain": {"sourceRange": {"sources": [grid(SH_SETTINGS, 28, 30, 0, 1)]}},
                        "series": {"sourceRange": {"sources": [grid(SH_SETTINGS, 28, 30, 1, 2)]}},
                        "pieHole": 0.5,
                    },
                },
                "position": {"overlayPosition": {
                    "anchorCell": {"sheetId": SH_DASH, "rowIndex": 23, "columnIndex": 0},
                    "widthPixels": 300, "heightPixels": 240,
                }},
            }
        }
    })
    # 4. Assets vs Liabilities column
    reqs.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Assets vs Liabilities",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "NO_LEGEND",
                        "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
                        "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_SETTINGS, 33, 35, 0, 1)]}}}],
                        "series": [{
                            "series": {"sourceRange": {"sources": [grid(SH_SETTINGS, 33, 35, 1, 2)]}},
                            "targetAxis": "LEFT_AXIS",
                        }],
                    },
                },
                "position": {"overlayPosition": {
                    "anchorCell": {"sheetId": SH_DASH, "rowIndex": 23, "columnIndex": 7},
                    "widthPixels": 300, "heightPixels": 240,
                }},
            }
        }
    })
    # 5. Multi-Year Net Worth column (Yearly Summary)
    reqs.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Multi-Year Net Worth",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "NO_LEGEND",
                        "headerCount": 1,
                        "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
                        "domains": [{"domain": {"sourceRange": {"sources": [grid(SH_YEARLY, 2, 13, 0, 1)]}}}],
                        "series": [{
                            "series": {"sourceRange": {"sources": [grid(SH_YEARLY, 2, 13, 3, 4)]}},
                            "targetAxis": "LEFT_AXIS",
                            "color": hex_to_rgb(PALETTE["mid"]),
                        }],
                    },
                },
                "position": {"overlayPosition": {
                    "anchorCell": {"sheetId": SH_YEARLY, "rowIndex": 14, "columnIndex": 0},
                    "widthPixels": 520, "heightPixels": 280,
                }},
            }
        }
    })
    return reqs


# =================================================================
# PHASE 6: FREEZE PANES
# =================================================================
def freeze_requests():
    return [
        freeze_request(SH_DASH, rows=1),
        freeze_request(SH_MONTHLY, rows=4, cols=3),
        freeze_request(SH_YEARLY, rows=3),
        freeze_request(SH_ACCOUNTS, rows=2),
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
    to_run = phases.items() if phase == "all" else [(phase, phases[phase])]
    for name, fn in to_run:
        reqs = fn()
        for batch in chunked(reqs, 40):
            ss.batch_update(sid, {"requests": batch})
        print(f"phase '{name}' done: {len(reqs)} requests")
