#!/usr/bin/env python3
"""
generate_dashboard.py

Generates Flow_Pressure_Dashboard.xlsx

Dashboard layout
────────────────
  Col A-B  │ Col C     │ Col D-E       │ Col F-I        │ Col G+ (chart)
  ─────────────────────────────────────────────────────────────────────────
  FLOW      │ spacer    │ PRESSURE      │ (data table)   │  dual-axis chart
  selector  │           │ 1/2/3 sel.    │                │  (floats right)
  list      │           │ list          │                │
  ─────────────────────────────────────────────────────────────────────────
  DATA TABLE (rows 27+)
    A=Date  B=FlowRaw   C=FlowAdj
    D=Pres1Raw  E=Pres1Adj  F=Pres2Raw  G=Pres2Adj  H=Pres3Raw  I=Pres3Adj

Controls:
  B2 = Flow Scaling Factor      E2 = Pressure Offset
  B3 = Selected Flow (↓ DV)     (leave blank to hide flow line)
  E3 = Pressure 1 selector
  E4 = Pressure 2 selector      (optional, leave blank to hide)
  E5 = Pressure 3 selector      (optional, leave blank to hide)
  Rows  5-20 = flow name list with conditional highlight
  Rows  7-22 = pressure name list with conditional highlight

Usage:
    python3 generate_dashboard.py
"""

import datetime
import io
import os
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule

# ── Colours ────────────────────────────────────────────────────────────────────
DARK_BLUE    = "1F4E79"
MID_BLUE     = "2E75B6"
LIGHT_BLUE   = "BDD7EE"
DARK_ORANGE  = "C55A11"
LIGHT_ORANGE = "FCE4D6"
WHITE        = "FFFFFF"
LIGHT_GRAY   = "F2F2F2"
DARK_GRAY    = "595959"
GREEN_DARK   = "375623"
GREEN_MID    = "70AD47"
LIGHT_GREEN  = "E2EFDA"
YELLOW_BG    = "FFF2CC"
PURPLE       = "7030A0"

# ── Sample data (taken directly from the problem statement) ────────────────────
FLOW_NAMES = [
    "AL012", "AL013", "AL014", "AL023", "AL028", "AL029",
    "AL035", "AL036", "AL038", "AL063", "AM005", "AM014",
    "AM015", "AM022", "AM024", "AM026",
]

SAMPLE_DATES = [
    datetime.datetime(2026, 12, 1) + datetime.timedelta(minutes=15 * i)
    for i in range(28)
]

FLOW_ROWS = [
    [3.168205, 2.204250, 2.665153, 3.436250, 3.225157, 2.574550, 3.670000, 2.038000, 3.130835, 4.021250, 22.68599, 19.81207, -999, 16.92209, 16.58575, 18.88019],
    [3.190769, 2.225250, 2.681334, 3.411250, 3.236456, 2.583137, 3.683750, 2.049000, 3.145557, 4.048750, 22.45106, 19.72593, -999, 17.10220, 18.43384, 15.47377],
    [3.216410, 2.246250, 2.674793, 3.387500, 3.262896, 2.597148, 3.707500, 2.060000, 3.148535, 4.073750, 22.20831, 19.66328, -999, 16.90642, 18.70008, 16.69538],
    [3.242051, 2.265750, 2.707717, 3.392500, 3.289562, 2.612967, 3.731250, 2.072000, 3.148144, 4.123750, 21.40173, 19.91387, -999, 16.70282, 18.12843, 13.89194],
    [3.265641, 2.286750, 2.735150, 3.433750, 3.293630, 2.629012, 3.757500, 2.084000, 3.140112, 4.150000, 21.87158, 19.90604, -999, 16.92992, 18.02663, 13.51606],
    [3.289231, 2.308500, 2.761070, 3.472500, 3.295212, 2.642345, 3.783750, 2.096000, 3.131763, 4.201250, 21.92640, 19.55365, -999, 17.04738, 19.00549, 14.44010],
    [3.312820, 2.331000, 2.780839, 3.520000, 3.299054, 2.654774, 3.811250, 2.109000, 3.122534, 4.228750, 21.77761, 19.82773, -999, 16.87510, 18.58262, 13.68834],
    [3.336410, 2.353500, 2.799314, 3.547500, 3.304251, 2.667655, 3.836250, 2.122000, 3.124414, 4.278750, 22.25529, 20.39155, -999, 16.96907, 18.61395, 13.22632],
    [3.361026, 2.376000, 2.818888, 3.568750, 3.311031, 2.679632, 3.863750, 2.134000, 3.126025, 4.305000, 21.56618, 19.71027, -999, 17.07087, 18.55913, 15.30149],
    [3.386667, 2.399250, 2.838779, 3.562500, 3.319166, 2.690253, 3.891250, 2.132000, 3.146167, 4.356250, 22.04386, 20.29758, -999, 17.27448, 18.80189, 13.70400],
    [3.408205, 2.421750, 2.859207, 3.532500, 3.327754, 2.700649, 3.920000, 2.126000, 3.177197, 4.382500, 21.87158, 20.21144, -999, 17.17267, 18.28505, 14.43227],
    [3.432820, 2.443500, 2.879171, 3.486250, 3.337245, 2.711044, 3.946250, 2.122000, 3.214893, 4.433750, 21.90290, 19.25608, -999, 16.87510, 17.63509, 15.83399],
    [3.459487, 2.466000, 2.899551, 3.433750, 3.346736, 2.721213, 3.983750, 2.118000, 3.255225, 4.462500, 22.41191, 19.23259, -999, 16.89076, 18.30854, 13.99374],
    [3.483077, 2.488500, 2.920345, 3.401250, 3.356454, 2.731382, 4.020000, 2.114000, 3.294092, 4.516250, 21.76195, 20.37589, -999, 16.99256, 18.44167, 14.43227],
    [3.509743, 2.510250, 2.940358, 3.396250, 3.366623, 2.741778, 4.056250, 2.109000, 3.317798, 4.541250, 21.66015, 19.38920, -999, 16.71848, 17.90134, 18.26155],
    [3.536410, 2.532000, 2.960762, 3.437500, 3.375436, 2.750817, 4.090000, 2.106000, 3.326123, 4.598750, 21.96555, 19.57714, -999, 16.86727, 18.25373, 14.00940],
    [3.557949, 2.553750, 2.980775, 3.505000, 3.382216, 2.759856, 4.126250, 2.101000, 3.323169, 4.630000, 21.89507, 19.89038, -999, 16.89076, 18.12060, 14.13470],
    [3.582564, 2.576250, 3.000666, 3.542500, 3.384476, 2.770930, 4.161250, 2.097000, 3.337646, 4.640000, 21.56618, 20.18012, -999, 17.06304, 18.63744, 13.35944],
    [3.608205, 2.597250, 3.020386, 3.566250, 3.388544, 2.780873, 4.197500, 2.092000, 3.355957, 4.621250, 21.82459, 20.15663, -999, 17.01606, 18.47299, 12.76430],
    [3.630769, 2.619750, 3.036445, 3.565000, 3.394193, 2.791042, 4.231250, 2.088000, 3.357056, 4.585000, 22.43540, 19.61630, -999, 16.89076, 18.33204, 13.22632],
    [3.655385, 2.640000, 3.027903, 3.533750, 3.398035, 2.799856, 4.265000, 2.083000, 3.353101, 4.567500, 21.37041, 20.14096, -999, 17.11786, 18.79405, 14.00157],
    [3.678974, 2.661000, 3.017896, 3.487500, 3.399164, 2.807991, 4.297500, 2.078000, 3.355786, 4.531250, 22.05169, 20.10181, -999, 17.24315, 17.46281, 18.08145],
    [3.700513, 2.677500, 3.007158, 3.435000, 3.406622, 2.844600, 4.328750, 2.073000, 3.358057, 4.512500, 21.56618, 19.55365, -999, 16.91426, 17.43932, 17.95615],
    [3.720000, 2.696250, 2.996370, 3.396250, 3.372951, 2.861775, 4.352500, 2.070000, 3.367798, 4.476250, 21.23728, 19.95302, -999, 17.01606, 17.38450, 19.17776],
    [3.737436, 2.712000, 2.986949, 3.387500, 3.340183, 2.857933, 4.371250, 2.062000, 3.381812, 4.456250, 20.77526, 19.89821, -999, 16.66367, 17.36101, 18.64527],
    [3.744615, 2.723250, 2.944190, 3.408750, 3.320748, 2.852058, 4.382500, 2.053000, 3.390039, 4.417500, 21.59750, 19.35005, -999, 16.69499, 17.40799, 18.93501],
    [3.749743, 2.730750, 2.942969, 3.447500, 3.304025, 2.842115, 4.386250, 2.056000, 3.391602, 4.396250, 21.48004, 19.64762, -999, 16.38959, 17.84652, 20.49335],
    [3.746667, 2.732250, 2.938479, 3.488750, 3.288884, 2.830589, 4.386250, 2.062000, 3.384717, 4.351250, 21.62099, 18.58262, -999, 16.37393, 18.44167, 20.41504],
]

PRES_ROWS = [
    [round(v * 0.97 + 0.3, 6) if v != -999 else -999 for v in row]
    for row in FLOW_ROWS
]

# ── Layout constants ───────────────────────────────────────────────────────────
#
#  Row 1   : Title
#  Row 2   : "Scaling Factor" label (A2) | value (B2) | "Pressure Offset" label (D2) | value (E2)
#  Row 3   : "Select Flow ▼" (A3)  | DROPDOWN (B3) | "Pressure 1 ▼" (D3) | DROPDOWN (E3)
#  Row 4   : "All Flows" header (A4:B4) | "Pressure 2 ▼" (D4) | DROPDOWN (E4)
#  Row 5+  : Flow name list starts (A5:B5) | "Pressure 3 ▼" (D5) | DROPDOWN (E5)
#  Row 6   : Flow list item 2 (A6:B6)   | "All Pressures" header (D6:E6)
#  Rows 7+ : Flow name list continues | Pressure name list starts
#
#  Leave B3 blank to hide Flow from the chart.
#  Leave E4 / E5 blank to show fewer than 3 pressure series.
#
#  Columns A-I are shared between the selector area (rows 1-24) and the
#  data table (rows DATA_HDR_ROW+).  Different row ranges, no conflict.
#
#  Chart: anchored at G1, floats to the right — does not overlap selector area.

LIST_SLOTS       = len(FLOW_NAMES)   # one visible row per sample name

# Flow selector list (left side, rows 5-20 for 16 names)
LIST_START_ROW   = 5
LIST_END_ROW     = LIST_START_ROW + LIST_SLOTS - 1   # = 20

# Pressure selectors are stacked vertically in column E
PRES_SEL_ROWS    = [3, 4, 5]                          # E3, E4, E5
PRES_LIST_HDR    = 6                                  # "All Pressures" header
PRES_LIST_START  = 7                                  # pressure name list start
PRES_LIST_END    = PRES_LIST_START + LIST_SLOTS - 1  # = 22

NOTE_ROW         = PRES_LIST_END + 1                 # = 23  (after both lists)
DATA_SECTION_ROW = NOTE_ROW + 2                      # = 25
DATA_HDR_ROW     = DATA_SECTION_ROW + 1              # = 26
DATA_START_ROW   = DATA_HDR_ROW + 1                  # = 27
DATA_OFFSET      = DATA_START_ROW - 2                # = 25  →  ROW()-25=2 at row 27
DATA_ROWS        = 100

# Column indices
COL_FLOW_LABEL   = 1   # A  – flow list names / Date in data table
COL_FLOW_CTRL    = 2   # B  – flow dropdown, scaling factor / Flow Raw in data table
COL_SPACER       = 3   # C  – Flow Adjusted in data table
COL_PRES_LABEL   = 4   # D  – pressure list names / Pressure 1 Raw in data table
COL_PRES_CTRL    = 5   # E  – pressure 1 dropdown, offset / Pressure 1 Adjusted in data table
COL_PRES2_RAW    = 6   # F  – Pressure 2 Raw in data table
COL_PRES2_ADJ    = 7   # G  – Pressure 2 Adjusted in data table
COL_PRES3_RAW    = 8   # H  – Pressure 3 Raw in data table
COL_PRES3_ADJ    = 9   # I  – Pressure 3 Adjusted in data table

CHART_ANCHOR     = "G1"
CHART_WIDTH_CM   = 20
CHART_HEIGHT_CM  = 14


# ── Style helpers ──────────────────────────────────────────────────────────────

def _thin():
    t = Side(style="thin")
    return Border(left=t, right=t, top=t, bottom=t)


def _medium():
    m = Side(style="medium")
    return Border(left=m, right=m, top=m, bottom=m)


def style_header(cell, text, bg=DARK_BLUE, fg=WHITE, bold=True, sz=10,
                 halign="center"):
    cell.value = text
    cell.fill = PatternFill(fill_type="solid", fgColor=bg)
    cell.font = Font(bold=bold, color=fg, size=sz)
    cell.alignment = Alignment(horizontal=halign, vertical="center")
    cell.border = _thin()


def style_label(cell, text, bold=False, fg="000000", sz=10,
                halign="left", italic=False):
    cell.value = text
    cell.font = Font(bold=bold, color=fg, size=sz, italic=italic)
    cell.alignment = Alignment(horizontal=halign, vertical="center")


def style_input(cell, value, bg=LIGHT_BLUE, fg="000000", bold=True, sz=11,
                num_fmt=None):
    cell.value = value
    cell.fill = PatternFill(fill_type="solid", fgColor=bg)
    cell.font = Font(bold=bold, color=fg, size=sz)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _medium()
    if num_fmt:
        cell.number_format = num_fmt


# ── Raw data sheets ────────────────────────────────────────────────────────────

def build_raw_sheet(ws, title, table_name, data_rows, dates):
    ws.title = title
    ws.row_dimensions[1].height = 22
    ws.column_dimensions["A"].width = 21

    # Headers
    style_header(ws.cell(1, 1), "Date", bg=DARK_BLUE)
    for ci, name in enumerate(FLOW_NAMES, start=2):
        style_header(ws.cell(1, ci), name, bg=MID_BLUE)
        ws.column_dimensions[get_column_letter(ci)].width = 13

    # Data rows
    for ri, (dt, row) in enumerate(zip(dates, data_rows), start=2):
        c = ws.cell(ri, 1, value=dt)
        c.number_format = "DD/MM/YYYY HH:MM"
        c.alignment = Alignment(horizontal="center", vertical="center")
        for ci, val in enumerate(row, start=2):
            dc = ws.cell(ri, ci, value=val)
            dc.number_format = "0.000000"
            dc.alignment = Alignment(horizontal="right")

    # Named Excel Table (makes Power Query setup one-click)
    last_col = get_column_letter(len(FLOW_NAMES) + 1)
    last_row = len(data_rows) + 1
    tbl = Table(displayName=table_name, ref=f"A1:{last_col}{last_row}")
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tbl)

    # Instructions
    note_row = last_row + 2
    nc = ws.cell(note_row, 1,
                 value=("INSTRUCTIONS: Delete the sample rows above (keep Row 1 headers), "
                        "then paste your data from Row 2. Date in Column A; "
                        "flow/pressure names as column headers. -999 = no-data."))
    nc.font = Font(italic=True, color=DARK_GRAY, size=9)
    nc.alignment = Alignment(wrap_text=True)
    ws.merge_cells(f"A{note_row}:{last_col}{note_row}")
    ws.row_dimensions[note_row].height = 30


# ── Dashboard sheet ────────────────────────────────────────────────────────────

def build_dashboard(ws, flow_names):
    ws.title = "Dashboard"

    # Column widths
    ws.column_dimensions["A"].width = 22   # flow names / date
    ws.column_dimensions["B"].width = 14   # flow ctrl / flow raw
    ws.column_dimensions["C"].width = 16   # flow adjusted
    ws.column_dimensions["D"].width = 22   # pressure names / pressure 1 raw
    ws.column_dimensions["E"].width = 14   # pressure 1 ctrl / pressure 1 adjusted
    ws.column_dimensions["F"].width = 14   # pressure 2 raw
    ws.column_dimensions["G"].width = 16   # pressure 2 adjusted
    ws.column_dimensions["H"].width = 14   # pressure 3 raw
    ws.column_dimensions["I"].width = 16   # pressure 3 adjusted

    # ── Row 1: Title ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 30
    tc = ws.cell(1, 1, value="Flow & Pressure Analysis Dashboard")
    tc.font = Font(bold=True, color=WHITE, size=14)
    tc.fill = PatternFill(fill_type="solid", fgColor=DARK_BLUE)
    tc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:F1")

    # ── Row 2: Scaling factor (left) | Pressure offset (right) ───────────────
    ws.row_dimensions[2].height = 26

    style_label(ws.cell(2, COL_FLOW_LABEL),
                "Flow Scaling Factor:", bold=True, sz=10)
    style_input(ws.cell(2, COL_FLOW_CTRL),
                1.0, bg=LIGHT_ORANGE, num_fmt="0.000")

    style_label(ws.cell(2, COL_PRES_LABEL),
                "Pressure Offset:", bold=True, sz=10)
    style_input(ws.cell(2, COL_PRES_CTRL),
                0.0, bg=LIGHT_GREEN, num_fmt="0.000")

    # ── Row 3: Flow selector (left) | Pressure 1 selector (right) ────────────
    ws.row_dimensions[3].height = 30

    style_label(ws.cell(3, COL_FLOW_LABEL),
                "Select Flow  ▼  (blank = hide)", bold=True, fg=DARK_BLUE, sz=11)
    flow_cell = ws.cell(3, COL_FLOW_CTRL)
    style_input(flow_cell, flow_names[0], bg=LIGHT_BLUE, sz=12)

    # Flow DV — allow blank so the user can leave it empty to hide the series
    dv_flow = DataValidation(
        type="list",
        formula1="'Raw Flow Data'!$B$1:$GR$1",
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=False,
    )
    ws.add_data_validation(dv_flow)
    dv_flow.add(flow_cell)

    style_label(ws.cell(3, COL_PRES_LABEL),
                "Pressure 1  ▼", bold=True, fg=DARK_ORANGE, sz=11)
    pres1_cell = ws.cell(3, COL_PRES_CTRL)
    style_input(pres1_cell, flow_names[0], bg=LIGHT_ORANGE, sz=12)

    dv_pres1 = DataValidation(
        type="list",
        formula1="'Raw Pressure Data'!$B$1:$GR$1",
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=False,
    )
    ws.add_data_validation(dv_pres1)
    dv_pres1.add(pres1_cell)

    # ── Row 4: "All Flows" header (left) | Pressure 2 selector (right) ───────
    ws.row_dimensions[4].height = 26

    fh = ws.cell(4, COL_FLOW_LABEL,
                 value="All Flows  (current selection highlighted)")
    fh.font = Font(bold=True, color=WHITE, size=9)
    fh.fill = PatternFill(fill_type="solid", fgColor=MID_BLUE)
    fh.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A4:B4")

    style_label(ws.cell(4, COL_PRES_LABEL),
                "Pressure 2  ▼  (optional)", bold=True, fg=DARK_ORANGE, sz=11)
    pres2_cell = ws.cell(4, COL_PRES_CTRL)
    style_input(pres2_cell, "", bg=LIGHT_ORANGE, sz=12)

    dv_pres2 = DataValidation(
        type="list",
        formula1="'Raw Pressure Data'!$B$1:$GR$1",
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=False,
    )
    ws.add_data_validation(dv_pres2)
    dv_pres2.add(pres2_cell)

    # ── Row 5: Flow list item 1 (left) | Pressure 3 selector (right) ─────────
    ws.row_dimensions[5].height = 26

    style_label(ws.cell(5, COL_PRES_LABEL),
                "Pressure 3  ▼  (optional)", bold=True, fg=DARK_ORANGE, sz=11)
    pres3_cell = ws.cell(5, COL_PRES_CTRL)
    style_input(pres3_cell, "", bg=LIGHT_ORANGE, sz=12)

    dv_pres3 = DataValidation(
        type="list",
        formula1="'Raw Pressure Data'!$B$1:$GR$1",
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=False,
    )
    ws.add_data_validation(dv_pres3)
    dv_pres3.add(pres3_cell)

    # ── Row 6: Flow list item 2 (left) | "All Pressures" header (right) ──────
    ws.row_dimensions[PRES_LIST_HDR].height = 20

    ph = ws.cell(PRES_LIST_HDR, COL_PRES_LABEL,
                 value="All Pressures  (highlighted = any of P1/P2/P3)")
    ph.font = Font(bold=True, color=WHITE, size=9)
    ph.fill = PatternFill(fill_type="solid", fgColor=DARK_ORANGE)
    ph.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(f"D{PRES_LIST_HDR}:E{PRES_LIST_HDR}")

    # ── Rows 5 – LIST_END_ROW (flow) / PRES_LIST_END (pressure): name lists ──
    # Flow list: rows LIST_START_ROW (5) to LIST_END_ROW (20)
    # Pressure list: rows PRES_LIST_START (7) to PRES_LIST_END (22)
    for i, name in enumerate(flow_names):
        # Flow list row
        rf = LIST_START_ROW + i
        ws.row_dimensions[rf].height = 18
        fc = ws.cell(rf, COL_FLOW_LABEL, value=name)
        fc.alignment = Alignment(horizontal="center", vertical="center")
        fc.border = _thin()
        ws.merge_cells(f"A{rf}:B{rf}")

        # Pressure list row (starts 2 rows later)
        rp = PRES_LIST_START + i
        ws.row_dimensions[rp].height = 18
        pc = ws.cell(rp, COL_PRES_LABEL, value=name)
        pc.alignment = Alignment(horizontal="center", vertical="center")
        pc.border = _thin()
        ws.merge_cells(f"D{rp}:E{rp}")

    # ── Conditional formatting — highlight selected items ─────────────────────
    hi_flow_fill = PatternFill(fill_type="solid", fgColor=LIGHT_BLUE)
    hi_flow_font = Font(bold=True, color=DARK_BLUE, size=10)
    ws.conditional_formatting.add(
        f"A{LIST_START_ROW}:B{LIST_END_ROW}",
        FormulaRule(
            formula=[f"$A{LIST_START_ROW}=$B$3"],
            fill=hi_flow_fill,
            font=hi_flow_font,
        ),
    )

    # Pressure CF: highlight if name matches ANY of the three pressure selectors
    hi_pres_fill = PatternFill(fill_type="solid", fgColor=LIGHT_ORANGE)
    hi_pres_font = Font(bold=True, color=DARK_ORANGE, size=10)
    ws.conditional_formatting.add(
        f"D{PRES_LIST_START}:E{PRES_LIST_END}",
        FormulaRule(
            formula=[f"OR($D{PRES_LIST_START}=$E$3,"
                     f"$D{PRES_LIST_START}=$E$4,"
                     f"$D{PRES_LIST_START}=$E$5)"],
            fill=hi_pres_fill,
            font=hi_pres_font,
        ),
    )

    # ── Note row ──────────────────────────────────────────────────────────────
    ws.row_dimensions[NOTE_ROW].height = 42
    nc = ws.cell(NOTE_ROW, 1,
                 value=("ℹ  Use Pressure 1 / 2 / 3 dropdowns (E3:E5) to overlay up to 3 pressure "
                        "series on the chart.  Leave E4 or E5 blank to show fewer series.  "
                        "Leave B3 blank to hide the flow line and show pressure only.  "
                        "After pasting your own data, right-click each selector cell → "
                        "Data Validation → update the Source to match your column headers.  "
                        "Values of -999 are treated as no-data and excluded."))
    nc.font = Font(italic=True, color=DARK_GRAY, size=9)
    nc.alignment = Alignment(wrap_text=True)
    ws.merge_cells(f"A{NOTE_ROW}:I{NOTE_ROW}")

    # ── DATA TABLE section banner ──────────────────────────────────────────────
    ws.row_dimensions[DATA_SECTION_ROW].height = 20
    sc = ws.cell(DATA_SECTION_ROW, 1,
                 value="FORMULA TABLE  —  updates automatically when you change the selections or adjustments above")
    sc.font = Font(bold=True, color=WHITE, size=9)
    sc.fill = PatternFill(fill_type="solid", fgColor=DARK_BLUE)
    sc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(f"A{DATA_SECTION_ROW}:I{DATA_SECTION_ROW}")

    # ── DATA TABLE column headers ──────────────────────────────────────────────
    ws.row_dimensions[DATA_HDR_ROW].height = 22
    hdr_cols = [
        (1, "Date",               DARK_BLUE),
        (2, "Flow (Raw)",         MID_BLUE),
        (3, "Flow Adjusted",      DARK_ORANGE),
        (4, "Pressure 1 (Raw)",   MID_BLUE),
        (5, "Pressure 1 Adj.",    GREEN_DARK),
        (6, "Pressure 2 (Raw)",   MID_BLUE),
        (7, "Pressure 2 Adj.",    GREEN_DARK),
        (8, "Pressure 3 (Raw)",   MID_BLUE),
        (9, "Pressure 3 Adj.",    GREEN_DARK),
    ]
    for ci, txt, bg in hdr_cols:
        style_header(ws.cell(DATA_HDR_ROW, ci), txt, bg=bg)

    # ── DATA TABLE formula rows ────────────────────────────────────────────────
    # DATA_OFFSET = 25  →  ROW() - 25 = 2 at row 27 (first data row)
    for r in range(DATA_START_ROW, DATA_START_ROW + DATA_ROWS):
        ws.row_dimensions[r].height = 15
        alt = (r % 2 == 0)

        # A: Date (from Raw Flow Data col A)
        ac = ws.cell(r, 1)
        ac.value = (
            f"=IFERROR("
            f"IF(INDEX('Raw Flow Data'!$A:$A,ROW()-{DATA_OFFSET})=\"\",\"\","
            f"INDEX('Raw Flow Data'!$A:$A,ROW()-{DATA_OFFSET})),\"\")"
        )
        ac.number_format = "DD/MM/YYYY HH:MM"
        ac.alignment = Alignment(horizontal="center")
        if alt:
            ac.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # B: Flow Raw (looked up by name in B3)
        bc = ws.cell(r, 2)
        bc.value = (
            f"=IFERROR("
            f"IF($B$3=\"\",\"\","
            f"INDEX('Raw Flow Data'!$A:$ZZ,ROW()-{DATA_OFFSET},"
            f"MATCH($B$3,'Raw Flow Data'!$1:$1,0))),\"\")"
        )
        bc.number_format = "0.000"
        bc.alignment = Alignment(horizontal="right")
        if alt:
            bc.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # C: Flow Adjusted  =  Flow Raw × Scaling Factor (B2)
        cc = ws.cell(r, 3)
        cc.value = f"=IF(OR(B{r}=\"\",B{r}=-999),\"\",B{r}*$B$2)"
        cc.number_format = "0.000"
        cc.alignment = Alignment(horizontal="right")
        cc.fill = PatternFill(fill_type="solid",
                              fgColor=LIGHT_ORANGE if alt else YELLOW_BG)

        # D: Pressure 1 Raw (looked up by name in E3)
        dc = ws.cell(r, 4)
        dc.value = (
            f"=IFERROR("
            f"IF($E$3=\"\",\"\","
            f"INDEX('Raw Pressure Data'!$A:$ZZ,ROW()-{DATA_OFFSET},"
            f"MATCH($E$3,'Raw Pressure Data'!$1:$1,0))),\"\")"
        )
        dc.number_format = "0.000"
        dc.alignment = Alignment(horizontal="right")
        if alt:
            dc.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # E: Pressure 1 Adjusted  =  Pressure Raw + Offset (E2)
        ec = ws.cell(r, 5)
        ec.value = f"=IF(OR(D{r}=\"\",D{r}=-999),\"\",D{r}+$E$2)"
        ec.number_format = "0.000"
        ec.alignment = Alignment(horizontal="right")
        ec.fill = PatternFill(fill_type="solid",
                              fgColor=LIGHT_GREEN if alt else "D6E4BC")

        # F: Pressure 2 Raw (looked up by name in E4)
        fc = ws.cell(r, 6)
        fc.value = (
            f"=IFERROR("
            f"IF($E$4=\"\",\"\","
            f"INDEX('Raw Pressure Data'!$A:$ZZ,ROW()-{DATA_OFFSET},"
            f"MATCH($E$4,'Raw Pressure Data'!$1:$1,0))),\"\")"
        )
        fc.number_format = "0.000"
        fc.alignment = Alignment(horizontal="right")
        if alt:
            fc.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # G: Pressure 2 Adjusted
        gc = ws.cell(r, 7)
        gc.value = f"=IF(OR(F{r}=\"\",F{r}=-999),\"\",F{r}+$E$2)"
        gc.number_format = "0.000"
        gc.alignment = Alignment(horizontal="right")
        gc.fill = PatternFill(fill_type="solid",
                              fgColor=LIGHT_GREEN if alt else "D6E4BC")

        # H: Pressure 3 Raw (looked up by name in E5)
        hc = ws.cell(r, 8)
        hc.value = (
            f"=IFERROR("
            f"IF($E$5=\"\",\"\","
            f"INDEX('Raw Pressure Data'!$A:$ZZ,ROW()-{DATA_OFFSET},"
            f"MATCH($E$5,'Raw Pressure Data'!$1:$1,0))),\"\")"
        )
        hc.number_format = "0.000"
        hc.alignment = Alignment(horizontal="right")
        if alt:
            hc.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # I: Pressure 3 Adjusted
        ic = ws.cell(r, 9)
        ic.value = f"=IF(OR(H{r}=\"\",H{r}=-999),\"\",H{r}+$E$2)"
        ic.number_format = "0.000"
        ic.alignment = Alignment(horizontal="right")
        ic.fill = PatternFill(fill_type="solid",
                              fgColor=LIGHT_GREEN if alt else "D6E4BC")

    # Copy-down hint
    hint_r = DATA_START_ROW + DATA_ROWS
    ws.row_dimensions[hint_r].height = 24
    hc = ws.cell(hint_r, 1,
                 value=(f"↑ Formulas cover {DATA_ROWS} rows (A{DATA_START_ROW}:I{hint_r - 1}). "
                        "For more data: select that range and copy-paste downward."))
    hc.font = Font(italic=True, color=DARK_GRAY, size=9)
    hc.alignment = Alignment(wrap_text=True)
    ws.merge_cells(f"A{hint_r}:I{hint_r}")

    # ── Chart ─────────────────────────────────────────────────────────────────
    _add_chart(ws)


def _add_chart(ws):
    """Dual-axis line chart — Flow Adjusted (primary left) + up to 3 Pressures (secondary right).
    This chart object is replaced by _patch_chart_xml; this just creates the xlsx entry."""

    # Primary chart — Flow Adjusted (col C)
    c1 = LineChart()
    c1.title = "Flow & Pressure Analysis"
    c1.y_axis.title = "Flow Adjusted"
    c1.y_axis.axId = 100
    c1.x_axis.axId = 100
    c1.style = 10
    c1.width  = CHART_WIDTH_CM
    c1.height = CHART_HEIGHT_CM

    last_row = DATA_START_ROW + DATA_ROWS - 1
    dates_ref = Reference(ws, min_col=1, max_col=1,
                          min_row=DATA_START_ROW, max_row=last_row)

    flow_ref = Reference(ws, min_col=COL_SPACER, max_col=COL_SPACER,
                         min_row=DATA_HDR_ROW, max_row=last_row)
    c1.add_data(flow_ref, titles_from_data=True)
    c1.set_categories(dates_ref)

    # Secondary chart — all 3 pressure series share the right Y axis
    c2 = LineChart()
    c2.y_axis.title  = "Pressure Adjusted"
    c2.y_axis.axId   = 200
    c2.y_axis.crosses = "max"
    c2.x_axis.axId   = 100
    c2.x_axis.delete = True

    for pcol in (COL_PRES_CTRL, COL_PRES2_ADJ, COL_PRES3_ADJ):
        ref = Reference(ws, min_col=pcol, max_col=pcol,
                        min_row=DATA_HDR_ROW, max_row=last_row)
        c2.add_data(ref, titles_from_data=True)
    c2.set_categories(dates_ref)

    c1 += c2

    try:
        colors = [MID_BLUE, DARK_ORANGE, GREEN_MID, PURPLE]
        for i, col in enumerate(colors):
            c1.series[i].graphicalProperties.line.solidFill = col
            c1.series[i].graphicalProperties.line.width = 20000
    except Exception:
        pass

    c1.anchor = CHART_ANCHOR
    ws.add_chart(c1)


def _build_correct_chart_xml():
    """Return a valid dual-axis line-chart XML string with 4 series and pre-computed
    numCache so the chart renders immediately on first open without requiring a
    formula recalculation cycle.

    Series layout:
      lineChart 1  (primary left Y axis 1002):
        idx=0  Flow Adjusted           col C  →  blue
      lineChart 2  (secondary right Y axis 1003):
        idx=1  Pressure 1 Adjusted     col E  →  orange
        idx=2  Pressure 2 Adjusted     col G  →  green   (blank = invisible)
        idx=3  Pressure 3 Adjusted     col I  →  purple  (blank = invisible)

    Leave B3 blank to hide Flow; leave E4/E5 blank to hide Pressure 2/3.
    """
    import datetime as _dt

    last_data_row = DATA_START_ROW + DATA_ROWS - 1
    date_col  = get_column_letter(COL_FLOW_LABEL)   # A
    flow_col  = get_column_letter(COL_SPACER)        # C
    p1_col    = get_column_letter(COL_PRES_CTRL)     # E
    p2_col    = get_column_letter(COL_PRES2_ADJ)     # G
    p3_col    = get_column_letter(COL_PRES3_ADJ)     # I

    def _ref(col, start=DATA_START_ROW, end=last_data_row):
        return f"Dashboard!${col}${start}:${col}${end}"

    def _hdr(col):
        return f"Dashboard!${col}${DATA_HDR_ROW}"

    # ── Pre-compute numCache values from the embedded sample data ────────────
    # Excel serial dates: epoch = 1899-12-30 for Python's perspective
    _epoch = _dt.datetime(1899, 12, 30)
    n = len(SAMPLE_DATES)

    def _date_pts():
        lines = []
        for i, d in enumerate(SAMPLE_DATES):
            delta = d - _epoch
            v = delta.days + delta.seconds / 86400.0
            lines.append(f'        <c:pt idx="{i}"><c:v>{v:.8f}</c:v></c:pt>')
        return "\n".join(lines)

    def _val_pts(column_index):
        """column_index: 0-based index into each data row."""
        lines = []
        for i, row in enumerate(FLOW_ROWS):
            try:
                v = row[column_index]
                if v == -999:
                    continue
                lines.append(f'        <c:pt idx="{i}"><c:v>{v}</c:v></c:pt>')
            except IndexError:
                pass
        return "\n".join(lines)

    def _pres_pts(column_index):
        lines = []
        for i, row in enumerate(PRES_ROWS):
            try:
                v = row[column_index]
                if v == -999:
                    continue
                lines.append(f'        <c:pt idx="{i}"><c:v>{v}</c:v></c:pt>')
            except IndexError:
                pass
        return "\n".join(lines)

    # Flow Adjusted = flow_raw × 1.0 (B2 default = 1.0), using FLOW_NAMES[0]=AL012 → col 0
    flow_pts  = _val_pts(0)
    # Pressure 1 Adjusted = pres_raw + 0.0 (E2 default = 0), using PRES_NAMES[0]=AL012 → col 0
    pres1_pts = _pres_pts(0)
    date_pts  = _date_pts()

    def _num_cache(fmt, pt_count, pts_xml):
        return (f"<c:numCache>"
                f"<c:formatCode>{fmt}</c:formatCode>"
                f"<c:ptCount val=\"{pt_count}\"/>\n"
                f"{pts_xml}\n"
                f"      </c:numCache>")

    # Dates cache (shared by all series)
    dates_cache   = _num_cache("m/d/yy h:mm", n, date_pts)
    flow_cache    = _num_cache("0.000", n, flow_pts)
    pres1_cache   = _num_cache("0.000", n, pres1_pts)
    # P2 and P3 start empty (selectors E4/E5 are blank by default)
    empty_cache   = '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numCache>'

    def _ser(idx, color, title_ref, dates_cache_xml, val_cache_xml):
        # kept for reference only — _series_xml is the actual builder used below
        pass

    # Build each series entry
    def _series_xml(idx, color, title_ref, val_ref, dates_cache_xml, val_cache_xml):
        return f"""\
        <c:ser>
          <c:idx val="{idx}"/>
          <c:order val="{idx}"/>
          <c:tx><c:strRef><c:f>{title_ref}</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="20000"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>
          <c:marker><c:symbol val="none"/></c:marker>
          <c:smooth val="0"/>
          <c:cat><c:numRef><c:f>{_ref(date_col)}</c:f>{dates_cache_xml}</c:numRef></c:cat>
          <c:val><c:numRef><c:f>{val_ref}</c:f>{val_cache_xml}</c:numRef></c:val>
        </c:ser>"""

    ser0 = _series_xml(0, MID_BLUE,    _hdr(flow_col), _ref(flow_col), dates_cache, flow_cache)
    ser1 = _series_xml(1, DARK_ORANGE, _hdr(p1_col),   _ref(p1_col),  dates_cache, pres1_cache)
    ser2 = _series_xml(2, GREEN_MID,   _hdr(p2_col),   _ref(p2_col),  dates_cache, empty_cache)
    ser3 = _series_xml(3, PURPLE,      _hdr(p3_col),   _ref(p3_col),  dates_cache, empty_cache)

    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="en-US"/>
  <c:style val="10"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>Flow &amp; Pressure Analysis</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
{ser0}
        <c:marker><c:symbol val="none"/></c:marker>
        <c:smooth val="0"/>
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:lineChart>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
{ser1}
{ser2}
{ser3}
        <c:marker><c:symbol val="none"/></c:marker>
        <c:smooth val="0"/>
        <c:axId val="1001"/>
        <c:axId val="1003"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="d/m/yy h:mm" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
        <c:auto val="1"/>
        <c:lblAlign val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:title>
          <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>Flow Adjusted</a:t></a:r></a:p></c:rich></c:tx>
          <c:overlay val="0"/>
        </c:title>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="1003"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="r"/>
        <c:majorGridlines/>
        <c:title>
          <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>Pressure Adjusted</a:t></a:r></a:p></c:rich></c:tx>
          <c:overlay val="0"/>
        </c:title>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
        <c:crosses val="max"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
</c:chartSpace>"""


def _patch_chart_xml(xlsx_path):
    """Replace the broken chart XML written by openpyxl with a valid version.

    Also writes [Content_Types].xml first in the ZIP, as required by the
    Open Packaging Convention (OPC/OOXML) specification.
    """
    correct_xml = _build_correct_chart_xml().encode("utf-8")
    tmp_path = xlsx_path + ".patching"
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            # OPC spec requires [Content_Types].xml to be the first entry
            ct_item = None
            for item in zin.infolist():
                if item.filename == "[Content_Types].xml":
                    ct_item = item
                    break
            if ct_item:
                zout.writestr(ct_item, zin.read(ct_item.filename))

            for item in zin.infolist():
                if item.filename == "[Content_Types].xml":
                    continue  # already written first
                if item.filename == "xl/charts/chart1.xml":
                    zout.writestr(item, correct_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))
    os.replace(tmp_path, xlsx_path)


# ── MOD output sheets ──────────────────────────────────────────────────────────

def build_mod_sheet(ws, title):
    ws.title = title
    ws.row_dimensions[1].height = 22
    ws.column_dimensions["A"].width = 21
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 20

    for ci, (h, bg) in enumerate(
        [("Date", DARK_BLUE), ("Name", DARK_BLUE), ("Adjusted Value", DARK_BLUE)],
        start=1,
    ):
        style_header(ws.cell(1, ci), h, bg=bg)

    ws.row_dimensions[2].height = 18
    ic = ws.cell(2, 1,
                 value=("Rows are appended here each time you run SaveToMOD. "
                        "History is preserved — existing rows are never overwritten."))
    ic.font = Font(italic=True, color=DARK_GRAY, size=9)
    ic.alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:C2")


# ── Instructions sheet ─────────────────────────────────────────────────────────

def build_instructions(ws):
    ws.title = "Instructions"
    ws.column_dimensions["A"].width = 115
    ws.sheet_view.showGridLines = False

    def section(row, text, bg=DARK_BLUE, sz=11, height=26):
        ws.row_dimensions[row].height = height
        c = ws.cell(row, 1, value=text)
        c.font = Font(bold=True, color=WHITE, size=sz)
        c.fill = PatternFill(fill_type="solid", fgColor=bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    def body(row, text, fg="000000", sz=10, indent=0, bold=False,
             italic=False, bg=None, height=17):
        ws.row_dimensions[row].height = height
        c = ws.cell(row, 1, value=text)
        c.font = Font(color=fg, size=sz, bold=bold, italic=italic)
        c.alignment = Alignment(horizontal="left", vertical="center",
                                indent=indent, wrap_text=True)
        if bg:
            c.fill = PatternFill(fill_type="solid", fgColor=bg)

    def blank(row, h=8):
        ws.row_dimensions[row].height = h

    r = 1

    section(r, "Flow & Pressure Dashboard — Instructions", bg=DARK_BLUE, sz=14, height=32)
    r += 1; blank(r); r += 1

    # ── 1. Quick start ────────────────────────────────────────────────────────
    section(r, "1.  QUICK START  (works immediately — no setup needed)", bg=MID_BLUE)
    r += 1
    for line in [
        "Step 1:  Paste your flow data into 'Raw Flow Data' (delete sample rows, keep Row 1 headers).",
        "Step 2:  Paste your pressure data into 'Raw Pressure Data' (same format).",
        "Step 3:  Go to the Dashboard sheet.",
        "Step 4:  Use 'Pressure 1 ▼' (cell E3) to pick the first pressure to display.",
        "         Optionally use 'Pressure 2 ▼' (E4) and 'Pressure 3 ▼' (E5) for more series.",
        "         Leave E4 / E5 blank to show fewer than 3 pressure lines on the chart.",
        "Step 5:  Use 'Select Flow ▼' (cell B3) to pick a flow.",
        "         Leave B3 blank to hide the Flow line and show only pressure(s).",
        "Step 6:  Adjust 'Flow Scaling Factor' (B2) and 'Pressure Offset' (E2) if needed.",
        "Step 7:  The chart (right side) and formula table (below) update instantly.",
        "Step 8:  When satisfied, run the SaveToMOD macro to store the adjusted data.",
        "",
        "NOTE:  After pasting your own data, right-click each selector cell (B3, E3-E5)",
        "       → Data Validation → update the Source to cover your column range,",
        "       e.g.  'Raw Flow Data'!$B$1:$BZ$1",
        "",
        "NOTE:  The formula table covers 100 rows (A27:I126). For longer datasets, select",
        "       that range and copy-paste downward as far as needed.",
        "",
        "NOTE:  -999 values are treated as no-data and are excluded from all calculations.",
    ]:
        body(r, line, indent=2)
        r += 1
    blank(r); r += 1

    # ── 2. Power Query (optional enhancement) ─────────────────────────────────
    section(r, "2.  POWER QUERY SETUP  (optional — recommended for very large datasets)",
            bg=DARK_ORANGE)
    r += 1
    for line in [
        "The Raw data sheets are already set up as Excel Tables (FlowData, PressureData).",
        "Power Query can load these tables, unpivot them, and merge them for use in PivotTables.",
        "",
        "Step 1:  Data tab → Get Data → From Table/Range → select the FlowData table.",
        "Step 2:  In Power Query Editor: select the Date column, then Home → Unpivot Other Columns.",
        "         Rename 'Attribute' → 'Flow Name',  'Value' → 'Flow Value'.",
        "Step 3:  Close & Load To… → Only Create Connection.  Name the query  FlowLong.",
        "Step 4:  Repeat for PressureData.  Name the query  PressureLong.",
        "Step 5:  Merge the two queries on Date + Name to get a combined table.",
        "Step 6:  Add calculated columns:  Flow Adjusted = [Flow Value] × scaling_factor",
        "                                  Pressure Adjusted = [Pressure Value] + offset",
        "Step 7:  Load the merged query to a sheet and build a PivotTable + Slicer on top of it.",
        "",
        "After pasting new data:  Data tab → Refresh All  (Ctrl+Alt+F5).",
    ]:
        body(r, line, indent=2)
        r += 1
    blank(r); r += 1

    # ── 3. PivotTable & Slicer ────────────────────────────────────────────────
    section(r, "3.  PIVOTTABLE + SLICER  (optional — for interactive multi-flow comparison)",
            bg=GREEN_DARK)
    r += 1
    for line in [
        "Once the Power Query merged table is loaded to a sheet:",
        "  • Insert → PivotTable",
        "  • Rows: Date    Values: Flow Adjusted, Pressure Adjusted",
        "  • PivotTable Analyze → Insert Slicer → tick 'Flow Name' → OK",
        "  • Click a flow name in the Slicer to filter instantly",
        "  • Insert → PivotChart → Line → add Secondary Axis to the Pressure series",
    ]:
        body(r, line, indent=2)
        r += 1
    blank(r); r += 1

    # ── 4. VBA Save button ────────────────────────────────────────────────────
    section(r, "4.  VBA SAVE BUTTON  (paste this into a Module — Alt+F11 → Insert → Module)",
            bg=PURPLE)
    r += 1
    body(r, "Assign the macro to a button on the Dashboard: Developer tab → Insert → Button (Form Control).",
         indent=2, italic=True, fg=DARK_GRAY)
    r += 1
    blank(r); r += 1

    vba = r"""' ═══════════════════════════════════════════════════════════════════════════
' SaveToMOD  —  appends current adjusted data to MOD Flow and MOD Pressure
' Reads:  B3 (selected flow), E3 (pressure 1), B2 (scaling), E2 (offset)
'         Data table rows 27+, col C = Flow Adjusted, col E = Pressure 1 Adjusted
'         (Pressure 2/3 in cols G/I are not saved by this macro — extend as needed)
' ═══════════════════════════════════════════════════════════════════════════
Sub SaveToMOD()

    Dim wsDash    As Worksheet
    Dim wsModFlow As Worksheet
    Dim wsModPres As Worksheet

    On Error Resume Next
    Set wsDash    = Worksheets("Dashboard")
    Set wsModFlow = Worksheets("MOD Flow")
    Set wsModPres = Worksheets("MOD Pressure")
    On Error GoTo 0

    If wsDash Is Nothing Or wsModFlow Is Nothing Or wsModPres Is Nothing Then
        MsgBox "Could not find Dashboard, MOD Flow, or MOD Pressure sheet.", _
               vbCritical, "Sheet Missing"
        Exit Sub
    End If

    ' ── Validate selections ──────────────────────────────────────────────────
    Dim flowName As String
    Dim presName As String
    flowName = Trim(wsDash.Range("B3").Value)
    presName = Trim(wsDash.Range("E3").Value)

    If flowName = "" Then
        MsgBox "Please select a Flow from the dropdown (cell B3).", _
               vbExclamation, "No Flow Selected"
        Exit Sub
    End If
    If presName = "" Then
        MsgBox "Please select a Pressure from the dropdown (cell E3).", _
               vbExclamation, "No Pressure Selected"
        Exit Sub
    End If

    ' ── Find last row with data ──────────────────────────────────────────────
    Const DATA_START As Long = 27
    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, "A").End(xlUp).Row

    If lastRow < DATA_START Then
        MsgBox "No data found in the formula table (rows 27 onwards).", _
               vbExclamation, "No Data"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ── Next empty row in each MOD tab (skip header + info row) ─────────────
    Dim nFlow As Long
    Dim nPres As Long
    nFlow = wsModFlow.Cells(wsModFlow.Rows.Count, "A").End(xlUp).Row + 1
    nPres = wsModPres.Cells(wsModPres.Rows.Count, "A").End(xlUp).Row + 1
    If nFlow < 3 Then nFlow = 3
    If nPres < 3 Then nPres = 3

    Dim saved As Long
    Dim i As Long

    For i = DATA_START To lastRow

        Dim dtVal     As Variant
        Dim flowAdj   As Variant
        Dim presAdj   As Variant

        dtVal   = wsDash.Cells(i, "A").Value   ' Date
        flowAdj = wsDash.Cells(i, "C").Value   ' Flow Adjusted   (col C)
        presAdj = wsDash.Cells(i, "E").Value   ' Pressure 1 Adjusted (col E)

        If dtVal = "" Then GoTo Skip

        If flowAdj <> "" Then
            wsModFlow.Cells(nFlow, "A").Value         = dtVal
            wsModFlow.Cells(nFlow, "A").NumberFormat  = "DD/MM/YYYY HH:MM"
            wsModFlow.Cells(nFlow, "B").Value         = flowName
            wsModFlow.Cells(nFlow, "C").Value         = flowAdj
            wsModFlow.Cells(nFlow, "C").NumberFormat  = "0.000"
            nFlow = nFlow + 1
        End If

        If presAdj <> "" Then
            wsModPres.Cells(nPres, "A").Value         = dtVal
            wsModPres.Cells(nPres, "A").NumberFormat  = "DD/MM/YYYY HH:MM"
            wsModPres.Cells(nPres, "B").Value         = presName
            wsModPres.Cells(nPres, "C").Value         = presAdj
            wsModPres.Cells(nPres, "C").NumberFormat  = "0.000"
            nPres = nPres + 1
        End If

        saved = saved + 1
Skip:
    Next i

    Application.ScreenUpdating = True

    If saved = 0 Then
        MsgBox "No data rows were saved — check the formula table has values.", _
               vbExclamation, "Nothing Saved"
    Else
        MsgBox "Saved " & saved & " rows." & Chr(10) & Chr(10) & _
               "  Flow:     " & flowName & Chr(10) & _
               "  Pressure: " & presName, _
               vbInformation, "Save Complete"
    End If

End Sub"""

    body(r, vba, sz=9, bg=LIGHT_GRAY, height=17)
    r += 1
    blank(r); r += 1

    # ── 5. Data format reference ──────────────────────────────────────────────
    section(r, "5.  DATA FORMAT REFERENCE", bg=DARK_GRAY)
    r += 1
    for line in [
        "Both Raw Data sheets expect this wide format:",
        "",
        "    Date               | AL012       | AL013       | AL014       | ...  ",
        "    12/01/2026 00:00   | 3.168205    | 2.204250    | 2.665153    | ...  ",
        "    12/01/2026 00:15   | 3.190769    | 2.225250    | 2.681334    | ...  ",
        "",
        "• Column A must contain a proper Date/Time value (not text).",
        "• Flow and pressure column names can be any mix of letters and numbers.",
        "• The flow names in Raw Flow Data and the pressure names in Raw Pressure Data",
        "  do NOT need to match — you select each independently on the Dashboard.",
        "• Use -999 for missing/no-data values — they are excluded from all calculations.",
        "• Data can be any length: months of 15-minute data = thousands of rows.",
        "• Columns F/G = Pressure 2 Raw/Adjusted  (driven by E4),",
        "  Columns H/I = Pressure 3 Raw/Adjusted  (driven by E5).",
        "• Leave E4 or E5 blank to leave those columns empty (no performance impact).",
    ]:
        body(r, line, indent=2)
        r += 1


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    wb = openpyxl.Workbook()

    ws_flow  = wb.active
    ws_pres  = wb.create_sheet()
    ws_dash  = wb.create_sheet()
    ws_mod_f = wb.create_sheet()
    ws_mod_p = wb.create_sheet()
    ws_instr = wb.create_sheet()

    build_raw_sheet(ws_flow, "Raw Flow Data",     "FlowData",     FLOW_ROWS, SAMPLE_DATES)
    build_raw_sheet(ws_pres, "Raw Pressure Data", "PressureData", PRES_ROWS, SAMPLE_DATES)
    build_dashboard(ws_dash, FLOW_NAMES)
    build_mod_sheet(ws_mod_f, "MOD Flow")
    build_mod_sheet(ws_mod_p, "MOD Pressure")
    build_instructions(ws_instr)

    wb.calculation.calcMode    = "auto"
    wb.calculation.fullCalcOnLoad = True

    out = "Flow_Pressure_Dashboard.xlsx"
    wb.save(out)
    _patch_chart_xml(out)
    print(f"Generated: {out}")


if __name__ == "__main__":
    main()
