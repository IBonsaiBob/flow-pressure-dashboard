#!/usr/bin/env python3
"""
generate_dashboard.py

Generates Flow_Pressure_Dashboard.xlsx — a Power Query & PivotTable-ready
Excel workbook for flow/pressure analysis.

Usage:
    python3 generate_dashboard.py

Outputs:
    Flow_Pressure_Dashboard.xlsx
"""

import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ── Colour palette ─────────────────────────────────────────────────────────────
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
YELLOW_DARK  = "BF8F00"
RED_DARK     = "C00000"
PURPLE       = "7030A0"

# ── Sample data (from the problem statement) ───────────────────────────────────
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

# Pressure sample — slightly offset from flow values
PRES_ROWS = [
    [round(v * 0.97 + 0.3, 6) if v != -999 else -999 for v in row]
    for row in FLOW_ROWS
]


# ── Style helpers ──────────────────────────────────────────────────────────────

def _thin():
    t = Side(style="thin")
    return Border(left=t, right=t, top=t, bottom=t)


def _medium():
    m = Side(style="medium")
    return Border(left=m, right=m, top=m, bottom=m)


def style_header(cell, text, bg=DARK_BLUE, fg=WHITE, bold=True, sz=10,
                 halign="center", border=True):
    cell.value = text
    cell.fill = PatternFill(fill_type="solid", fgColor=bg)
    cell.font = Font(bold=bold, color=fg, size=sz)
    cell.alignment = Alignment(horizontal=halign, vertical="center",
                               wrap_text=False)
    if border:
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


# ── Sheet builders ─────────────────────────────────────────────────────────────

def build_raw_sheet(ws, title, table_name, data_rows, dates):
    """Populate a raw data sheet with sample data and a named Excel Table."""
    ws.title = title

    # ── Headers ──
    ws.row_dimensions[1].height = 22
    ws.column_dimensions["A"].width = 21

    style_header(ws.cell(1, 1), "Date", bg=DARK_BLUE)
    for ci, name in enumerate(FLOW_NAMES, start=2):
        style_header(ws.cell(1, ci), name, bg=MID_BLUE)
        ws.column_dimensions[get_column_letter(ci)].width = 13

    # ── Data ──
    for ri, (dt, row) in enumerate(zip(dates, data_rows), start=2):
        c = ws.cell(ri, 1, value=dt)
        c.number_format = "DD/MM/YYYY HH:MM"
        c.alignment = Alignment(horizontal="center", vertical="center")
        for ci, val in enumerate(row, start=2):
            dc = ws.cell(ri, ci, value=val)
            dc.number_format = "0.000000"
            dc.alignment = Alignment(horizontal="right")

    # ── Named Table ──
    last_col = get_column_letter(len(FLOW_NAMES) + 1)
    last_row = len(data_rows) + 1
    tbl = Table(displayName=table_name, ref=f"A1:{last_col}{last_row}")
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True, showColumnStripes=False,
        showFirstColumn=False, showLastColumn=False,
    )
    ws.add_table(tbl)

    # ── Paste-zone note ──
    note_row = last_row + 2
    nc = ws.cell(note_row, 1,
                 value=(
                     "INSTRUCTIONS: Delete the sample rows above (keep Row 1 headers), "
                     "then paste your data starting at Row 2. "
                     "Keep Date in Column A; flow/pressure names as column headers. "
                     "Values of -999 are treated as no-data."
                 ))
    nc.font = Font(italic=True, color=DARK_GRAY, size=9)
    nc.alignment = Alignment(wrap_text=True)
    ws.merge_cells(f"A{note_row}:{last_col}{note_row}")
    ws.row_dimensions[note_row].height = 32

    # ── Power Query note ──
    pq_row = note_row + 1
    pc = ws.cell(pq_row, 1,
                 value=(
                     "POWER QUERY TIP: This sheet is already set up as an Excel Table "
                     f'("{table_name}"). To connect Power Query: '
                     "Data → Get Data → From Table/Range → select this table. "
                     "See the Instructions sheet for the full walkthrough."
                 ))
    pc.font = Font(italic=True, color=MID_BLUE, size=9)
    pc.alignment = Alignment(wrap_text=True)
    ws.merge_cells(f"A{pq_row}:{last_col}{pq_row}")
    ws.row_dimensions[pq_row].height = 32


def build_dashboard(ws, flow_names):
    """Build the main Dashboard sheet."""
    ws.title = "Dashboard"

    # Column widths
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 2   # spacer before chart

    # ── Row 1: Main title ──
    ws.row_dimensions[1].height = 30
    tc = ws.cell(1, 1, value="Flow & Pressure Analysis Dashboard")
    tc.font = Font(bold=True, color=WHITE, size=15)
    tc.fill = PatternFill(fill_type="solid", fgColor=DARK_BLUE)
    tc.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("A1:F1")

    # ── Row 2: section labels ──
    ws.row_dimensions[2].height = 18
    cc = ws.cell(2, 1, value="CONTROLS")
    cc.font = Font(bold=True, color=WHITE, size=10)
    cc.fill = PatternFill(fill_type="solid", fgColor=MID_BLUE)
    cc.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A2:B2")

    pc = ws.cell(2, 3, value="PARAMETERS  (used by Power Query)")
    pc.font = Font(bold=True, color=WHITE, size=10)
    pc.fill = PatternFill(fill_type="solid", fgColor=DARK_ORANGE)
    pc.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("C2:F2")

    # ── Row 3: Flow selector ──
    ws.row_dimensions[3].height = 24
    style_label(ws.cell(3, 1), "Select Flow:", bold=True, sz=10)

    flow_cell = ws.cell(3, 2)
    style_input(flow_cell, flow_names[0], bg=LIGHT_BLUE, sz=11)

    # Data Validation dropdown — references the Raw Flow Data header row
    # This covers up to 200 flow columns; blank entries show at bottom
    dv = DataValidation(
        type="list",
        formula1="'Raw Flow Data'!$B$1:$GR$1",
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Invalid Flow",
        error="Please select a flow name from the list.",
    )
    ws.add_data_validation(dv)
    dv.add(flow_cell)

    # ── Row 3 right: Scaling Factor parameter ──
    style_label(ws.cell(3, 3), "Flow Scaling Factor:", bold=True, sz=10)
    sf_cell = ws.cell(3, 4)
    style_input(sf_cell, 1.0, bg=LIGHT_ORANGE, sz=11, num_fmt="0.000")
    style_label(ws.cell(3, 5), "Multiply flow by this value", italic=True,
                fg=DARK_GRAY, sz=9)

    # Name B4 area for documentation (just label it)
    ws.cell(3, 6, value="← Change me").font = Font(italic=True,
                                                    color=DARK_ORANGE, sz=9,
                                                    size=9)

    # ── Row 4: Pressure Offset ──
    ws.row_dimensions[4].height = 24
    style_label(ws.cell(4, 1), "ℹ  -999 = no-data (excluded)",
                italic=True, fg=DARK_GRAY, sz=9)

    style_label(ws.cell(4, 3), "Pressure Offset:", bold=True, sz=10)
    po_cell = ws.cell(4, 4)
    style_input(po_cell, 0.0, bg=LIGHT_GREEN, sz=11, num_fmt="0.000")
    style_label(ws.cell(4, 5), "Add this value to pressure", italic=True,
                fg=DARK_GRAY, sz=9)

    # Give cells named-range style comments for Power Query
    ws.cell(4, 6, value="← Change me").font = Font(italic=True,
                                                    color=GREEN_MID,
                                                    size=9)

    # Named ranges for Power Query to reference
    # D3 = Scaling Factor, D4 = Pressure Offset
    # (documented in Instructions sheet)

    # ── Row 5: tip bar ──
    ws.row_dimensions[5].height = 18
    tip = ws.cell(5, 1,
                  value=(
                      "💡  For large datasets use Power Query + PivotTable — "
                      "see the Instructions sheet for a full step-by-step guide."
                  ))
    tip.font = Font(italic=True, color=MID_BLUE, size=9)
    ws.merge_cells("A5:F5")

    # ── Row 6: tip 2 ──
    ws.row_dimensions[6].height = 18
    tip2 = ws.cell(6, 1,
                   value=(
                       "⚠  After pasting your own data: update the flow dropdown — "
                       "right-click cell B3 → Data Validation → update the Source range "
                       "to match your actual column headers."
                   ))
    tip2.font = Font(italic=True, color=DARK_ORANGE, size=9)
    ws.merge_cells("A6:F6")

    # ── Row 7: spacer ──
    ws.row_dimensions[7].height = 6

    # ── Row 8: Parameters table label (for PQ) ──
    ws.row_dimensions[8].height = 18
    ph = ws.cell(8, 3, value="PARAMETER TABLE  (Power Query reads D9:D10)")
    ph.font = Font(bold=True, color=WHITE, size=9)
    ph.fill = PatternFill(fill_type="solid", fgColor=DARK_ORANGE)
    ph.alignment = Alignment(horizontal="center")
    ws.merge_cells("C8:F8")

    # ── Row 9-10: Named Parameter Table ──
    ws.row_dimensions[9].height = 20
    ws.row_dimensions[10].height = 20
    for col, header, bg in [(3, "Parameter", DARK_ORANGE),
                             (4, "Value",     DARK_ORANGE),
                             (5, "Notes",     DARK_ORANGE)]:
        style_header(ws.cell(9, col), header, bg=bg)

    # Row 10: Scaling Factor param
    ws.cell(10, 3, value="Flow Scaling Factor").font = Font(sz=10)
    ws.cell(10, 3).border = _thin()
    sf_link = ws.cell(10, 4)
    sf_link.value = "=D3"
    sf_link.number_format = "0.000"
    sf_link.border = _thin()
    sf_link.alignment = Alignment(horizontal="center")
    ws.cell(10, 5, value="Linked to control above (D3)").font = Font(
        italic=True, color=DARK_GRAY, size=9)
    ws.cell(10, 5).border = _thin()

    # Row 11: Pressure Offset param
    ws.row_dimensions[11].height = 20
    ws.cell(11, 3, value="Pressure Offset").font = Font(sz=10)
    ws.cell(11, 3).border = _thin()
    po_link = ws.cell(11, 4)
    po_link.value = "=D4"
    po_link.number_format = "0.000"
    po_link.border = _thin()
    po_link.alignment = Alignment(horizontal="center")
    ws.cell(11, 5, value="Linked to control above (D4)").font = Font(
        italic=True, color=DARK_GRAY, size=9)
    ws.cell(11, 5).border = _thin()

    # Parameter Table (named "Parameters" for Power Query)
    param_tbl = Table(displayName="Parameters", ref="C9:E11")
    param_tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight11", showRowStripes=True)
    ws.add_table(param_tbl)

    # ── Row 12: spacer ──
    ws.row_dimensions[12].height = 6

    # ── Row 13: Data table section header ──
    ws.row_dimensions[13].height = 18
    dh = ws.cell(13, 1,
                 value="FORMULA TABLE  (live preview — updates when you change B3 / D3 / D4)")
    dh.font = Font(bold=True, color=WHITE, size=9)
    dh.fill = PatternFill(fill_type="solid", fgColor=DARK_BLUE)
    dh.alignment = Alignment(horizontal="left")
    ws.merge_cells("A13:F13")

    # ── Row 14: Column headers ──
    ws.row_dimensions[14].height = 22
    headers = [
        ("Date",               DARK_BLUE),
        ("Flow (Raw)",         MID_BLUE),
        ("Pressure (Raw)",     MID_BLUE),
        ("Flow Adjusted",      DARK_ORANGE),
        ("Pressure Adjusted",  GREEN_DARK),
    ]
    for ci, (hdr_text, bg) in enumerate(headers, start=1):
        style_header(ws.cell(14, ci), hdr_text, bg=bg)

    # ── Rows 15 – 114: Formula rows (covers 100 data points) ──
    # ROW()-14 maps row 15 → raw data row 2 (first data row; row 1 = headers)
    # General pattern: INDEX(sheet!$A:$ZZ, ROW()-13, col_match)
    # ROW()-13 = 2 at row 15  ✓
    OFFSET = 13   # ROW() - OFFSET = raw sheet row index (2 = first data row)
    DATA_ROWS = 100

    for r in range(15, 15 + DATA_ROWS):
        ws.row_dimensions[r].height = 15
        raw_row = r - OFFSET   # e.g. 15-13=2 → raw sheet row 2

        # ── Col A: Date ──
        dc = ws.cell(r, 1)
        dc.value = (
            f"=IFERROR("
            f"IF(INDEX('Raw Flow Data'!$A:$A,ROW()-{OFFSET})=\"\",\"\","
            f"INDEX('Raw Flow Data'!$A:$A,ROW()-{OFFSET})),\"\")"
        )
        dc.number_format = "DD/MM/YYYY HH:MM"
        dc.alignment = Alignment(horizontal="center")
        if r % 2 == 0:
            dc.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # ── Col B: Flow Raw ──
        bc = ws.cell(r, 2)
        bc.value = (
            f"=IFERROR("
            f"IF($B$3=\"\",\"\","
            f"INDEX('Raw Flow Data'!$A:$ZZ,ROW()-{OFFSET},"
            f"MATCH($B$3,'Raw Flow Data'!$1:$1,0))),\"\")"
        )
        bc.number_format = "0.000"
        bc.alignment = Alignment(horizontal="right")
        if r % 2 == 0:
            bc.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # ── Col C: Pressure Raw ──
        cc = ws.cell(r, 3)
        cc.value = (
            f"=IFERROR("
            f"IF($B$3=\"\",\"\","
            f"INDEX('Raw Pressure Data'!$A:$ZZ,ROW()-{OFFSET},"
            f"MATCH($B$3,'Raw Pressure Data'!$1:$1,0))),\"\")"
        )
        cc.number_format = "0.000"
        cc.alignment = Alignment(horizontal="right")
        if r % 2 == 0:
            cc.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)

        # ── Col D: Flow Adjusted  = Flow × Scaling Factor ──
        ec = ws.cell(r, 4)
        ec.value = (
            f"=IF(OR(B{r}=\"\",B{r}=-999),\"\",B{r}*$D$3)"
        )
        ec.number_format = "0.000"
        ec.alignment = Alignment(horizontal="right")
        ec.fill = PatternFill(fill_type="solid",
                              fgColor=LIGHT_ORANGE if r % 2 == 0 else YELLOW_BG)

        # ── Col E: Pressure Adjusted = Pressure + Offset ──
        fc = ws.cell(r, 5)
        fc.value = (
            f"=IF(OR(C{r}=\"\",C{r}=-999),\"\",C{r}+$D$4)"
        )
        fc.number_format = "0.000"
        fc.alignment = Alignment(horizontal="right")
        fc.fill = PatternFill(fill_type="solid",
                              fgColor=LIGHT_GREEN if r % 2 == 0 else "D6E4BC")

    # ── Row after table: copy-down hint ──
    hint_row = 15 + DATA_ROWS
    ws.row_dimensions[hint_row].height = 18
    hc = ws.cell(hint_row, 1,
                 value=(
                     f"↑ Formula rows cover {DATA_ROWS} data points. "
                     "For more rows: select A15:E114, copy, paste below. "
                     "For very large datasets use Power Query (see Instructions)."
                 ))
    hc.font = Font(italic=True, color=DARK_GRAY, size=9)
    hc.alignment = Alignment(wrap_text=True)
    ws.merge_cells(f"A{hint_row}:F{hint_row}")
    ws.row_dimensions[hint_row].height = 28

    # ── Chart (anchored at H1, dual-axis line chart) ──
    _add_chart(ws, data_start_row=14, data_end_row=14 + DATA_ROWS)


def _add_chart(ws, data_start_row, data_end_row):
    """Add a dual-axis line chart to the Dashboard, anchored at column H."""
    # Primary chart — Flow Adjusted (col D)
    c1 = LineChart()
    c1.title = "Flow & Pressure Analysis"
    c1.y_axis.title = "Flow Adjusted"
    c1.y_axis.axId = 100
    c1.x_axis.axId = 100
    c1.style = 10
    c1.height = 14
    c1.width = 22

    flow_ref = Reference(ws, min_col=4, max_col=4,
                         min_row=data_start_row, max_row=data_end_row)
    c1.add_data(flow_ref, titles_from_data=True)

    dates_ref = Reference(ws, min_col=1, max_col=1,
                          min_row=data_start_row + 1, max_row=data_end_row)
    c1.set_categories(dates_ref)

    # Secondary chart — Pressure Adjusted (col E)
    c2 = LineChart()
    c2.y_axis.title = "Pressure Adjusted"
    c2.y_axis.axId = 200
    c2.y_axis.crosses = "max"   # right-hand axis
    c2.x_axis.axId = 100
    c2.x_axis.delete = True     # suppress duplicate x-axis

    pres_ref = Reference(ws, min_col=5, max_col=5,
                         min_row=data_start_row, max_row=data_end_row)
    c2.add_data(pres_ref, titles_from_data=True)
    c2.set_categories(dates_ref)

    # Merge into dual-axis chart
    c1 += c2

    # Series colours
    try:
        c1.series[0].graphicalProperties.line.solidFill = MID_BLUE
        c1.series[0].graphicalProperties.line.width = 18000   # 1.8 pt
        c1.series[1].graphicalProperties.line.solidFill = DARK_ORANGE
        c1.series[1].graphicalProperties.line.width = 18000
    except Exception:
        pass  # cosmetic only — chart still works without colour overrides

    c1.anchor = "H1"
    ws.add_chart(c1)


def build_mod_sheet(ws, title):
    """Build a MOD output sheet (flow or pressure)."""
    ws.title = title

    ws.row_dimensions[1].height = 22
    ws.column_dimensions["A"].width = 21
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 20

    headers = ["Date", "Flow Name", "Adjusted Value"]
    bg_colors = [DARK_BLUE, DARK_BLUE, DARK_BLUE]
    for ci, (h, bg) in enumerate(zip(headers, bg_colors), start=1):
        style_header(ws.cell(1, ci), h, bg=bg)

    # Instruction row
    ws.row_dimensions[2].height = 18
    ic = ws.cell(2, 1,
                 value=(
                     "Data is appended here each time you run the SaveToMOD macro "
                     "(history is preserved — rows are never overwritten)."
                 ))
    ic.font = Font(italic=True, color=DARK_GRAY, size=9)
    ic.alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:C2")

    # Date format for column A
    ws.column_dimensions["A"].width = 21


def build_instructions(ws):
    """Build the Instructions sheet with Power Query, PivotTable, and VBA guides."""
    ws.title = "Instructions"
    ws.column_dimensions["A"].width = 110
    ws.sheet_view.showGridLines = False

    def section(row, text, bg=DARK_BLUE, fg=WHITE, sz=12, height=26):
        ws.row_dimensions[row].height = height
        c = ws.cell(row, 1, value=text)
        c.font = Font(bold=True, color=fg, size=sz)
        c.fill = PatternFill(fill_type="solid", fgColor=bg)
        c.alignment = Alignment(horizontal="left", vertical="center",
                                indent=1)

    def body(row, text, fg="000000", sz=10, indent=0, bold=False,
             italic=False, bg=None, height=16):
        ws.row_dimensions[row].height = height
        c = ws.cell(row, 1, value=text)
        c.font = Font(color=fg, size=sz, bold=bold, italic=italic)
        c.alignment = Alignment(horizontal="left", vertical="center",
                                indent=indent, wrap_text=True)
        if bg:
            c.fill = PatternFill(fill_type="solid", fgColor=bg)

    def blank(row, height=8):
        ws.row_dimensions[row].height = height

    r = 1

    # ── Title ──
    section(r, "Flow & Pressure Dashboard — Instructions & Setup Guide",
            bg=DARK_BLUE, sz=14, height=32)
    r += 1
    blank(r); r += 1

    # ── Section 1: Quick Start ──
    section(r, "1.  QUICK START  (Formula-Based — works immediately, no setup needed)",
            bg=MID_BLUE, sz=11)
    r += 1
    for line in [
        "Step 1:  Go to the 'Raw Flow Data' sheet.  Delete the sample rows (keep Row 1 headers).  Paste your flow data from Row 2 onwards.",
        "Step 2:  Go to the 'Raw Pressure Data' sheet and do the same with your pressure data.",
        "Step 3:  Go to the 'Dashboard' sheet.  Click cell B3 and pick a flow from the dropdown.",
        "Step 4:  Adjust 'Flow Scaling Factor' (cell D3) and 'Pressure Offset' (cell D4) as needed.",
        "Step 5:  The formula table (rows 15 onwards) and chart update automatically.",
        "Note:    The formula table covers 100 rows. For more data rows, select A15:E114 and copy-paste downward.",
        "Note:    After pasting your own flow names, right-click cell B3 → Data Validation → update the",
        "         Source range to match your column headers (e.g. 'Raw Flow Data'!$B$1:$BZ$1).",
    ]:
        body(r, line, indent=2, height=17)
        r += 1
    blank(r); r += 1

    # ── Section 2: Power Query ──
    section(r, "2.  POWER QUERY SETUP  (recommended for large datasets)",
            bg=DARK_ORANGE, sz=11)
    r += 1
    body(r, "Power Query loads and transforms your raw data so PivotTables and charts always reflect the latest paste.", indent=2, height=18)
    r += 1
    blank(r); r += 1

    body(r, "2a.  Load the Flow data query", bold=True, sz=10, indent=2)
    r += 1
    for line in [
        "1. Go to the 'Raw Flow Data' sheet and click anywhere inside the table.",
        "2. Data tab → Get Data → From Table/Range.  Power Query Editor opens.",
        "3. In Power Query Editor:  Home → Unpivot Other Columns  (select the Date column first, then Unpivot).",
        "   This converts wide format (one column per flow) into long format: Date | Flow Name | Value.",
        "4. Rename the columns:  'Attribute' → 'Flow Name',  'Value' → 'Flow Value'.",
        "5. Home → Close & Load To…  →  'Only Create Connection'  (tick 'Add to Data Model' if you like).",
        "6. Name the query  FlowLong.",
    ]:
        body(r, line, indent=4, height=17)
        r += 1
    blank(r); r += 1

    body(r, "2b.  Load the Pressure data query  (same steps)", bold=True, sz=10, indent=2)
    r += 1
    for line in [
        "1. Go to the 'Raw Pressure Data' sheet and repeat steps 1-5 above.",
        "2. Rename columns:  'Attribute' → 'Flow Name',  'Value' → 'Pressure Value'.",
        "3. Name the query  PressureLong.",
    ]:
        body(r, line, indent=4, height=17)
        r += 1
    blank(r); r += 1

    body(r, "2c.  Load the Parameters table  (scaling factor and offset)", bold=True, sz=10, indent=2)
    r += 1
    for line in [
        "1. Go to the 'Dashboard' sheet and click anywhere inside the orange Parameters table (C9:E11).",
        "2. Data tab → Get Data → From Table/Range.",
        "3. No transformations needed — just click  Close & Load To… → Only Create Connection.",
        "4. Name the query  Parameters.",
    ]:
        body(r, line, indent=4, height=17)
        r += 1
    blank(r); r += 1

    body(r, "2d.  Merge queries into one combined table", bold=True, sz=10, indent=2)
    r += 1
    for line in [
        "1. Data tab → Get Data → Combine Queries → Merge.",
        "2. Select FlowLong as the left table; PressureLong as the right table.",
        "3. Match on:  Date  AND  Flow Name  (hold Ctrl to select both columns).",
        "4. Expand the merged column to include 'Pressure Value'.",
        "5. Add a custom column for Flow Adjusted:",
        '   = if [Flow Value] = -999 then null else [Flow Value] * Parameters{[Parameter="Flow Scaling Factor"]}[Value]',
        "6. Add a custom column for Pressure Adjusted:",
        '   = if [Pressure Value] = -999 then null else [Pressure Value] + Parameters{[Parameter="Pressure Offset"]}[Value]',
        "7. Home → Close & Load To… → Table  (load to a new sheet, e.g. 'PQ Output').",
        "8. Name the query  CombinedAdjusted.",
        "",
        "To refresh after pasting new data:  Data tab → Refresh All  (or Ctrl+Alt+F5).",
    ]:
        body(r, line, indent=4, height=17)
        r += 1
    blank(r); r += 1

    # ── Section 3: PivotTable + Slicer ──
    section(r, "3.  PIVOTTABLE + SLICER SETUP  (interactive filtering)",
            bg=GREEN_DARK, sz=11)
    r += 1
    body(r, "Once you have the CombinedAdjusted Power Query output loaded to a sheet, follow these steps:", indent=2, height=18)
    r += 1
    blank(r); r += 1

    body(r, "3a.  Create the PivotTable", bold=True, sz=10, indent=2)
    r += 1
    for line in [
        "1. Click anywhere in the CombinedAdjusted table on the 'PQ Output' sheet.",
        "2. Insert → PivotTable → New Worksheet (or Existing Worksheet).",
        "3. In the PivotTable Field List:",
        "   • Drag 'Date' to ROWS.",
        "   • Drag 'Flow Adjusted' to VALUES  (set to Sum or Average).",
        "   • Drag 'Pressure Adjusted' to VALUES  (set to Sum or Average).",
        "4. Right-click a date cell → Group → by Hours or Minutes to control time grouping.",
    ]:
        body(r, line, indent=4, height=17)
        r += 1
    blank(r); r += 1

    body(r, "3b.  Add a Flow Name Slicer", bold=True, sz=10, indent=2)
    r += 1
    for line in [
        "1. Click anywhere in the PivotTable.",
        "2. PivotTable Analyze tab → Insert Slicer.",
        "3. Tick 'Flow Name' → OK.",
        "4. Click a flow name in the Slicer to filter the PivotTable (and chart) to that flow.",
        "   Hold Ctrl to select multiple flows.",
    ]:
        body(r, line, indent=4, height=17)
        r += 1
    blank(r); r += 1

    body(r, "3c.  Create a PivotChart", bold=True, sz=10, indent=2)
    r += 1
    for line in [
        "1. Click anywhere in the PivotTable.",
        "2. PivotTable Analyze tab → PivotChart → Line → Line with Markers.",
        "3. To add a secondary axis for Pressure:",
        "   Right-click the Pressure Adjusted series → Format Data Series → Secondary Axis.",
        "4. The chart updates automatically when you click a flow in the Slicer.",
    ]:
        body(r, line, indent=4, height=17)
        r += 1
    blank(r); r += 1

    # ── Section 4: VBA Save Button ──
    section(r, "4.  VBA SAVE BUTTON CODE  (paste this yourself)",
            bg=PURPLE, sz=11)
    r += 1
    body(r,
         "Press Alt+F11 to open the VBA editor.  Insert → Module, then paste the code below.",
         indent=2, height=18)
    r += 1
    blank(r); r += 1

    vba_code = r"""' ─────────────────────────────────────────────────────────────────────────────
' SaveToMOD  —  copies the current adjusted table on Dashboard to the MOD tabs
' Assign this macro to a button on the Dashboard sheet.
' ─────────────────────────────────────────────────────────────────────────────
Sub SaveToMOD()

    Dim wsDash      As Worksheet
    Dim wsModFlow   As Worksheet
    Dim wsModPres   As Worksheet

    Set wsDash    = Worksheets("Dashboard")
    Set wsModFlow = Worksheets("MOD Flow")
    Set wsModPres = Worksheets("MOD Pressure")

    ' ── Validate ──────────────────────────────────────────────────────────────
    Dim flowName As String
    flowName = Trim(wsDash.Range("B3").Value)

    If flowName = "" Then
        MsgBox "Please select a flow from the dropdown (cell B3) before saving.", _
               vbExclamation, "No Flow Selected"
        Exit Sub
    End If

    ' ── Find the last data row on Dashboard ──────────────────────────────────
    Const DATA_START As Long = 15        ' first formula row
    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, "A").End(xlUp).Row

    If lastRow < DATA_START Then
        MsgBox "No data found in the formula table (rows 15 onwards).", _
               vbExclamation, "No Data"
        Exit Sub
    End If

    ' ── Find next empty row in each MOD tab ──────────────────────────────────
    Dim nextFlow As Long
    Dim nextPres As Long
    nextFlow = wsModFlow.Cells(wsModFlow.Rows.Count, "A").End(xlUp).Row + 1
    nextPres = wsModPres.Cells(wsModPres.Rows.Count, "A").End(xlUp).Row + 1
    If nextFlow < 3 Then nextFlow = 3    ' skip header + info rows
    If nextPres < 3 Then nextPres = 3

    ' ── Disable screen updates for speed ─────────────────────────────────────
    Application.ScreenUpdating = False

    Dim i           As Long
    Dim savedRows   As Long
    savedRows = 0

    For i = DATA_START To lastRow

        Dim dateVal     As Variant
        Dim flowAdj     As Variant
        Dim presAdj     As Variant

        dateVal = wsDash.Cells(i, "A").Value
        flowAdj = wsDash.Cells(i, "D").Value   ' Flow Adjusted
        presAdj = wsDash.Cells(i, "E").Value   ' Pressure Adjusted

        ' Skip rows where date or both values are empty
        If dateVal = "" Or (flowAdj = "" And presAdj = "") Then GoTo NextRow

        ' Write to MOD Flow
        If flowAdj <> "" Then
            wsModFlow.Cells(nextFlow, "A").Value = dateVal
            wsModFlow.Cells(nextFlow, "A").NumberFormat = "DD/MM/YYYY HH:MM"
            wsModFlow.Cells(nextFlow, "B").Value = flowName
            wsModFlow.Cells(nextFlow, "C").Value = flowAdj
            wsModFlow.Cells(nextFlow, "C").NumberFormat = "0.000"
            nextFlow = nextFlow + 1
        End If

        ' Write to MOD Pressure
        If presAdj <> "" Then
            wsModPres.Cells(nextPres, "A").Value = dateVal
            wsModPres.Cells(nextPres, "A").NumberFormat = "DD/MM/YYYY HH:MM"
            wsModPres.Cells(nextPres, "B").Value = flowName
            wsModPres.Cells(nextPres, "C").Value = presAdj
            wsModPres.Cells(nextPres, "C").NumberFormat = "0.000"
            nextPres = nextPres + 1
        End If

        savedRows = savedRows + 1
NextRow:
    Next i

    Application.ScreenUpdating = True

    If savedRows = 0 Then
        MsgBox "No data rows were saved. Make sure the formula table has data.", _
               vbExclamation, "Nothing Saved"
    Else
        MsgBox "Saved " & savedRows & " rows for flow: " & flowName & Chr(10) & _
               "  → MOD Flow       (" & (nextFlow - 3) & " rows total)" & Chr(10) & _
               "  → MOD Pressure   (" & (nextPres - 3) & " rows total)", _
               vbInformation, "Save Complete"
    End If

End Sub"""

    body(r, vba_code, sz=9, bg=LIGHT_GRAY, height=18)
    r += 1
    blank(r); r += 1

    body(r,
         "To add a button:  Developer tab → Insert → Button (Form Control) → draw on Dashboard → assign macro SaveToMOD.",
         indent=2, height=18, italic=True, fg=DARK_GRAY)
    r += 1
    body(r,
         "If the Developer tab is not visible:  File → Options → Customise Ribbon → tick Developer.",
         indent=2, height=18, italic=True, fg=DARK_GRAY)
    r += 1
    blank(r); r += 1

    # ── Section 5: Data format reference ──
    section(r, "5.  DATA FORMAT REFERENCE", bg=DARK_GRAY, sz=11)
    r += 1
    for line in [
        "Both Raw Data sheets expect this wide format:",
        "",
        "    Date               | AL012       | AL013       | AL014       | ...  ",
        "    12/01/2026 00:00   | 3.168205    | 2.204250    | 2.665153    | ...  ",
        "    12/01/2026 00:15   | 3.190769    | 2.225250    | 2.681334    | ...  ",
        "",
        "• Column A must be a Date/Time value (not text).",
        "• Flow/pressure names can be any combination of letters and numbers (AL012, AM005, AF037, etc.).",
        "• Use -999 to represent missing / no-data values — they are excluded from all calculations.",
        "• Data can be any length (months of 15-minute intervals = thousands of rows).",
        "• Both sheets must use the same flow/pressure names (column headers must match).",
    ]:
        body(r, line, indent=2, sz=10, height=17)
        r += 1


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    wb = openpyxl.Workbook()

    # Sheet order
    ws_flow = wb.active
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

    # Ensure formulas recalculate on open
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    out_path = "Flow_Pressure_Dashboard.xlsx"
    wb.save(out_path)
    print(f"Generated: {out_path}")


if __name__ == "__main__":
    main()
