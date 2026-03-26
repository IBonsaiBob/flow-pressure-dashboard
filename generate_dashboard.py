#!/usr/bin/env python3
"""
generate_dashboard.py

Updates 'Model Build Dashboard v1.21.xlsx' by applying:
  - Removal of MOD Flow and MOD Pressure sheets
  - Removal of the "💾 Save Rest" button cell (Dashboard K7)
  - Updated Instructions sheet with the new SaveOneSensor VBA
    (saves adjusted values back into Raw tabs; no longer writes to MOD sheets)

Also writes companion VBA text files:
  VBA_Module1_SaveSensor.txt  — standard module (SaveOneSensor subroutine)
  VBA_Dashboard_Sheet.txt     — Dashboard sheet event handler

Usage:
    pip install openpyxl   # only needed the first time
    python3 generate_dashboard.py
"""

import io
import os
import re
import zipfile

SOURCE = "Model Build Dashboard v1.21.xlsx"

# ---------------------------------------------------------------------------
# VBA source text
# ---------------------------------------------------------------------------

VBA_MODULE = """\
' ===========================================================================
' STANDARD MODULE CODE
' Paste into a new Module (Alt+F11 -> Insert -> Module)
' ===========================================================================
' SaveOneSensor  - applies Scale / Offset / Dt for ONE sensor and writes the
'                  adjusted values back into the corresponding Raw data tab,
'                  overwriting that sensor column in place.
'
' isFlow : True  -> Flow sensor   (Raw Flow Data,     Scale * value,  col D Dt)
'          False -> Pressure      (Raw Pressure Data, value + Offset, col I Dt)
' sRow   : selector row index 1-20
'
' Dashboard selector rows 3-22:
'   Col B (2) = Flow name    Col C (3) = Scale    Col D (4) = Flow Dt
'   Col G (7) = Pres name    Col H (8) = Offset   Col I (9) = Pres Dt
' ===========================================================================
Sub SaveOneSensor(isFlow As Boolean, sRow As Long)

    Const SEL_START  As Long = 3
    Const FLOW_NAME  As Long = 2   ' B
    Const FLOW_SCALE As Long = 3   ' C
    Const FLOW_DT    As Long = 4   ' D
    Const PRES_NAME  As Long = 7   ' G
    Const PRES_OFF   As Long = 8   ' H
    Const PRES_DT    As Long = 9   ' I

    Dim wsDash  As Worksheet
    Dim wsRaw   As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim sensorName  As String
    Dim scaleFactor As Double
    Dim offset      As Double
    Dim dt          As Long
    Dim dashRow     As Long
    dashRow = SEL_START + sRow - 1

    If isFlow Then
        sensorName = Trim(wsDash.Cells(dashRow, FLOW_NAME).Value)
        If sensorName = "" Then
            MsgBox "No sensor selected in Flow row " & sRow, vbExclamation
            Exit Sub
        End If
        scaleFactor = ToDouble(wsDash.Cells(dashRow, FLOW_SCALE).Value)
        If scaleFactor = 0 Then scaleFactor = 1
        dt = ToLong(wsDash.Cells(dashRow, FLOW_DT).Value)
        Set wsRaw = Worksheets("Raw Flow Data")
    Else
        sensorName = Trim(wsDash.Cells(dashRow, PRES_NAME).Value)
        If sensorName = "" Then
            MsgBox "No sensor selected in Pressure row " & sRow, vbExclamation
            Exit Sub
        End If
        offset = ToDouble(wsDash.Cells(dashRow, PRES_OFF).Value)
        dt = ToLong(wsDash.Cells(dashRow, PRES_DT).Value)
        Set wsRaw = Worksheets("Raw Pressure Data")
    End If

    ' --- Find sensor column in raw sheet ---
    Dim lastHdrCol As Long
    lastHdrCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    Dim sensorCol As Long: sensorCol = 0
    Dim k As Long
    For k = 2 To lastHdrCol
        If Trim(wsRaw.Cells(1, k).Value) = sensorName Then
            sensorCol = k
            Exit For
        End If
    Next k
    If sensorCol = 0 Then
        MsgBox "Sensor '" & sensorName & "' not found in " & wsRaw.Name, _
               vbExclamation, "Sensor Not Found"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    ' --- Apply transformation and write back into the Raw tab ---
    Application.ScreenUpdating = False
    Dim totalSaved As Long: totalSaved = 0
    Dim rawVal As Variant
    Dim srcRow As Long
    Dim j As Long

    For j = 2 To lastRow
        srcRow = j - dt
        If srcRow >= 2 And srcRow <= lastRow Then
            rawVal = wsRaw.Cells(srcRow, sensorCol).Value
            If IsNumeric(rawVal) And CDbl(rawVal) <> -999 Then
                If isFlow Then
                    wsRaw.Cells(j, sensorCol).Value = CDbl(rawVal) * scaleFactor
                Else
                    wsRaw.Cells(j, sensorCol).Value = CDbl(rawVal) + offset
                End If
                wsRaw.Cells(j, sensorCol).NumberFormat = "0.000"
                totalSaved = totalSaved + 1
            End If
        End If
    Next j

    Application.ScreenUpdating = True
    MsgBox "Saved " & totalSaved & " values for '" & sensorName & _
           "' into " & wsRaw.Name & ".", vbInformation, "Save Complete"
End Sub

' ===========================================================================
' Safe numeric helpers
' ===========================================================================
Private Function ToDouble(v As Variant) As Double
    If IsNumeric(v) Then ToDouble = CDbl(v)
End Function

Private Function ToLong(v As Variant) As Long
    If IsNumeric(v) Then ToLong = CLng(v)
End Function
"""

VBA_SHEET = """\
' ===========================================================================
' DASHBOARD SHEET MODULE CODE
' Paste into the Dashboard sheet module (double-click Sheet1 in Project tree)
' ===========================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Const FLOW_SAVE_COL As Long = 5    ' E - flow 💾 cells
    Const PRES_SAVE_COL As Long = 10   ' J - pres 💾 cells
    Const SEL_START     As Long = 3    ' first selector row
    Const SEL_END       As Long = 22   ' last selector row

    If Target.Count > 1 Then Exit Sub

    If Target.Column = FLOW_SAVE_COL And _
       Target.Row >= SEL_START And Target.Row <= SEL_END Then
        Application.EnableEvents = False
        Target.Offset(0, -1).Select
        Application.EnableEvents = True
        SaveOneSensor True, Target.Row - SEL_START + 1

    ElseIf Target.Column = PRES_SAVE_COL And _
           Target.Row >= SEL_START And Target.Row <= SEL_END Then
        Application.EnableEvents = False
        Target.Offset(0, -1).Select
        Application.EnableEvents = True
        SaveOneSensor False, Target.Row - SEL_START + 1
    End If
End Sub
"""


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def _escape(text):
    """Escape special XML characters for use in inline string cells."""
    return (
        text
        .replace("&",  "&amp;")
        .replace("<",  "&lt;")
        .replace(">",  "&gt;")
        .replace('"',  "&quot;")
        .replace("'",  "&apos;")
    )


def _cell(ref, style, text):
    """Return an inline-string <c> element."""
    return (
        f'<c r="{ref}" s="{style}" t="inlineStr">'
        f'<is><t xml:space="preserve">{_escape(text)}</t></is>'
        f"</c>"
    )


# ---------------------------------------------------------------------------
# Style indices (from the existing styles.xml — do not change)
# ---------------------------------------------------------------------------
# S_TITLE    = 5   dark blue bg (1F4E79), bold white 14pt
# S_HDR1     = 64  mid blue bg  (2E75B6), bold white 11pt  → Quick Start
# S_HDR2     = 66  orange bg    (C55A11), bold white 11pt  → Power Query
# S_HDR3     = 67  dark green   (375623), bold white 11pt  → PivotTable
# S_HDR4     = 68  purple bg    (7030A0), bold white 11pt  → VBA
# S_HDR5     = 71  dark gray    (595959), bold white 11pt  → Data Format
# S_BODY     = 65  default bg,  black 10pt, wrap, indent=2
# S_CODE     = 69  light gray   (F2F2F2), black 9pt,  wrap
# S_ITALIC   = 70  default bg,  gray  10pt, wrap, indent=2

S_TITLE  = 5
S_HDR1   = 64
S_HDR2   = 66
S_HDR3   = 67
S_HDR4   = 68
S_HDR5   = 71
S_BODY   = 65
S_CODE   = 69
S_ITALIC = 70


# ---------------------------------------------------------------------------
# Build new Instructions sheet XML
# ---------------------------------------------------------------------------

def _build_instructions_xml():
    """
    Return the XML for the updated Instructions worksheet.

    Uses inline strings (t="inlineStr") so no changes to sharedStrings.xml
    are required.  Style indices reference the workbook's existing styles.xml.
    """
    rows = []   # list of (row_number, height, cell_list)

    def sec(r, txt, style, h=26):
        rows.append((r, h, [_cell("A" + str(r), style, txt)]))

    def body(r, txt, h=17):
        rows.append((r, h, [_cell("A" + str(r), S_BODY, txt)]))

    def blank(r, h=8):
        rows.append((r, h, []))

    r = 1
    sec(r, "Flow & Pressure Dashboard — Instructions", S_TITLE, h=32); r += 1
    blank(r); r += 1

    # ── 1. Quick Start ────────────────────────────────────────────────────────
    sec(r, "1.  QUICK START  (works immediately — no setup needed)", S_HDR1); r += 1
    for line in [
        "Step 1:  Paste your flow data into 'Raw Flow Data' (delete sample rows, keep Row 1 headers).",
        "Step 2:  Paste your pressure data into 'Raw Pressure Data' (same format).",
        "Step 3:  Go to the Dashboard sheet.",
        "Step 4:  Rows 3-22 (col B) are the 20 flow selectors — pick a sensor from the dropdown.",
        "         Leave unused rows blank to hide that series.",
        "Step 5:  Rows 3-22 (col G) are the 20 pressure selectors — same idea.",
        "Step 6:  Adjust the Scale (col C) for each flow row and the Offset (col H) for",
        "         each pressure row independently.  Default Scale = 1.000, Offset = 0.000.",
        "Step 7:  The input cell for each series is coloured to match its chart line.",
        "         Flow lines and pressure lines are both solid.",
        "Step 8:  Use the Chart Controls panel (cols K-L, top right of the Dashboard):",
        "         \u2022 Start Date / End Date \u2014 enter dates to filter the formula table and chart.",
        "           Leave blank to show all available data.  Dates must exist in 'Raw Flow Data'.",
        "Step 9:  Each flow row (col D) and each pressure row (col I) has its own \u0394t cell.",
        "         Enter an integer to shift that series in time:",
        "         +2 = read from 2 timesteps later;  -3 = read from 3 timesteps earlier.",
        "         Use this to align sensors with different transit / delay times.",
        "Step 10: Click a \U0001f4be cell (col E for flow, col J for pressure) to save that sensor.",
        "         Scale / Offset / \u0394t are applied and the adjusted values are written directly",
        "         into the corresponding Raw tab, overwriting the original column in place.",
        "         IMPORTANT: Keep a backup of your raw data before clicking Save.",
        "",
        "NOTE:  Up to 20 flow series (left Y-axis, blue/teal shades) and 20 pressure series",
        "       (right Y-axis, warm/cool shades) are shown simultaneously.",
        "",
        "NOTE:  The formula table covers 2000 rows. For longer datasets, select",
        "       that range and copy-paste downward as far as needed.",
        "",
        "NOTE:  -999 values are treated as no-data and are excluded from all calculations.",
    ]:
        body(r, line); r += 1
    blank(r); r += 1

    # ── 2. Power Query ────────────────────────────────────────────────────────
    sec(r, "2.  POWER QUERY SETUP  (optional \u2014 recommended for very large datasets)", S_HDR2); r += 1
    for line in [
        "The Raw data sheets are already set up as Excel Tables (FlowData, PressureData).",
        "Power Query can load these tables, unpivot them, and merge them for use in PivotTables.",
        "",
        "Step 1:  Data tab \u2192 Get Data \u2192 From Table/Range \u2192 select the FlowData table.",
        "Step 2:  In Power Query Editor: select the Date column, then Home \u2192 Unpivot Other Columns.",
        "         Rename 'Attribute' \u2192 'Flow Name',  'Value' \u2192 'Flow Value'.",
        "Step 3:  Close & Load To\u2026 \u2192 Only Create Connection.  Name the query  FlowLong.",
        "Step 4:  Repeat for PressureData.  Name the query  PressureLong.",
        "Step 5:  Merge the two queries on Date + Name to get a combined table.",
        "Step 6:  Add calculated columns:  Flow Adjusted = [Flow Value] \u00d7 scaling_factor",
        "                                  Pressure Adjusted = [Pressure Value] + offset",
        "Step 7:  Load the merged query to a sheet and build a PivotTable + Slicer on top of it.",
        "",
        "After pasting new data:  Data tab \u2192 Refresh All  (Ctrl+Alt+F5).",
    ]:
        body(r, line); r += 1
    blank(r); r += 1

    # ── 3. PivotTable & Slicer ────────────────────────────────────────────────
    sec(r, "3.  PIVOTTABLE + SLICER  (optional \u2014 for interactive multi-flow comparison)", S_HDR3); r += 1
    for line in [
        "Once the Power Query merged table is loaded to a sheet:",
        "  \u2022 Insert \u2192 PivotTable",
        "  \u2022 Rows: Date    Values: Flow Adjusted, Pressure Adjusted",
        "  \u2022 PivotTable Analyze \u2192 Insert Slicer \u2192 tick 'Flow Name' \u2192 OK",
        "  \u2022 Click a flow name in the Slicer to filter instantly",
        "  \u2022 Insert \u2192 PivotChart \u2192 Line \u2192 add Secondary Axis to the Pressure series",
    ]:
        body(r, line); r += 1
    blank(r); r += 1

    # ── 4. VBA Save macros ────────────────────────────────────────────────────
    sec(r, "4.  VBA SAVE MACROS  (saves adjusted data back into the Raw tabs)", S_HDR4); r += 1
    for line in [
        "Two plain-text files are generated alongside this workbook:",
        "  VBA_Module1_SaveSensor.txt  \u2014 standard module code",
        "  VBA_Dashboard_Sheet.txt     \u2014 Dashboard sheet module code",
        "",
        "Step A:  Press Alt+F11 to open the VBA editor.",
        "Step B:  Click Insert \u2192 Module.  Open VBA_Module1_SaveSensor.txt in",
        "         Notepad, press Ctrl+A then Ctrl+C, and paste into the new module.",
        "Step C:  In the Project tree double-click 'Sheet1 (Dashboard)'.",
        "         Open VBA_Dashboard_Sheet.txt in Notepad, press Ctrl+A then",
        "         Ctrl+C, and paste into that sheet module.",
        "Step D:  Close the VBA editor and save the file as .xlsm.",
        "",
        "IMPORTANT: The \U0001f4be buttons overwrite the sensor column in the Raw tab.",
        "           Keep a backup of your original data before clicking Save.",
        "",
        "NOTE: Always use the .txt files \u2014 do NOT copy from this cell.",
        "      Excel wraps cell content in quotes, which corrupts VBA syntax.",
    ]:
        body(r, line); r += 1
    blank(r); r += 1

    rows.append((r, 17, [_cell("A" + str(r), S_CODE, VBA_MODULE)])); r += 1
    blank(r); r += 1

    sec(r,
        "DASHBOARD SHEET MODULE CODE  "
        "(open VBA_Dashboard_Sheet.txt in Notepad, copy, paste here)",
        S_HDR4); r += 1
    rows.append((r, 17,
        [_cell("A" + str(r), S_ITALIC,
               "Open VBA_Dashboard_Sheet.txt in Notepad, press Ctrl+A, Ctrl+C, then paste into "
               "the Dashboard sheet module (double-click 'Sheet1 (Dashboard)' in the Project tree). "
               "This makes the \U0001f4be cells (col E for flow, col J for pressure, rows 3\u201322) "
               "respond to a single click.")])); r += 1
    blank(r); r += 1

    rows.append((r, 17, [_cell("A" + str(r), S_CODE, VBA_SHEET)])); r += 1
    blank(r); r += 1

    # ── 5. Data Format Reference ──────────────────────────────────────────────
    sec(r, "5.  DATA FORMAT REFERENCE", S_HDR5); r += 1
    for line in [
        "Both Raw Data sheets use this wide format:",
        "",
        "    Date               | AL012       | AL013       | AL014       | ...  ",
        "    01/12/2026 00:00   | 3.168205    | 2.204250    | 2.665153    | ...  ",
        "    01/12/2026 00:15   | 3.190769    | 2.225250    | 2.681334    | ...  ",
        "",
        "  Dates are displayed in UK format: DD/MM/YYYY HH:MM",
        "",
        "\u2022 Column A must contain a proper Date/Time value (not text).",
        "\u2022 Flow and pressure column names can be any mix of letters and numbers.",
        "\u2022 The names in Raw Flow Data and Raw Pressure Data do NOT need to match",
        "  \u2014 you select each independently on the Dashboard.",
        "\u2022 Use -999 for missing/no-data values \u2014 they are excluded from all calculations.",
        "\u2022 Data can be any length: months of 15-minute data = thousands of rows.",
        "\u2022 To paste your own data into a raw sheet: delete the sample data rows",
        "  (keep row 1 headers), then paste starting from row 2.",
        "",
        "Dashboard formula table (rows 26+):",
        "  Col A     = Date",
        "  Cols B-U  = Flow 1-20 Adjusted  (Name in B3-B22 \u00d7 Scale in C3-C22 + \u0394t in D3-D22)",
        "  Cols V-AO = Pres 1-20 Adjusted  (Name in G3-G22 + Offset in H3-H22 + \u0394t in I3-I22)",
        "",
        "\U0001f4be Save buttons (col E = flow, col J = pressure):",
        "  Clicking \U0001f4be applies the current Scale / Offset / \u0394t and overwrites that",
        "  sensor's column in the Raw tab.  Keep a backup before saving.",
        "",
        "Leave a Name cell blank to hide that series (formula returns empty, not plotted).",
    ]:
        body(r, line); r += 1

    # ── Assemble worksheet XML ─────────────────────────────────────────────────
    sheetdata_lines = []
    for row_num, height, cells in rows:
        if cells:
            cell_xml = "".join(cells)
            sheetdata_lines.append(
                f'<row r="{row_num}" spans="1:1" ht="{height}" customHeight="1"'
                f' x14ac:dyDescent="0.25">{cell_xml}</row>'
            )
        else:
            sheetdata_lines.append(
                f'<row r="{row_num}" spans="1:1" ht="{height}" customHeight="1"'
                f' x14ac:dyDescent="0.25"/>'
            )

    last_row = rows[-1][0] if rows else 1
    sheetdata = "\n".join(sheetdata_lines)

    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' mc:Ignorable="x14ac xr xr2 xr3"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"'
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"'
        ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"'
        ' xr:uid="{00000000-0001-0000-0500-000000000000}">'
        '<sheetPr codeName="Sheet6"/>'
        f'<dimension ref="A1:A{last_row}"/>'
        '<sheetViews>'
        '<sheetView showGridLines="0" workbookViewId="0">'
        '<selection activeCell="A1" sqref="A1"/>'
        '</sheetView>'
        '</sheetViews>'
        '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
        '<cols><col min="1" max="1" width="115" customWidth="1"/></cols>'
        f'<sheetData>{sheetdata}</sheetData>'
        '</worksheet>'
    )
    return xml


# ---------------------------------------------------------------------------
# XML patching helpers
# ---------------------------------------------------------------------------

def _remove_mod_sheets_from_workbook_xml(xml):
    """Strip MOD Flow and MOD Pressure <sheet> entries from workbook.xml."""
    xml = re.sub(
        r'<sheet\s+name="MOD Flow"[^/]*/>\s*',
        "",
        xml,
    )
    xml = re.sub(
        r'<sheet\s+name="MOD Pressure"[^/]*/>\s*',
        "",
        xml,
    )
    return xml


def _remove_mod_rels(xml):
    """Strip rId5 and rId6 (MOD sheets) from workbook.xml.rels."""
    xml = re.sub(
        r'<Relationship\s+Id="rId5"[^/]*/>\s*',
        "",
        xml,
    )
    xml = re.sub(
        r'<Relationship\s+Id="rId6"[^/]*/>\s*',
        "",
        xml,
    )
    return xml


def _remove_mod_content_types(xml):
    """Strip MOD sheet overrides from [Content_Types].xml."""
    xml = re.sub(
        r'<Override\s+PartName="/xl/worksheets/sheet5\.xml"[^/]*/>\s*',
        "",
        xml,
    )
    xml = re.sub(
        r'<Override\s+PartName="/xl/worksheets/sheet6\.xml"[^/]*/>\s*',
        "",
        xml,
    )
    return xml


def _clear_save_rest_cell(xml):
    """
    Clear the value of cell K7 (Save Rest button) in the Dashboard sheet XML.
    The cell keeps its style but its shared-string value is removed.
    """
    # Replace  <c r="K7" s="74" t="s"><v>30</v></c>
    # with     <c r="K7" s="74"/>
    xml = re.sub(
        r'<c r="K7" s="\d+"[^>]*>.*?</c>',
        lambda m: re.sub(r'\s+t="[^"]*"', "", m.group(0).split(">")[0]) + "/>",
        xml,
        flags=re.DOTALL,
    )
    return xml


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not os.path.isfile(SOURCE):
        raise FileNotFoundError(
            f"'{SOURCE}' not found in the current directory.\n"
            "Run this script from the repository root."
        )

    print(f"Reading {SOURCE} …")
    out_buf = io.BytesIO()

    # Files from the original ZIP that belong to the MOD sheets — skip these.
    SKIP = {
        "xl/worksheets/sheet5.xml",   # MOD Flow
        "xl/worksheets/sheet6.xml",   # MOD Pressure
        "xl/calcChain.xml",           # stale calc order; Excel will rebuild
    }

    # Pre-build the new Instructions XML once (it's expensive).
    new_instr_xml = _build_instructions_xml().encode("utf-8")

    with zipfile.ZipFile(SOURCE, "r") as zin, \
         zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            name = item.filename

            if name in SKIP:
                print(f"  removing  {name}")
                continue

            data = zin.read(name)

            if name == "xl/workbook.xml":
                data = _remove_mod_sheets_from_workbook_xml(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (removed MOD sheet entries)")

            elif name == "xl/_rels/workbook.xml.rels":
                data = _remove_mod_rels(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (removed MOD rels)")

            elif name == "[Content_Types].xml":
                data = _remove_mod_content_types(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (removed MOD content types)")

            elif name == "xl/worksheets/sheet1.xml":
                data = _clear_save_rest_cell(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (cleared Save Rest cell K7)")

            elif name == "xl/worksheets/sheet7.xml":
                data = new_instr_xml
                print(f"  replaced  {name}  (updated Instructions)")

            zout.writestr(item, data)

    print(f"\nWriting updated workbook …")
    with open(SOURCE, "wb") as fh:
        fh.write(out_buf.getvalue())
    print(f"Saved: {SOURCE}")

    # ── Write VBA companion text files ────────────────────────────────────────
    out_dir = os.path.dirname(os.path.abspath(SOURCE))
    for filename, content in [
        ("VBA_Module1_SaveSensor.txt", VBA_MODULE),
        ("VBA_Dashboard_Sheet.txt",    VBA_SHEET),
    ]:
        path = os.path.join(out_dir, filename)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(content)
        print(f"Generated: {filename}")

    print("\nDone.")
    print()
    print("Next steps:")
    print("  1. Open 'Model Build Dashboard v1.21.xlsx' in Excel.")
    print("  2. Install the VBA macros (see Instructions sheet, section 4).")
    print("  3. Re-save the file as .xlsm to retain the macros.")


if __name__ == "__main__":
    main()
