#!/usr/bin/env python3
"""
generate_plots.py
=================
Rebuilds the Plots sheet and adds a "Plot Archive" sheet to
  Model Build Dashboard v1.33.xlsx

Run:
    python3 generate_plots.py

Outputs:
    Model Build Dashboard v1.33.xlsx   (updated in place)
    VBA_Plots_Sheet.txt                (paste into Plots sheet module)
    VBA_Module2_Plots.txt              (paste into a new Module2)
"""

import io, os, re, zipfile

SOURCE = "Model Build Dashboard v1.33.xlsx"

# ── New numFmt ID ────────────────────────────────────────────────────────────
NUMFMT_HHmm = 200       # "HH:MM" — safe user-defined range (164-400)

# ── Style indices: re-used from existing styles.xml ─────────────────────────
#   xf[1]  dark blue fill (#1F4E79), white bold, thin border, center
#   xf[6]  medium blue fill (#2E75B6), white bold 9pt, thin border, center
#   xf[30] green fill (#70AD47), white bold 11pt, medium border, center
#   xf[0]  default (no fill, no border)
S_DARK_HDR  = 1
S_MED_HDR   = 6
S_GREEN_BTN = 30
S_DEFAULT   = 0

# ── Style indices: ADDED by this script ─────────────────────────────────────
#   xf[86] yellow fill (#FFFFF2CC), medium border, left, wrap
#   xf[87] no fill, HH:MM numFmt, medium border, center
#   xf[88] light-yellow fill (#FFFFEB9C), thin border, center, wrap
S_YELLOW_IN = 86
S_TIME_CELL = 87
S_PASTE_LBL = 88

# ── VBA: Plots worksheet event handlers ─────────────────────────────────────
VBA_PLOTS_SHEET = """\
' ============================================================
' Plots Worksheet — Event Handlers
' How to install:
'   1. Alt+F11 to open the VBA editor.
'   2. In the Project pane, expand "Microsoft Excel Objects".
'   3. Double-click "Plots".
'   4. Paste this entire block, replacing any existing code.
' ============================================================

Private Sub Worksheet_Activate()
    ' Create the chart automatically the first time this sheet is visited.
    If Me.ChartObjects.Count = 0 Then
        Call CreatePlotsChart
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' B2 date changed → re-fill timestamps, auto-fill sensor data, refresh chart.
    If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
        Application.EnableEvents = False
        On Error Resume Next
        Call FillTimestamps
        Call AutoFillSensorData
        Call RefreshPlotsChart
        Application.EnableEvents = True   ' restored while On Error Resume Next is active
        On Error GoTo 0
    End If

    ' B6:K6 sensor header row changed → rebuild chart series dynamically.
    ' Clearing a header removes that series; typing a new one adds it.
    If Not Intersect(Target, Me.Range("B6:K6")) Is Nothing Then
        Application.EnableEvents = False
        On Error Resume Next
        Call RefreshPlotsChart
        Application.EnableEvents = True   ' restored while On Error Resume Next is active
        On Error GoTo 0
    End If

    ' E2 chart title changed → update the live chart title.
    If Not Intersect(Target, Me.Range("E2")) Is Nothing Then
        If Me.ChartObjects.Count > 0 Then
            With Me.ChartObjects(1).Chart
                .HasTitle = True
                .ChartTitle.Text = Trim(CStr(Me.Range("E2").Value))
            End With
        End If
    End If

    ' I2 Y-axis label changed → update the live chart Y-axis title.
    If Not Intersect(Target, Me.Range("I2")) Is Nothing Then
        If Me.ChartObjects.Count > 0 Then
            Dim yLabel As String
            yLabel = Trim(CStr(Me.Range("I2").Value))
            With Me.ChartObjects(1).Chart.Axes(xlValue)
                .HasTitle = (yLabel <> "")
                If yLabel <> "" Then .AxisTitle.Text = yLabel
            End With
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Allow merged-cell ranges through the guard (same fix as Dashboard).
    If Target.Count > 1 Then
        If Not Target.MergeCells Then Exit Sub
    End If
    Dim cell As Range
    Set cell = Target.Cells(1, 1)

    ' K2 = "SAVE PLOTS" button.
    If cell.Row = 2 And cell.Column = 11 Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        Call SavePlots
    End If

    ' L2 = "REFRESH CHART" button.
    ' Also called automatically when headers (B6:K6) change or a date is entered.
    ' Click this after manually pasting sensor names or editing column headers.
    If cell.Row = 2 And cell.Column = 12 Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        Call RefreshPlotsChart
    End If
End Sub
"""

# ── VBA: Module2 — Plots procedures ─────────────────────────────────────────
VBA_MODULE2_PLOTS = """\
' ============================================================
' Module2 — Plots Tab Procedures
' How to install:
'   1. Alt+F11 to open the VBA editor.
'   2. Insert → Module.
'   3. Paste this entire block into the new module.
' ============================================================

' ---------------------------------------------------------------------------
' FillTimestamps
' Fills A7:A102 with 15-minute interval times for the day entered in B2.
' The column is formatted as HH:MM so only the time portion is shown.
' ---------------------------------------------------------------------------
Sub FillTimestamps()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Plots")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim dateStr As String
    dateStr = Trim(CStr(ws.Range("B2").Value))
    If dateStr = "" Then Exit Sub

    Dim baseDate As Date
    On Error Resume Next
    baseDate = CDate(dateStr)
    On Error GoTo 0
    If baseDate = 0 Then
        MsgBox "Invalid date in B2. Please use DD/MM/YY format.", _
               vbExclamation, "Date Error"
        Exit Sub
    End If

    ' Write 96 timestamps: 00:00, 00:15, ... 23:45
    Dim i As Long
    For i = 0 To 95
        ws.Cells(7 + i, 1).Value = baseDate + CDbl(i * 15) / 1440
        ws.Cells(7 + i, 1).NumberFormat = "HH:MM"
    Next i
End Sub

' ---------------------------------------------------------------------------
' CreatePlotsChart
' Creates a line chart to the right of the data table, sized for an A4
' landscape document (18 cm wide x 11 cm tall — report-ready proportions).
'
' Builds ONE series per non-empty sensor header in B6:K6.
' Column A (timestamps) is used ONLY as the X-axis; it never appears as a
' series or in the legend.
' ---------------------------------------------------------------------------
Sub CreatePlotsChart()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Plots")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Remove any existing chart named PlotsChart.
    Dim chtObj As ChartObject
    For Each chtObj In ws.ChartObjects
        If chtObj.Name = "PlotsChart" Then chtObj.Delete: Exit For
    Next chtObj

    ' Position: left edge at column M (column 13), top at the header row (row 6).
    Dim L As Double: L = ws.Columns("M").Left
    Dim T As Double: T = ws.Rows(6).Top

    ' 18 cm x 11 cm — report-quality landscape proportions.
    ' 1 cm = 28.3465 pt.
    Dim W As Double: W = 18 * 28.3465
    Dim H As Double: H = 11 * 28.3465

    Set chtObj = ws.ChartObjects.Add(L, T, W, H)
    chtObj.Name = "PlotsChart"

    With chtObj.Chart
        .ChartType = xlLine

        ' Build series manually — one per non-empty sensor header (B6:K6).
        ' Column A is the X-axis only; it must NOT appear as a series.
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim c As Long
        For c = 2 To 11   ' columns B(2) to K(11)
            Dim hdr As String
            hdr = Trim(CStr(ws.Cells(6, c).Value))
            If hdr <> "" Then
                Dim sr As Series
                Set sr = .SeriesCollection.NewSeries
                sr.Name    = hdr
                sr.XValues = ws.Range("A7:A102")
                sr.Values  = ws.Range(ws.Cells(7, c), ws.Cells(102, c))
            End If
        Next c

        ' Title from the input cell.
        .HasTitle = True
        .ChartTitle.Text = Trim(CStr(ws.Range("E2").Value))

        ' Category (time) axis: HH:MM label every hour (every 4th of 96 fifteen-min steps).
        With .Axes(xlCategory)
            .TickLabels.NumberFormat = "HH:MM"
            .TickLabelSpacing = 4
            .TickLabelPosition = xlTickLabelPositionLow
        End With

        ' Value (Y) axis: title from I2 if provided.
        Dim yLabel As String
        yLabel = Trim(CStr(ws.Range("I2").Value))
        With .Axes(xlValue)
            .HasTitle = (yLabel <> "")
            If yLabel <> "" Then .AxisTitle.Text = yLabel
        End With

        ' Legend at the bottom.
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' Clean white chart area.
        .PlotArea.Interior.Color    = RGB(255, 255, 255)
        .ChartArea.Interior.Color   = RGB(255, 255, 255)
        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Weight    = xlThin
        .ChartArea.Border.Color     = RGB(180, 180, 180)
    End With

    ' Add a guide label for the software-plot paste area to the RIGHT of the chart.
    ' Column T (col 20) is the start of the paste zone (see row 5 guide labels).
    Dim pasteL As Double: pasteL = ws.Columns("T").Left

    ' Remove any stale paste label from a previous run.
    On Error Resume Next
    ws.Shapes("PastePlotLabel").Delete
    On Error GoTo 0

    Dim txtBox As Shape
    Set txtBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        pasteL, T, W, 22)
    txtBox.Name = "PastePlotLabel"
    With txtBox.TextFrame2
        .TextRange.Text = _
            "PASTE SOFTWARE PLOT IMAGE HERE (click here, then Ctrl+V)"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    With txtBox.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 102, 0)   ' orange — visually distinct from chart area
    End With
    txtBox.Line.Visible = msoFalse
End Sub

' ---------------------------------------------------------------------------
' RefreshPlotsChart
' Rebuilds the series collection of the existing PlotsChart without
' recreating the whole chart object.  Called when sensor headers (B6:K6)
' change or after an auto-fill.  If no chart exists yet, delegates to
' CreatePlotsChart.
' ---------------------------------------------------------------------------
Sub RefreshPlotsChart()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Plots")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' If no chart exists yet, do a full create.
    If ws.ChartObjects.Count = 0 Then
        Call CreatePlotsChart
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ws.ChartObjects(1).Chart

    With cht
        ' --- Rebuild series from current headers (B6:K6) ---
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim c As Long
        For c = 2 To 11   ' columns B(2) to K(11)
            Dim hdr As String
            hdr = Trim(CStr(ws.Cells(6, c).Value))
            If hdr <> "" Then
                Dim sr As Series
                Set sr = .SeriesCollection.NewSeries
                sr.Name    = hdr
                sr.XValues = ws.Range("A7:A102")
                sr.Values  = ws.Range(ws.Cells(7, c), ws.Cells(102, c))
            End If
        Next c

        ' --- Sync title and Y-axis label from control panel ---
        Dim chartTitle As String
        chartTitle = Trim(CStr(ws.Range("E2").Value))
        .HasTitle = True
        .ChartTitle.Text = chartTitle

        Dim yLabel As String
        yLabel = Trim(CStr(ws.Range("I2").Value))
        With .Axes(xlValue)
            .HasTitle = (yLabel <> "")
            If yLabel <> "" Then .AxisTitle.Text = yLabel
        End With
    End With
End Sub

' ---------------------------------------------------------------------------
' SavePlots
' 1. Exports the Excel chart as  {date}_{title}_Chart.png
' 2. Exports any pasted software-plot image as  {date}_{title}_SoftwarePlot.png
' 3. Archives data + metadata to the "Plot Archive" sheet.
' 4. Saves the session note to Point Index column M.
' ---------------------------------------------------------------------------
Sub SavePlots()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Plots")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' ── Read control-panel values ─────────────────────────────────────────
    Dim exportPath As String
    exportPath = Trim(CStr(ws.Range("B3").Value))
    If exportPath = "" Then
        MsgBox "Please enter an export folder path in cell B3.", _
               vbExclamation, "Missing Export Path"
        Exit Sub
    End If
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"

    Dim chartTitle As String
    chartTitle = Trim(CStr(ws.Range("E2").Value))
    If chartTitle = "" Then chartTitle = "Plot"

    Dim dateStr As String
    dateStr = Trim(CStr(ws.Range("B2").Value))

    ' Convert DD/MM/YY (or DD/MM/YYYY) → YYYYMMDD for filenames.
    Dim fileDate As String
    Dim dp() As String
    dp = Split(Replace(dateStr, "-", "/"), "/")
    If UBound(dp) = 2 Then
        Dim yr As String: yr = Trim(dp(2))
        If Len(yr) = 2 Then yr = "20" & yr
        fileDate = yr & Format(CInt(Trim(dp(1))), "00") & _
                        Format(CInt(Trim(dp(0))), "00")
    Else
        fileDate = Format(Now, "YYYYMMDD")
    End If

    ' Remove characters that are illegal in Windows filenames.
    Dim safeTitle As String
    safeTitle = chartTitle
    Dim badChar As Variant
    For Each badChar In Array("/", "\\", ":", "*", "?", Chr(34), "<", ">", "|")
        safeTitle = Replace(safeTitle, CStr(badChar), "-")
    Next badChar
    safeTitle = Replace(safeTitle, " ", "_")

    Dim saved As String: saved = ""

    ' ── 1. Export the Excel chart as PNG ──────────────────────────────────
    Dim chartFile As String
    chartFile = exportPath & fileDate & "_" & safeTitle & "_Chart.png"
    If ws.ChartObjects.Count > 0 Then
        ws.ChartObjects(1).Chart.Export Filename:=chartFile, FilterName:="PNG"
        saved = saved & "  Chart:  " & chartFile & vbNewLine
    Else
        saved = saved & "  (No chart found — open the Plots sheet first to auto-create it)" & vbNewLine
    End If

    ' ── 2. Export pasted software-plot image as PNG ───────────────────────
    ' Find the first picture shape that is NOT the chart or label textbox.
    Dim shp As Shape
    Dim imgFile As String
    imgFile = exportPath & fileDate & "_" & safeTitle & "_SoftwarePlot.png"
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            ' Export via a temporary off-screen chart object (standard VBA trick).
            Dim tmpC As ChartObject
            Set tmpC = ws.ChartObjects.Add(0, 0, shp.Width, shp.Height)
            shp.CopyPicture xlScreen, xlPicture
            tmpC.Chart.Paste
            tmpC.Chart.Export Filename:=imgFile, FilterName:="PNG"
            tmpC.Delete
            saved = saved & "  Image:  " & imgFile & vbNewLine
            Exit For
        End If
    Next shp

    ' ── 3. Archive data and note ──────────────────────────────────────────
    Call ArchivePlotData(ws, chartTitle, dateStr)
    saved = saved & "  Data archived to 'Plot Archive' sheet." & vbNewLine

    MsgBox "Save complete:" & vbNewLine & vbNewLine & saved, _
           vbInformation, "Plots Saved"
End Sub

' ---------------------------------------------------------------------------
' ArchivePlotData
' Appends the current day's timestamps, sensor readings, chart title,
' date and session notes to the "Plot Archive" sheet.
' Also writes the session note to Point Index column M.
' ---------------------------------------------------------------------------
Sub ArchivePlotData(ws As Worksheet, chartTitle As String, dateStr As String)
    Const ARCH_SHEET As String = "Plot Archive"
    Const FIRST_ROW  As Long   = 7    ' first data row on the Plots sheet
    Const DATA_ROWS  As Long   = 96   ' one day at 15-min intervals
    Const SEN_COLS   As Long   = 10   ' sensor columns B(2) through K(11)

    ' ── Get or create the Plot Archive sheet ──────────────────────────────
    Dim wsA As Worksheet
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets(ARCH_SHEET)
    On Error GoTo 0

    If wsA Is Nothing Then
        Set wsA = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsA.Name = ARCH_SHEET
        ' Header row
        wsA.Cells(1, 1).Value = "Archived"
        wsA.Cells(1, 2).Value = "Date"
        wsA.Cells(1, 3).Value = "Chart Title"
        wsA.Cells(1, 4).Value = "Notes"
        wsA.Cells(1, 5).Value = "Time"
        Dim h As Long
        For h = 1 To SEN_COLS
            wsA.Cells(1, 5 + h).Value = ws.Cells(6, 1 + h).Value
        Next h
        wsA.Rows(1).Font.Bold = True
    End If

    ' Get session notes from B4.
    Dim notes As String
    notes = Trim(CStr(ws.Range("B4").Value))

    ' Find the first empty row in the archive.
    Dim nxtRow As Long
    nxtRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    If nxtRow < 2 Then nxtRow = 2

    Dim archDate As String
    archDate = Format(Now, "DD/MM/YYYY HH:MM")

    ' Copy 96 rows of data.
    Dim i As Long, j As Long
    For i = 0 To DATA_ROWS - 1
        wsA.Cells(nxtRow + i, 1).Value = archDate
        wsA.Cells(nxtRow + i, 2).Value = dateStr
        wsA.Cells(nxtRow + i, 3).Value = chartTitle
        wsA.Cells(nxtRow + i, 4).Value = notes
        wsA.Cells(nxtRow + i, 5).Value = ws.Cells(FIRST_ROW + i, 1).Value
        wsA.Cells(nxtRow + i, 5).NumberFormat = "HH:MM"
        For j = 1 To SEN_COLS
            wsA.Cells(nxtRow + i, 5 + j).Value = _
                ws.Cells(FIRST_ROW + i, 1 + j).Value
        Next j
    Next i

    ' ── Save note to Point Index column M ─────────────────────────────────
    Call SaveNotesToPointIndex(dateStr, chartTitle, notes)
End Sub

' ---------------------------------------------------------------------------
' SaveNotesToPointIndex
' Appends a session note (date + chart title + notes text) to column M
' of the "Point Index" sheet.  Column L is used by the Save button on the
' Dashboard; column M is the new Plots-session notes column.
' ---------------------------------------------------------------------------
Sub SaveNotesToPointIndex(dateStr As String, _
                          chartTitle As String, _
                          notes As String)
    Dim wsPI As Worksheet
    On Error Resume Next
    Set wsPI = ThisWorkbook.Sheets("Point Index")
    On Error GoTo 0
    If wsPI Is Nothing Then Exit Sub

    ' Add header in M1 if the column is still empty.
    If Trim(CStr(wsPI.Cells(1, 13).Value)) = "" Then
        wsPI.Cells(1, 13).Value = "Plot Notes"
        wsPI.Cells(1, 13).Font.Bold = True
    End If

    ' Append to the next empty row in column M.
    Dim piRow As Long
    piRow = wsPI.Cells(wsPI.Rows.Count, 13).End(xlUp).Row + 1
    If piRow < 2 Then piRow = 2
    wsPI.Cells(piRow, 13).Value = _
        "[" & dateStr & "]  " & chartTitle & "  —  " & notes
End Sub

' ---------------------------------------------------------------------------
' CopySelectionToPlotsHeaders
' Reads non-empty text values from the current selection (on any sheet) and
' writes them as sensor column headers into B6:K6 on the Plots sheet, then
' calls RefreshPlotsChart so the chart updates immediately.
'
' Typical workflow:
'   1. Open the Point Index sheet (or any sheet with sensor names).
'   2. Select one or more cells that contain sensor / point reference names.
'   3. Press Ctrl+Shift+V  — or click the REFRESH CHART button after pasting
'      names manually into B6:K6.
'
' The shortcut is registered automatically when the workbook opens; see the
' Workbook_Open event in ThisWorkbook.
' ---------------------------------------------------------------------------
Sub CopySelectionToPlotsHeaders()
    Dim wsPlots As Worksheet
    On Error Resume Next
    Set wsPlots = ThisWorkbook.Sheets("Plots")
    On Error GoTo 0
    If wsPlots Is Nothing Then
        MsgBox "Plots sheet not found.", vbExclamation, "Copy to Plots"
        Exit Sub
    End If

    ' Collect up to 10 non-empty text values from the current selection.
    Dim vals(9) As String
    Dim nVals As Long: nVals = 0
    Dim cell As Range
    For Each cell In Application.Selection
        Dim v As String
        v = Trim(CStr(cell.Value))
        If v <> "" And v <> "0" Then
            If nVals < 10 Then
                vals(nVals) = v
                nVals = nVals + 1
            End If
        End If
    Next cell

    If nVals = 0 Then
        MsgBox "No sensor names found in the current selection." & vbNewLine & _
               "Select one or more sensor name cells, then press Ctrl+Shift+V.", _
               vbInformation, "Copy to Plots Headers"
        Exit Sub
    End If

    ' Write the names into B6:K6, clearing any existing headers first.
    Application.EnableEvents = False
    On Error Resume Next
    wsPlots.Range("B6:K6").ClearContents
    Dim i As Long
    For i = 0 To nVals - 1
        wsPlots.Cells(6, 2 + i).Value = vals(i)
    Next i
    Application.EnableEvents = True
    On Error GoTo 0

    ' Rebuild the chart series for the new headers.
    Call RefreshPlotsChart

    MsgBox nVals & " sensor name(s) copied to Plots B6:K6." & vbNewLine & _
           "Switch to the Plots sheet to see the updated chart.", _
           vbInformation, "Headers Updated"
End Sub

' ---------------------------------------------------------------------------
' AutoFillSensorData
' Searches "Raw Pressure Data" then "Raw Flow Data" for readings that match
' the date in B2.  Writes sensor names to B6:K6 and values to B7:K102.
' Existing values are cleared first; any cell can be overwritten manually.
' Silently skips (leaves the data area empty) if no matching data is found.
' ---------------------------------------------------------------------------
Sub AutoFillSensorData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Plots")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim baseDate As Date
    On Error Resume Next
    baseDate = CDate(Trim(CStr(ws.Range("B2").Value)))
    On Error GoTo 0
    If CDbl(baseDate) = 0 Then Exit Sub

    ' Clear sensor headers and data ready for a fresh fill.
    ws.Range("B6:K6").ClearContents
    ws.Range("B7:K102").ClearContents

    ' Try Raw Pressure Data first, then Raw Flow Data.
    Dim rawSheets(1) As String
    rawSheets(0) = "Raw Pressure Data"
    rawSheets(1) = "Raw Flow Data"

    Dim wsRaw As Worksheet
    Dim si As Integer
    For si = 0 To 1
        Set wsRaw = Nothing
        On Error Resume Next
        Set wsRaw = ThisWorkbook.Sheets(rawSheets(si))
        On Error GoTo 0
        If Not wsRaw Is Nothing Then
            If FillFromRawSheet(ws, wsRaw, baseDate) Then Exit For
        End If
    Next si
End Sub

' ---------------------------------------------------------------------------
' FillFromRawSheet  (helper for AutoFillSensorData)
' Finds rows in wsRaw whose timestamps fall on baseDate, copies up to 10
' sensor columns to wsPlots rows 6 (headers) and 7-102 (data).
' Matches each 15-min slot to the nearest raw timestamp within +/-7.5 minutes.
' Returns True when at least one data value was written.
' ---------------------------------------------------------------------------
Function FillFromRawSheet(wsPlots As Worksheet, _
                          wsRaw   As Worksheet, _
                          baseDate As Date) As Boolean
    FillFromRawSheet = False

    Dim lastRow As Long: lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim lastCol As Long: lastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then Exit Function

    Dim baseDbl As Double: baseDbl = CDbl(baseDate)

    ' ── Find the first row whose date portion equals baseDate ─────────────
    Dim startRow As Long: startRow = 0
    Dim r As Long
    Dim tsDbl As Double
    For r = 2 To lastRow
        On Error Resume Next
        tsDbl = CDbl(wsRaw.Cells(r, 1).Value)
        On Error GoTo 0
        If Int(tsDbl) = Int(baseDbl) Then
            startRow = r
            Exit For
        End If
    Next r
    If startRow = 0 Then Exit Function

    ' ── Count consecutive rows that belong to this date ───────────────────
    Dim dayCount As Long: dayCount = 0
    For r = startRow To lastRow
        On Error Resume Next
        tsDbl = CDbl(wsRaw.Cells(r, 1).Value)
        On Error GoTo 0
        If Int(tsDbl) = Int(baseDbl) Then
            dayCount = dayCount + 1
        Else
            Exit For
        End If
    Next r
    If dayCount = 0 Then Exit Function

    ' ── Number of sensor columns to copy (max 10) ─────────────────────────
    Dim sensorCols As Long
    sensorCols = Application.WorksheetFunction.Min(lastCol - 1, 10)

    ' ── Copy sensor names to Plots row 6 ──────────────────────────────────
    Dim c As Long
    For c = 1 To sensorCols
        wsPlots.Cells(6, 1 + c).Value = wsRaw.Cells(1, 1 + c).Value
    Next c

    ' ── Load day rows into arrays for fast per-slot lookup ────────────────
    Dim rawTS()   As Double
    Dim rawVals() As Variant
    ReDim rawTS(1 To dayCount)
    ReDim rawVals(1 To dayCount, 1 To sensorCols)
    Dim di As Long: di = 1
    For r = startRow To startRow + dayCount - 1
        On Error Resume Next
        rawTS(di) = CDbl(wsRaw.Cells(r, 1).Value)
        On Error GoTo 0
        For c = 1 To sensorCols
            rawVals(di, c) = wsRaw.Cells(r, 1 + c).Value
        Next c
        di = di + 1
    Next r

    ' ── For each 15-min slot find the nearest raw row (within +/-7.5 min) ─
    Dim halfStep As Double: halfStep = 7.5 / 1440
    Dim i As Long
    Dim slotDbl  As Double
    Dim bestIdx  As Long
    Dim bestDiff As Double
    Dim dIdx     As Long
    Dim diff     As Double
    For i = 0 To 95
        slotDbl  = baseDbl + CDbl(i * 15) / 1440
        bestIdx  = 0
        bestDiff = halfStep + 1
        For dIdx = 1 To dayCount
            diff = Abs(rawTS(dIdx) - slotDbl)
            If diff < bestDiff Then
                bestDiff = diff
                bestIdx  = dIdx
            End If
        Next dIdx
        If bestIdx > 0 And bestDiff <= halfStep Then
            For c = 1 To sensorCols
                wsPlots.Cells(7 + i, 1 + c).Value = rawVals(bestIdx, c)
            Next c
            FillFromRawSheet = True
        End If
    Next i
End Function
"""

# ── VBA: ThisWorkbook — Workbook event handlers ──────────────────────────────
VBA_THIS_WORKBOOK = """\
' ============================================================
' ThisWorkbook — Workbook Event Handlers
' How to install:
'   1. Alt+F11 to open the VBA editor.
'   2. In the Project pane, expand "Microsoft Excel Objects".
'   3. Double-click "ThisWorkbook".
'   4. Paste this entire block, replacing any existing code.
' ============================================================

Private Sub Workbook_Open()
    ' Register Ctrl+Shift+V as a shortcut to push the currently selected
    ' sensor names (from any sheet, e.g. Point Index) into the Plots sheet
    ' header row (B6:K6), then refresh the chart automatically.
    '
    ' Usage:
    '   1. On the Point Index sheet, select one or more cells containing
    '      sensor / point-reference names.
    '   2. Press Ctrl+Shift+V.
    '   3. Switch to the Plots sheet — the chart will reflect the new headers.
    Application.OnKey "^+V", "CopySelectionToPlotsHeaders"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Deregister the custom shortcut so it does not persist in other workbooks.
    Application.OnKey "^+V"
End Sub
"""


# ── XML helpers ──────────────────────────────────────────────────────────────

def _col_letter(n: int) -> str:
    """Convert 1-based column number to Excel letter(s)."""
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _ref(col: int, row: int) -> str:
    return f"{_col_letter(col)}{row}"


def _escape(text: str) -> str:
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


def _str_cell(col: int, row: int, style: int, value: str) -> str:
    """Inline-string cell."""
    return (f'<c r="{_ref(col, row)}" s="{style}" t="inlineStr">'
            f'<is><t>{_escape(value)}</t></is></c>')


def _empty_cell(col: int, row: int, style: int) -> str:
    return f'<c r="{_ref(col, row)}" s="{style}"/>'


# ── Style patcher ────────────────────────────────────────────────────────────

_STYLES_MARKER = "<!--PLOTS_STYLES_ADDED-->"


def _add_new_styles(styles_xml: str) -> str:
    """
    Idempotent: adds numFmt 200 and three new xf entries (indices 86-88)
    to styles.xml.  A comment marker prevents double-application.
    """
    if _STYLES_MARKER in styles_xml:
        return styles_xml   # already patched

    # 1. Add numFmt 200 = "HH:MM"
    new_numfmt = f'<numFmt numFmtId="{NUMFMT_HHmm}" formatCode="HH:MM"/>'
    if f'numFmtId="{NUMFMT_HHmm}"' not in styles_xml:
        styles_xml = styles_xml.replace("</numFmts>",
                                        new_numfmt + "</numFmts>")
        styles_xml = re.sub(
            r'<numFmts count="(\d+)">',
            lambda m: f'<numFmts count="{int(m.group(1)) + 1}">',
            styles_xml,
        )

    # 2. Add three new xf entries.
    new_xfs = (
        # S_YELLOW_IN = 86: yellow fill (fillId=7), medium border, left, wrap
        '<xf numFmtId="0" fontId="0" fillId="7" borderId="2" xfId="0"'
        ' applyFill="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="left" vertical="center" wrapText="1"/></xf>',

        # S_TIME_CELL = 87: no fill, HH:MM numFmt, medium border, center
        f'<xf numFmtId="{NUMFMT_HHmm}" fontId="0" fillId="0" borderId="2"'
        ' xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center"/></xf>',

        # S_PASTE_LBL = 88: light-yellow fill (fillId=48), thin border, center, wrap
        '<xf numFmtId="0" fontId="9" fillId="48" borderId="1" xfId="0"'
        ' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>',
    )

    styles_xml = styles_xml.replace(
        "</cellXfs>",
        _STYLES_MARKER + "".join(new_xfs) + "</cellXfs>",
    )
    styles_xml = re.sub(
        r'<cellXfs count="(\d+)">',
        lambda m: f'<cellXfs count="{int(m.group(1)) + len(new_xfs)}">',
        styles_xml,
    )
    return styles_xml


# ── Plots sheet XML builder ──────────────────────────────────────────────────

def _build_plots_sheet_xml() -> str:
    """
    Returns the complete XML for a clean Plots worksheet.

    Layout
    ──────
    Row 1  : Title bar  (A1:K1 merged)
    Row 2  : Date (B2) | Chart Title (E2:G2 merged) | Y-Axis Label (I2:J2 merged) | SAVE PLOTS button (K2)
    Row 3  : Export Path (B3:J3 merged)
    Row 4  : Session Notes (B4:J4 merged, taller row)
    Row 5  : Separator / right-panel guide label
    Row 6  : Column headers  Time | Sensor 1 … Sensor 10
    Rows 7-102 : 96 data rows (one full day at 15-minute intervals)
    """
    rows = []
    merges = []

    # ── Row 1 — title ────────────────────────────────────────────────────────
    r = [_str_cell(1, 1, S_DARK_HDR, "\U0001f4ca  PLOTS \u2014 Final Report")]
    for c in range(2, 12):
        r.append(_empty_cell(c, 1, S_DARK_HDR))
    rows.append(f'<row r="1" ht="24" customHeight="1">{"".join(r)}</row>')
    merges.append('<mergeCell ref="A1:K1"/>')

    # ── Row 2 — Date | Chart Title | Y-Axis Label | Save button ────────────────
    r = [
        _str_cell(1, 2, S_DEFAULT, "Date (DD/MM/YY):"),   # A2 label
        _str_cell(2, 2, S_YELLOW_IN, ""),                  # B2 date input ← date
        _str_cell(3, 2, S_DEFAULT, ""),                    # C2 gap
        _str_cell(4, 2, S_DEFAULT, "Chart Title:"),        # D2 label
        _str_cell(5, 2, S_YELLOW_IN, ""),                  # E2 title input ← chart title (E2:G2 merged)
        _empty_cell(6, 2, S_YELLOW_IN),                    # F2
        _empty_cell(7, 2, S_YELLOW_IN),                    # G2
        _str_cell(8, 2, S_DEFAULT, "Y-Axis Label:"),       # H2 label
        _str_cell(9, 2, S_YELLOW_IN, ""),                  # I2 y-axis input ← y label (I2:J2 merged)
        _empty_cell(10, 2, S_YELLOW_IN),                   # J2
        _str_cell(11, 2, S_GREEN_BTN, "\U0001f4be  SAVE PLOTS"),   # K2 button
        _str_cell(12, 2, S_MED_HDR,   "\u21ba  REFRESH CHART"),    # L2 button
    ]
    rows.append(f'<row r="2" ht="22" customHeight="1">{"".join(r)}</row>')
    merges.append('<mergeCell ref="E2:G2"/>')
    merges.append('<mergeCell ref="I2:J2"/>')

    # ── Row 3 — Export path ───────────────────────────────────────────────────
    r = [
        _str_cell(1, 3, S_DEFAULT, "Export Path:"),        # A3 label
        _str_cell(2, 3, S_YELLOW_IN, ""),                  # B3 path input ← export path
        _empty_cell(3, 3, S_YELLOW_IN),
        _empty_cell(4, 3, S_YELLOW_IN),
        _empty_cell(5, 3, S_YELLOW_IN),
        _empty_cell(6, 3, S_YELLOW_IN),
        _empty_cell(7, 3, S_YELLOW_IN),
        _empty_cell(8, 3, S_YELLOW_IN),
        _empty_cell(9, 3, S_YELLOW_IN),
        _empty_cell(10, 3, S_YELLOW_IN),
        _empty_cell(11, 3, S_DEFAULT),                     # K3 gap
    ]
    rows.append(f'<row r="3" ht="22" customHeight="1">{"".join(r)}</row>')
    merges.append('<mergeCell ref="B3:J3"/>')

    # ── Row 4 — Session notes ────────────────────────────────────────────────
    r = [
        _str_cell(1, 4, S_DEFAULT, "Session Notes:"),      # A4 label
        _str_cell(2, 4, S_YELLOW_IN, ""),                  # B4 notes input ← notes
        _empty_cell(3, 4, S_YELLOW_IN),
        _empty_cell(4, 4, S_YELLOW_IN),
        _empty_cell(5, 4, S_YELLOW_IN),
        _empty_cell(6, 4, S_YELLOW_IN),
        _empty_cell(7, 4, S_YELLOW_IN),
        _empty_cell(8, 4, S_YELLOW_IN),
        _empty_cell(9, 4, S_YELLOW_IN),
        _empty_cell(10, 4, S_YELLOW_IN),
        _empty_cell(11, 4, S_DEFAULT),
    ]
    rows.append(f'<row r="4" ht="36" customHeight="1">{"".join(r)}</row>')
    merges.append('<mergeCell ref="B4:J4"/>')

    # ── Row 5 — separator + two right-panel guide labels ─────────────────────
    r = [_empty_cell(c, 5, S_DEFAULT) for c in range(1, 12)]
    # Chart area label: M5:S5 (cols 13-19) — blue header
    r.append(_str_cell(13, 5, S_MED_HDR,
                        "\u2193  EXCEL CHART AREA  \u2193"))
    for c in range(14, 20):
        r.append(_empty_cell(c, 5, S_MED_HDR))
    # Paste area label: T5:AF5 (cols 20-32) — yellow label
    r.append(_str_cell(20, 5, S_PASTE_LBL,
                        "\u2193  PASTE SOFTWARE PLOT IMAGE HERE  \u2193"))
    for c in range(21, 33):
        r.append(_empty_cell(c, 5, S_PASTE_LBL))
    rows.append(f'<row r="5" ht="18" customHeight="1">{"".join(r)}</row>')
    merges.append('<mergeCell ref="M5:S5"/>')
    merges.append('<mergeCell ref="T5:AF5"/>')

    # ── Row 6 — column headers ────────────────────────────────────────────────
    r = [_str_cell(1, 6, S_MED_HDR, "Time")]
    for s in range(1, 11):
        r.append(_str_cell(1 + s, 6, S_MED_HDR, f"Sensor {s}"))
    rows.append(f'<row r="6" ht="18" customHeight="1">{"".join(r)}</row>')

    # ── Rows 7-102 — data (96 rows = one day at 15-min intervals) ────────────
    for rn in range(7, 103):
        r = [_empty_cell(1, rn, S_TIME_CELL)]           # A: timestamp
        for c in range(2, 12):
            r.append(_empty_cell(c, rn, S_DEFAULT))     # B-K: sensor data
        rows.append(f'<row r="{rn}">{"".join(r)}</row>')

    # ── Assemble worksheet ────────────────────────────────────────────────────
    merge_block = (
        f'<mergeCells count="{len(merges)}">{"".join(merges)}</mergeCells>'
    )
    cols = (
        "<cols>"
        '<col min="1"  max="1"  width="13"  customWidth="1"/>'  # A  timestamps
        '<col min="2"  max="11" width="12"  customWidth="1"/>'  # B-K sensor data
        '<col min="12" max="12" width="16"  customWidth="1"/>'  # L  REFRESH CHART button
        '<col min="13" max="19" width="14"  customWidth="1"/>'  # M-S chart zone
        '<col min="20" max="32" width="10"  customWidth="1"/>'  # T-AF paste zone
        "</cols>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' mc:Ignorable="x14ac xr xr2 xr3"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        '<dimension ref="A1:AF102"/>'
        '<sheetViews>'
        '<sheetView tabSelected="1" workbookViewId="0">'
        '<selection activeCell="B2" sqref="B2"/>'
        '</sheetView>'
        '</sheetViews>'
        '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
        f'{cols}'
        f'<sheetData>{"".join(rows)}</sheetData>'
        f'{merge_block}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"'
        ' header="0.3" footer="0.3"/>'
        '</worksheet>'
    )


def _build_plot_archive_sheet_xml() -> str:
    """
    Returns XML for the Plot Archive worksheet — header row only.
    Data is appended at run time by ArchivePlotData() in Module2.
    """
    headers = (
        ["Archived", "Date", "Chart Title", "Notes", "Time"]
        + [f"Sensor {i}" for i in range(1, 11)]
    )
    r = [_str_cell(c + 1, 1, S_MED_HDR, h) for c, h in enumerate(headers)]
    last_col = _col_letter(len(headers))
    cols = (
        "<cols>"
        '<col min="1"  max="1"  width="20" customWidth="1"/>'  # Archived
        '<col min="2"  max="2"  width="12" customWidth="1"/>'  # Date
        '<col min="3"  max="3"  width="22" customWidth="1"/>'  # Chart Title
        '<col min="4"  max="4"  width="30" customWidth="1"/>'  # Notes
        '<col min="5"  max="5"  width="10" customWidth="1"/>'  # Time
        f'<col min="6"  max="{len(headers)}" width="12" customWidth="1"/>'
        "</cols>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{last_col}1"/>'
        '<sheetViews>'
        '<sheetView workbookViewId="0">'
        '<selection activeCell="A2" sqref="A2"/>'
        '</sheetView>'
        '</sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'{cols}'
        f'<sheetData><row r="1" ht="18" customHeight="1">{"".join(r)}</row></sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"'
        ' header="0.3" footer="0.3"/>'
        '</worksheet>'
    )


# ── Workbook patchers ─────────────────────────────────────────────────────────

_ARCHIVE_SHEET_ID  = 9
_ARCHIVE_RID       = "rId12"
_ARCHIVE_FILE      = "worksheets/sheet7.xml"
_ARCHIVE_SHEET_TAG = (
    f'<sheet name="Plot Archive" sheetId="{_ARCHIVE_SHEET_ID}"'
    f' r:id="{_ARCHIVE_RID}"/>'
)
_WB_MARKER = "<!--PLOT_ARCHIVE_ADDED-->"


def _patch_workbook_xml(wb_xml: str) -> str:
    # Remove stale external-references block (broken network-drive link).
    wb_xml = re.sub(r'<externalReferences>.*?</externalReferences>', '',
                    wb_xml, flags=re.DOTALL)
    if _WB_MARKER in wb_xml:
        return wb_xml
    return wb_xml.replace(
        "</sheets>",
        _WB_MARKER + _ARCHIVE_SHEET_TAG + "</sheets>",
    )


def _patch_workbook_rels(rels_xml: str) -> str:
    # Remove the externalLink relationship (broken network-drive link).
    rels_xml = re.sub(
        r'<Relationship[^>]*Type="[^"]*externalLink[^"]*"[^>]*/>\s*',
        '', rels_xml,
    )
    if _ARCHIVE_RID in rels_xml:
        return rels_xml
    new_rel = (
        f'<Relationship Id="{_ARCHIVE_RID}"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
        f' Target="{_ARCHIVE_FILE}"/>'
    )
    return rels_xml.replace("</Relationships>", new_rel + "</Relationships>")


def _patch_content_types(ct_xml: str) -> str:
    # Remove stale externalLink content-type entry.
    ct_xml = re.sub(
        r'<Override[^>]*externalLink[^>]*/>\s*',
        '', ct_xml,
    )
    archive_path = f'/xl/{_ARCHIVE_FILE}'
    if archive_path in ct_xml:
        return ct_xml
    new_override = (
        f'<Override PartName="{archive_path}"'
        ' ContentType="application/vnd.openxmlformats-officedocument'
        '.spreadsheetml.worksheet+xml"/>'
    )
    return ct_xml.replace("</Types>", new_override + "</Types>")


# ── Minimal sheet rels (no drawing for the fresh Plots sheet) ────────────────

_EMPTY_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '</Relationships>'
)


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    if not os.path.isfile(SOURCE):
        print(f"ERROR: {SOURCE!r} not found in the current directory.")
        return

    print(f"Processing {SOURCE} …\n")

    # Pre-build the new sheet XMLs once.
    new_plots_xml    = _build_plots_sheet_xml().encode("utf-8")
    new_archive_xml  = _build_plot_archive_sheet_xml().encode("utf-8")

    out_buf = io.BytesIO()

    with zipfile.ZipFile(SOURCE, "r") as zin, \
         zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:

        existing = {item.filename for item in zin.infolist()}

        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)

            # ── Plots sheet (sheet3.xml) — full replacement ───────────────
            if name == "xl/worksheets/sheet3.xml":
                data = new_plots_xml
                print(f"  replaced  {name}  (clean Plots layout)")

            # ── Plots sheet rels — strip old drawing reference ────────────
            elif name == "xl/worksheets/_rels/sheet3.xml.rels":
                data = _EMPTY_RELS.encode("utf-8")
                print(f"  replaced  {name}  (removed stale drawing reference)")

            # ── styles.xml — add new xf entries ──────────────────────────
            elif name == "xl/styles.xml":
                data = _add_new_styles(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (added S_YELLOW_IN/S_TIME_CELL/S_PASTE_LBL styles)")

            # ── workbook.xml — register Plot Archive sheet ────────────────
            elif name == "xl/workbook.xml":
                data = _patch_workbook_xml(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (added Plot Archive sheet entry)")

            # ── workbook rels — add Plot Archive relationship ─────────────
            elif name == "xl/_rels/workbook.xml.rels":
                data = _patch_workbook_rels(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (added Plot Archive rId12)")

            # ── [Content_Types].xml — register new worksheet ──────────────
            elif name == "[Content_Types].xml":
                data = _patch_content_types(data.decode("utf-8")).encode("utf-8")
                print(f"  patched   {name}  (added Plot Archive content type)")

            # ── drop stale calc chain (Excel rebuilds it on open) ─────────
            elif name == "xl/calcChain.xml":
                print(f"  removed   {name}  (Excel will rebuild)")
                continue

            # ── drop broken external-link files ───────────────────────────
            elif name.startswith("xl/externalLinks/"):
                print(f"  removed   {name}  (stale network-drive link)")
                continue

            zout.writestr(item, data)

        # ── Write Plot Archive sheet if sheet7.xml doesn't exist yet ─────
        archive_file = f"xl/{_ARCHIVE_FILE}"
        if archive_file not in existing:
            zout.writestr(archive_file, new_archive_xml)
            print(f"  added     {archive_file}  (new Plot Archive sheet)")

    # ── Save workbook ─────────────────────────────────────────────────────────
    print(f"\nWriting updated workbook …")
    with open(SOURCE, "wb") as fh:
        fh.write(out_buf.getvalue())
    print(f"Saved: {SOURCE}")

    # ── Write VBA companion text files ────────────────────────────────────────
    out_dir = os.path.dirname(os.path.abspath(SOURCE))
    for filename, content in [
        ("VBA_Plots_Sheet.txt",    VBA_PLOTS_SHEET),
        ("VBA_Module2_Plots.txt",  VBA_MODULE2_PLOTS),
        ("VBA_ThisWorkbook.txt",   VBA_THIS_WORKBOOK),
    ]:
        path = os.path.join(out_dir, filename)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(content)
        print(f"Generated: {filename}")

    print()
    print("Done.")
    print()
    print("Next steps:")
    print("  1. Open 'Model Build Dashboard v1.33.xlsx' in Excel.")
    print("  2. Press Alt+F11 to open the VBA editor.")
    print("  3. Paste VBA_Plots_Sheet.txt into the Plots sheet module")
    print("     (double-click 'Plots' under Microsoft Excel Objects).")
    print("  4. Insert a new standard Module (Insert → Module) and paste")
    print("     VBA_Module2_Plots.txt into it.")
    print("  5. Paste VBA_ThisWorkbook.txt into the ThisWorkbook module")
    print("     (double-click 'ThisWorkbook' under Microsoft Excel Objects).")
    print("  6. Save the file as .xlsm to retain the macros.")
    print("  7. Close and reopen the workbook so Workbook_Open fires and")
    print("     registers the Ctrl+Shift+V shortcut.")
    print()
    print("Plots sheet usage:")
    print("  • B2  — Enter the date (DD/MM/YY); timestamps fill A7:A102 automatically.")
    print("  • E2  — Enter the chart title; the chart updates live.")
    print("  • I2  — Enter the Y-axis label (e.g. 'Flow (L/s)' or 'Pressure (bar)'); updates live.")
    print("  • B3  — Enter the export folder path (e.g. C:\\Reports\\).")
    print("  • B4  — Enter session notes (saved to Point Index col M on Save).")
    print("  • B6:K6 — Sensor column headers (auto-filled by date; edit manually to swap sensors).")
    print("  • B7:K102 — Sensor data (auto-filled by date; paste manually to override).")
    print("  • K2  — Click SAVE PLOTS to export PNGs and archive data.")
    print("  • L2  — Click REFRESH CHART to rebuild chart series from current B6:K6 headers.")
    print()
    print("Dynamic sensor workflow:")
    print("  Option A — Auto-fill: enter a date in B2; sensors from Raw Data fill automatically.")
    print("  Option B — Manual headers: type or paste sensor names into B6:K6, then click L2.")
    print("  Option C — From Point Index: select sensor names in Point Index sheet,")
    print("             then press Ctrl+Shift+V to copy them to B6:K6 and refresh the chart.")


if __name__ == "__main__":
    main()
