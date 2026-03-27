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
'                  For pressure sensors the elevation in col J is also written
'                  back to the "Point Index" tab.
'
' isFlow : True  -> Flow sensor   (Raw Flow Data,     Scale * value,  col D Dt)
'          False -> Pressure      (Raw Pressure Data, value + Offset, col I Dt)
' sRow   : selector row index 1-20
'
' Dashboard selector rows 3-22:
'   Col B (2) = Flow name    Col C (3) = Scale    Col D (4) = Flow Dt
'   Col G (7) = Pres name    Col H (8) = Offset   Col I (9) = Pres Dt
'   Col J (10) = Elevation Z (m) — auto-populated from Point Index tab
' ===========================================================================
Sub SaveOneSensor(isFlow As Boolean, sRow As Long)

    Const SEL_START  As Long = 3
    Const FLOW_NAME  As Long = 2   ' B
    Const FLOW_SCALE As Long = 3   ' C
    Const FLOW_DT    As Long = 4   ' D
    Const PRES_NAME  As Long = 7   ' G
    Const PRES_OFF   As Long = 8   ' H
    Const PRES_DT    As Long = 9   ' I
    Const PRES_ELEV  As Long = 10  ' J

    Dim wsDash  As Worksheet
    Dim wsRaw   As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim sensorName  As String
    Dim scaleFactor As Double
    Dim offset      As Double
    Dim dt          As Long
    Dim dashRow     As Long
    Dim cellVal     As Variant
    dashRow = SEL_START + sRow - 1

    If isFlow Then
        sensorName = Trim(wsDash.Cells(dashRow, FLOW_NAME).Value)
        If sensorName = "" Then
            MsgBox "No sensor selected in Flow row " & sRow, vbExclamation
            Exit Sub
        End If
        cellVal = wsDash.Cells(dashRow, FLOW_SCALE).Value
        If IsNumeric(cellVal) Then scaleFactor = CDbl(cellVal)
        If scaleFactor = 0 Then scaleFactor = 1
        cellVal = wsDash.Cells(dashRow, FLOW_DT).Value
        If IsNumeric(cellVal) Then dt = CLng(cellVal)
        Set wsRaw = Worksheets("Raw Flow Data")
    Else
        sensorName = Trim(wsDash.Cells(dashRow, PRES_NAME).Value)
        If sensorName = "" Then
            MsgBox "No sensor selected in Pressure row " & sRow, vbExclamation
            Exit Sub
        End If
        cellVal = wsDash.Cells(dashRow, PRES_OFF).Value
        If IsNumeric(cellVal) Then offset = CDbl(cellVal)
        cellVal = wsDash.Cells(dashRow, PRES_DT).Value
        If IsNumeric(cellVal) Then dt = CLng(cellVal)
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

    ' --- Read entire sensor column into memory (one round-trip) ---
    Dim dataRows As Long: dataRows = lastRow - 1   ' rows 2..lastRow
    Dim srcArr  As Variant
    Dim outArr  As Variant
    srcArr = wsRaw.Range(wsRaw.Cells(2, sensorCol), _
                         wsRaw.Cells(lastRow, sensorCol)).Value  ' 2-D array (n,1)
    outArr = srcArr   ' copy preserves original; we overwrite valid entries below

    ' --- Process in memory ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim totalSaved As Long: totalSaved = 0
    Dim rawVal As Variant
    Dim srcIdx As Long
    Dim j As Long

    For j = 1 To dataRows             ' outArr is 1-based
        srcIdx = j - dt               ' dt shift (array-relative)
        If srcIdx >= 1 And srcIdx <= dataRows Then
            rawVal = srcArr(srcIdx, 1)
            If IsNumeric(rawVal) And CDbl(rawVal) <> -999 Then
                If isFlow Then
                    outArr(j, 1) = CDbl(rawVal) * scaleFactor
                Else
                    outArr(j, 1) = CDbl(rawVal) + offset
                End If
                totalSaved = totalSaved + 1
            End If
        End If
    Next j

    ' --- Write back and format in two round-trips (vs. 2*n previously) ---
    With wsRaw.Range(wsRaw.Cells(2, sensorCol), wsRaw.Cells(lastRow, sensorCol))
        .Value = outArr
        .NumberFormat = "0.000"
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' --- Capture old elevation from Point Index BEFORE it is overwritten ---
    Dim elevForNote As Variant
    If Not isFlow Then elevForNote = GetElevationFromPointIndex(sensorName)

    ' --- For pressure sensors: write elevation back to Point Index tab ---
    If Not isFlow Then
        SaveElevationToPointIndex sRow
        ' If +Z toggle is ON, re-apply elevation to this sensor's formula-table
        ' column so the chart reflects the new offset / time-shift immediately.
        RefreshElevatedColumnIfOn sRow, wsDash, dashRow
    End If

    ' --- Mark the save button green and add a timestamped note ---
    Dim adjParam As Double
    adjParam = IIf(isFlow, scaleFactor, offset)
    MarkSaved isFlow, sRow, sensorName, adjParam, dt

    ' --- Write compact data note to Point Index tab col K ---
    Dim adjLabel As String
    Dim zPart As String
    If isFlow Then
        adjLabel = "Scale: " & Format(adjParam, "0.000")
        zPart = ""
    Else
        adjLabel = "Offset: " & Format(adjParam, "0.000")
        If IsNumeric(elevForNote) Then
            zPart = " | Z: " & Format(CDbl(elevForNote), "0.###")
        Else
            zPart = ""
        End If
    End If
    Dim piNote As String
    piNote = Format(Now, "dd/mm/yyyy HH:mm") & _
             zPart & " | " & adjLabel & " | Dt: " & dt
    WriteDataNoteToPointIndex sensorName, piNote

    MsgBox "Saved " & totalSaved & " values for '" & sensorName & _
           "' into " & wsRaw.Name & ".", vbInformation, "Save Complete"
End Sub

' ===========================================================================
' MarkSaved  - highlights the 💾 save-button cell green and attaches a note
'              recording the save timestamp, sensor name, and parameters used.
'              Called by SaveOneSensor after a successful save.
' isFlow    : True = flow (col E button), False = pressure (col K button)
' sRow      : selector row index 1-20
' sensorName: name written to raw sheet
' adjVal    : scaleFactor (flow) or offset (pressure) actually applied
' dt        : time-step shift applied
' ===========================================================================
Sub MarkSaved(isFlow As Boolean, sRow As Long, sensorName As String, _
              adjVal As Double, dt As Long)

    Const SEL_START As Long = 3
    Const FLOW_SAVE As Long = 5   ' E
    Const PRES_SAVE As Long = 11  ' K

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim saveCol As Long
    Dim adjLabel As String
    If isFlow Then
        saveCol   = FLOW_SAVE
        adjLabel  = "Scale: " & Format(adjVal, "0.000")
    Else
        saveCol   = PRES_SAVE
        adjLabel  = "Offset: " & Format(adjVal, "0.000")
    End If

    Dim saveCell As Range
    Set saveCell = wsDash.Cells(SEL_START + sRow - 1, saveCol)

    ' Green fill to show the row has been saved
    saveCell.Interior.Color = RGB(198, 239, 206)

    ' Add / replace note with save details
    On Error Resume Next
    saveCell.Comment.Delete
    On Error GoTo 0

    Dim noteText As String
    noteText = "Saved: " & Format(Now, "dd/mm/yyyy HH:mm") & Chr(10) & _
               "Sensor: " & sensorName & Chr(10) & _
               adjLabel & Chr(10) & _
               "Dt: " & dt

    With saveCell.AddComment(noteText)
        .Shape.Width  = 185
        .Shape.Height = 65
    End With
End Sub

' ===========================================================================
' ClearSavedMark  - resets the 💾 save-button cell appearance when the user
'                   picks a new sensor, so the row shows its "unsaved" state:
'                     • sensor name present  → light yellow (not yet saved)
'                     • sensor name cleared  → no fill (default/blank)
'                   Also removes any existing hover-note on the button.
' isFlow : True = flow (col E button), False = pressure (col K button)
' sRow   : selector row index 1-20
' ===========================================================================
Sub ClearSavedMark(isFlow As Boolean, sRow As Long)

    Const SEL_START As Long = 3
    Const FLOW_NAME As Long = 2   ' B
    Const FLOW_SAVE As Long = 5   ' E
    Const PRES_NAME As Long = 7   ' G
    Const PRES_SAVE As Long = 11  ' K

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim dashRow As Long
    dashRow = SEL_START + sRow - 1

    Dim sensorName As String
    sensorName = Trim(wsDash.Cells(dashRow, IIf(isFlow, FLOW_NAME, PRES_NAME)).Value)

    Dim saveCell As Range
    Set saveCell = wsDash.Cells(dashRow, IIf(isFlow, FLOW_SAVE, PRES_SAVE))

    If sensorName = "" Then
        saveCell.Interior.ColorIndex = -4142   ' No sensor — restore default colour
    Else
        saveCell.Interior.Color = RGB(255, 235, 156)   ' Sensor present, not yet saved — light yellow
    End If

    On Error Resume Next
    saveCell.Comment.Delete
    On Error GoTo 0
End Sub

' ===========================================================================
' PopulateElevation  - looks up Z (m) from the "Point Index" tab for the
'                      selected pressure sensor and writes it to col J.
'                      Called automatically when a pressure sensor is chosen.
' sRow : selector row index 1-20
' ===========================================================================
Sub PopulateElevation(sRow As Long)

    Const SEL_START As Long = 3
    Const PRES_NAME As Long = 7   ' G
    Const PRES_ELEV As Long = 10  ' J

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim dashRow As Long
    dashRow = SEL_START + sRow - 1

    Dim sensorName As String
    sensorName = Trim(wsDash.Cells(dashRow, PRES_NAME).Value)

    ' Clear elevation if no sensor selected
    If sensorName = "" Then
        wsDash.Cells(dashRow, PRES_ELEV).ClearContents
        Exit Sub
    End If

    ' Find the "Point Index" worksheet
    Dim wsIdx As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Trim(ws.Name)) = "point index" Then
            Set wsIdx = ws
            Exit For
        End If
    Next ws
    If wsIdx Is Nothing Then Exit Sub

    ' Find "Point Ref" and "Z (m)" column headers
    Dim lastHdrCol As Long
    lastHdrCol = wsIdx.Cells(1, wsIdx.Columns.Count).End(xlToLeft).Column
    Dim pointRefCol As Long: pointRefCol = 0
    Dim zCol As Long: zCol = 0
    Dim c As Long
    For c = 1 To lastHdrCol
        Select Case Trim(wsIdx.Cells(1, c).Value)
            Case "Point Ref": pointRefCol = c
            Case "Z (m)":     zCol = c
        End Select
    Next c
    If pointRefCol = 0 Or zCol = 0 Then Exit Sub

    ' Look up the sensor and populate elevation
    Dim lastRow As Long
    lastRow = wsIdx.Cells(wsIdx.Rows.Count, pointRefCol).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsIdx.Cells(r, pointRefCol).Value) = sensorName Then
            wsDash.Cells(dashRow, PRES_ELEV).Value = wsIdx.Cells(r, zCol).Value
            Exit Sub
        End If
    Next r

    ' Sensor not found in Point Index — leave elevation blank
    wsDash.Cells(dashRow, PRES_ELEV).ClearContents
End Sub

' ===========================================================================
' PopulateAllElevations  - calls PopulateElevation for every pressure row
'                          that has a sensor selected.  Safe to call at any
'                          time (e.g. on sheet activate, or to refresh after
'                          the Point Index has been updated).
' ===========================================================================
Sub PopulateAllElevations()

    Const SEL_START As Long = 3
    Const SEL_END   As Long = 22
    Const PRES_NAME As Long = 7   ' G

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim r As Long
    For r = SEL_START To SEL_END
        PopulateElevation r - SEL_START + 1
    Next r

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ===========================================================================
' SaveElevationToPointIndex  - writes the elevation value in col J back to
'                              the "Point Index" tab Z (m) column.
'                              Called by SaveOneSensor for pressure sensors.
' sRow : selector row index 1-20
' ===========================================================================
Sub SaveElevationToPointIndex(sRow As Long)

    Const SEL_START As Long = 3
    Const PRES_NAME As Long = 7   ' G
    Const PRES_ELEV As Long = 10  ' J

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim dashRow As Long
    dashRow = SEL_START + sRow - 1

    Dim sensorName As String
    sensorName = Trim(wsDash.Cells(dashRow, PRES_NAME).Value)
    If sensorName = "" Then Exit Sub

    Dim elevVal As Variant
    elevVal = wsDash.Cells(dashRow, PRES_ELEV).Value
    If Not IsNumeric(elevVal) Then Exit Sub

    ' Find the "Point Index" worksheet
    Dim wsIdx As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Trim(ws.Name)) = "point index" Then
            Set wsIdx = ws
            Exit For
        End If
    Next ws
    If wsIdx Is Nothing Then
        MsgBox "Worksheet 'Point Index' not found.", vbExclamation
        Exit Sub
    End If

    ' Find "Point Ref" and "Z (m)" column headers
    Dim lastHdrCol As Long
    lastHdrCol = wsIdx.Cells(1, wsIdx.Columns.Count).End(xlToLeft).Column
    Dim pointRefCol As Long: pointRefCol = 0
    Dim zCol As Long: zCol = 0
    Dim c As Long
    For c = 1 To lastHdrCol
        Select Case Trim(wsIdx.Cells(1, c).Value)
            Case "Point Ref": pointRefCol = c
            Case "Z (m)":     zCol = c
        End Select
    Next c
    If pointRefCol = 0 Or zCol = 0 Then
        MsgBox "Could not find 'Point Ref' or 'Z (m)' column in 'Point Index'.", vbExclamation
        Exit Sub
    End If

    ' Find and update the matching row
    Dim lastRow As Long
    lastRow = wsIdx.Cells(wsIdx.Rows.Count, pointRefCol).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsIdx.Cells(r, pointRefCol).Value) = sensorName Then
            wsIdx.Cells(r, zCol).Value = CDbl(elevVal)
            Exit Sub
        End If
    Next r

    MsgBox "Sensor '" & sensorName & "' not found in 'Point Index'.", vbExclamation
End Sub

' ===========================================================================
' GetElevationFromPointIndex  - returns the current Z (m) value stored in the
'                               "Point Index" tab for the given sensor, or Empty
'                               if the sensor / sheet / column cannot be found.
'                               Used to record the OLD elevation in the save note
'                               before SaveElevationToPointIndex overwrites it.
' sensorName : sensor name to look up in the "Point Ref" column
' ===========================================================================
Function GetElevationFromPointIndex(sensorName As String) As Variant

    GetElevationFromPointIndex = Empty

    Dim wsIdx As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Trim(ws.Name)) = "point index" Then
            Set wsIdx = ws
            Exit For
        End If
    Next ws
    If wsIdx Is Nothing Then Exit Function

    Dim lastHdrCol As Long
    lastHdrCol = wsIdx.Cells(1, wsIdx.Columns.Count).End(xlToLeft).Column
    Dim pointRefCol As Long: pointRefCol = 0
    Dim zCol As Long: zCol = 0
    Dim c As Long
    For c = 1 To lastHdrCol
        Select Case Trim(wsIdx.Cells(1, c).Value)
            Case "Point Ref": pointRefCol = c
            Case "Z (m)":     zCol = c
        End Select
    Next c
    If pointRefCol = 0 Or zCol = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = wsIdx.Cells(wsIdx.Rows.Count, pointRefCol).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsIdx.Cells(r, pointRefCol).Value) = sensorName Then
            GetElevationFromPointIndex = wsIdx.Cells(r, zCol).Value
            Exit Function
        End If
    Next r

End Function

' ===========================================================================
' WriteDataNoteToPointIndex  - appends a compact single-line save note to the
'                              "Data Notes" column of the matching row in Point
'                              Index. The column is found by header name; if it
'                              does not exist it is created after the last column.
' sensorName : sensor name to look up in the "Point Ref" column
' noteText   : compact note string, e.g. "26/03/2026 14:35 | Z: 33.95 | Offset: 10.000 | Dt: 4"
' ===========================================================================
Sub WriteDataNoteToPointIndex(sensorName As String, noteText As String)

    ' Find the "Point Index" worksheet
    Dim wsIdx As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Trim(ws.Name)) = "point index" Then
            Set wsIdx = ws
            Exit For
        End If
    Next ws
    If wsIdx Is Nothing Then
        MsgBox "Worksheet 'Point Index' not found.", vbExclamation
        Exit Sub
    End If

    ' Find "Point Ref" and "Data Notes" column headers dynamically
    Dim lastHdrCol As Long
    lastHdrCol = wsIdx.Cells(1, wsIdx.Columns.Count).End(xlToLeft).Column
    Dim pointRefCol  As Long: pointRefCol  = 0
    Dim dataNotesCol As Long: dataNotesCol = 0
    Dim c As Long
    For c = 1 To lastHdrCol
        Select Case Trim(wsIdx.Cells(1, c).Value)
            Case "Point Ref":  pointRefCol  = c
            Case "Data Notes": dataNotesCol = c
        End Select
    Next c
    If pointRefCol = 0 Then
        MsgBox "Could not find 'Point Ref' column in 'Point Index'.", vbExclamation
        Exit Sub
    End If

    ' If "Data Notes" column doesn't exist yet, create it after last used column
    If dataNotesCol = 0 Then
        dataNotesCol = lastHdrCol + 1
        wsIdx.Cells(1, dataNotesCol).Value = "Data Notes"
    End If

    ' Find the matching sensor row and append the note
    Dim lastRow As Long
    lastRow = wsIdx.Cells(wsIdx.Rows.Count, pointRefCol).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsIdx.Cells(r, pointRefCol).Value) = sensorName Then
            Dim existingText As String
            existingText = wsIdx.Cells(r, dataNotesCol).Value
            If existingText = "" Then
                wsIdx.Cells(r, dataNotesCol).Value = noteText
            Else
                wsIdx.Cells(r, dataNotesCol).Value = existingText & Chr(10) & Chr(10) & noteText
            End If
            wsIdx.Cells(r, dataNotesCol).WrapText = True
            Exit Sub
        End If
    Next r

    ' Sensor not found in Point Index
    MsgBox "Sensor '" & sensorName & "' not found in 'Point Index'." & Chr(10) & _
           "Not in Point Index", vbExclamation, "Not in Point Index"
End Sub

' ===========================================================================
' RefreshElevatedColumnIfOn  - if the +Z toggle is currently ON, refreshes
'                               the formula-table column for one pressure
'                               sensor by:
'                                 1. temporarily restoring its pressure formula
'                                 2. forcing recalculation
'                                 3. adding the Z(m) elevation offset back
'                                 4. writing the result as static values
'
'                               Called after a sensor is saved (💾) or when
'                               the user edits Offset / Δt while +Z is ON so
'                               the chart updates immediately.
' sRow    : selector row index 1-20
' wsDash  : Dashboard worksheet reference
' dashRow : absolute row on Dashboard (SEL_START + sRow - 1)
' ===========================================================================
Sub RefreshElevatedColumnIfOn(sRow As Long, wsDash As Worksheet, dashRow As Long)

    Const TOGGLE_COL    As Long = 13  ' M
    Const TOGGLE_ROW    As Long = 7
    Const FT_START_ROW  As Long = 26  ' first formula-table data row
    Const PRES_FT_FIRST As Long = 22  ' V = column 22
    Const PRES_ELEV     As Long = 10  ' J

    Dim toggleState As String
    toggleState = Trim(CStr(wsDash.Cells(TOGGLE_ROW, TOGGLE_COL).Value))
    If LCase(toggleState) <> "+z on" Then Exit Sub

    Dim elevVal As Variant
    elevVal = wsDash.Cells(dashRow, PRES_ELEV).Value
    If Not IsNumeric(elevVal) Then Exit Sub
    Dim elevZ As Double: elevZ = CDbl(elevVal)

    Dim lastFTRow As Long
    lastFTRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row
    If lastFTRow < FT_START_ROW Then Exit Sub

    Dim ftCol As Long
    ftCol = PRES_FT_FIRST + (sRow - 1)

    Dim ftrng As Range
    Set ftrng = wsDash.Range(wsDash.Cells(FT_START_ROW, ftCol), _
                             wsDash.Cells(lastFTRow, ftCol))

    ' Temporarily restore the formula so it recalculates against current raw data
    Dim q As String: q = Chr(34)
    Dim n As String: n = CStr(dashRow)
    ftrng.Formula = "=IFERROR(IF($G$" & n & "=" & q & q & ",NA()," & _
        "IF($N$5+ROW()-26>$N$6,NA()," & _
        "IF(INDEX('Raw Pressure Data'!$A:$ZZ," & _
        "$N$5+ROW()-26-$I$" & n & "," & _
        "MATCH($G$" & n & ",'Raw Pressure Data'!$1:$1,0))=-999,NA()," & _
        "INDEX('Raw Pressure Data'!$A:$ZZ," & _
        "$N$5+ROW()-26-$I$" & n & "," & _
        "MATCH($G$" & n & ",'Raw Pressure Data'!$1:$1,0))" & _
        "+$H$" & n & "))),NA())"

    Application.Calculate

    ' Read recalculated values, add elevation, write back as static values
    Dim arr As Variant
    arr = ftrng.Value
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If IsNumeric(arr(i, 1)) Then
            arr(i, 1) = CDbl(arr(i, 1)) + elevZ
        End If
    Next i
    ftrng.Value = arr
End Sub

' ===========================================================================
' ToggleElevationAdjust  - adds or removes the elevation Z (m) offset from
'                          all active pressure columns in the formula table
'                          for chart display purposes only.
'
' Toggle state is stored as the value of cell M7 on the Dashboard:
'   "+Z OFF"  -> elevation not applied (normal pressure display)
'   "+Z ON"   -> elevation applied (pressure + Z shown on chart)
'
' When turning ON : formula-table cells are overwritten with value + elevation.
' When turning OFF: the original pressure formulas are rebuilt from the
'                   known template so the chart returns to base values.
'
' NOTE: While the toggle is ON the pressure columns contain static values.
'       Saving a sensor (💾) or editing its Offset / Δt will automatically
'       refresh the column via RefreshElevatedColumnIfOn.
' ===========================================================================
Sub ToggleElevationAdjust()

    Const SEL_START     As Long = 3
    Const SEL_END       As Long = 22
    Const PRES_NAME     As Long = 7   ' G
    Const PRES_ELEV     As Long = 10  ' J
    Const FT_START_ROW  As Long = 26  ' first formula-table data row
    Const PRES_FT_FIRST As Long = 22  ' V  — first pressure formula-table column
    Const TOGGLE_COL    As Long = 13  ' M
    Const TOGGLE_ROW    As Long = 7   ' row of the toggle button (M7)

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    ' Determine current toggle state
    Dim currentState As String
    currentState = Trim(CStr(wsDash.Cells(TOGGLE_ROW, TOGGLE_COL).Value))
    Dim turningOn As Boolean
    turningOn = (LCase(currentState) <> "+z on")

    ' Find last formula-table data row (date in col A)
    Dim lastFTRow As Long
    lastFTRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row
    If lastFTRow < FT_START_ROW Then
        MsgBox "No formula table data found.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim pIdx As Long   ' 0-based pressure sensor index (0 = sensor 1 in row 3)
    For pIdx = 0 To (SEL_END - SEL_START)
        Dim dashRow As Long
        dashRow = SEL_START + pIdx

        Dim sensorName As String
        sensorName = Trim(wsDash.Cells(dashRow, PRES_NAME).Value)
        If sensorName = "" Then GoTo NextSensor

        Dim elevVal As Variant
        elevVal = wsDash.Cells(dashRow, PRES_ELEV).Value
        If Not IsNumeric(elevVal) Then GoTo NextSensor

        Dim elev As Double
        elev = CDbl(elevVal)
        If elev = 0 Then GoTo NextSensor

        Dim ftCol As Long
        ftCol = PRES_FT_FIRST + pIdx   ' V for sensor 1, W for sensor 2, ...

        Dim rng As Range
        Set rng = wsDash.Range(wsDash.Cells(FT_START_ROW, ftCol), _
                               wsDash.Cells(lastFTRow, ftCol))

        If turningOn Then
            ' Evaluate formulas, add elevation, store as static values
            Dim arr As Variant
            arr = rng.Value
            Dim i As Long
            For i = 1 To UBound(arr, 1)
                If IsNumeric(arr(i, 1)) Then
                    arr(i, 1) = CDbl(arr(i, 1)) + elev
                End If
            Next i
            rng.Value = arr
        Else
            ' Rebuild original pressure formulas from the known template.
            ' Use Chr(34) to embed double-quote chars without nested VBA string escapes.
            Dim q As String: q = Chr(34)
            Dim n As String
            n = CStr(dashRow)
            rng.Formula = "=IFERROR(IF($G$" & n & "=" & q & q & ",NA()," & _
                "IF($N$5+ROW()-26>$N$6,NA()," & _
                "IF(INDEX('Raw Pressure Data'!$A:$ZZ," & _
                "$N$5+ROW()-26-$I$" & n & "," & _
                "MATCH($G$" & n & ",'Raw Pressure Data'!$1:$1,0))=-999,NA()," & _
                "INDEX('Raw Pressure Data'!$A:$ZZ," & _
                "$N$5+ROW()-26-$I$" & n & "," & _
                "MATCH($G$" & n & ",'Raw Pressure Data'!$1:$1,0))" & _
                "+$H$" & n & "))),NA())"
        End If

NextSensor:
    Next pIdx

    ' Update toggle button label
    wsDash.Cells(TOGGLE_ROW, TOGGLE_COL).Value = IIf(turningOn, "+Z ON", "+Z OFF")

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox IIf(turningOn, _
        "Elevation added to all pressure series. Click again to remove.", _
        "Elevation removed. Pressure formulas restored."), _
        vbInformation, "Elevation Toggle"
End Sub

' ===========================================================================
' ExportOnePRN  - exports the PRN file for the sensor in one dashboard row.
'
' Dashboard layout (rows 3-22):
'   Col B (2)  = Flow name       Col F (6)  = Flow PRN button
'   Col G (7)  = Pres name       Col L (12) = Pres PRN button
'   Col M (13) = Client Name label   Col N (14) = Client Name value (row 8)
'                                    Col N (14) = Export Path value  (row 9)
'   Col M (13) = Export All PRNs button (row 10)
'
' isFlow : True  = flow row (col B name, Raw Flow Data)
'          False = pressure/depth row (col G name, Raw Pressure Data)
' sRow   : selector row index 1-20
' ===========================================================================
Sub ExportOnePRN(isFlow As Boolean, sRow As Long)

    Const SEL_START  As Long = 3
    Const FLOW_NAME  As Long = 2   ' B
    Const PRES_NAME  As Long = 7   ' G
    Const CLIENT_ROW As Long = 8   ' N8 = client name value
    Const PATH_ROW   As Long = 9   ' N9 = export path value
    Const INPUT_COL  As Long = 14  ' N

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim dashRow As Long
    dashRow = SEL_START + sRow - 1

    Dim sensorName As String
    If isFlow Then
        sensorName = Trim(wsDash.Cells(dashRow, FLOW_NAME).Value)
        If sensorName = "" Then
            MsgBox "No sensor selected in Flow row " & sRow, vbExclamation
            Exit Sub
        End If
    Else
        sensorName = Trim(wsDash.Cells(dashRow, PRES_NAME).Value)
        If sensorName = "" Then
            MsgBox "No sensor selected in Pressure row " & sRow, vbExclamation
            Exit Sub
        End If
    End If

    Dim clientName As String
    Dim exportPath As String
    clientName = Trim(CStr(wsDash.Cells(CLIENT_ROW, INPUT_COL).Value))
    exportPath  = Trim(CStr(wsDash.Cells(PATH_ROW,   INPUT_COL).Value))

    If exportPath = "" Then
        MsgBox "Please enter an Export Path in cell N9 on the Dashboard.", _
               vbExclamation, "Export Path Missing"
        Exit Sub
    End If
    If Right(exportPath, 1) <> "\\" Then exportPath = exportPath & "\\"
    If clientName = "" Then clientName = "Wessex Logger"

    Dim chanType As String
    Dim chanUnit As String
    If isFlow Then
        chanType = "flow": chanUnit = "l/s"
    Else
        Dim sType As String
        sType = GetSensorTypeFromPointIndex(sensorName)
        chanType = sType: chanUnit = "m"
    End If

    Dim wsRaw As Worksheet
    If isFlow Then
        Set wsRaw = Worksheets("Raw Flow Data")
    Else
        Set wsRaw = Worksheets("Raw Pressure Data")
    End If

    On Error GoTo ErrHandler
    WriteOnePRNFile sensorName, chanType, chanUnit, _
                    GetLoggerIDFromPointIndex(sensorName), _
                    wsRaw, clientName, exportPath
    MsgBox "Exported: " & exportPath & sensorName & ".prn", _
           vbInformation, "PRN Export"
    Exit Sub
ErrHandler:
    MsgBox "Export failed for '" & sensorName & "':" & Chr(10) & Err.Description, _
           vbExclamation, "PRN Export Error"
End Sub

' ===========================================================================
' ExportAllPRNs  - exports a PRN file for every column in both Raw data tabs.
'                  Flow columns -> type "flow".
'                  Pressure columns -> type "pressure" or "depth" per Point Index.
' ===========================================================================
Sub ExportAllPRNs()

    Const CLIENT_ROW As Long = 8
    Const PATH_ROW   As Long = 9
    Const INPUT_COL  As Long = 14  ' N

    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")

    Dim clientName As String
    Dim exportPath As String
    clientName = Trim(CStr(wsDash.Cells(CLIENT_ROW, INPUT_COL).Value))
    exportPath  = Trim(CStr(wsDash.Cells(PATH_ROW,   INPUT_COL).Value))

    If exportPath = "" Then
        MsgBox "Please enter an Export Path in cell N9 on the Dashboard.", _
               vbExclamation, "Export Path Missing"
        Exit Sub
    End If
    If Right(exportPath, 1) <> "\\" Then exportPath = exportPath & "\\"
    If clientName = "" Then clientName = "Wessex Logger"

    Dim wsFlow As Worksheet
    Dim wsPres As Worksheet
    On Error Resume Next
    Set wsFlow = Worksheets("Raw Flow Data")
    Set wsPres = Worksheets("Raw Pressure Data")
    On Error GoTo 0

    Dim exported As Long: exported = 0
    Dim errCount As Long: errCount = 0
    Dim errList  As String: errList = ""

    Application.ScreenUpdating = False

    If Not wsFlow Is Nothing Then
        Dim lastColF As Long
        lastColF = wsFlow.Cells(1, wsFlow.Columns.Count).End(xlToLeft).Column
        Dim j As Long
        For j = 2 To lastColF
            Dim sName As String
            sName = Trim(wsFlow.Cells(1, j).Value)
            If sName <> "" Then
                On Error Resume Next
                WriteOnePRNFile sName, "flow", "l/s", _
                                GetLoggerIDFromPointIndex(sName), _
                                wsFlow, clientName, exportPath
                If Err.Number <> 0 Then
                    errCount = errCount + 1
                    errList = errList & Chr(10) & "  " & sName & ": " & Err.Description
                    Err.Clear
                Else
                    exported = exported + 1
                End If
                On Error GoTo 0
            End If
        Next j
    End If

    If Not wsPres Is Nothing Then
        Dim lastColP As Long
        lastColP = wsPres.Cells(1, wsPres.Columns.Count).End(xlToLeft).Column
        Dim k As Long
        For k = 2 To lastColP
            sName = Trim(wsPres.Cells(1, k).Value)
            If sName <> "" Then
                Dim presType As String
                presType = GetSensorTypeFromPointIndex(sName)
                On Error Resume Next
                WriteOnePRNFile sName, presType, "m", _
                                GetLoggerIDFromPointIndex(sName), _
                                wsPres, clientName, exportPath
                If Err.Number <> 0 Then
                    errCount = errCount + 1
                    errList = errList & Chr(10) & "  " & sName & ": " & Err.Description
                    Err.Clear
                Else
                    exported = exported + 1
                End If
                On Error GoTo 0
            End If
        Next k
    End If

    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Export complete: " & exported & " PRN file(s) written."
    If errCount > 0 Then
        msg = msg & Chr(10) & errCount & " error(s):" & errList
        MsgBox msg, vbExclamation, "PRN Export"
    Else
        MsgBox msg, vbInformation, "PRN Export"
    End If
End Sub

' ===========================================================================
' WriteOnePRNFile  - writes one .prn file for a single sensor.
'   Silently overwrites any existing file with the same name.
'   sensorName : exact column header in wsRaw
'   chanType   : "flow", "pressure", or "depth"
'   chanUnit   : "l/s" or "m"
'   loggerID   : string, "0" if blank/unknown
'   wsRaw      : Raw Flow Data or Raw Pressure Data worksheet
'   clientName : text for the Title line
'   exportPath : folder path ending in "\\"
' ===========================================================================
Private Sub WriteOnePRNFile(sensorName As String, chanType As String, _
                             chanUnit As String, loggerID As String, _
                             wsRaw As Worksheet, clientName As String, _
                             exportPath As String)

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
        Err.Raise vbObjectError + 1, "WriteOnePRNFile", _
                  "Sensor '" & sensorName & "' not found in " & wsRaw.Name
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    Dim dataRows As Long: dataRows = lastRow - 1
    If dataRows < 1 Then Exit Sub

    Dim firstDate As Variant
    firstDate = wsRaw.Cells(2, 1).Value
    Dim dateStr As String: dateStr = Format(Now, "DD/MM/YY")
    Dim timeStr As String: timeStr = "00:00"
    If IsDate(firstDate) Then
        dateStr = Format(CDate(firstDate), "DD/MM/YY")
        timeStr = Format(CDate(firstDate), "HH:MM")
    End If
    If loggerID = "" Then loggerID = "0"

    Dim srcArr As Variant
    srcArr = wsRaw.Range(wsRaw.Cells(2, sensorCol), _
                          wsRaw.Cells(lastRow, sensorCol)).Value

    Dim filePath As String
    filePath = exportPath & sensorName & ".prn"
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum

    Dim q As String: q = Chr(34)
    Print #fileNum, q & "text" & q & "," & q & "Title:" & clientName & ": ID - " & sensorName & q
    Print #fileNum, q & "text" & q & "," & q & "Site:0" & q
    Print #fileNum, q & "text" & q & "," & q & "Logger:" & loggerID & q
    Print #fileNum, q & "ch" & q & ",1," & q & chanType & q & "," & q & chanUnit & q
    Print #fileNum, q & "time" & q & "," & dateStr & "," & timeStr & ",15," & q & "min" & q & "," & dataRows

    Dim i As Long
    Dim rawVal As Variant
    Dim outVal As Double
    Dim valStr As String
    Dim padLen As Long
    For i = 1 To dataRows
        rawVal = srcArr(i, 1)
        If IsNumeric(rawVal) Then
            outVal = IIf(CDbl(rawVal) = -999, -99, CDbl(rawVal))
        Else
            outVal = -99
        End If
        valStr = Format(outVal, "0.000")
        padLen = 9 - Len(valStr)
        If padLen > 0 Then valStr = String(padLen, " ") & valStr
        Print #fileNum, valStr
    Next i

    Close #fileNum
End Sub

' ===========================================================================
' GetLoggerIDFromPointIndex  - returns the "Logger ID" for a sensor,
'                               or "0" if the column / sensor is not found.
' ===========================================================================
Function GetLoggerIDFromPointIndex(sensorName As String) As String

    GetLoggerIDFromPointIndex = "0"

    Dim wsIdx As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Trim(ws.Name)) = "point index" Then
            Set wsIdx = ws
            Exit For
        End If
    Next ws
    If wsIdx Is Nothing Then Exit Function

    Dim lastHdrCol As Long
    lastHdrCol = wsIdx.Cells(1, wsIdx.Columns.Count).End(xlToLeft).Column
    Dim pointRefCol As Long: pointRefCol = 0
    Dim loggerCol   As Long: loggerCol   = 0
    Dim c As Long
    For c = 1 To lastHdrCol
        Select Case LCase(Trim(wsIdx.Cells(1, c).Value))
            Case "point ref": pointRefCol = c
            Case "logger id": loggerCol   = c
        End Select
    Next c
    If pointRefCol = 0 Or loggerCol = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = wsIdx.Cells(wsIdx.Rows.Count, pointRefCol).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsIdx.Cells(r, pointRefCol).Value) = sensorName Then
            Dim logVal As Variant
            logVal = wsIdx.Cells(r, loggerCol).Value
            If IsEmpty(logVal) Or Trim(CStr(logVal)) = "" Then
                GetLoggerIDFromPointIndex = "0"
            Else
                GetLoggerIDFromPointIndex = Trim(CStr(logVal))
            End If
            Exit Function
        End If
    Next r
End Function

' ===========================================================================
' GetSensorTypeFromPointIndex  - returns "depth" if any cell in the sensor row
'                                 in Point Index contains "depth" (partial,
'                                 case-insensitive). Returns "pressure" otherwise.
' ===========================================================================
Function GetSensorTypeFromPointIndex(sensorName As String) As String

    GetSensorTypeFromPointIndex = "pressure"

    Dim wsIdx As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Trim(ws.Name)) = "point index" Then
            Set wsIdx = ws
            Exit For
        End If
    Next ws
    If wsIdx Is Nothing Then Exit Function

    Dim lastHdrCol As Long
    lastHdrCol = wsIdx.Cells(1, wsIdx.Columns.Count).End(xlToLeft).Column
    Dim pointRefCol As Long: pointRefCol = 0
    Dim c As Long
    For c = 1 To lastHdrCol
        If LCase(Trim(wsIdx.Cells(1, c).Value)) = "point ref" Then
            pointRefCol = c
            Exit For
        End If
    Next c
    If pointRefCol = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = wsIdx.Cells(wsIdx.Rows.Count, pointRefCol).End(xlUp).Row
    Dim lastCol As Long
    lastCol = wsIdx.Cells(1, wsIdx.Columns.Count).End(xlToLeft).Column
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsIdx.Cells(r, pointRefCol).Value) = sensorName Then
            For c = 1 To lastCol
                If InStr(1, LCase(Trim(CStr(wsIdx.Cells(r, c).Value))), "depth") > 0 Then
                    GetSensorTypeFromPointIndex = "depth"
                    Exit Function
                End If
            Next c
            Exit Function
        End If
    Next r
End Function
"""

VBA_SHEET = """\
' ===========================================================================
' DASHBOARD SHEET MODULE CODE
' Paste into the Dashboard sheet module (double-click Sheet1 in Project tree)
' ===========================================================================

' Worksheet_Activate: re-populates col J elevation for every sensor that is
'                     already selected whenever the user switches to this sheet.
'                     This covers sensors that were already set when the file
'                     was opened (Worksheet_Change never fires for those).
Private Sub Worksheet_Activate()
    PopulateAllElevations
End Sub

' Worksheet_Change: auto-populates col J elevation when a pressure sensor
'                   name is selected or pasted into col G (rows 3-22).
'                   Also resets the \U0001f4be button colour when a sensor name changes.
'                   When +Z is ON, live-refreshes the chart column if Offset (H)
'                   or \u0394t (I) is edited.
'                   Handles both single-cell selection and multi-cell paste.
'                   EnableEvents is always restored even if an inner call errors.
Private Sub Worksheet_Change(ByVal Target As Range)

    Const SEL_START As Long = 3
    Const SEL_END   As Long = 22
    Const FLOW_NAME As Long = 2   ' B
    Const PRES_NAME As Long = 7   ' G
    Const PRES_OFF  As Long = 8   ' H
    Const PRES_DT   As Long = 9   ' I

    Dim cell As Range
    For Each cell In Target
        If cell.Row >= SEL_START And cell.Row <= SEL_END Then
            If cell.Column = PRES_NAME Then
                Application.EnableEvents = False
                On Error Resume Next
                ClearSavedMark False, cell.Row - SEL_START + 1
                PopulateElevation cell.Row - SEL_START + 1
                On Error GoTo 0
                Application.EnableEvents = True
            ElseIf cell.Column = FLOW_NAME Then
                Application.EnableEvents = False
                On Error Resume Next
                ClearSavedMark True, cell.Row - SEL_START + 1
                On Error GoTo 0
                Application.EnableEvents = True
            ElseIf cell.Column = PRES_OFF Or cell.Column = PRES_DT Then
                ' When +Z is ON live-refresh the elevated column for this sensor
                Application.EnableEvents = False
                On Error Resume Next
                RefreshElevatedColumnIfOn cell.Row - SEL_START + 1, Me, cell.Row
                On Error GoTo 0
                Application.EnableEvents = True
            End If
        End If
    Next cell
End Sub

' Worksheet_SelectionChange: handles the \U0001f4be click-to-save buttons, PRN export
'                            buttons (col F for flow, col L for pressure),
'                            the elevation toggle button (M7), and the
'                            Export All PRNs button (M10).
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Const FLOW_SAVE_COL   As Long = 5    ' E \u2014 flow \U0001f4be cells
    Const FLOW_PRN_COL    As Long = 6    ' F \u2014 flow PRN export cells
    Const PRES_SAVE_COL   As Long = 11   ' K \u2014 pres \U0001f4be cells
    Const PRES_PRN_COL    As Long = 12   ' L \u2014 pres PRN export cells
    Const ELEV_TOGGLE_COL As Long = 13   ' M
    Const ELEV_TOGGLE_ROW As Long = 7    ' M7 \u2014 elevation toggle button
    Const EXPORT_ALL_COL  As Long = 13   ' M
    Const EXPORT_ALL_ROW  As Long = 10   ' M10 \u2014 Export All PRNs button
    Const SEL_START       As Long = 3    ' first selector row
    Const SEL_END         As Long = 22   ' last selector row

    ' Merged-cell ranges have Count > 1; allow them through if they are a
    ' single merged area (e.g. M7:N7 for the +Z toggle, M10:N10 for All PRNs).
    ' For everything else a multi-cell selection is ignored.
    If Target.Count > 1 Then
        If Not Target.MergeCells Then Exit Sub
    End If

    ' Use the top-left cell of the selection/merge so that column/row checks
    ' work correctly regardless of which cell inside the merged area was clicked.
    Dim cell As Range
    Set cell = Target.Cells(1, 1)

    If cell.Column = FLOW_SAVE_COL And _
       cell.Row >= SEL_START And cell.Row <= SEL_END Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        SaveOneSensor True, cell.Row - SEL_START + 1

    ElseIf cell.Column = FLOW_PRN_COL And _
           cell.Row >= SEL_START And cell.Row <= SEL_END Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        ExportOnePRN True, cell.Row - SEL_START + 1

    ElseIf cell.Column = PRES_SAVE_COL And _
           cell.Row >= SEL_START And cell.Row <= SEL_END Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        SaveOneSensor False, cell.Row - SEL_START + 1

    ElseIf cell.Column = PRES_PRN_COL And _
           cell.Row >= SEL_START And cell.Row <= SEL_END Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        ExportOnePRN False, cell.Row - SEL_START + 1

    ElseIf cell.Column = ELEV_TOGGLE_COL And _
           cell.Row = ELEV_TOGGLE_ROW Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        ToggleElevationAdjust

    ElseIf cell.Column = EXPORT_ALL_COL And _
           cell.Row = EXPORT_ALL_ROW Then
        Application.EnableEvents = False
        cell.Offset(0, -1).Select
        Application.EnableEvents = True
        ExportAllPRNs
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
        "Step 8:  Use the Chart Controls panel (cols M-N, top right of the Dashboard):",
        "         \u2022 Start Date / End Date \u2014 enter dates to filter the formula table and chart.",
        "           Leave blank to show all available data.  Dates must exist in 'Raw Flow Data'.",
        "Step 9:  Each flow row (col D) and each pressure row (col I) has its own \u0394t cell.",
        "         Enter an integer to shift that series in time:",
        "         +2 = read from 2 timesteps later;  -3 = read from 3 timesteps earlier.",
        "         Use this to align sensors with different transit / delay times.",
        "Step 10: Col J (Elevation) is auto-populated from the 'Point Index' tab when a",
        "         pressure sensor name is chosen.  You can override the value directly.",
        "         Clicking the pressure \U0001f4be (col K) also writes the col J elevation back",
        "         to the 'Point Index' tab Z (m) column.",
        "Step 11: Click a \U0001f4be cell (col E for flow, col K for pressure) to save that sensor.",
        "         Scale / Offset / \u0394t are applied and the adjusted values are written directly",
        "         into the corresponding Raw tab, overwriting the original column in place.",
        "         IMPORTANT: Keep a backup of your raw data before clicking Save.",
        "Step 12: Click the '+Z OFF' / '+Z ON' button (cell M7) to toggle elevation adjustment",
        "         for all active pressure series on the chart.",
        "         When ON, the chart displays  pressure + Z (m)  for every selected sensor.",
        "         This is a display-only toggle — raw data is never modified.",
        "         NOTE: While ON, pressure columns hold static values.  Turn OFF before",
        "         making changes to raw data, then turn ON again to refresh the display.",
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
    ]:
        body(r, line); r += 1
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
        "Dashboard selector columns (rows 3-22):",
        "  Col G = Pressure sensor name   Col H = Offset   Col I = \u0394t",
        "  Col J = Elevation Z (m)        Col K = \U0001f4be Save",
        "  Col M = Chart Controls labels  Col N = Chart Controls values",
        "  M7    = '+Z OFF' / '+Z ON' elevation toggle button",
        "",
        "\U0001f4be Save buttons (col E = flow, col K = pressure):",
        "  Clicking \U0001f4be applies the current Scale / Offset / \u0394t and overwrites that",
        "  sensor's column in the Raw tab.  For pressure sensors the elevation in",
        "  col J is also written back to the 'Point Index' tab.  Keep a backup before saving.",
        "",
        "Elevation column (col J):",
        "  Auto-populated from the 'Point Index' tab (column 'Point Ref' matched to",
        "  'Z (m)') when a pressure sensor name is chosen.  Can be overridden manually.",
        "  Changes are saved to Point Index when the \U0001f4be button is clicked.",
        "",
        "Elevation toggle (cell M7):",
        "  Click '+Z OFF' to add Z (m) to all pressure series on the chart (display only).",
        "  Click '+Z ON' to restore the original pressure formulas.",
        "  Raw data is never modified by this toggle.",
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
    Only acts on the original Save Rest shared-string (index 30); leaves
    all other K7 content untouched to support idempotent re-runs.
    """
    # Replace the specific original pattern:
    #   <c r="K7" s="74" t="s"><v>30</v></c>  (shared string 30 = "💾 Save Rest")
    # with the empty self-closing form:
    #   <c r="K7" s="74"/>
    return xml.replace(
        '<c r="K7" s="74" t="s"><v>30</v></c>',
        '<c r="K7" s="74"/>',
    )


def _z_elev_formula(r):
    """Return the INDEX-MATCH Excel formula that looks up the Z (m) elevation
    for the pressure sensor in dashboard row *r* from the Point Index sheet.
    Using dynamic column lookup so it is robust to Point Index column changes.
    """
    # Note: OOXML <f> elements must NOT include the leading '=' sign.
    return (
        "IFERROR(INDEX('Point Index'!$A:$K,"
        f"MATCH($G{r},'Point Index'!$A:$A,0),"
        "MATCH(\"Z (m)\",'Point Index'!$1:$1,0)),\"\")"
    )


def _add_elevation_column(xml):
    """
    Modify the Dashboard sheet XML to add the elevation column (col J) and
    shift the pressure save button from J to K, with chart-controls moving
    from K/L to L/M.

    Changes applied:
      1. Replace $L$5 / $L$6 formula refs → $M$5 / $M$6 throughout.
      2. Column widths: J wider (8), K narrow (5), L=14, 13-42 stays 13.
      3. Row 1: extend title merge and spans to include M1.
      4. Header row 2: J2=Z(m), K2=💾, L2=Chart Controls.
      5. Rows 3-4: J→elevation, K→💾, chart-label K→L, date-value L→M.
      6. Rows 5-6: J→elevation, K→💾, label K→L, formula L→M (updated refs).
      7. Row 7: J→elevation, K→💾, L→toggle "+Z OFF".
      8. Rows 8-22: J→elevation, add K→💾.
      9. Update merge cells (A1:L1→A1:M1, K2:L2→L2:M2, remove K7:L7).
    """
    # ── 1. Global formula reference update ──────────────────────────────────
    xml = xml.replace("$L$5", "$M$5").replace("$L$6", "$M$6")

    # ── 2. Column widths ─────────────────────────────────────────────────────
    xml = xml.replace(
        '<col min="10" max="10" width="5" customWidth="1"/>',
        '<col min="10" max="10" width="8" customWidth="1"/>',
    )
    xml = xml.replace(
        '<col min="11" max="11" width="14" customWidth="1"/>',
        '<col min="11" max="11" width="5" customWidth="1"/>'
        '<col min="12" max="12" width="14" customWidth="1"/>',
    )
    xml = xml.replace(
        '<col min="12" max="42" width="13" customWidth="1"/>',
        '<col min="13" max="42" width="13" customWidth="1"/>',
    )

    # ── 3. Row 1: extend title to cover new M column ──────────────────────────
    xml = re.sub(
        r'<row r="1" spans="1:12"',
        '<row r="1" spans="1:13"',
        xml,
        count=1,
    )
    # Add M1 only when it does not already follow L1 (idempotent)
    xml = re.sub(
        r'<c r="L1" s="73"/>(?!<c r="M1")',
        '<c r="L1" s="73"/><c r="M1" s="73"/>',
        xml,
    )

    # ── 4. Header row 2 ──────────────────────────────────────────────────────
    # J2: 💾(22) → Z(m)(1351), same style s=7
    # K2: Chart Controls(25) s=74 → 💾(22) s=7
    # L2: empty s=73 → Chart Controls(25) s=74
    xml = xml.replace(
        '<c r="J2" s="7" t="s"><v>22</v></c>'
        '<c r="K2" s="74" t="s"><v>25</v></c>'
        '<c r="L2" s="73"/>',
        '<c r="J2" s="7" t="s"><v>1351</v></c>'
        '<c r="K2" s="7" t="s"><v>22</v></c>'
        '<c r="L2" s="74" t="s"><v>25</v></c>',
    )

    # ── 5. Rows 3-4: chart-label K→L, date-value L→M ─────────────────────────
    xml = xml.replace(
        '<c r="J3" s="16" t="s"><v>22</v></c>'
        '<c r="K3" s="17" t="s"><v>26</v></c>'
        '<c r="L3" s="18"><v>46056</v></c>',
        '<c r="J3" s="15"><f>' + _z_elev_formula(3) + '</f></c>'
        '<c r="K3" s="16" t="s"><v>22</v></c>'
        '<c r="L3" s="17" t="s"><v>26</v></c>'
        '<c r="M3" s="18"><v>46056</v></c>',
    )
    xml = xml.replace(
        '<c r="J4" s="16" t="s"><v>22</v></c>'
        '<c r="K4" s="17" t="s"><v>27</v></c>'
        '<c r="L4" s="18"><v>46058</v></c>',
        '<c r="J4" s="15"><f>' + _z_elev_formula(4) + '</f></c>'
        '<c r="K4" s="16" t="s"><v>22</v></c>'
        '<c r="L4" s="17" t="s"><v>27</v></c>'
        '<c r="M4" s="18"><v>46058</v></c>',
    )

    # ── 6. Rows 5-6: label K→L, formula L→M (updated $M$3/$M$4) ─────────────
    m5_formula = (
        "IF($M$3=\"\",2,IFERROR(MATCH($M$3,"
        "'Raw Flow Data'!$A$2:$A$50001,1)+1,2))"
    )
    xml = re.sub(
        r'<c r="J5" s="16" t="s"><v>22</v></c>'
        r'<c r="K5" s="23" t="s"><v>28</v></c>'
        r'<c r="L5" s="24">.*?</c>',
        (
            '<c r="J5" s="15"><f>' + _z_elev_formula(5) + '</f></c>'
            '<c r="K5" s="16" t="s"><v>22</v></c>'
            '<c r="L5" s="23" t="s"><v>28</v></c>'
            f'<c r="M5" s="24"><f>{m5_formula}</f><v>2</v></c>'
        ),
        xml,
        flags=re.DOTALL,
    )
    m6_formula = (
        "IF($M$4=\"\",9999999,IFERROR(MATCH($M$4,"
        "'Raw Flow Data'!$A$2:$A$50001,1)+1,9999999))"
    )
    xml = re.sub(
        r'<c r="J6" s="16" t="s"><v>22</v></c>'
        r'<c r="K6" s="23" t="s"><v>29</v></c>'
        r'<c r="L6" s="24">.*?</c>',
        (
            '<c r="J6" s="15"><f>' + _z_elev_formula(6) + '</f></c>'
            '<c r="K6" s="16" t="s"><v>22</v></c>'
            '<c r="L6" s="23" t="s"><v>29</v></c>'
            f'<c r="M6" s="24"><f>{m6_formula}</f><v>9999999</v></c>'
        ),
        xml,
        flags=re.DOTALL,
    )

    # ── 7. Row 7: handle both cleared K7 (s="74"/>) and original Save Rest ──────
    # State 1: K7 already empty (previous script run)
    xml = xml.replace(
        '<c r="J7" s="16" t="s"><v>22</v></c>'
        '<c r="K7" s="74"/>'
        '<c r="L7" s="73"/>',
        '<c r="J7" s="15"><f>' + _z_elev_formula(7) + '</f></c>'
        '<c r="K7" s="16" t="s"><v>22</v></c>'
        '<c r="L7" s="74" t="inlineStr"><is><t>+Z OFF</t></is></c>',
    )
    # State 2: K7 still has original Save Rest shared-string (fresh Excel)
    xml = re.sub(
        r'<c r="J7" s="16" t="s"><v>22</v></c>'
        r'<c r="K7" s="\d+" t="s"><v>\d+</v></c>'
        r'<c r="L7" s="\d+"/>',
        '<c r="J7" s="15"><f>' + _z_elev_formula(7) + '</f></c>'
        '<c r="K7" s="16" t="s"><v>22</v></c>'
        '<c r="L7" s="74" t="inlineStr"><is><t>+Z OFF</t></is></c>',
        xml,
    )

    # ── 8. Rows 8-22: J→elevation formula, add K=💾 ───────────────────────────
    for r in range(8, 23):
        xml = xml.replace(
            f'<c r="J{r}" s="16" t="s"><v>22</v></c>',
            f'<c r="J{r}" s="15"><f>{_z_elev_formula(r)}</f></c>'
            f'<c r="K{r}" s="16" t="s"><v>22</v></c>',
        )

    # ── 9. Spans for rows 3-6 (new M column) ─────────────────────────────────
    for r in (3, 4, 5, 6):
        xml = xml.replace(
            f'<row r="{r}" spans="1:12"',
            f'<row r="{r}" spans="1:13"',
        )

    # ── 9b. Fallback: ensure J cells have elevation formula ──────────────────
    # If rows 3-22 J cells are still self-closing blanks (previous script run),
    # replace them with formula cells.
    for r in range(3, 23):
        xml = xml.replace(
            f'<c r="J{r}" s="15"/>',
            f'<c r="J{r}" s="15"><f>{_z_elev_formula(r)}</f></c>',
        )

    # ── 10. Merge cells ───────────────────────────────────────────────────────
    # Title row: A1:L1 → A1:M1
    xml = xml.replace('<mergeCell ref="A1:L1"/>', '<mergeCell ref="A1:M1"/>')
    # Chart Controls header: K2:L2 → L2:M2
    xml = xml.replace('<mergeCell ref="K2:L2"/>', '<mergeCell ref="L2:M2"/>')
    # Save Rest button merge no longer needed (K7 and L7 are separate)
    xml = xml.replace('<mergeCell ref="K7:L7"/>', '')
    # Update merge-cell count (6 → 5)
    xml = xml.replace('<mergeCells count="6">', '<mergeCells count="5">')

    return xml


# ---------------------------------------------------------------------------
# Styles patch – add F/G divider border
# ---------------------------------------------------------------------------

def _patch_styles_add_divider_border(xml):
    """
    Add a new border (index 4) with a thick right edge in dark navy to styles.xml,
    then add two new cellXf styles (80 = s7-variant, 81 = s16-variant) that use
    this border.  These are applied to F column cells to create a visible divider
    between the Flow section and the Pressure section.

    Idempotent: skips if <borders count="5"> already present.
    """
    if '<borders count="5">' in xml:
        return xml

    new_border = (
        '<border>'
        '<left style="thin"><color auto="1"/></left>'
        '<right style="thick"><color rgb="FF1F3864"/></right>'
        '<top style="thin"><color auto="1"/></top>'
        '<bottom style="thin"><color auto="1"/></bottom>'
        '<diagonal/>'
        '</border>'
    )
    xml = xml.replace('<borders count="4">', '<borders count="5">')
    xml = xml.replace('</borders>', new_border + '</borders>', 1)

    new_xf_80 = (
        '<xf numFmtId="0" fontId="3" fillId="4" borderId="4" xfId="0"'
        ' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center"/></xf>'
    )
    new_xf_81 = (
        '<xf numFmtId="0" fontId="7" fillId="4" borderId="4" xfId="0"'
        ' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center"/></xf>'
    )
    xml = xml.replace('<cellXfs count="80">', '<cellXfs count="82">')
    xml = xml.replace('</cellXfs>', new_xf_80 + new_xf_81 + '</cellXfs>', 1)
    return xml


# ---------------------------------------------------------------------------
# PRN button layout and Point Index Logger ID
# ---------------------------------------------------------------------------

def _add_prn_buttons(xml):
    """
    Repurpose column F (was '#' row numbers) as Flow PRN export buttons and
    column L (was chart-controls labels) as Pres PRN export buttons.
    Move the chart-controls panel to M-N (rows 3-10), removing blank column M.
    Add Client Name (M8/N8), Export Path (M9/N9), and 'All PRNs' button (M10).

    Handles two input states idempotently:
      State A: fresh file output from _add_elevation_column (chart controls at L/M)
      State B: previously-processed file (chart controls at N/O, from old script run)

    Column layout after this function:
      F  (6)  = Flow PRN button — thick right border (styles 80/81)
      L  (12) = Pres PRN button
      M  (13) = Chart Controls labels
      N  (14) = Chart Controls values

    VBA constants: TOGGLE_COL 12(L)/14(N) → 13(M);  INPUT_COL 15(O) → 14(N)
    """
    # Idempotent guard: skip only when both M7 toggle and correct F-column styles
    # are already in place (all F3-F22 use s=81).
    if ('r="M7"' in xml
            and '+Z ' in xml[xml.find('r="M7"'):xml.find('r="M7"')+120]
            and '<c r="F3" s="81"' in xml):
        return xml

    # ── 1. Global formula reference normalisation ─────────────────────────────
    # Handles State A ($M$n) and State B ($O$n) → final $N$n
    xml = (xml
        .replace("$M$3", "$N$3").replace("$O$3", "$N$3")
        .replace("$M$4", "$N$4").replace("$O$4", "$N$4")
        .replace("$M$5", "$N$5").replace("$O$5", "$N$5")
        .replace("$M$6", "$N$6").replace("$O$6", "$N$6")
    )

    # ── 2. Column widths ──────────────────────────────────────────────────────
    # State A: single range 13-42 → split to M(13)=14, N+(14)=13
    xml = xml.replace(
        '<col min="13" max="42" width="13" customWidth="1"/>',
        '<col min="13" max="13" width="14" customWidth="1"/>'
        '<col min="14" max="42" width="13" customWidth="1"/>',
    )
    # State B: already 3-part (13=13, 14=14, 15-42=13) → collapse to 2-part
    xml = xml.replace(
        '<col min="13" max="13" width="13" customWidth="1"/>'
        '<col min="14" max="14" width="14" customWidth="1"/>'
        '<col min="15" max="42" width="13" customWidth="1"/>',
        '<col min="13" max="13" width="14" customWidth="1"/>'
        '<col min="14" max="42" width="13" customWidth="1"/>',
    )

    # ── 3. Row 2: F2 header → PRN with divider; add M2 Chart Controls ─────────
    # State A: F2 has "#" shared string
    xml = xml.replace(
        '<c r="F2" s="7" t="s"><v>18</v></c>',
        '<c r="F2" s="80" t="inlineStr"><is><t>PRN</t></is></c>',
    )
    # State B: F2 already "PRN" inlineStr with s=7 → update style to 80
    xml = xml.replace(
        '<c r="F2" s="7" t="inlineStr"><is><t>PRN</t></is></c>',
        '<c r="F2" s="80" t="inlineStr"><is><t>PRN</t></is></c>',
    )
    # State A: L2 has "Chart Controls" shared string → L2=PRN + M2=Chart Controls
    xml = xml.replace(
        '<c r="L2" s="74" t="s"><v>25</v></c>',
        '<c r="L2" s="7" t="inlineStr"><is><t>PRN</t></is></c>'
        '<c r="M2" s="74" t="inlineStr"><is><t>Chart Controls</t></is></c>',
    )
    # State B: N2 already has Chart Controls inlineStr → rename to M2
    xml = xml.replace(
        '<c r="N2" s="74" t="inlineStr"><is><t>Chart Controls</t></is></c>',
        '<c r="M2" s="74" t="inlineStr"><is><t>Chart Controls</t></is></c>',
    )
    xml = xml.replace('<row r="2" spans="1:12"', '<row r="2" spans="1:14"')

    # ── 4. Rows 3-4: Start/End date labels and values ─────────────────────────
    for r, ss_label, date_val in [(3, 26, 46056), (4, 27, 46058)]:
        # State A: L{r}=label(shared-str), M{r}=date → L=PRN, M=label, N=date
        xml = xml.replace(
            f'<c r="L{r}" s="17" t="s"><v>{ss_label}</v></c>'
            f'<c r="M{r}" s="18"><v>{date_val}</v></c>',
            f'<c r="L{r}" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
            f'<c r="M{r}" s="17" t="s"><v>{ss_label}</v></c>'
            f'<c r="N{r}" s="18"><v>{date_val}</v></c>',
        )
        # State B: N{r}=label, O{r}=date → M=label, N=date
        xml = xml.replace(
            f'<c r="N{r}" s="17" t="s"><v>{ss_label}</v></c>'
            f'<c r="O{r}" s="18"><v>{date_val}</v></c>',
            f'<c r="M{r}" s="17" t="s"><v>{ss_label}</v></c>'
            f'<c r="N{r}" s="18"><v>{date_val}</v></c>',
        )
        # F3/F4: replace row number cell (State A) with PRN button
        xml = re.sub(
            rf'<c r="F{r}" s="\d+"><v>{r - 2}</v></c>',
            f'<c r="F{r}" s="81" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
            xml,
        )

    # ── 5. Row 5: * s.row: label and start-row formula ────────────────────────
    n5_formula = (
        "IF($N$3=\"\",2,IFERROR(MATCH($N$3,"
        "'Raw Flow Data'!$A$2:$A$50001,1)+1,2))"
    )
    # State A: J5=blank/formula, K5=💾, L5=label, M5=formula
    xml = re.sub(
        r'<c r="J5" s="15">(?:<f>[^<]*</f>)?</c>'
        r'<c r="K5" s="16" t="s"><v>22</v></c>'
        r'<c r="L5" s="23" t="s"><v>28</v></c>'
        r'<c r="M5" s="24">.*?</c>',
        (
            '<c r="J5" s="15"><f>' + _z_elev_formula(5) + '</f></c>'
            '<c r="K5" s="16" t="s"><v>22</v></c>'
            '<c r="L5" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
            '<c r="M5" s="23" t="s"><v>28</v></c>'
            f'<c r="N5" s="24"><f>{n5_formula}</f><v>2</v></c>'
        ),
        xml,
        flags=re.DOTALL,
    )
    # State B: N5=label, O5=formula → M5=label, N5=formula
    xml = re.sub(
        r'<c r="N5" s="23" t="s"><v>28</v></c>'
        r'<c r="O5" s="24">.*?</c>',
        (
            '<c r="M5" s="23" t="s"><v>28</v></c>'
            f'<c r="N5" s="24"><f>{n5_formula}</f><v>2</v></c>'
        ),
        xml,
        flags=re.DOTALL,
    )
    xml = re.sub(
        r'<c r="F5" s="\d+"><v>3</v></c>',
        '<c r="F5" s="81" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
        xml,
    )

    # ── 6. Row 6: * e.row: label and end-row formula ──────────────────────────
    n6_formula = (
        "IF($N$4=\"\",9999999,IFERROR(MATCH($N$4,"
        "'Raw Flow Data'!$A$2:$A$50001,1)+1,9999999))"
    )
    # State A
    xml = re.sub(
        r'<c r="J6" s="15">(?:<f>[^<]*</f>)?</c>'
        r'<c r="K6" s="16" t="s"><v>22</v></c>'
        r'<c r="L6" s="23" t="s"><v>29</v></c>'
        r'<c r="M6" s="24">.*?</c>',
        (
            '<c r="J6" s="15"><f>' + _z_elev_formula(6) + '</f></c>'
            '<c r="K6" s="16" t="s"><v>22</v></c>'
            '<c r="L6" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
            '<c r="M6" s="23" t="s"><v>29</v></c>'
            f'<c r="N6" s="24"><f>{n6_formula}</f><v>9999999</v></c>'
        ),
        xml,
        flags=re.DOTALL,
    )
    # State B
    xml = re.sub(
        r'<c r="N6" s="23" t="s"><v>29</v></c>'
        r'<c r="O6" s="24">.*?</c>',
        (
            '<c r="M6" s="23" t="s"><v>29</v></c>'
            f'<c r="N6" s="24"><f>{n6_formula}</f><v>9999999</v></c>'
        ),
        xml,
        flags=re.DOTALL,
    )
    xml = re.sub(
        r'<c r="F6" s="\d+"><v>4</v></c>',
        '<c r="F6" s="81" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
        xml,
    )

    # Update spans for rows 3-6 → 1:14
    for r in (3, 4, 5, 6):
        xml = xml.replace(f'<row r="{r}" spans="1:13"', f'<row r="{r}" spans="1:14"')
        xml = xml.replace(f'<row r="{r}" spans="1:15"', f'<row r="{r}" spans="1:14"')

    # ── 7. Row 7: toggle button → M7 (merged M7:N7), L7 → PRN ────────────────
    # State A: L7="+Z OFF/ON" toggle (from _add_elevation_column)
    for toggle_val in ('+Z OFF', '+Z ON'):
        xml = xml.replace(
            '<c r="J7" s="15"><f>' + _z_elev_formula(7) + '</f></c>'
            '<c r="K7" s="16" t="s"><v>22</v></c>'
            f'<c r="L7" s="74" t="inlineStr"><is><t>{toggle_val}</t></is></c>',
            '<c r="J7" s="15"><f>' + _z_elev_formula(7) + '</f></c>'
            '<c r="K7" s="16" t="s"><v>22</v></c>'
            '<c r="L7" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
            f'<c r="M7" s="74" t="inlineStr"><is><t>{toggle_val}</t></is></c>',
        )
    # State B: N7 has toggle → rename to M7, ensure L7=PRN exists
    for toggle_val in ('+Z OFF', '+Z ON'):
        # With L7 already PRN (either style)
        xml = re.sub(
            r'<c r="L7" s="\d+" t="inlineStr"><is><t>' + '\U0001f4c4' + r'</t></is></c>'
            + re.escape(f'<c r="N7" s="74" t="inlineStr"><is><t>{toggle_val}</t></is></c>'),
            '<c r="L7" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
            f'<c r="M7" s="74" t="inlineStr"><is><t>{toggle_val}</t></is></c>',
            xml,
        )
        # Without L7 (rare, bare N7)
        xml = xml.replace(
            f'<c r="N7" s="74" t="inlineStr"><is><t>{toggle_val}</t></is></c>',
            '<c r="L7" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
            f'<c r="M7" s="74" t="inlineStr"><is><t>{toggle_val}</t></is></c>',
        )
    # Normalize L7 to standard PRN style in case a previous partial run set it wrong
    xml = re.sub(
        r'<c r="L7" s="\d+" t="inlineStr"><is><t>' + '\U0001f4c4' + r'</t></is></c>',
        '<c r="L7" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
        xml,
    )
    xml = re.sub(
        r'<c r="F7" s="\d+"><v>5</v></c>',
        '<c r="F7" s="81" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
        xml,
    )
    xml = xml.replace('<row r="7" spans="1:12"', '<row r="7" spans="1:14"')
    xml = xml.replace('<row r="7" spans="1:14"', '<row r="7" spans="1:14"')  # noop

    # ── 8. Rows 8-22: F→PRN, ensure L→PRN ────────────────────────────────────
    for r in range(3, 23):
        val = r - 2
        xml = re.sub(
            rf'<c r="F{r}" s="\d+"><v>{val}</v></c>',
            f'<c r="F{r}" s="81" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
            xml,
        )
        # Update F style 16→81 if already a PRN button (State B re-run)
        xml = xml.replace(
            f'<c r="F{r}" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
            f'<c r="F{r}" s="81" t="inlineStr"><is><t>\U0001f4c4</t></is></c>',
        )
        if r >= 8:
            # Ensure L PRN button exists (add after K if missing)
            k_cell = f'<c r="K{r}" s="16" t="s"><v>22</v></c>'
            l_cell = f'<c r="L{r}" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
            if l_cell not in xml:
                xml = xml.replace(k_cell, k_cell + l_cell)

    # ── 9. Rows 8/9/10: Client Name, Export Path, All PRNs ───────────────────
    for r, label in [(8, 'Client Name:'), (9, 'Export Path:')]:
        # State B: N{r}=label, O{r}=input → M{r}=label, N{r}=input
        xml = xml.replace(
            f'<c r="N{r}" s="17" t="inlineStr"><is><t>{label}</t></is></c>'
            f'<c r="O{r}" s="18"/>',
            f'<c r="M{r}" s="17" t="inlineStr"><is><t>{label}</t></is></c>'
            f'<c r="N{r}" s="18"/>',
        )
        # State A: add M{r}+N{r} after L{r} (no N{r}/O{r} yet)
        l_btn = f'<c r="L{r}" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
        m_cell = f'<c r="M{r}" s="17" t="inlineStr"><is><t>{label}</t></is></c>'
        n_cell = f'<c r="N{r}" s="18"/>'
        if m_cell not in xml:
            xml = xml.replace(l_btn, l_btn + m_cell + n_cell)

    # Row 10: All PRNs
    # State B: N10 → M10
    xml = xml.replace(
        '<c r="N10" s="74" t="inlineStr"><is><t>\U0001f4c4 All PRNs</t></is></c>',
        '<c r="M10" s="74" t="inlineStr"><is><t>\U0001f4c4 All PRNs</t></is></c>',
    )
    # State A: add M10 after L10
    l10 = '<c r="L10" s="16" t="inlineStr"><is><t>\U0001f4c4</t></is></c>'
    m10 = '<c r="M10" s="74" t="inlineStr"><is><t>\U0001f4c4 All PRNs</t></is></c>'
    if m10 not in xml:
        xml = xml.replace(l10, l10 + m10)

    # ── 10. Spans for rows 7-10 ───────────────────────────────────────────────
    for r in (8, 9, 10):
        for old_span in ('1:12', '1:14', '1:15'):
            xml = xml.replace(f'<row r="{r}" spans="{old_span}"',
                              f'<row r="{r}" spans="1:14"')

    # ── 11. Merge cells ────────────────────────────────────────────────────────
    # State A: L2:M2 → M2:N2
    xml = xml.replace('<mergeCell ref="L2:M2"/>', '<mergeCell ref="M2:N2"/>')
    # State B: N2:O2 → M2:N2
    xml = xml.replace('<mergeCell ref="N2:O2"/>', '<mergeCell ref="M2:N2"/>')

    # Add M7:N7 and M10:N10 if not already present
    if '<mergeCell ref="M7:N7"/>' not in xml:
        xml = re.sub(
            r'<mergeCells count="(\d+)">',
            lambda m: f'<mergeCells count="{int(m.group(1)) + 2}">',
            xml, count=1,
        )
        xml = xml.replace(
            '</mergeCells>',
            '<mergeCell ref="M7:N7"/><mergeCell ref="M10:N10"/></mergeCells>',
            1,
        )

    return xml


def _add_point_index_logger_id(xml):
    """
    Insert a 'Logger ID' column at column B in the Point Index sheet,
    shifting the existing columns B-J to C-K.

    Applied idempotently: if the B1 cell already holds 'Logger ID' the
    function returns xml unchanged.
    """
    # Idempotent guard
    if 'Logger ID' in xml and 'r="B1"' in xml:
        b1 = re.search(r'<c r="B1"[^>]*>.*?</c>|<c r="B1"[^/]*/>', xml, re.DOTALL)
        if b1 and 'Logger ID' in b1.group():
            return xml

    # Shift column letters B→C, C→D, …, J→K in all cell r="Xn" attributes.
    # Process in reverse order to avoid double-shifting.
    for old_col, new_col in zip('JIHGFEDCB', 'KJIHGFEDC'):
        xml = re.sub(
            rf'(<c r="){old_col}(\d+")',
            lambda m, nc=new_col: f'{m.group(1)}{nc}{m.group(2)}',
            xml,
        )

    # Insert Logger ID header at B1 (after A1 = "Point Ref")
    xml = xml.replace(
        '<c r="A1" s="79" t="s"><v>529</v></c>',
        '<c r="A1" s="79" t="s"><v>529</v></c>'
        '<c r="B1" s="79" t="inlineStr"><is><t>Logger ID</t></is></c>',
    )

    # Update row spans from "1:10" to "1:11"
    xml = re.sub(r' spans="1:10"', ' spans="1:11"', xml)

    # Update column width range
    xml = xml.replace(
        '<col min="1" max="10" width="20.140625" customWidth="1"/>',
        '<col min="1" max="11" width="20.140625" customWidth="1"/>',
    )

    # Update sheet dimension (A1:Jn → A1:Kn)
    xml = re.sub(r'(<dimension ref="A1:)J(\d+")', r'\1K\2', xml)

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
                txt = _clear_save_rest_cell(data.decode("utf-8"))
                txt = _add_elevation_column(txt)
                txt = _add_prn_buttons(txt)
                data = txt.encode("utf-8")
                print(f"  patched   {name}  (elevation formula, PRN buttons, chart-controls shift)")

            elif name == "xl/styles.xml":
                txt = _patch_styles_add_divider_border(data.decode("utf-8"))
                data = txt.encode("utf-8")
                print(f"  patched   {name}  (F/G divider border added)")

            elif name == "xl/worksheets/sheet2.xml":
                txt = _add_point_index_logger_id(data.decode("utf-8"))
                data = txt.encode("utf-8")
                print(f"  patched   {name}  (Logger ID column added to Point Index)")

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
    print()
    print("New features in this build:")
    print("  \u2022 Col F (rows 3-22): PRN export button for each flow sensor (thick right border = divider).")
    print("  \u2022 Col L (rows 3-22): PRN export button for each pressure sensor.")
    print("  \u2022 Chart Controls panel shifted from L-M to M-N (blank column M removed).")
    print("  \u2022 Cell M7 \u2014 elevation toggle button (+Z OFF / +Z ON) spanning M7:N7.")
    print("  \u2022 M8/N8 \u2014 Client Name input (used in PRN title line).")
    print("  \u2022 M9/N9 \u2014 Export Path input (folder for PRN files).")
    print("  \u2022 M10  \u2014 'All PRNs' button spanning M10:N10.")
    print("  \u2022 Point Index: 'Logger ID' column added at column B.")
    print("  \u2022 Col J: elevation auto-populated via INDEX-MATCH formula (no VBA needed).")


if __name__ == "__main__":
    main()
