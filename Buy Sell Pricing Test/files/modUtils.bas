Attribute VB_Name = "modUtils"
Option Explicit

' ADO constants restated here since we use late binding (no type library reference)
Public Const adCmdStoredProc As Long = 4
Public Const adParamInput As Long = 1
Public Const adParamOutput As Long = 2
Public Const adParamInputOutput As Long = 3
Public Const adParamReturnValue As Long = 4
Public Const adStateOpen As Long = 1

' Binds a value onto a SQL parameter, matching either "@ParamName" or "ParamName".
' Passing Null explicitly sends NULL (used for the BuyDate/SellDate split).
Public Sub SetParamValue(cmd As Object, ByVal paramName As String, ByVal val As Variant)
    Dim p As Object
    On Error Resume Next
    Set p = Nothing
    Set p = cmd.Parameters("@" & paramName)
    If p Is Nothing Then Set p = cmd.Parameters(paramName)
    On Error GoTo 0

    If Not p Is Nothing Then
        If p.Direction = adParamInput Or p.Direction = adParamInputOutput Then
            If IsNull(val) Then
                p.Value = Null
            ElseIf IsEmpty(val) Or Len(Trim(CStr(val))) = 0 Then
                p.Value = Null
            Else
                p.Value = val
            End If
        End If
    End If
End Sub

Public Function FindColumn(ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(ws.Cells(1, c).Value) = Trim(headerName) Then
            FindColumn = c
            Exit Function
        End If
    Next c
    FindColumn = 0
End Function

' Reads TestCases!<headerName> for row r. Returns "" if the column or value is missing.
Public Function GetColumnValue(ws As Worksheet, ByVal r As Long, ByVal headerName As String) As Variant
    Dim c As Long
    c = FindColumn(ws, headerName)
    If c = 0 Then
        GetColumnValue = ""
    Else
        GetColumnValue = ws.Cells(r, c).Value
    End If
End Function

' Two-column key/value lookup on a sheet named sheetName (col A = key, col B = value).
' Returns the original key unchanged if no match is found or the sheet doesn't exist,
' so a missing lookup entry fails loudly at the SQL layer instead of silently as "".
Public Function LookupValue(ByVal sheetName As String, ByVal key As String) As Variant
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        LookupValue = key
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(UCase(CStr(ws.Cells(r, 1).Value))) = Trim(UCase(key)) Then
            LookupValue = ws.Cells(r, 2).Value
            Exit Function
        End If
    Next r
    LookupValue = key
End Function

' Multi-column table lookup: finds the row on sheetName where the column headed
' keyColumnHeader matches keyValue, then returns the value from the column headed
' valueColumnHeader on that same row. Used for reference tables like
' Lookup_PortfolioPricingParams or Lookup_Expenses, where one key (e.g. ReportID)
' drives many different SQL parameters, each pulling a different column.
' Returns "" if the sheet/columns/key aren't found, so a bad lookup shows up as an
' empty/NULL parameter rather than silently reusing the raw key.
Public Function LookupTableValue(ByVal sheetName As String, ByVal keyColumnHeader As String, ByVal keyValue As Variant, ByVal valueColumnHeader As String) As Variant
    Dim ws As Worksheet, keyCol As Long, valCol As Long, r As Long, lastRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        LookupTableValue = ""
        Exit Function
    End If

    keyCol = FindColumn(ws, keyColumnHeader)
    valCol = FindColumn(ws, valueColumnHeader)
    If keyCol = 0 Or valCol = 0 Then
        LookupTableValue = ""
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(UCase(CStr(ws.Cells(r, keyCol).Value))) = Trim(UCase(CStr(keyValue))) Then
            LookupTableValue = ws.Cells(r, valCol).Value
            Exit Function
        End If
    Next r
    LookupTableValue = ""
End Function

' Finds a header on row 1 of ws, or appends a new column with that header if it
' doesn't exist yet. Used to build the Results sheet's Out_<field> columns on the fly.
Public Function GetOrCreateColumn(ws As Worksheet, ByVal header As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If ws.Cells(1, c).Value = header Then
            GetOrCreateColumn = c
            Exit Function
        End If
    Next c

    lastCol = lastCol + 1
    With ws.Cells(1, lastCol)
        .Value = header
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(55, 86, 35)
        .HorizontalAlignment = -4108 ' xlCenter
    End With
    ws.Columns(lastCol).ColumnWidth = 18
    GetOrCreateColumn = lastCol
End Function

Public Function SafeVal(ByVal v As Variant) As Variant
    If IsNull(v) Then
        SafeVal = ""
    Else
        SafeVal = v
    End If
End Function

Public Sub ClearResults(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = 0
    End If
    ' Drop any dynamically-added Out_ columns from a previous run (fixed columns are 1-7)
    ' so the sheet doesn't accumulate stale columns from an older SP version.
    Dim c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = lastCol To 9 Step -1
        ws.Columns(c).Delete
    Next c
End Sub
