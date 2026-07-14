Attribute VB_Name = "modTestRunner"
Option Explicit

Private Const RES_COL_TESTID As Long = 1
Private Const RES_COL_SP As Long = 2
Private Const RES_COL_TS As Long = 3
Private Const RES_COL_DUR As Long = 4
Private Const RES_COL_STATUS As Long = 5
Private Const RES_COL_ERR As Long = 6
Private Const RES_COL_EXPECTED As Long = 7
Private Const RES_COL_ERRNUM As Long = 8
Private Const RES_FIRST_DYNAMIC_COL As Long = 9

' Entry point - Alt+F8 or wire to a button on TestCases.
Public Sub RunAllTests()
    Dim wsTC As Worksheet, wsRes As Worksheet, wsMap As Worksheet
    Dim conn As Object
    Dim lastRow As Long, r As Long, resultRow As Long
    Dim colEnabled As Long, colTestID As Long

    Set wsTC = ThisWorkbook.Sheets("TestCases")
    Set wsRes = ThisWorkbook.Sheets("Results")
    Set wsMap = ThisWorkbook.Sheets("ParamMap")

    ClearResults wsRes

    Set conn = GetDBConnection()
    If conn Is Nothing Then Exit Sub

    colTestID = FindColumn(wsTC, "Test Case ID")
    colEnabled = FindColumn(wsTC, "Enabled")
    lastRow = wsTC.Cells(wsTC.Rows.Count, colTestID).End(xlUp).Row

    Application.ScreenUpdating = False
    resultRow = 2

    For r = 2 To lastRow
        If colEnabled = 0 Or UCase(Trim(wsTC.Cells(r, colEnabled).Value)) = "Y" Then
            Application.StatusBar = "Running " & wsTC.Cells(r, colTestID).Value & " ..."
            RunSingleTest conn, wsTC, wsMap, r, wsRes, resultRow
            resultRow = resultRow + 1
        End If
    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If conn.State = adStateOpen Then conn.Close
    MsgBox "Test run complete: " & (resultRow - 2) & " test(s) executed. See Results sheet.", vbInformation
End Sub

' Runs just the TestCases row the cursor is on (ignores Enabled).
Public Sub RunSelectedTest()
    Dim wsTC As Worksheet, wsRes As Worksheet, wsMap As Worksheet
    Dim conn As Object, r As Long, nextResultRow As Long

    If ActiveSheet.Name <> "TestCases" Then
        MsgBox "Click a cell on the TestCases sheet, on the row you want to run.", vbExclamation
        Exit Sub
    End If
    r = ActiveCell.Row
    If r < 2 Then Exit Sub

    Set wsTC = ThisWorkbook.Sheets("TestCases")
    Set wsRes = ThisWorkbook.Sheets("Results")
    Set wsMap = ThisWorkbook.Sheets("ParamMap")

    Set conn = GetDBConnection()
    If conn Is Nothing Then Exit Sub

    nextResultRow = wsRes.Cells(wsRes.Rows.Count, 1).End(xlUp).Row + 1
    If nextResultRow < 2 Then nextResultRow = 2

    RunSingleTest conn, wsTC, wsMap, r, wsRes, nextResultRow
    If conn.State = adStateOpen Then conn.Close

    MsgBox "Test complete. See Results sheet.", vbInformation
End Sub

Private Sub RunSingleTest(conn As Object, wsTC As Worksheet, wsMap As Worksheet, ByVal r As Long, wsRes As Worksheet, ByVal resultRow As Long)
    Dim testID As String, spName As String, expected As String
    Dim cmd As Object, rs As Object
    Dim startTime As Double
    Dim executionOK As Boolean, status As String, errMsg As String, errNum As String
    Dim stepName As String

    testID = GetColumnValue(wsTC, r, "Test Case ID")
    expected = Trim(UCase(CStr(GetColumnValue(wsTC, r, "Expected Result"))))
    spName = ResolveStoredProcName(wsTC, r)

    startTime = Timer
    executionOK = True
    errMsg = ""
    errNum = ""
    Set rs = Nothing
    stepName = "Resolving stored procedure name"

    On Error GoTo TestErr

    If Len(spName) = 0 Then Err.Raise vbObjectError + 1, , "Could not resolve a stored procedure name for this row."

    stepName = "Creating ADO command"
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = spName
        stepName = "Refreshing SP parameter list (Parameters.Refresh)"
        .Parameters.Refresh   ' auto-discovers the SP's real parameter list from SQL Server
    End With

    stepName = "Binding parameters from ParamMap"
    BindParamsFromMap cmd, wsMap, wsTC, r

    stepName = "Executing stored procedure"
    Set rs = cmd.Execute

    ' Paste every bit of SQL output into dynamic Out_<name> columns, regardless of pass/fail,
    ' so you always see what the stored procedure actually returned. Runs BEFORE the row is
    ' written below, so a failure here still lands in the Status/ErrorMessage cells instead
    ' of being silently overwritten by an earlier "success" state.
    stepName = "Reading result sets (NextRecordset loop)"
    Dim rsCurrent As Object, rsNext As Object
    Dim rsGuard As Integer
    Set rsCurrent = rs
    rsGuard = 0
    Do While Not rsCurrent Is Nothing
        If rsCurrent.State = adStateOpen Then
            DumpRecordset wsRes, resultRow, rsCurrent
        End If
        Set rsNext = rsCurrent.NextRecordset   ' also closes rsCurrent for us
        Set rsCurrent = rsNext
        rsGuard = rsGuard + 1
        If rsGuard > 15 Then Exit Do   ' safety net against a runaway loop
    Loop

    ' Output/return parameters are only reliably populated once every result set has
    ' been fully read, so read them last.
    stepName = "Reading output/return parameters"
    DumpOutputParams wsRes, resultRow, cmd

    GoTo WriteResult

TestErr:
    Dim capturedErrNum As Long, capturedErrDesc As String
    capturedErrNum = Err.Number
    capturedErrDesc = Err.Description
    executionOK = False
    errMsg = "[" & stepName & "] " & GetFullSqlErrorText(conn, capturedErrDesc)
    errNum = CStr(capturedErrNum)
    Resume WriteResult

WriteResult:
    On Error Resume Next

    ' Decide status against Expected Result (PASS/FAIL semantics), or just log if blank.
    If expected = "PASS" Then
        status = IIf(executionOK, "PASS", "FAIL")
    ElseIf expected = "FAIL" Then
        status = IIf(Not executionOK, "PASS", "FAIL")   ' negative test: erroring IS the pass case
    Else
        status = IIf(executionOK, "SUCCESS", "ERROR")
    End If

    With wsRes
        .Cells(resultRow, RES_COL_TESTID).Value = testID
        .Cells(resultRow, RES_COL_SP).Value = spName
        .Cells(resultRow, RES_COL_TS).Value = Now
        .Cells(resultRow, RES_COL_DUR).Value = Format((Timer - startTime) * 1000, "0") & " ms"
        .Cells(resultRow, RES_COL_STATUS).Value = status
        .Cells(resultRow, RES_COL_ERR).Value = errMsg
        .Cells(resultRow, RES_COL_EXPECTED).Value = expected
        .Cells(resultRow, RES_COL_ERRNUM).Value = errNum

        Select Case status
            Case "PASS", "SUCCESS": .Cells(resultRow, RES_COL_STATUS).Interior.Color = RGB(198, 239, 206)
            Case "FAIL": .Cells(resultRow, RES_COL_STATUS).Interior.Color = RGB(255, 199, 206)
            Case "ERROR": .Cells(resultRow, RES_COL_STATUS).Interior.Color = RGB(255, 235, 156)
        End Select
    End With

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    On Error GoTo 0
End Sub

Private Function ResolveStoredProcName(wsTC As Worksheet, ByVal r As Long) As String
    Dim spName As String, pattern As String, buyOrSale As String
    spName = Trim(CStr(GetColumnValue(wsTC, r, "Stored Procedure (optional)")))
    If Len(spName) > 0 Then
        ResolveStoredProcName = spName
        Exit Function
    End If

    pattern = GetConfigValue("StoredProcPattern")
    buyOrSale = Trim(CStr(GetColumnValue(wsTC, r, "Buy or Sale")))
    ResolveStoredProcName = Replace(pattern, "{BuyOrSale}", buyOrSale)
End Function

' Walks the ParamMap sheet and binds every row's SQL parameter using the mapping rules:
' SourceType Column/Config/Fixed/Auto, optional LookupSheet value translation, and an
' optional ApplyWhenColumn/ApplyWhenValue condition (used for @BuyDate vs @SellDate).
Private Sub BindParamsFromMap(cmd As Object, wsMap As Worksheet, wsTC As Worksheet, ByVal r As Long)
    Dim lastRow As Long, m As Long
    Dim sqlParam As String, sourceType As String, source As String
    Dim lookupSheet As String, applyWhenCol As String, applyWhenVal As String
    Dim rawVal As Variant, applies As Boolean

    lastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row

    For m = 2 To lastRow
        sqlParam = Trim(wsMap.Cells(m, 1).Value)
        If Len(sqlParam) = 0 Then GoTo NextParam
        sourceType = Trim(wsMap.Cells(m, 2).Value)
        source = Trim(CStr(wsMap.Cells(m, 3).Value))
        lookupSheet = Trim(wsMap.Cells(m, 4).Value)
        Dim lookupKeyColumn As String, lookupValueColumn As String
        lookupKeyColumn = Trim(wsMap.Cells(m, 5).Value)
        lookupValueColumn = Trim(wsMap.Cells(m, 6).Value)
        applyWhenCol = Trim(wsMap.Cells(m, 7).Value)
        applyWhenVal = Trim(wsMap.Cells(m, 8).Value)

        applies = True
        If Len(applyWhenCol) > 0 Then
            Dim actualCond As String
            actualCond = Trim(UCase(CStr(GetColumnValue(wsTC, r, applyWhenCol))))
            If actualCond <> Trim(UCase(applyWhenVal)) Then applies = False
        End If

        If Not applies Then
            SetParamValue cmd, Mid(sqlParam, 2), Null
            GoTo NextParam
        End If

        Select Case UCase(sourceType)
            Case "COLUMN"
                rawVal = GetColumnValue(wsTC, r, source)
                ' @Vendor-style fallback to a Config default when the sheet cell is blank
                If Len(Trim(CStr(rawVal))) = 0 Then
                    Dim fallback As String
                    fallback = GetConfigValue("Default" & Replace(Mid(sqlParam, 2), " ", ""))
                    If Len(fallback) > 0 Then rawVal = fallback
                End If
            Case "CONFIG"
                rawVal = GetConfigValue(source)
            Case "FIXED"
                rawVal = source
            Case "AUTO"
                rawVal = GenerateAutoValue(source, r)
            Case Else
                rawVal = ""
        End Select

        If Len(lookupSheet) > 0 Then
            If Len(lookupKeyColumn) > 0 And Len(lookupValueColumn) > 0 Then
                ' Full table lookup: rawVal (from Source above) is the KEY - e.g. Report ID -
                ' find that row on lookupSheet and pull the named value column from it.
                ' Used for reference tables like Lookup_PortfolioPricingParams / Lookup_Expenses
                ' where one key drives many different SQL parameters.
                rawVal = LookupTableValue(lookupSheet, lookupKeyColumn, rawVal, lookupValueColumn)
            ElseIf Len(Trim(CStr(rawVal))) > 0 Then
                ' Simple 2-column translation (e.g. Lookup_YesNo, Lookup_CalculationType).
                rawVal = LookupValue(lookupSheet, CStr(rawVal))
            End If
        End If

        SetParamValue cmd, Mid(sqlParam, 2), rawVal
NextParam:
    Next m
End Sub

Private Function GenerateAutoValue(ByVal kind As String, ByVal r As Long) As Variant
    Select Case UCase(kind)
        Case "GUID"
            ' Not a true GUID (no extra library needed) - unique enough per test run for a @Guid tag.
            GenerateAutoValue = CLng(Format(Now, "hhnnss")) * 1000 + r
        Case Else
            GenerateAutoValue = ""
    End Select
End Function

Private Sub DumpOutputParams(wsRes As Worksheet, ByVal resultRow As Long, cmd As Object)
    Dim p As Object, outCol As Long
    For Each p In cmd.Parameters
        If p.Direction = adParamOutput Or p.Direction = adParamInputOutput Then
            outCol = GetOrCreateColumn(wsRes, "Out_" & Mid(p.Name, 2))
            wsRes.Cells(resultRow, outCol).Value = SafeVal(p.Value)
        ElseIf p.Direction = adParamReturnValue Then
            outCol = GetOrCreateColumn(wsRes, "Out_ReturnValue")
            wsRes.Cells(resultRow, outCol).Value = SafeVal(p.Value)
        End If
    Next p
End Sub

' Your pricing SPs return a long/key-value shaped result set:
'   Metric               | Value        [ | DataType ]
'   Expected loss        | 0.021045399
'   Coupon rate          | 0.0525
'   ...
' rather than one wide row. This pivots each Metric row into its own
' Out_<Metric name> column (e.g. Out_Expected loss, Out_Coupon rate, Out_Yield...),
' so every metric the SP returns shows up as a column on Results, not just PASS/FAIL.
' Some SPs (e.g. arapl_getSummaryPricing) also return a Methodology column, where the
' SAME metric name repeats under different methodologies with different values
' ("Capital Cost (Tail)" under both "Capital Consumption Approach" and "Marginal Capital
' Approach"). When a Methodology column is present, the output column is keyed on
' "Methodology - Metric" instead of just Metric, so those don't overwrite each other.
' Falls back to a plain wide dump of row 1 if the result set isn't shaped this way at all.
Private Sub DumpRecordset(wsRes As Worksheet, ByVal resultRow As Long, rs As Object)
    If rs Is Nothing Then Exit Sub
    If rs.State <> adStateOpen Then Exit Sub
    If rs.EOF Then Exit Sub

    Dim colMetric As Long, colValue As Long, colType As Long, colMethodology As Long
    Dim fld As Object, idx As Long

    colMetric = -1: colValue = -1: colType = -1: colMethodology = -1
    idx = 0
    For Each fld In rs.Fields
        Select Case True
            Case InStr(1, fld.Name, "Methodology", vbTextCompare) > 0: colMethodology = idx
            Case InStr(1, fld.Name, "Metric", vbTextCompare) > 0: colMetric = idx
            Case InStr(1, fld.Name, "Value", vbTextCompare) > 0: colValue = idx
            Case InStr(1, fld.Name, "DataType", vbTextCompare) > 0: colType = idx
        End Select
        idx = idx + 1
    Next fld

    Dim outCol As Long, metricName As String, metricVal As Variant, outHeader As String, methodologyName As String

    If colMetric >= 0 And colValue >= 0 Then
        ' Key-value shape: one output column per Metric row, looped across ALL rows.
        Do While Not rs.EOF
            metricName = Trim(CStr(rs.Fields(colMetric).Value))
            metricVal = rs.Fields(colValue).Value
            If colMethodology >= 0 Then
                methodologyName = Trim(CStr(rs.Fields(colMethodology).Value))
                outHeader = "Out_" & methodologyName & " - " & metricName
            Else
                outHeader = "Out_" & metricName
            End If
            If Len(metricName) > 0 Then
                outCol = GetOrCreateColumn(wsRes, outHeader)
                wsRes.Cells(resultRow, outCol).Value = SafeVal(metricVal)
            End If
            rs.MoveNext
        Loop
    Else
        ' Fallback: plain wide row (one column per field, row 1 only).
        For Each fld In rs.Fields
            outCol = GetOrCreateColumn(wsRes, "Out_" & fld.Name)
            wsRes.Cells(resultRow, outCol).Value = SafeVal(fld.Value)
        Next fld
        rs.MoveNext
        If Not rs.EOF Then
            outCol = GetOrCreateColumn(wsRes, "Out_Note")
            wsRes.Cells(resultRow, outCol).Value = "SP returned multiple rows - only row 1 shown here"
        End If
    End If
End Sub

' ============================================================================
' arapl_getSummaryPricing - separate test suite (own sheets, same engine)
' ============================================================================
' This SP is fixed (no Buy/Sell variant, no StoredProcPattern needed - just put
' the SP name directly in TestCases_Pricing!"Stored Procedure (optional)" on each
' row). Everything else - parameter binding, lookups, output pivoting including
' the Methodology+Metric composite keys above - reuses the exact same engine.

Public Sub RunAllTests_SummaryPricing()
    Dim wsTC As Worksheet, wsRes As Worksheet, wsMap As Worksheet
    Dim conn As Object
    Dim lastRow As Long, r As Long, resultRow As Long
    Dim colEnabled As Long, colTestID As Long

    Set wsTC = ThisWorkbook.Sheets("TestCases_Pricing")
    Set wsRes = ThisWorkbook.Sheets("Results_Pricing")
    Set wsMap = ThisWorkbook.Sheets("ParamMap_Pricing")

    ClearResults wsRes

    Set conn = GetDBConnection()
    If conn Is Nothing Then Exit Sub

    colTestID = FindColumn(wsTC, "Test Case ID")
    colEnabled = FindColumn(wsTC, "Enabled")
    lastRow = wsTC.Cells(wsTC.Rows.Count, colTestID).End(xlUp).Row

    Application.ScreenUpdating = False
    resultRow = 2

    For r = 2 To lastRow
        If colEnabled = 0 Or UCase(Trim(wsTC.Cells(r, colEnabled).Value)) = "Y" Then
            Application.StatusBar = "Running " & wsTC.Cells(r, colTestID).Value & " ..."
            RunSingleTest conn, wsTC, wsMap, r, wsRes, resultRow
            resultRow = resultRow + 1
        End If
    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If conn.State = adStateOpen Then conn.Close
    MsgBox "Test run complete: " & (resultRow - 2) & " test(s) executed. See Results_Pricing sheet.", vbInformation
End Sub

Public Sub RunSelectedTest_SummaryPricing()
    Dim wsTC As Worksheet, wsRes As Worksheet, wsMap As Worksheet
    Dim conn As Object, r As Long, nextResultRow As Long

    If ActiveSheet.Name <> "TestCases_Pricing" Then
        MsgBox "Click a cell on the TestCases_Pricing sheet, on the row you want to run.", vbExclamation
        Exit Sub
    End If
    r = ActiveCell.Row
    If r < 2 Then Exit Sub

    Set wsTC = ThisWorkbook.Sheets("TestCases_Pricing")
    Set wsRes = ThisWorkbook.Sheets("Results_Pricing")
    Set wsMap = ThisWorkbook.Sheets("ParamMap_Pricing")

    Set conn = GetDBConnection()
    If conn Is Nothing Then Exit Sub

    nextResultRow = wsRes.Cells(wsRes.Rows.Count, 1).End(xlUp).Row + 1
    If nextResultRow < 2 Then nextResultRow = 2

    RunSingleTest conn, wsTC, wsMap, r, wsRes, nextResultRow
    If conn.State = adStateOpen Then conn.Close

    MsgBox "Test complete. See Results_Pricing sheet.", vbInformation
End Sub
