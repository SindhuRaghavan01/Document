Attribute VB_Name = "M_Driver"
Option Explicit

'===================================================================================
'  M_Driver  -  STANDALONE DRIVER (lives in its own fresh .xlsm)
'
'  Driver sheet layout (row 1 = headers, data from row 2):
'   A TestCaseID        F Vendor           K IncludeExpired   P Output Location
'   B PortfolioID       G Fronting         L AEPType          Q Result (icon)
'   C Import Status     H SecurityInstr    M AEPCoTVARType
'   D Import Time       I AsOfDate         N Parameters (Ready/Missing)
'   E Run Report?       J RunOff           O Report Status
'     (Yes/No dropdown)
'
'  Columns F:M are pre-filled with defaults after Step 1, but are yours to edit
'  before running Step 3 - each row's own values are what gets sent to the
'  Report Output Template workbook's TestCases sheet.
'
'    STEP 1 - Step1_PushAllTestCasesToSQL
'             Loops every Test Case ID in 01_Portfolio, opens the Import
'             Template workbook, populates its paste-in sheets, pushes to SQL,
'             and fills in columns A:N here with results + default parameters.
'
'    STEP 3 - Step3_RunReport
'             Refreshes the Parameters Ready/Missing flag, writes one row per
'             case where Run Report? = Yes AND Parameters = Ready into the
'             Report Output Template workbook's TestCases sheet (using YOUR edited values,
'             not hardcoded defaults), then calls that workbook's own
'             MasterAutomation_OUTPUT.RunAllCases macro. Afterwards, reads its
'             ErrorLog sheet to mark each row here Completed or Failed.
'
'  INSTALL
'  1. New blank workbook -> Save As -> Excel Macro-Enabled Workbook (.xlsm)
'  2. Alt+F11 -> File -> Import File -> this .bas
'  3. Alt+F8 -> run BrowseForPaths (or open the "Settings" sheet - auto-created the
'     first time you run anything - and type/paste the 4 paths into column B
'     yourself). No VBA editing needed for this part.
'  4. Run RunEverything, or Step1_PushAllTestCasesToSQL then Step3_RunReport,
'     from Alt+F8 (or wire buttons to them)
'===================================================================================

'------------------------------- CONFIG -------------------------------------------
' No paths hardcoded here - set them on the "Settings" sheet in this workbook
' instead (auto-created the first time you run anything). Click "Browse..." next
' to each row there, or just type/paste the path into the cell.
Const SETTINGS_SHEET As String = "Settings"
Const DRIVER_SHEET As String = "Driver"
'------------------------------------------------------------------------------------

Const TC_DATA_START_ROW As Long = 6   ' data starts on row 6 of every 0x_... test-case tab

' Driver sheet column layout
Const DCOL_TESTCASEID As Long = 1
Const DCOL_PORTFOLIOID As Long = 2
Const DCOL_IMPORTSTATUS As Long = 3
Const DCOL_IMPORTTIME As Long = 4
Const DCOL_RUNREPORT As Long = 5
Const DCOL_VENDOR As Long = 6
Const DCOL_FRONTING As Long = 7
Const DCOL_SECINSTR As Long = 8
Const DCOL_ASOFDATE As Long = 9
Const DCOL_RUNOFF As Long = 10
Const DCOL_INCLUDEEXPIRED As Long = 11
Const DCOL_AEPTYPE As Long = 12
Const DCOL_AEPCOTVARTYPE As Long = 13
Const DCOL_PARAMETERS As Long = 14
Const DCOL_REPORTSTATUS As Long = 15
Const DCOL_OUTPUTLOCATION As Long = 16
Const DCOL_RESULT As Long = 17
Const DCOL_EXPORT As Long = 18   ' "Export?" Yes/No - which completed cases to include in Step 4
Const DCOL_INCLUDEINRUN As Long = 19   ' "Include in Test Run?" Yes/No - governs the ENTIRE pipeline
                                        ' (import + SQL push + report), not just the report step.
                                        ' Preserved across Step1 re-runs so it survives the Driver
                                        ' sheet being rebuilt.

' Where Step4_ExportResults saves its output workbook(s) - set on the Settings sheet,
' must already exist as a folder
' (EXPORT_FOLDER_PATH is defined as a Function further down, alongside the other 3 paths)

' The 10 AC_ sheets RunAllCases builds in the Report Output Template workbook
Private Function AC_SHEET_NAMES() As Variant
    AC_SHEET_NAMES = Array( _
        "AC_Region Peril", "AC_BoundPositions", "AC_Bound Portfolio Report", _
        "AC_PnL Report", "AC_Portfolio Metrics", _
        "AC_GS Limit by SecInstr", "AC_GS PML Summary", "AC_GS Issuer", _
        "AC_AAL Monthly", "AC_AAL Premium")
End Function

'-----------------------------------------------------------------------------------
' Derives PortfolioID directly from the Test Case ID (e.g. "TC_PF_001" -> "001"),
' rather than trusting 07_Bound_Portfolio's own PortfolioID column - guarantees a
' unique PortfolioID per test case even if that source column repeats/duplicates.
' Takes everything after the last "_", then keeps only its trailing digits
' (preserving leading zeros). Falls back to the raw Test Case ID if that pattern
' doesn't match.
'-----------------------------------------------------------------------------------
' ---- Testing-framework safety constants ----
' Real PortfolioIDs in SQL are small sequential ints (see BoundPortfolio.PortfolioID,
' type int). Test PortfolioIDs are offset into a range no real portfolio would ever
' reach, so a test run can NEVER collide with, and therefore never modify, real data.
Const TEST_PORTFOLIOID_BASE As Long = 900000000
' Prefix applied to BoundPortfolio!Name so test data is identifiable at a glance in
' SQL and can be safely filtered by a cleanup stored procedure (WHERE Name LIKE 'Test\_%').
Const TEST_NAME_PREFIX As String = "Test_"

'-----------------------------------------------------------------------------------
' Derives a PortfolioID from the Test Case ID's trailing digits (e.g. "TC_PF_001" ->
' 1), then offsets it by TEST_PORTFOLIOID_BASE (e.g. -> 900000001). The offset is
' the actual safety mechanism - it guarantees this can never equal a real
' PortfolioID, so pushing/cleaning up test data can never touch real records even
' if the underlying digits (1, 2, 3...) would otherwise collide with real IDs.
'-----------------------------------------------------------------------------------
Private Function DerivePortfolioID(tcID As String) As Long
    Dim pos As Long, i As Long, s As String, digitsOnly As String

    pos = InStrRev(tcID, "_")
    If pos > 0 And pos < Len(tcID) Then
        s = Mid$(tcID, pos + 1)
    Else
        s = tcID
    End If

    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) Like "#" Then
            digitsOnly = Mid$(s, i, 1) & digitsOnly
        Else
            Exit For
        End If
    Next i

    If digitsOnly <> "" Then
        DerivePortfolioID = 900000000 + CLng(digitsOnly)   ' 900000000 = safety offset, hardcoded directly (see comment above)
    Else
        ' No trailing digits found - fall back to a stable hash of the whole
        ' TestCaseID so it's still deterministic and still safely offset.
        Dim h As Long, ch As Long
        h = 0
        For i = 1 To Len(s)
            ch = Asc(Mid$(s, i, 1))
            h = (h * 31 + ch) Mod 1000000
        Next i
        DerivePortfolioID = 900000000 + h
    End If
End Function

'===================================================================================
' The four configurable paths, read live from the "Settings" sheet each time they're
' used - so editing a cell in Excel is all it takes, no VBA editing required.
' Named as zero-argument Functions so every existing call site (e.g.
' GetOrOpenWorkbook(TESTCASE_WORKBOOK_PATH, ...)) keeps working unchanged.
'===================================================================================
Private Function TESTCASE_WORKBOOK_PATH() As String
    TESTCASE_WORKBOOK_PATH = GetSetting_("Test Case Workbook Path")
End Function

Private Function TEMPLATE_WORKBOOK_PATH() As String
    TEMPLATE_WORKBOOK_PATH = GetSetting_("Import Template Workbook Path")
End Function

Private Function UITESTING_WORKBOOK_PATH() As String
    UITESTING_WORKBOOK_PATH = GetSetting_("Report Output Template Path")
End Function

Private Function EXPORT_FOLDER_PATH() As String
    Dim p As String
    p = GetSetting_("Export Folder Path")
    If p <> "" And Right$(p, 1) <> "\" Then p = p & "\"
    EXPORT_FOLDER_PATH = p
End Function

' Reads the "PnL Only? (Yes/No)" toggle on the Settings sheet. Defaults to False
' (run everything) if it's ever blank, so a missing/cleared cell fails safe.
Private Function IsPnLOnly() As Boolean
    Dim v As String
    On Error Resume Next
    v = UCase$(Trim$(CStr(GetOrCreateSettingsSheet(ThisWorkbook).Range("B6").Value)))
    On Error GoTo 0
    IsPnLOnly = (v = "YES")
End Function

'-----------------------------------------------------------------------------------
' Reads one labeled value from the "Settings" sheet (column A = label, column B =
' value), creating the sheet with all four rows blank if it doesn't exist yet.
' Raises a clear error if the requested setting is still blank, so you get pointed
' straight at the Settings sheet instead of a cryptic "file not found" later on.
'-----------------------------------------------------------------------------------
Private Function GetSetting_(label As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim v As String

    Set ws = GetOrCreateSettingsSheet(ThisWorkbook)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) = label Then
            v = Trim$(CStr(ws.Cells(r, 2).Value))
            If v = "" Then
                Err.Raise vbObjectError + 1, , "Setting '" & label & "' is blank on the '" & _
                    SETTINGS_SHEET & "' sheet - please fill it in (or click its Browse... button) " & _
                    "before running this."
            End If
            GetSetting_ = v
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 1, , "Setting '" & label & "' was not found on the '" & _
        SETTINGS_SHEET & "' sheet."
End Function

'-----------------------------------------------------------------------------------
' Creates the Settings sheet (label/value/browse-button layout) the first time
' anything needs it. Safe to call repeatedly - does nothing if it already exists.
'-----------------------------------------------------------------------------------
Private Function GetOrCreateSettingsSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim labels As Variant, i As Long

    On Error Resume Next
    Set ws = wb.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set GetOrCreateSettingsSheet = ws
        Exit Function
    End If

    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = SETTINGS_SHEET

    ws.Range("A1:C1").Value = Array("Setting", "Value (type/paste a path, or click Browse)", "")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(68, 114, 196)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)

    labels = Array("Test Case Workbook Path", "Import Template Workbook Path", _
                    "Report Output Template Path", "Export Folder Path", "PnL Only? (Yes/No)")
    For i = 0 To UBound(labels)
        ws.Cells(i + 2, 1).Value = labels(i)
        ws.Cells(i + 2, 1).Font.Bold = True
        ws.Cells(i + 2, 2).Interior.Color = RGB(255, 255, 0)   ' yellow = editable input
        If labels(i) Like "PnL Only*" Then
            ws.Cells(i + 2, 2).Value = "No"
            With ws.Cells(i + 2, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
            End With
        Else
            ws.Cells(i + 2, 3).Value = "Browse..."
            ws.Cells(i + 2, 3).Interior.Color = RGB(217, 217, 217)
            ws.Cells(i + 2, 3).Font.Underline = True
            ws.Cells(i + 2, 3).HorizontalAlignment = xlCenter
        End If
    Next i

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 70
    ws.Columns("C").ColumnWidth = 12

    ws.Range("A7").Value = "Tip: the 'Browse...' cells are just labels (Excel doesn't let a plain " & _
        "cell run code on click). Run BrowseForPaths from Alt+F8 instead, or Insert a Shape and " & _
        "assign it to BrowseForPaths to get a real clickable button."
    ws.Range("A7:C7").Merge
    ws.Range("A7").Font.Italic = True
    ws.Range("A7").Font.Size = 9
    ws.Range("A7").WrapText = True
    ws.Rows(7).RowHeight = 30

    Set GetOrCreateSettingsSheet = ws
End Function

'-----------------------------------------------------------------------------------
' Convenience: pops a file/folder picker for each of the 4 settings and writes
' whatever you choose straight into the Settings sheet. Run from Alt+F8, or wire a
' button to it (Insert -> Shape -> right-click -> Assign Macro -> BrowseForPaths).
'-----------------------------------------------------------------------------------
Sub BrowseForPaths()
    Dim ws As Worksheet
    Dim fd As FileDialog
    Dim result As String

    Set ws = GetOrCreateSettingsSheet(ThisWorkbook)

    result = PickFile("Select the Test Case workbook", "Excel Files (*.xlsx),*.xlsx")
    If result <> "" Then ws.Range("B2").Value = result

    result = PickFile("Select the Import Template workbook", "Excel Macro-Enabled Files (*.xlsm),*.xlsm")
    If result <> "" Then ws.Range("B3").Value = result

    result = PickFile("Select the Report Output Template workbook", "Excel Macro-Enabled Files (*.xlsm),*.xlsm")
    If result <> "" Then ws.Range("B4").Value = result

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select the Export Folder"
    If fd.Show = -1 Then ws.Range("B5").Value = fd.SelectedItems(1) & "\"

    MsgBox "Paths updated on the '" & SETTINGS_SHEET & "' sheet.", vbInformation
End Sub

Private Function PickFile(title As String, filterSpec As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = title
    fd.Filters.Clear
    fd.Filters.Add Split(filterSpec, ",")(0), Split(filterSpec, ",")(1)
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function

' Result icons (built with ChrW so they paste correctly regardless of editor encoding)
Private Function ICON_OK() As String: ICON_OK = ChrW(&H2714): End Function        ' check mark
Private Function ICON_FAIL() As String: ICON_FAIL = ChrW(&H274C): End Function    ' cross mark
Private Function ICON_WAIT() As String: ICON_WAIT = ChrW(&H25CF): End Function    ' dot
Private Function ICON_SKIP() As String: ICON_SKIP = ChrW(&H23ED): End Function    ' skip
Private Function ICON_RUN() As String: ICON_RUN = ChrW(&H21BB): End Function      ' running arrow

'===================================================================================
' Run everything in one go
'===================================================================================
Sub RunEverything()
    Step1_PushAllTestCasesToSQL
    Step3_RunReport
End Sub

'===================================================================================
' STEP 1 - list every Test Case ID, push each to SQL, pre-fill default report
' parameters (editable afterwards), log Import Status/Time against it
'===================================================================================
Sub Step1_PushAllTestCasesToSQL()

    Dim tcWB As Workbook, tplWB As Workbook
    Dim tcWS As Worksheet, drvWS As Worksheet
    Dim uniqueIDs As Collection
    Dim v As Variant
    Dim n As Long, lastRow As Long, r As Long
    Dim tcID As String
    Dim startTime As Double
    Dim wasOpenTC As Boolean, wasOpenTpl As Boolean
    Dim priorIncludeFlags As Object

    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    Set tcWB = GetOrOpenWorkbook(TESTCASE_WORKBOOK_PATH, wasOpenTC, readOnly:=True)
    Set tplWB = GetOrOpenWorkbook(TEMPLATE_WORKBOOK_PATH, wasOpenTpl, readOnly:=False)
    Set tcWS = tcWB.Worksheets("01_Portfolio")
    Set drvWS = GetOrCreateDriverSheet(ThisWorkbook)

    ' ---- capture existing "Include in Test Run?" flags BEFORE clearing, so a
    ' rebuild doesn't wipe out choices already made - this is what lets someone
    ' exclude a test case from the whole pipeline without editing 01_Portfolio ----
    Set priorIncludeFlags = CreateObject("Scripting.Dictionary")
    lastRow = drvWS.Cells(drvWS.Rows.Count, DCOL_TESTCASEID).End(xlUp).Row
    For r = 2 To lastRow
        Dim priorID As String
        priorID = Trim$(CStr(drvWS.Cells(r, DCOL_TESTCASEID).Value))
        If priorID <> "" Then
            priorIncludeFlags(priorID) = Trim$(CStr(drvWS.Cells(r, DCOL_INCLUDEINRUN).Value))
        End If
    Next r

    ' ---- collect unique, non-blank Test Case IDs, in the order they appear ----
    Set uniqueIDs = New Collection
    lastRow = tcWS.Cells(tcWS.Rows.Count, 1).End(xlUp).Row
    For r = TC_DATA_START_ROW To lastRow
        tcID = Trim$(CStr(tcWS.Cells(r, 1).Value))
        If tcID <> "" Then
            On Error Resume Next
            uniqueIDs.Add tcID, tcID
            On Error GoTo CleanFail
        End If
    Next r

    ' ---- (re)build the Driver sheet list, one row per Test Case ID, restoring
    ' each one's prior Include-in-Test-Run flag (defaults to Yes if new) ----
    ClearBelowHeader drvWS, 2
    n = 0
    For Each v In uniqueIDs
        n = n + 1
        drvWS.Cells(n + 1, DCOL_TESTCASEID).Value = CStr(v)
        drvWS.Cells(n + 1, DCOL_REPORTSTATUS).Value = "Not Run"
        drvWS.Cells(n + 1, DCOL_RESULT).Value = ICON_WAIT()
        If priorIncludeFlags.Exists(CStr(v)) And priorIncludeFlags(CStr(v)) <> "" Then
            drvWS.Cells(n + 1, DCOL_INCLUDEINRUN).Value = priorIncludeFlags(CStr(v))
        Else
            drvWS.Cells(n + 1, DCOL_INCLUDEINRUN).Value = "Yes"
        End If
    Next v
    ApplyRunReportDropdown drvWS, 2, n + 1
    ApplyIncludeInRunDropdown drvWS, 2, n + 1

    ' ---- process each row: populate template, push to SQL, fill defaults -
    ' rows flagged Include in Test Run? = No are skipped entirely, not even
    ' touched in the Import Template ----
    For r = 2 To n + 1
        tcID = CStr(drvWS.Cells(r, DCOL_TESTCASEID).Value)

        If UCase$(Trim$(CStr(drvWS.Cells(r, DCOL_INCLUDEINRUN).Value))) = "NO" Then
            drvWS.Cells(r, DCOL_IMPORTSTATUS).Value = "Skipped (Include in Test Run? = No)"
            drvWS.Cells(r, DCOL_RUNREPORT).Value = "No"
            drvWS.Cells(r, DCOL_EXPORT).Value = "No"
            drvWS.Cells(r, DCOL_IMPORTTIME).Value = Now
            GoTo NextRow
        End If

        If Left$(tcID, 5) <> "Test_" Then
            drvWS.Cells(r, DCOL_IMPORTSTATUS).Value = "Warning: TestCaseID doesn't start with 'Test_' " & _
                "- proceeding anyway (Name/PortfolioID are still tagged as test data)"
        End If

        On Error Resume Next
        Err.Clear
        ProcessOneTestCase tplWB, tcWB, drvWS, r, tcID
        If Err.Number <> 0 Then
            drvWS.Cells(r, DCOL_IMPORTSTATUS).Value = "Failed - " & Err.Description
            drvWS.Cells(r, DCOL_RUNREPORT).Value = "No"
        End If
        On Error GoTo CleanFail

        drvWS.Cells(r, DCOL_IMPORTTIME).Value = Now
        ColorImportStatus drvWS, r
        DoEvents
NextRow:
    Next r

    RefreshParametersStatus drvWS
    tplWB.Save

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    If Not wasOpenTC And Not tcWB Is Nothing Then tcWB.Close SaveChanges:=False
    ' Leave tplWB open - handy to eyeball red validation cells if something failed

    MsgBox n & " test case(s) processed in " & Format(Timer - startTime, "0.0") & _
           " sec." & vbCrLf & "Review/edit parameters in '" & DRIVER_SHEET & _
           "', then run Step3_RunReport.", vbInformation, "Step 1 complete"
    Exit Sub

CleanFail:
    MsgBox "Step 1 stopped: " & Err.Description, vbCritical
    Resume CleanExit

End Sub

'-----------------------------------------------------------------------------------
' Populate the (external) template for one test case, push to SQL, write the
' import result + default report parameters into this workbook's Driver row.
'-----------------------------------------------------------------------------------
Private Sub ProcessOneTestCase(tplWB As Workbook, tcWB As Workbook, drvWS As Worksheet, _
                                drvRow As Long, tcID As String)

    Dim portfolioID As Variant, reportDateVal As Variant, vendorVal As Variant
    Dim runOffVal As Variant, includeExpiredVal As Variant, asOfDateVal As Variant
    Dim newReportID As Long, rowsWritten As Long

    ClearTemplateSheets tplWB

    PopulateGuidelines tplWB, tcWB, tcID
    PopulateProfitAndLoss tplWB, tcWB, tcID
    PopulateBoundPortfolio tplWB, tcWB, tcID, portfolioID, reportDateVal, vendorVal, _
                            runOffVal, includeExpiredVal, asOfDateVal
    rowsWritten = PopulateBoundPortfolioReport(tplWB, tcWB, tcID)

    drvWS.Cells(drvRow, DCOL_PORTFOLIOID).Value = portfolioID

    If portfolioID = "" Or IsEmpty(portfolioID) Then
        drvWS.Cells(drvRow, DCOL_IMPORTSTATUS).Value = "Failed - no matching row in 07_Bound_Portfolio"
        drvWS.Cells(drvRow, DCOL_RUNREPORT).Value = "No"
        drvWS.Cells(drvRow, DCOL_EXPORT).Value = "No"
        Exit Sub
    End If

    If rowsWritten = 0 Then
        drvWS.Cells(drvRow, DCOL_IMPORTSTATUS).Value = "Failed - no matching 03_Mapping/02_Transaction rows"
        drvWS.Cells(drvRow, DCOL_RUNREPORT).Value = "No"
        drvWS.Cells(drvRow, DCOL_EXPORT).Value = "No"
        Exit Sub
    End If

    ' ---- safe existence check (read-only) - if this test PortfolioID already
    ' has data in SQL from a prior run, skip rather than blindly re-insert and
    ' create duplicates. This does NOT delete/modify anything - that requires
    ' confirming the FK relationships first (see the cleanup stored procedure
    ' discussion). Re-run the cleanup SP, or manually clear this PortfolioID's
    ' rows, then re-run this test case. ----
    If PortfolioAlreadyExists(tplWB, portfolioID) Then
        drvWS.Cells(drvRow, DCOL_IMPORTSTATUS).Value = "Skipped - PortfolioID " & portfolioID & _
            " already exists in SQL (re-run of a prior test?). Run the cleanup stored " & _
            "procedure first if you want fresh data for this case."
        drvWS.Cells(drvRow, DCOL_RUNREPORT).Value = "No"
        drvWS.Cells(drvRow, DCOL_EXPORT).Value = "No"
        ClearTemplateSheets tplWB
        Exit Sub
    End If

    ' ---- push to SQL via the Import Template's own macro (different file,
    ' so Application.Run is correct; qualified with its module name because
    ' that workbook has two same-named BoundPortfolio Subs) ----
    Application.Run "'" & tplWB.Name & "'!InputBoundPortfolio.BoundPortfolio"

    newReportID = GetLatestReportID(tplWB, portfolioID, reportDateVal, vendorVal)

    If newReportID > 0 Then
        drvWS.Cells(drvRow, DCOL_IMPORTSTATUS).Value = "Imported"
        drvWS.Cells(drvRow, DCOL_RUNREPORT).Value = "Yes"
        drvWS.Cells(drvRow, DCOL_EXPORT).Value = "Yes"

        ' ---- default report parameters (editable afterwards) - RunOff,
        ' IncludeExpired, and AsOfDate now come from 07_Bound_Portfolio itself
        ' rather than a fixed guess ----
        drvWS.Cells(drvRow, DCOL_VENDOR).Value = vendorVal
        drvWS.Cells(drvRow, DCOL_FRONTING).Value = "all"
        drvWS.Cells(drvRow, DCOL_SECINSTR).Value = "all"
        drvWS.Cells(drvRow, DCOL_ASOFDATE).Value = asOfDateVal
        drvWS.Cells(drvRow, DCOL_RUNOFF).Value = runOffVal
        drvWS.Cells(drvRow, DCOL_INCLUDEEXPIRED).Value = includeExpiredVal
        drvWS.Cells(drvRow, DCOL_AEPTYPE).Value = "AEPTVAR"
        drvWS.Cells(drvRow, DCOL_AEPCOTVARTYPE).Value = "AEPCOTVAR"
    Else
        drvWS.Cells(drvRow, DCOL_IMPORTSTATUS).Value = "Failed - validation failed, check " & _
            "template sheets for red-highlighted cells"
        drvWS.Cells(drvRow, DCOL_RUNREPORT).Value = "No"
    End If

    ClearTemplateSheets tplWB   ' leave the template clean for the next test case

End Sub

'===================================================================================
' Recompute the Parameters column (Ready / Missing) for every row from whatever
' is currently in F:M - call this any time after you hand-edit parameters, or
' automatically before Step 3 runs.
'===================================================================================
Sub RefreshParametersStatus(Optional drvWS As Worksheet)
    Dim lastRow As Long, r As Long
    If drvWS Is Nothing Then Set drvWS = GetOrCreateDriverSheet(ThisWorkbook)
    lastRow = drvWS.Cells(drvWS.Rows.Count, DCOL_TESTCASEID).End(xlUp).Row
    For r = 2 To lastRow
        If RowHasAllRequiredParams(drvWS, r) Then
            drvWS.Cells(r, DCOL_PARAMETERS).Value = "Ready"
            drvWS.Cells(r, DCOL_PARAMETERS).Font.Color = RGB(0, 128, 0)
        Else
            drvWS.Cells(r, DCOL_PARAMETERS).Value = "Missing"
            drvWS.Cells(r, DCOL_PARAMETERS).Font.Color = RGB(200, 0, 0)
        End If
    Next r
End Sub

Private Function RowHasAllRequiredParams(drvWS As Worksheet, r As Long) As Boolean
    Dim c As Variant
    For Each c In Array(DCOL_VENDOR, DCOL_FRONTING, DCOL_SECINSTR, DCOL_ASOFDATE, _
                         DCOL_RUNOFF, DCOL_INCLUDEEXPIRED, DCOL_AEPTYPE, DCOL_AEPCOTVARTYPE)
        If Trim$(CStr(drvWS.Cells(r, CLng(c)).Value)) = "" Then
            RowHasAllRequiredParams = False
            Exit Function
        End If
    Next c
    RowHasAllRequiredParams = True
End Function

'===================================================================================
' STEP 3 - rebuild the Report Output Template workbook's TestCases sheet from every row here
' where Run Report? = Yes AND Parameters = Ready (using YOUR edited values), run
' its report macro, then read results back per row.
'===================================================================================
Sub Step3_RunReport()

    Dim drvWS As Worksheet, tcsWS As Worksheet
    Dim uiWB As Workbook
    Dim lastDrvRow As Long, r As Long, outRow As Long
    Dim wasOpenUI As Boolean
    Dim n As Long
    Dim eligibleRows As Collection
    Dim rv As Variant

    Set drvWS = GetOrCreateDriverSheet(ThisWorkbook)
    RefreshParametersStatus drvWS

    Set uiWB = GetOrOpenWorkbook(UITESTING_WORKBOOK_PATH, wasOpenUI, readOnly:=False)
    Set tcsWS = uiWB.Worksheets("TestCases")
    ClearBelowHeader tcsWS, 2

    lastDrvRow = drvWS.Cells(drvWS.Rows.Count, DCOL_TESTCASEID).End(xlUp).Row
    outRow = 2
    Set eligibleRows = New Collection

    For r = 2 To lastDrvRow
        If UCase$(Trim$(CStr(drvWS.Cells(r, DCOL_RUNREPORT).Value))) = "YES" And _
           CStr(drvWS.Cells(r, DCOL_IMPORTSTATUS).Value) = "Imported" And _
           CStr(drvWS.Cells(r, DCOL_PARAMETERS).Value) = "Ready" Then

            With tcsWS
                .Cells(outRow, 1).Value = drvWS.Cells(r, DCOL_PORTFOLIOID).Value
                .Cells(outRow, 2).Value = drvWS.Cells(r, DCOL_ASOFDATE).Value       ' ReportDate
                .Cells(outRow, 3).Value = drvWS.Cells(r, DCOL_VENDOR).Value
                .Cells(outRow, 4).Value = drvWS.Cells(r, DCOL_FRONTING).Value
                .Cells(outRow, 5).Value = drvWS.Cells(r, DCOL_SECINSTR).Value
                .Cells(outRow, 6).Value = drvWS.Cells(r, DCOL_ASOFDATE).Value
                .Cells(outRow, 7).Value = drvWS.Cells(r, DCOL_RUNOFF).Value
                .Cells(outRow, 8).Value = drvWS.Cells(r, DCOL_INCLUDEEXPIRED).Value
                .Cells(outRow, 9).Value = ""                                        ' ExportID
                .Cells(outRow, 10).Value = drvWS.Cells(r, DCOL_AEPTYPE).Value
                .Cells(outRow, 11).Value = drvWS.Cells(r, DCOL_AEPCOTVARTYPE).Value
            End With

            drvWS.Cells(r, DCOL_REPORTSTATUS).Value = "Running..."
            drvWS.Cells(r, DCOL_RESULT).Value = ICON_RUN()
            eligibleRows.Add r

            outRow = outRow + 1
        ElseIf UCase$(Trim$(CStr(drvWS.Cells(r, DCOL_RUNREPORT).Value))) = "NO" Then
            drvWS.Cells(r, DCOL_REPORTSTATUS).Value = "Skipped"
            drvWS.Cells(r, DCOL_RESULT).Value = ICON_SKIP()
        ElseIf CStr(drvWS.Cells(r, DCOL_IMPORTSTATUS).Value) <> "Imported" Then
            drvWS.Cells(r, DCOL_REPORTSTATUS).Value = "Skipped - not imported"
            drvWS.Cells(r, DCOL_RESULT).Value = ICON_SKIP()
        ElseIf CStr(drvWS.Cells(r, DCOL_PARAMETERS).Value) <> "Ready" Then
            drvWS.Cells(r, DCOL_REPORTSTATUS).Value = "Skipped - missing parameters"
            drvWS.Cells(r, DCOL_RESULT).Value = ICON_SKIP()
        End If
    Next r

    n = eligibleRows.Count
    If n = 0 Then
        MsgBox "No rows are eligible to run (need Run Report?=Yes, Import Status=Imported, " & _
               "Parameters=Ready).", vbExclamation
        If Not wasOpenUI Then uiWB.Close SaveChanges:=False
        Exit Sub
    End If

    uiWB.Save
    Application.ScreenUpdating = True: DoEvents   ' let the "Running..." rows paint before the long call

    On Error GoTo ReportFail
    Application.Run "'" & uiWB.Name & "'!MasterAutomation_OUTPUT.RunAllCases", IsPnLOnly()
    On Error GoTo 0

    ' ---- read back per-case outcome from the Report Output Template workbook's ErrorLog ----
    Dim failedIDs As Object
    Set failedIDs = GetFailedCaseIDs(uiWB)

    Dim tcIDLookup As String, pid As String
    For Each rv In eligibleRows
        tcIDLookup = CStr(drvWS.Cells(CLng(rv), DCOL_TESTCASEID).Value)
        pid = CStr(drvWS.Cells(CLng(rv), DCOL_PORTFOLIOID).Value)
        If failedIDs.Exists(tcIDLookup) Then
            drvWS.Cells(CLng(rv), DCOL_REPORTSTATUS).Value = "Failed - " & failedIDs(tcIDLookup)
            drvWS.Cells(CLng(rv), DCOL_RESULT).Value = ICON_FAIL()
            drvWS.Cells(CLng(rv), DCOL_OUTPUTLOCATION).Value = ""
        Else
            drvWS.Cells(CLng(rv), DCOL_REPORTSTATUS).Value = "Completed"
            drvWS.Cells(CLng(rv), DCOL_RESULT).Value = ICON_OK()
            drvWS.Cells(CLng(rv), DCOL_OUTPUTLOCATION).Value = _
                "AC_ tabs in " & uiWB.Name & " (filter TestCaseID=" & tcIDLookup & ")"
        End If
    Next rv

    uiWB.Save
    ' Leave uiWB open so you can review the AC_ tabs straight away

    MsgBox n & " case(s) sent to the report step. See '" & DRIVER_SHEET & _
           "' for Completed/Failed status per row.", vbInformation, "Report step complete"
    Exit Sub

ReportFail:
    For Each rv In eligibleRows
        drvWS.Cells(CLng(rv), DCOL_REPORTSTATUS).Value = "Failed - report macro error"
        drvWS.Cells(CLng(rv), DCOL_RESULT).Value = ICON_FAIL()
    Next rv
    MsgBox "Report step failed: " & Err.Description & vbCrLf & _
           "Check that 'MasterAutomation_OUTPUT.RunAllCases' still exists in the Report Output Template workbook.", _
           vbCritical
End Sub

'-----------------------------------------------------------------------------------
' Reads the Report Output Template workbook's "ErrorLog" sheet (written by RunAllCases /
' LogError) and returns a Dictionary of PortfolioID -> most recent failure reason,
' for every case that failed on this run.
'-----------------------------------------------------------------------------------
' NOTE: ErrorLog's real layout, confirmed from InitErrorLog/LogError, is
' Timestamp (A) | CaseID (B) | Tab/Step (C) | Message (D). This function
' previously read column A (Timestamp) instead of column B (CaseID) - meaning
' the failure lookup below could never match anything, so cases likely always
' showed "Completed" regardless of whether they actually failed. Fixed here.
' Also now keyed by TestCaseID (since the patched RunAllCases logs TestCaseID,
' not the raw PortfolioID, into column B) rather than PortfolioID.
Private Function GetFailedCaseIDs(uiWB As Workbook) As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim tcID As String, reason As String

    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = uiWB.Worksheets("ErrorLog")
    On Error GoTo 0
    If ws Is Nothing Then
        Set GetFailedCaseIDs = dict
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        tcID = Trim$(CStr(ws.Cells(r, 2).Value))    ' CaseID column
        reason = Trim$(CStr(ws.Cells(r, 4).Value))  ' Message column
        If tcID <> "" Then dict(tcID) = reason   ' later rows overwrite earlier ones - keeps the latest attempt
    Next r

    Set GetFailedCaseIDs = dict
End Function

'===================================================================================
' STEP 4 - export results for whichever rows have Export? = Yes (and Import
' Status = Imported, Parameters = Ready). Mode is read from the Driver sheet's
' "EXPORT CONTROLS" panel (U2):
'
'   "Separate"     -> re-runs the report ONE case at a time, in isolation (AC_
'                     tabs cleared before each), then saves a full copy of the
'                     whole Report Output Template workbook as OutputTemplate_<TestCaseID>.xlsx.
'                     Every file is a complete, untouched snapshot for that case
'                     only - nothing gets filtered, nothing gets overwritten by
'                     another case's data.
'
'   "Consolidated" -> re-runs the report for every selected case together in ONE
'                     batch, then saves a single full copy of the workbook as
'                     OutputTemplate_<TestCaseIDs>.xlsx (or a count+timestamp if
'                     too many cases to fit in a filename).
'
' Both re-run the actual stored-procedure call rather than reusing whatever is
' currently sitting in the AC_ tabs, so the exported file always matches exactly
' what you ticked in Export? - independent of whatever Step3_RunReport did earlier.
'===================================================================================
Sub Step4_ExportResults()

    Dim drvWS As Worksheet, uiWB As Workbook
    Dim wasOpenUI As Boolean
    Dim mode As String
    Dim lastRow As Long, r As Long
    Dim selected As Collection

    Set drvWS = GetOrCreateDriverSheet(ThisWorkbook)
    RefreshParametersStatus drvWS
    mode = UCase$(Trim$(CStr(drvWS.Range("V2").Value)))
    If mode <> "CONSOLIDATED" Then mode = "SEPARATE"

    Set uiWB = GetOrOpenWorkbook(UITESTING_WORKBOOK_PATH, wasOpenUI, readOnly:=False)

    Set selected = New Collection
    lastRow = drvWS.Cells(drvWS.Rows.Count, DCOL_TESTCASEID).End(xlUp).Row
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(drvWS.Cells(r, DCOL_EXPORT).Value))) = "YES" And _
           CStr(drvWS.Cells(r, DCOL_IMPORTSTATUS).Value) = "Imported" And _
           CStr(drvWS.Cells(r, DCOL_PARAMETERS).Value) = "Ready" Then
            selected.Add r
        End If
    Next r

    If selected.Count = 0 Then
        MsgBox "No rows are eligible to export (need Export?=Yes, Import Status=Imported, " & _
               "Parameters=Ready).", vbExclamation
        If Not wasOpenUI Then uiWB.Close SaveChanges:=False
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If mode = "CONSOLIDATED" Then
        RunBatchAndSnapshot uiWB, drvWS, selected
    Else
        Dim rv As Variant
        For Each rv In selected
            RunSingleCaseAndSnapshot uiWB, drvWS, CLng(rv)
        Next rv
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    uiWB.Save
    If Not wasOpenUI Then uiWB.Close SaveChanges:=False

    MsgBox selected.Count & " test case(s) exported (" & mode & " mode). See '" & DRIVER_SHEET & _
           "' Output Location column for the file path(s).", vbInformation, "Step 4 complete"

End Sub

'-----------------------------------------------------------------------------------
' Separate mode, one row: clear AC_/ErrorLog, push just this case into TestCases,
' run the report macro in isolation, and save a full workbook copy named with its
' Test Case ID. No filtering - the copy IS that case's complete, isolated result.
'-----------------------------------------------------------------------------------
Private Sub RunSingleCaseAndSnapshot(uiWB As Workbook, drvWS As Worksheet, drvRow As Long)
    Dim tcsWS As Worksheet
    Dim tcID As String
    Dim outPath As String
    Dim failedIDs As Object

    tcID = CStr(drvWS.Cells(drvRow, DCOL_TESTCASEID).Value)

    ClearACAndErrorLog uiWB

    Set tcsWS = uiWB.Worksheets("TestCases")
    ClearBelowHeader tcsWS, 2
    WriteTestCasesRow tcsWS, 2, drvWS, drvRow

    drvWS.Cells(drvRow, DCOL_REPORTSTATUS).Value = "Running..."
    drvWS.Cells(drvRow, DCOL_RESULT).Value = ICON_RUN()

    On Error GoTo Fail
    Application.Run "'" & uiWB.Name & "'!MasterAutomation_OUTPUT.RunAllCases", IsPnLOnly()
    On Error GoTo 0

    Set failedIDs = GetFailedCaseIDs(uiWB)
    If failedIDs.Exists(tcID) Then
        drvWS.Cells(drvRow, DCOL_REPORTSTATUS).Value = "Failed - " & failedIDs(tcID)
        drvWS.Cells(drvRow, DCOL_RESULT).Value = ICON_FAIL()
        drvWS.Cells(drvRow, DCOL_OUTPUTLOCATION).Value = ""
        Exit Sub
    End If

    outPath = EXPORT_FOLDER_PATH & "OutputTemplate_" & tcID & ".xlsx"
    On Error GoTo SaveFail
    uiWB.SaveCopyAs outPath
    On Error GoTo 0

    drvWS.Cells(drvRow, DCOL_REPORTSTATUS).Value = "Completed"
    drvWS.Cells(drvRow, DCOL_RESULT).Value = ICON_OK()
    drvWS.Cells(drvRow, DCOL_OUTPUTLOCATION).Value = outPath
    Exit Sub

Fail:
    drvWS.Cells(drvRow, DCOL_REPORTSTATUS).Value = "Failed - report macro error: " & Err.Description
    drvWS.Cells(drvRow, DCOL_RESULT).Value = ICON_FAIL()
    Exit Sub

SaveFail:
    drvWS.Cells(drvRow, DCOL_OUTPUTLOCATION).Value = "Export FAILED - " & Err.Description & _
        " (check the 'Export Folder Path' on the Settings sheet points to a real, existing folder: " & _
        EXPORT_FOLDER_PATH & ")"
End Sub

'-----------------------------------------------------------------------------------
' Consolidated mode: clear AC_/ErrorLog once, push every selected case into
' TestCases together, run the report macro once for all of them, then save ONE
' full workbook copy containing exactly that selected set.
'-----------------------------------------------------------------------------------
Private Sub RunBatchAndSnapshot(uiWB As Workbook, drvWS As Worksheet, selected As Collection)
    Dim tcsWS As Worksheet
    Dim rv As Variant, outRow As Long, drvRow As Long
    Dim tcIDs As String, tcIDLookup As String
    Dim outPath As String
    Dim failedIDs As Object

    ClearACAndErrorLog uiWB

    Set tcsWS = uiWB.Worksheets("TestCases")
    ClearBelowHeader tcsWS, 2
    outRow = 2
    For Each rv In selected
        drvRow = CLng(rv)
        WriteTestCasesRow tcsWS, outRow, drvWS, drvRow
        drvWS.Cells(drvRow, DCOL_REPORTSTATUS).Value = "Running..."
        drvWS.Cells(drvRow, DCOL_RESULT).Value = ICON_RUN()
        If tcIDs = "" Then
            tcIDs = CStr(drvWS.Cells(drvRow, DCOL_TESTCASEID).Value)
        Else
            tcIDs = tcIDs & "_" & CStr(drvWS.Cells(drvRow, DCOL_TESTCASEID).Value)
        End If
        outRow = outRow + 1
    Next rv

    On Error GoTo Fail
    Application.Run "'" & uiWB.Name & "'!MasterAutomation_OUTPUT.RunAllCases", IsPnLOnly()
    On Error GoTo 0

    Set failedIDs = GetFailedCaseIDs(uiWB)

    If Len(tcIDs) <= 100 Then
        outPath = EXPORT_FOLDER_PATH & "OutputTemplate_" & tcIDs & ".xlsx"
    Else
        outPath = EXPORT_FOLDER_PATH & "OutputTemplate_Consolidated_" & selected.Count & _
                  "cases_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    End If

    On Error GoTo SaveFail
    uiWB.SaveCopyAs outPath
    On Error GoTo 0

    For Each rv In selected
        drvRow = CLng(rv)
        tcIDLookup = CStr(drvWS.Cells(drvRow, DCOL_TESTCASEID).Value)
        If failedIDs.Exists(tcIDLookup) Then
            drvWS.Cells(drvRow, DCOL_REPORTSTATUS).Value = "Failed - " & failedIDs(tcIDLookup)
            drvWS.Cells(drvRow, DCOL_RESULT).Value = ICON_FAIL()
            drvWS.Cells(drvRow, DCOL_OUTPUTLOCATION).Value = ""
        Else
            drvWS.Cells(drvRow, DCOL_REPORTSTATUS).Value = "Completed"
            drvWS.Cells(drvRow, DCOL_RESULT).Value = ICON_OK()
            drvWS.Cells(drvRow, DCOL_OUTPUTLOCATION).Value = outPath
        End If
    Next rv
    Exit Sub

Fail:
    For Each rv In selected
        drvWS.Cells(CLng(rv), DCOL_REPORTSTATUS).Value = "Failed - report macro error: " & Err.Description
        drvWS.Cells(CLng(rv), DCOL_RESULT).Value = ICON_FAIL()
    Next rv
    Exit Sub

SaveFail:
    For Each rv In selected
        drvWS.Cells(CLng(rv), DCOL_OUTPUTLOCATION).Value = "Export FAILED - " & Err.Description & _
            " (check the 'Export Folder Path' on the Settings sheet points to a real, existing folder: " & _
            EXPORT_FOLDER_PATH & ")"
    Next rv
End Sub

'-----------------------------------------------------------------------------------
' Writes one Driver row's parameters into a TestCases row, in the 11-column order
' MasterAutomation_OUTPUT.RunAllCases expects.
'-----------------------------------------------------------------------------------
Private Sub WriteTestCasesRow(tcsWS As Worksheet, outRow As Long, drvWS As Worksheet, drvRow As Long)
    With tcsWS
        .Cells(outRow, 1).Value = drvWS.Cells(drvRow, DCOL_PORTFOLIOID).Value   ' required by the stored procedure - do not change
        .Cells(outRow, 2).Value = drvWS.Cells(drvRow, DCOL_ASOFDATE).Value    ' ReportDate
        .Cells(outRow, 3).Value = drvWS.Cells(drvRow, DCOL_VENDOR).Value
        .Cells(outRow, 4).Value = drvWS.Cells(drvRow, DCOL_FRONTING).Value
        .Cells(outRow, 5).Value = drvWS.Cells(drvRow, DCOL_SECINSTR).Value
        .Cells(outRow, 6).Value = drvWS.Cells(drvRow, DCOL_ASOFDATE).Value
        .Cells(outRow, 7).Value = drvWS.Cells(drvRow, DCOL_RUNOFF).Value
        .Cells(outRow, 8).Value = drvWS.Cells(drvRow, DCOL_INCLUDEEXPIRED).Value
        .Cells(outRow, 9).Value = ""                                          ' ExportID
        .Cells(outRow, 10).Value = drvWS.Cells(drvRow, DCOL_AEPTYPE).Value
        .Cells(outRow, 11).Value = drvWS.Cells(drvRow, DCOL_AEPCOTVARTYPE).Value
        .Cells(outRow, 12).Value = drvWS.Cells(drvRow, DCOL_TESTCASEID).Value  ' TestCaseID (col L) - used by
                                                                                ' the patched RunAllCases to tag
                                                                                ' every AC_ row/ErrorLog entry,
                                                                                ' instead of the raw PortfolioID
    End With
End Sub

'-----------------------------------------------------------------------------------
' Deletes every AC_... sheet and the ErrorLog sheet so the next RunAllCases call
' starts from a clean slate - required for true per-case isolation in Separate
' mode, and to keep a Consolidated snapshot limited to exactly what was selected.
'-----------------------------------------------------------------------------------
Private Sub ClearACAndErrorLog(uiWB As Workbook)
    Dim ws As Worksheet, i As Long
    For i = uiWB.Worksheets.Count To 1 Step -1
        Set ws = uiWB.Worksheets(i)
        If Left$(ws.Name, 3) = "AC_" Or ws.Name = "ErrorLog" Then
            ws.Delete
        End If
    Next i
End Sub

'===================================================================================
' Data validation dropdown (Yes/No) on the Run Report? column
'===================================================================================
Private Sub ApplyRunReportDropdown(drvWS As Worksheet, firstRow As Long, lastRow As Long)
    Dim rng As Range
    If lastRow < firstRow Then Exit Sub
    Set rng = drvWS.Range(drvWS.Cells(firstRow, DCOL_RUNREPORT), drvWS.Cells(lastRow, DCOL_RUNREPORT))
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Yes,No"
    On Error GoTo 0

    Set rng = drvWS.Range(drvWS.Cells(firstRow, DCOL_EXPORT), drvWS.Cells(lastRow, DCOL_EXPORT))
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Yes,No"
    On Error GoTo 0
End Sub

Private Sub ApplyIncludeInRunDropdown(drvWS As Worksheet, firstRow As Long, lastRow As Long)
    Dim rng As Range
    If lastRow < firstRow Then Exit Sub
    Set rng = drvWS.Range(drvWS.Cells(firstRow, DCOL_INCLUDEINRUN), drvWS.Cells(lastRow, DCOL_INCLUDEINRUN))
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Yes,No"
    On Error GoTo 0
End Sub

Private Sub ColorImportStatus(drvWS As Worksheet, r As Long)
    Dim v As String
    v = CStr(drvWS.Cells(r, DCOL_IMPORTSTATUS).Value)
    If v = "Imported" Then
        drvWS.Cells(r, DCOL_IMPORTSTATUS).Interior.Color = RGB(198, 239, 206)
    ElseIf InStr(1, v, "Failed", vbTextCompare) > 0 Then
        drvWS.Cells(r, DCOL_IMPORTSTATUS).Interior.Color = RGB(255, 199, 206)
    End If
End Sub

'===================================================================================
' 04_Portfolio_Guidelines  ->  "Guidelines"   (D:I -> A:F, columns line up 1:1)
'===================================================================================
Private Sub PopulateGuidelines(tplWB As Workbook, tcWB As Workbook, tcID As String)
    Dim src As Worksheet, tgt As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Set src = tcWB.Worksheets("04_Portfolio_Guidelines")
    Set tgt = tplWB.Worksheets("Guidelines")

    lastRow = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    outRow = 2
    For r = TC_DATA_START_ROW To lastRow
        If Trim$(CStr(src.Cells(r, 1).Value)) = tcID Then
            tgt.Cells(outRow, 1).Resize(1, 6).Value = src.Cells(r, 4).Resize(1, 6).Value
            outRow = outRow + 1
        End If
    Next r
End Sub

'===================================================================================
' 06_Portfolio_Expenses  ->  "Profit and Loss"  (one row per test case)
'===================================================================================
Private Sub PopulateProfitAndLoss(tplWB As Workbook, tcWB As Workbook, tcID As String)
    Dim src As Worksheet, tgt As Worksheet
    Dim lastRow As Long, r As Long
    Set src = tcWB.Worksheets("06_Portfolio_Expenses")
    Set tgt = tplWB.Worksheets("Profit and Loss")

    lastRow = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    For r = TC_DATA_START_ROW To lastRow
        If Trim$(CStr(src.Cells(r, 1).Value)) = tcID Then
            With tgt
                .Cells(2, 1).Value = src.Cells(r, 3).Value    ' Management Fee
                .Cells(2, 2).Value = src.Cells(r, 4).Value    ' Carry
                .Cells(2, 3).Value = src.Cells(r, 6).Value    ' Hurdle (Hard)
                .Cells(2, 4).Value = src.Cells(r, 7).Value    ' Fronting Fee
                .Cells(2, 5).Value = src.Cells(r, 8).Value    ' Acquisition Expense
                .Cells(2, 6).Value = src.Cells(r, 9).Value    ' FET
                .Cells(2, 7).Value = src.Cells(r, 10).Value   ' Investment Inc
                .Cells(2, 8).Value = src.Cells(r, 11).Value   ' Op/Org Expenses
                .Cells(2, 9).Value = src.Cells(r, 5).Value    ' HurdleType
                .Cells(2, 10).Value = src.Cells(r, 12).Value  ' Risk Free Rate
                .Cells(2, 11).Value = src.Cells(r, 13).Value  ' QS Attritional LR US Wind
                .Cells(2, 12).Value = src.Cells(r, 14).Value  ' QS Attritional LR Worldwide
                .Cells(2, 13).Value = src.Cells(r, 15).Value  ' Pro Forma Rate Change
            End With
            Exit For
        End If
    Next r
End Sub

'===================================================================================
' 07_Bound_Portfolio  ->  "BoundPortfolio"   (one row per test case)
'===================================================================================
Private Sub PopulateBoundPortfolio(tplWB As Workbook, tcWB As Workbook, tcID As String, _
                                    ByRef portfolioID As Variant, ByRef reportDateVal As Variant, _
                                    ByRef vendorVal As Variant, ByRef runOffVal As Variant, _
                                    ByRef includeExpiredVal As Variant, ByRef asOfDateVal As Variant)
    Dim src As Worksheet, tgt As Worksheet
    Dim lastRow As Long, r As Long

    portfolioID = Empty
    Set src = tcWB.Worksheets("07_Bound_Portfolio")
    Set tgt = tplWB.Worksheets("BoundPortfolio")

    lastRow = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    For r = TC_DATA_START_ROW To lastRow
        If Trim$(CStr(src.Cells(r, 1).Value)) = tcID Then
            ' PortfolioID is DERIVED from the Test Case ID itself (e.g. TC_PF_001 ->
            ' "001"), not read from column C - that source column isn't reliably
            ' unique per test case, which was causing every case to collide on the
            ' same SQL PortfolioID and overwrite each other's data.
            portfolioID = DerivePortfolioID(tcID)
            vendorVal = src.Cells(r, 5).Value
            reportDateVal = src.Cells(r, 13).Value
            If reportDateVal = "" Then reportDateVal = src.Cells(r, 12).Value
            If reportDateVal = "" Then reportDateVal = Date

            ' ---- report parameters that 07_Bound_Portfolio already carries per
            ' test case - used as the Driver sheet's defaults instead of a fixed
            ' 0/1/ReportDate guess ----
            runOffVal = src.Cells(r, 7).Value          ' RunOff (G)
            If runOffVal = "" Then runOffVal = 0
            includeExpiredVal = src.Cells(r, 8).Value  ' IncludeExpired (H)
            If includeExpiredVal = "" Then includeExpiredVal = 1
            asOfDateVal = src.Cells(r, 12).Value        ' AsOfDate (L)
            If asOfDateVal = "" Then asOfDateVal = reportDateVal

            With tgt
                .Cells(2, 1).Value = portfolioID
                .Cells(2, 2).Value = vendorVal
                .Cells(2, 3).Value = "Test_" & src.Cells(r, 4).Value   ' Test_ + 07_Bound_Portfolio's own Name
                .Cells(2, 4).Value = src.Cells(r, 9).Value
                .Cells(2, 5).Value = src.Cells(r, 10).Value
                .Cells(2, 6).Value = src.Cells(r, 6).Value
                .Cells(2, 7).Value = reportDateVal
                .Cells(2, 8).Value = src.Cells(r, 14).Value
            End With
            Exit For
        End If
    Next r
End Sub

'===================================================================================
' 03_Mapping (joined to 02_Transaction on TransactionID)  ->  "Bound Portfolio Report"
'===================================================================================
Private Function PopulateBoundPortfolioReport(tplWB As Workbook, tcWB As Workbook, tcID As String) As Long
    Dim mapWS As Worksheet, trxWS As Worksheet, tgt As Worksheet
    Dim mapLastRow As Long, trxLastRow As Long
    Dim mr As Long, tr As Long, outRow As Long
    Dim trxID As String, ilsName As String
    Dim foundTrxRow As Long

    Set mapWS = tcWB.Worksheets("03_Mapping")
    Set trxWS = tcWB.Worksheets("02_Transaction")
    Set tgt = tplWB.Worksheets("Bound Portfolio Report")

    mapLastRow = mapWS.Cells(mapWS.Rows.Count, 1).End(xlUp).Row
    trxLastRow = trxWS.Cells(trxWS.Rows.Count, 1).End(xlUp).Row
    outRow = 2

    For mr = TC_DATA_START_ROW To mapLastRow
        If Trim$(CStr(mapWS.Cells(mr, 1).Value)) = tcID Then
            trxID = Trim$(CStr(mapWS.Cells(mr, 3).Value))
            ilsName = Trim$(CStr(mapWS.Cells(mr, 4).Value))

            foundTrxRow = 0
            For tr = TC_DATA_START_ROW To trxLastRow
                If Trim$(CStr(trxWS.Cells(tr, 1).Value)) = tcID And _
                   Trim$(CStr(trxWS.Cells(tr, 3).Value)) = trxID Then
                    foundTrxRow = tr
                    Exit For
                End If
            Next tr

            If foundTrxRow > 0 Then
                WriteBoundDetailRow tgt, outRow, ilsName, trxWS, foundTrxRow
                outRow = outRow + 1
            End If
        End If
    Next mr

    PopulateBoundPortfolioReport = outRow - 2
End Function

Private Sub WriteBoundDetailRow(tgt As Worksheet, outRow As Long, ilsName As String, _
                                 trxWS As Worksheet, trxRow As Long)
    Dim buyDate As Variant
    buyDate = trxWS.Cells(trxRow, 10).Value

    With tgt
        .Cells(outRow, 3).Value = ilsName
        .Cells(outRow, 4).Value = trxWS.Cells(trxRow, 7).Value
        .Cells(outRow, 5).Value = trxWS.Cells(trxRow, 6).Value
        .Cells(outRow, 6).Value = trxWS.Cells(trxRow, 8).Value
        .Cells(outRow, 11).Value = trxWS.Cells(trxRow, 12).Value
        .Cells(outRow, 12).Value = buyDate
        .Cells(outRow, 15).Value = trxWS.Cells(trxRow, 14).Value
        .Cells(outRow, 18).Value = trxWS.Cells(trxRow, 11).Value
        .Cells(outRow, 20).Value = trxWS.Cells(trxRow, 13).Value
        .Cells(outRow, 21).Value = trxWS.Cells(trxRow, 17).Value
        .Cells(outRow, 22).Value = trxWS.Cells(trxRow, 9).Value
        .Cells(outRow, 23).Value = buyDate
        .Cells(outRow, 26).Value = trxWS.Cells(trxRow, 19).Value
        .Cells(outRow, 29).Value = trxWS.Cells(trxRow, 18).Value

        .Cells(outRow, 1).Value = "OPt" & (outRow - 1)
        .Cells(outRow, 2).Value = "ci" & (outRow - 1)
        .Cells(outRow, 7).Value = ""
        .Cells(outRow, 8).Value = "LOD"
        .Cells(outRow, 9).Value = ""
        .Cells(outRow, 10).Value = ""
        .Cells(outRow, 13).Value = DateAdd("yyyy", 4, buyDate)
        .Cells(outRow, 14).Value = ""
        .Cells(outRow, 16).Value = 0
        .Cells(outRow, 17).Value = 0
        .Cells(outRow, 19).Value = ""
        .Cells(outRow, 24).Value = ""
        .Cells(outRow, 25).Value = ""
        .Cells(outRow, 27).Value = 0
        .Cells(outRow, 28).Value = 0
        .Cells(outRow, 30).Value = 0
        .Cells(outRow, 31).Value = 0
        .Cells(outRow, 32).Value = 0
        .Cells(outRow, 33).Value = "Auto-loaded by M_Driver"
        .Cells(outRow, 34).Value = 0
        .Cells(outRow, 35).Value = buyDate
        .Cells(outRow, 36).Value = buyDate
        .Cells(outRow, 37).Value = buyDate
        .Cells(outRow, 38).Value = buyDate
        .Cells(outRow, 39).Value = buyDate
        .Cells(outRow, 40).Value = "Indemnity"
        .Cells(outRow, 41).Value = "Occurrence"
        .Cells(outRow, 42).Value = 0
        .Cells(outRow, 43).Value = 0
        .Cells(outRow, 44).Value = 0
        .Cells(outRow, 45).Value = 1
    End With
End Sub

'===================================================================================
' Clear every paste-in sheet in the (external) Import Template - content AND
' any stale red validation-failure highlighting from a previous run
'===================================================================================
Private Sub ClearTemplateSheets(tplWB As Workbook)
    ClearBelowHeader tplWB.Worksheets("Guidelines"), 2
    ClearBelowHeader tplWB.Worksheets("Profit and Loss"), 2
    ClearBelowHeader tplWB.Worksheets("BoundPortfolio"), 2
    ClearBelowHeader tplWB.Worksheets("Bound Portfolio Report"), 2
    ClearBelowHeader tplWB.Worksheets("Historical Events Adjustment"), 2
End Sub

Private Sub ClearBelowHeader(ws As Worksheet, firstDataRow As Long)
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    If lastRow >= firstDataRow Then
        With ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastRow, lastCol))
            .ClearContents
            .Interior.ColorIndex = xlColorIndexNone
        End With
    End If
End Sub

'===================================================================================
' Look up the ID (=ReportID) BoundPortfolio.BoundPortfolio() just inserted for this
' PortfolioID / ReportDate / Vendor. Returns 0 if nothing was inserted.
'===================================================================================
Private Function GetLatestReportID(tplWB As Workbook, portfolioID As Variant, _
                                    reportDateVal As Variant, vendorVal As Variant) As Long
    Dim conn As Object, rs As Object
    Dim server As String, database As String
    Dim sql As String

    On Error GoTo Fail
    server = tplWB.Worksheets("Setup").Cells(7, 4).Value
    database = tplWB.Worksheets("Setup").Cells(8, 4).Value

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Driver={SQL Server};Server=" & server & ";Database=" & database & ";Trusted_Connection=Yes;"

    sql = "SELECT MAX(ID) AS ID FROM BoundPortfolio WHERE [PortfolioID] = '" & portfolioID & _
          "' AND [ReportDate] = '" & Format(reportDateVal, "yyyy-mm-dd") & _
          "' AND [Vendor] = '" & vendorVal & "'"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn
    If Not rs.EOF Then
        If Not IsNull(rs("ID").Value) Then GetLatestReportID = rs("ID").Value
    End If
    rs.Close
    conn.Close
    Exit Function

Fail:
    GetLatestReportID = 0
End Function

'-----------------------------------------------------------------------------------
' Safe, read-only check: does this PortfolioID already have a row in BoundPortfolio?
' Used to skip re-pushing a test case that's already in SQL, rather than creating
' duplicates. This is intentionally NOT an upsert (no DELETE/UPDATE) until the
' foreign-key relationships between BoundPortfolio/BoundPortfolioDetails/Guidelines/
' BoundPortfolioSettings are confirmed - see the cleanup stored procedure notes.
' On any connection error, fails safe by returning False (proceeds with the push)
' rather than silently blocking every run if SQL is briefly unreachable.
'-----------------------------------------------------------------------------------
Private Function PortfolioAlreadyExists(tplWB As Workbook, portfolioID As Variant) As Boolean
    Dim conn As Object, rs As Object
    Dim server As String, database As String
    Dim sql As String

    On Error GoTo Fail
    server = tplWB.Worksheets("Setup").Cells(7, 4).Value
    database = tplWB.Worksheets("Setup").Cells(8, 4).Value

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Driver={SQL Server};Server=" & server & ";Database=" & database & ";Trusted_Connection=Yes;"

    sql = "SELECT COUNT(*) AS Cnt FROM BoundPortfolio WHERE [PortfolioID] = " & CLng(portfolioID)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn
    If Not rs.EOF Then
        PortfolioAlreadyExists = (rs("Cnt").Value > 0)
    End If
    rs.Close
    conn.Close
    Exit Function

Fail:
    PortfolioAlreadyExists = False   ' fail safe - don't block the run over a transient connection issue
End Function

'===================================================================================
' Driver sheet helpers (this workbook's own sheet)
'===================================================================================
Private Function GetOrCreateDriverSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(DRIVER_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = DRIVER_SHEET
        ws.Range("A1:S1").Value = Array("TestCaseID", "PortfolioID", "Import Status", "Import Time", _
            "Run Report?", "Vendor", "Fronting", "SecurityInstrument", "AsOfDate", "RunOff", _
            "IncludeExpired", "AEPType", "AEPCoTVARType", "Parameters", "Report Status", _
            "Output Location", "Result", "Export?", "Include in Test Run?")
        ws.Range("A1:S1").Font.Bold = True
        ws.Range("A1:S1").Interior.Color = RGB(68, 114, 196)
        ws.Range("A1:S1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A:S").AutoFit
        ws.Columns(DCOL_OUTPUTLOCATION).ColumnWidth = 45
        ws.Rows(1).AutoFilter

        ' ---- Export control panel, off to the side so it doesn't collide with the table ----
        ws.Range("U1").Value = "EXPORT CONTROLS"
        ws.Range("U1").Font.Bold = True
        ws.Range("U2").Value = "Mode:"
        ws.Range("V2").Value = "Separate"
        With ws.Range("V2").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Separate,Consolidated"
        End With
        ws.Range("U3").Value = "Run:"
        ws.Range("V3").Value = "Step4_ExportResults (Alt+F8)"
        ws.Range("V3").Font.Italic = True
        ws.Columns("U:V").AutoFit
    End If
    Set GetOrCreateDriverSheet = ws
End Function

'===================================================================================
' Workbook-open helper: reuses the workbook if it's already open
'===================================================================================
Private Function GetOrOpenWorkbook(path As String, ByRef wasAlreadyOpen As Boolean, readOnly As Boolean) As Workbook
    Dim wb As Workbook
    Dim fname As String
    fname = Mid$(path, InStrRev(path, "\") + 1)

    On Error Resume Next
    Set wb = Workbooks(fname)
    On Error GoTo 0

    If Not wb Is Nothing Then
        wasAlreadyOpen = True
    Else
        wasAlreadyOpen = False
        Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=readOnly)
    End If
    Set GetOrOpenWorkbook = wb
End Function
