Attribute VB_Name = "MasterAutomation_OUTPUT"
Option Explicit

'==================================================================================
' HOW TO INSTALL / USE (in the OUTPUT / "Bound_Report" workbook)
' 1. TestCases sheet: row 1 = headers, one row per case from row 2, columns A:K
'    matching Inputs!A2:A12 top to bottom (PortfolioID, ReportDate, Vendor,
'    Fronting, SecurityInstrument, AsOfDate, RunOff, IncludeExpired, ExportID,
'    AEPType, AEPCoTVARType).
' 2. Alt+F11 -> select ALL the code in this module -> Delete -> paste this ENTIRE
'    file over it (all 6 procedures - RunAllCases, CallStoredProcedureWithRetry,
'    AppendUsedRangeBlock, AppendFixedRangeBlock, InitErrorLog, LogError). Don't
'    paste just one procedure - RunAllCases calls all the others.
' 3. Alt+F8 -> run RunAllCases.
' 4. Failures no longer pop up a dialog per error - they're written to an
'    "ErrorLog" sheet instead, so the run doesn't stop for you to click OK
'    repeatedly. Check that sheet after the run finishes.
' 5. If the stored procedure call fails with a connection/"object is closed"
'    type error, it's retried automatically (see RETRY_COUNT / RETRY_WAIT_SECONDS
'    below) before that case is logged as failed and skipped.
' 6. EP Curve and HistoricalEvents are intentionally excluded from SOURCE_SHEETS -
'    they're very large tabs. Add them back into the array if you want them.
' 7. Every row in every AC_ sheet is now tagged with the PortfolioID in column A
'    (the tab's own original content is shifted to start at column B), so you
'    can filter/sort/pivot by PortfolioID once you have many cases stacked.
'    If you already have AC_ sheets from an OLDER version of this macro (with
'    "=== Case X ===" label rows instead), delete them first so they rebuild
'    cleanly in the new format - don't mix old and new layouts on the same tab.
' 9. Each case's block now starts right at that tab's real header row - the
'    leading title/blank rows above the header (e.g. report titles) are
'    skipped entirely, so cases stack with no wasted blank rows between them.
'    If you have AC_ sheets from an EARLIER version of this macro, delete them
'    first so they rebuild cleanly - don't mix old and new layouts.
' 10. P&L, Guidelines Summary, and AAL by Month are now each SPLIT into their
'     own dedicated AC_ sheets, based on the exact rows/columns of your current
'     template:
'       AC_PnL Report              <- P&L rows 3-20 (cols A-L)
'       AC_Portfolio Metrics       <- P&L rows 23-40 (cols A-C)
'       AC_GS Limit by SecInstr    <- Guidelines Summary rows 1-8 (cols A-G)
'       AC_GS PML Summary          <- Guidelines Summary rows 10-34 (cols A-G)
'       AC_GS Issuer               <- Guidelines Summary row 40 and below (cols A-E)
'       AC_AAL Monthly             <- AAL by Month, columns A-M
'       AC_AAL Premium             <- AAL by Month, columns P-R
'     Region Peril, BoundPositions, and Bound Portfolio Report are unchanged -
'     each still gets one single AC_ sheet as before.
'     If your team edits these source tabs again (adds/removes rows or moves
'     things around), the row/column numbers in the Call AppendFixedRangeBlock
'     lines inside RunAllCases will need updating to match.
' 11. RunAllCases now takes an optional pnlOnly parameter (defaults to False).
'     Call it plain (RunAllCases) or with pnlOnly:=False for the full report as
'     always. Call it with pnlOnly:=True to build ONLY "AC_PnL Report" and skip
'     every other section - useful for a quick check without waiting on the rest.
' 12. TestCases sheet now has a 12th column, L = TestCaseID. Column A is still
'     PortfolioID (required by the stored procedure - do not remove it), but
'     every AC_ sheet row and every ErrorLog entry is now tagged/labeled with
'     TestCaseID (column L) instead of the raw PortfolioID, since PortfolioID
'     alone isn't a human-readable/reliable identifier. If you drive this sheet
'     by hand rather than through the M_Driver workbook, make sure column L is
'     filled in for every row.
' 13. AC_PnL Report no longer includes the return-period header row (the one
'     reading "1.111 yr, 2 yr, 5 yr, 10 yr, 25 yr, 50 yr, 100 yr, 250 yr, 500 yr,
'     1000 yr") - it's excluded from the copy entirely, not just left unhighlighted.
'==================================================================================

Const RETRY_COUNT As Integer = 2                 ' extra attempts if the SP call fails
Const RETRY_WAIT_SECONDS As Integer = 5          ' pause before each retry
Const POST_CALL_WAIT_SECONDS As Integer = 3      ' pause after a successful SP call, before reading results

Sub RunAllCases(Optional pnlOnly As Boolean = False)

    Dim wsCases As Worksheet, wsInputs As Worksheet
    Dim lastRow As Long, i As Long
    Dim SOURCE_SHEETS As Variant
    Dim HEADER_ROW_OFFSET As Variant
    Dim s As Long
    Dim caseID As Variant
    Dim spOK As Boolean
    Dim successCount As Long, failCount As Long

    On Error Resume Next
    Set wsCases = ThisWorkbook.Sheets("TestCases")
    On Error GoTo 0
    If wsCases Is Nothing Then
        MsgBox "Could not find a sheet named 'TestCases'. Please create it first.", vbCritical
        Exit Sub
    End If

    Set wsInputs = ThisWorkbook.Sheets("Inputs")
    InitErrorLog

    SOURCE_SHEETS = Array( _
        "Region Peril", _
        "BoundPositions", _
        "Bound Portfolio Report" _
    )

    ' Row (relative to the top of each tab's used range, 1 = first row) where the
    ' real column headers sit on that tab.
    HEADER_ROW_OFFSET = Array(3, 4, 3)

    lastRow = wsCases.Cells(wsCases.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No cases found in TestCases sheet (need data starting row 2).", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    successCount = 0
    failCount = 0

    For i = 2 To lastRow

        ' caseID = the label used to tag every AC_ row and ErrorLog entry -
        ' TestCaseID (column L), human-readable and unique. This is DIFFERENT
        ' from the PortfolioID sent to the stored procedure below (column A),
        ' which must stay a real PortfolioID for the SQL lookup to work.
        caseID = wsCases.Cells(i, "L").Value
        If Trim$(CStr(caseID)) = "" Then caseID = wsCases.Cells(i, "A").Value  ' fallback if column L is blank

        ' ---- 1. Push this case's inputs into the Inputs sheet ----
        wsInputs.Range("B2").Value = wsCases.Cells(i, "A").Value  ' PortfolioID - required by the SP, unchanged
        wsInputs.Range("B3").Value = wsCases.Cells(i, "B").Value  ' ReportDate
        wsInputs.Range("B4").Value = wsCases.Cells(i, "C").Value  ' Vendor
        wsInputs.Range("B5").Value = wsCases.Cells(i, "D").Value  ' Fronting
        wsInputs.Range("B6").Value = wsCases.Cells(i, "E").Value  ' SecurityInstrument
        wsInputs.Range("B7").Value = wsCases.Cells(i, "F").Value  ' AsOfDate
        wsInputs.Range("B8").Value = wsCases.Cells(i, "G").Value  ' RunOff
        wsInputs.Range("B9").Value = wsCases.Cells(i, "H").Value  ' IncludeExpired
        wsInputs.Range("B10").Value = wsCases.Cells(i, "I").Value ' ExportID
        wsInputs.Range("B11").Value = wsCases.Cells(i, "J").Value ' AEPType
        wsInputs.Range("B12").Value = wsCases.Cells(i, "K").Value ' AEPCoTVARType

        ' ---- 2. Run the SAME macro the "Run Stored Procedure" button calls,
        '         retrying on failure (handles the "object is closed" connection error) ----
        spOK = CallStoredProcedureWithRetry(caseID)

        If Not spOK Then
            failCount = failCount + 1
            GoTo NextCase
        End If

        ' ---- 3. Append the "yellow" tabs as-is (single AC_ sheet each) ----
        ' Skipped entirely when pnlOnly is True.
        If Not pnlOnly Then
            For s = LBound(SOURCE_SHEETS) To UBound(SOURCE_SHEETS)
                Call AppendUsedRangeBlock(caseID, CStr(SOURCE_SHEETS(s)), CInt(HEADER_ROW_OFFSET(s)))
            Next s
        End If

        ' ---- 4. Split the "blue" tabs into their own dedicated AC_ sheets ----
        ' P&L -> AC_PnL Report ALWAYS runs, pnlOnly or not.
        ' excludeRowInSource:=4 drops the return-period header row ("1.111 yr,
        ' 2 yr, 5 yr, ... 1000 yr") entirely from the output - not just
        ' unhighlighted, genuinely not copied. numHeaderRows is 1 now since
        ' only the first header row (row 3) remains.
        Call AppendFixedRangeBlock(caseID, "P&L", "AC_PnL Report", _
             3, 20, False, 1, 12, 1, 1, excludeRowInSource:=4)

        ' Everything below here is skipped when pnlOnly is True.
        If Not pnlOnly Then
            Call AppendFixedRangeBlock(caseID, "P&L", "AC_Portfolio Metrics", _
                 23, 40, False, 1, 3, 3, 1)

            ' Guidelines Summary -> three sub-tables
            Call AppendFixedRangeBlock(caseID, "Guidelines Summary", "AC_GS Limit by SecInstr", _
                 1, 8, False, 1, 7, 2, 1)
            Call AppendFixedRangeBlock(caseID, "Guidelines Summary", "AC_GS PML Summary", _
                 10, 34, False, 1, 7, 2, 1)
            Call AppendFixedRangeBlock(caseID, "Guidelines Summary", "AC_GS Issuer", _
                 40, 0, True, 1, 5, 1, 1)

            ' AAL by Month -> same rows, two different column ranges
            Call AppendFixedRangeBlock(caseID, "AAL by Month", "AC_AAL Monthly", _
                 2, 0, True, 1, 13, 1, 1)
            Call AppendFixedRangeBlock(caseID, "AAL by Month", "AC_AAL Premium", _
                 2, 0, True, 16, 18, 1, 1)
        End If

        successCount = successCount + 1

NextCase:
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox successCount & " case(s) completed successfully, " & failCount & _
           " case(s) failed." & vbCrLf & _
           "See sheets prefixed 'AC_' for stacked results, and the 'ErrorLog' " & _
           "sheet for details on any failures.", vbInformation

End Sub


' Calls Sheet13.extractDataandCallSP, retrying a few times if it errors
' (e.g. "Operation is not allowed when the object is closed"). Returns True
' if the call eventually succeeded, False if it failed on every attempt.
' All failures are logged - nothing pops up a dialog.
Function CallStoredProcedureWithRetry(caseID As Variant) As Boolean

    Dim attempt As Integer
    Dim lastErrNum As Long, lastErrDesc As String
    Dim wsInputs As Worksheet
    Dim exportIDBefore As Variant, exportIDAfter As Variant

    CallStoredProcedureWithRetry = False
    Set wsInputs = ThisWorkbook.Sheets("Inputs")

    For attempt = 1 To (RETRY_COUNT + 1)

        exportIDBefore = wsInputs.Range("B10").Value

        On Error Resume Next
        Err.Clear
        Call Sheet13.extractDataandCallSP
        lastErrNum = Err.Number
        lastErrDesc = Err.Description
        On Error GoTo 0

        exportIDAfter = wsInputs.Range("B10").Value

        ' The vendor macro can swallow its own DB errors and show its own MsgBox,
        ' so Err.Number alone isn't reliable proof of success. Also require
        ' ExportID (Inputs!B10) to have actually changed/populated as a sign
        ' the stored procedure really ran and wrote back a result.
        If lastErrNum = 0 And Not IsEmpty(exportIDAfter) And exportIDAfter <> exportIDBefore Then
            DoEvents
            Application.Wait Now + TimeValue("0:00:0" & POST_CALL_WAIT_SECONDS)
            CallStoredProcedureWithRetry = True
            Exit Function
        Else
            If lastErrNum <> 0 Then
                LogError caseID, "Sheet13.extractDataandCallSP", _
                         "Attempt " & attempt & " of " & (RETRY_COUNT + 1) & " - " & _
                         "VBA error " & lastErrNum & " - " & lastErrDesc
            Else
                LogError caseID, "Sheet13.extractDataandCallSP", _
                         "Attempt " & attempt & " of " & (RETRY_COUNT + 1) & " - " & _
                         "No VBA error, but ExportID did not change (still " & exportIDAfter & _
                         "). The stored procedure likely failed internally " & _
                         "(e.g. its own 'SP failed' popup) and swallowed the error."
            End If
            If attempt <= RETRY_COUNT Then
                DoEvents
                Application.Wait Now + TimeValue("0:00:0" & RETRY_WAIT_SECONDS)
            End If
        End If

    Next attempt

    ' All attempts failed
    LogError caseID, "Sheet13.extractDataandCallSP", _
             "GAVE UP after " & (RETRY_COUNT + 1) & " attempts - case skipped entirely, " & _
             "no tabs will be appended for this case."

End Function


' Copies srcName's data starting from its real header row (per headerRowOffset)
' down through its last used row - into a same-named "AC_<srcName>" sheet, with a
' PortfolioID column (A) tagging EVERY row of this case's block. The leading
' title/blank rows above the header (e.g. report titles) are intentionally
' skipped so cases stack cleanly with no wasted blank rows between them.
' Creates the AC_ sheet (with a "PortfolioID" header in A1) on first use.
' The header row of every case's block gets a color fill + bold.
' Any failure here is caught and LOGGED (no popup) - it does NOT stop the rest of the run.
Sub AppendUsedRangeBlock(caseID As Variant, srcName As String, headerRowOffset As Integer)

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim dstName As String
    Dim nextRow As Long
    Dim numRows As Long, numCols As Long
    Dim isNewSheet As Boolean
    Dim srcTopRow As Long, srcBottomRow As Long, srcLeftCol As Long, srcRightCol As Long
    Dim srcRange As Range

    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsSrc = ThisWorkbook.Sheets(srcName)
    On Error GoTo ErrHandler
    If wsSrc Is Nothing Then
        LogError caseID, srcName, "Source sheet not found - skipped."
        Exit Sub
    End If

    dstName = "AC_" & srcName
    If Len(dstName) > 31 Then dstName = Left(dstName, 31)

    isNewSheet = False
    On Error Resume Next
    Set wsDst = ThisWorkbook.Sheets(dstName)
    On Error GoTo ErrHandler
    If wsDst Is Nothing Then
        Set wsDst = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDst.Name = dstName
        isNewSheet = True
    End If

    On Error Resume Next
    wsDst.Unprotect
    On Error GoTo ErrHandler

    If isNewSheet Then
        wsDst.Cells(1, 1).Value = "TestCaseID"
        wsDst.Cells(1, 1).Interior.Color = RGB(68, 114, 196)
        wsDst.Cells(1, 1).Font.Color = RGB(255, 255, 255)
        wsDst.Cells(1, 1).Font.Bold = True
        nextRow = 2
    Else
        nextRow = wsDst.Cells(wsDst.Rows.Count, "B").End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
    End If

    If nextRow < 1 Or nextRow > wsDst.Rows.Count Then
        LogError caseID, srcName, "Calculated row (" & nextRow & ") out of range - skipped."
        Exit Sub
    End If

    ' Start the copy at the real header row (skip title/blank rows above it)
    srcTopRow = wsSrc.UsedRange.Row + headerRowOffset - 1
    srcBottomRow = wsSrc.UsedRange.Row + wsSrc.UsedRange.Rows.Count - 1
    srcLeftCol = wsSrc.UsedRange.Column
    srcRightCol = wsSrc.UsedRange.Column + wsSrc.UsedRange.Columns.Count - 1

    If srcTopRow > srcBottomRow Then
        LogError caseID, srcName, "Header row offset (" & headerRowOffset & ") is beyond the sheet's used range - skipped."
        Exit Sub
    End If

    ' Trim fully-blank rows off the BOTTOM only (leaves any internal structural
    ' gaps within the report - e.g. between sub-tables - untouched, since those
    ' are consistent, genuine parts of the report layout, not artifacts).
    Do While srcBottomRow > srcTopRow
        If Application.WorksheetFunction.CountA(wsSrc.Range(wsSrc.Cells(srcBottomRow, srcLeftCol), wsSrc.Cells(srcBottomRow, srcRightCol))) = 0 Then
            srcBottomRow = srcBottomRow - 1
        Else
            Exit Do
        End If
    Loop

    Set srcRange = wsSrc.Range(wsSrc.Cells(srcTopRow, srcLeftCol), wsSrc.Cells(srcBottomRow, srcRightCol))
    numRows = srcRange.Rows.Count
    numCols = srcRange.Columns.Count

    srcRange.Copy
    wsDst.Cells(nextRow, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    ' Tag every row of this block with the PortfolioID in column A
    wsDst.Range(wsDst.Cells(nextRow, 1), wsDst.Cells(nextRow + numRows - 1, 1)).Value = caseID

    ' Highlight this case's header row - it's now always the FIRST row of the block,
    ' since we started the copy at the header itself.
    With wsDst.Range(wsDst.Cells(nextRow, 1), wsDst.Cells(nextRow, numCols + 1))
        .Interior.Color = RGB(221, 235, 247)   ' light blue fill
        .Font.Bold = True
    End With

    Exit Sub

ErrHandler:
    Application.CutCopyMode = False
    LogError caseID, srcName, "Error " & Err.Number & " - " & Err.Description
    Resume Next

End Sub


' Copies a FIXED row/column range from srcName (e.g. one sub-table within a tab
' that has several) into its own dedicated "destName" AC_ sheet, with a
' PortfolioID column (A) tagging every row - same as AppendUsedRangeBlock, but
' for a specific sub-region instead of the whole tab's used range.
'
' rowStart/rowEnd: the fixed row numbers on the SOURCE sheet to capture.
' dynamicEnd: if True, ignore rowEnd and instead find the last used row
'             (scanning down colStart) - use this when a sub-table's row count
'             varies by case (e.g. a variable-length Issuer or instrument list).
' colStart/colEnd: fixed column numbers on the source sheet to capture.
' headerRowOffset: which row WITHIN this block (1 = first row of the block) is
'                  the real column-header row, for highlighting.
' numHeaderRows: how many consecutive rows starting at headerRowOffset make up
'                the header (1 for a normal single header row, 2 for a tab like
'                PnL Report that has a two-row header - percentile row + period row).
' excludeRowInSource: an absolute row number on the SOURCE sheet to drop from the
'                copy entirely (not just skip highlighting) - e.g. a header row
'                you don't want appearing in the output at all. 0 = none excluded.
Sub AppendFixedRangeBlock(caseID As Variant, srcName As String, destName As String, _
                          rowStart As Long, rowEnd As Long, dynamicEnd As Boolean, _
                          colStart As Long, colEnd As Long, _
                          headerRowOffset As Integer, numHeaderRows As Integer, _
                          Optional excludeRowInSource As Long = 0)

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim dstName As String
    Dim nextRow As Long
    Dim numRows As Long, numCols As Long
    Dim isNewSheet As Boolean
    Dim srcTopRow As Long, srcBottomRow As Long
    Dim srcRange As Range
    Dim hRow As Integer

    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsSrc = ThisWorkbook.Sheets(srcName)
    On Error GoTo ErrHandler
    If wsSrc Is Nothing Then
        LogError caseID, destName, "Source sheet '" & srcName & "' not found - skipped."
        Exit Sub
    End If

    dstName = destName
    If Len(dstName) > 31 Then dstName = Left(dstName, 31)

    isNewSheet = False
    On Error Resume Next
    Set wsDst = ThisWorkbook.Sheets(dstName)
    On Error GoTo ErrHandler
    If wsDst Is Nothing Then
        Set wsDst = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDst.Name = dstName
        isNewSheet = True
    End If

    On Error Resume Next
    wsDst.Unprotect
    On Error GoTo ErrHandler

    If isNewSheet Then
        wsDst.Cells(1, 1).Value = "TestCaseID"
        wsDst.Cells(1, 1).Interior.Color = RGB(68, 114, 196)
        wsDst.Cells(1, 1).Font.Color = RGB(255, 255, 255)
        wsDst.Cells(1, 1).Font.Bold = True
        nextRow = 2
    Else
        nextRow = wsDst.Cells(wsDst.Rows.Count, "B").End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
    End If

    If nextRow < 1 Or nextRow > wsDst.Rows.Count Then
        LogError caseID, destName, "Calculated row (" & nextRow & ") out of range - skipped."
        Exit Sub
    End If

    srcTopRow = rowStart

    If dynamicEnd Then
        srcBottomRow = wsSrc.Cells(wsSrc.Rows.Count, colStart).End(xlUp).Row
        If srcBottomRow < srcTopRow Then srcBottomRow = srcTopRow
    Else
        srcBottomRow = rowEnd
    End If

    ' Trim fully-blank rows off the BOTTOM only
    Do While srcBottomRow > srcTopRow
        If Application.WorksheetFunction.CountA(wsSrc.Range(wsSrc.Cells(srcBottomRow, colStart), wsSrc.Cells(srcBottomRow, colEnd))) = 0 Then
            srcBottomRow = srcBottomRow - 1
        Else
            Exit Do
        End If
    Loop

    Set srcRange = wsSrc.Range(wsSrc.Cells(srcTopRow, colStart), wsSrc.Cells(srcBottomRow, colEnd))

    ' Drop excludeRowInSource entirely, if it falls within this range - build a
    ' Union of the rows above and below it. Excel pastes a Union of whole,
    ' column-aligned row-ranges contiguously (no gap left for the skipped row).
    If excludeRowInSource > 0 And excludeRowInSource >= srcTopRow And excludeRowInSource <= srcBottomRow Then
        Dim topPart As Range, bottomPart As Range
        Set topPart = Nothing
        Set bottomPart = Nothing
        If excludeRowInSource > srcTopRow Then
            Set topPart = wsSrc.Range(wsSrc.Cells(srcTopRow, colStart), wsSrc.Cells(excludeRowInSource - 1, colEnd))
        End If
        If excludeRowInSource < srcBottomRow Then
            Set bottomPart = wsSrc.Range(wsSrc.Cells(excludeRowInSource + 1, colStart), wsSrc.Cells(srcBottomRow, colEnd))
        End If
        If Not topPart Is Nothing And Not bottomPart Is Nothing Then
            Set srcRange = Union(topPart, bottomPart)
        ElseIf Not topPart Is Nothing Then
            Set srcRange = topPart
        ElseIf Not bottomPart Is Nothing Then
            Set srcRange = bottomPart
        End If
    End If

    numCols = colEnd - colStart + 1
    If srcRange.Areas.Count > 1 Then
        numRows = 0
        Dim areaIdx As Long
        For areaIdx = 1 To srcRange.Areas.Count
            numRows = numRows + srcRange.Areas(areaIdx).Rows.Count
        Next areaIdx
    Else
        numRows = srcRange.Rows.Count
    End If

    srcRange.Copy
    wsDst.Cells(nextRow, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    ' Tag every row of this block with the PortfolioID in column A
    wsDst.Range(wsDst.Cells(nextRow, 1), wsDst.Cells(nextRow + numRows - 1, 1)).Value = caseID

    ' Highlight the header row(s) for this case's block
    For hRow = 0 To numHeaderRows - 1
        If headerRowOffset + hRow <= numRows Then
            With wsDst.Range(wsDst.Cells(nextRow + headerRowOffset - 1 + hRow, 1), _
                              wsDst.Cells(nextRow + headerRowOffset - 1 + hRow, numCols + 1))
                .Interior.Color = RGB(221, 235, 247)   ' light blue fill
                .Font.Bold = True
            End With
        End If
    Next hRow

    Exit Sub

ErrHandler:
    Application.CutCopyMode = False
    LogError caseID, destName, "Error " & Err.Number & " - " & Err.Description
    Resume Next

End Sub


' ---- Simple error log sheet, no popups ----

Sub InitErrorLog()
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("ErrorLog")
    On Error GoTo 0
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "ErrorLog"
        wsLog.Range("A1:D1").Value = Array("Timestamp", "CaseID", "Tab/Step", "Message")
    End If
End Sub

Sub LogError(caseID As Variant, tabName As String, message As String)
    Dim wsLog As Worksheet
    Dim r As Long
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("ErrorLog")
    On Error GoTo 0
    If wsLog Is Nothing Then
        InitErrorLog
        Set wsLog = ThisWorkbook.Sheets("ErrorLog")
    End If
    r = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(r, 1).Value = Now
    wsLog.Cells(r, 2).Value = caseID
    wsLog.Cells(r, 3).Value = tabName
    wsLog.Cells(r, 4).Value = message
End Sub
