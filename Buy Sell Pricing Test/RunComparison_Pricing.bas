Attribute VB_Name = "modComparisonMapped"
Option Explicit

' ============================================================
' Same approach as the BuySell tool's modComparison, adapted for
' cases where the Expected/Actual headers do NOT share a simple
' prefix (e.g. "Capital Cost (Tail) [CapitalConsumptionApproach]"
' vs "Out_Capital Consumption Approach - Capital Cost (Tail)").
' Column pairs are read explicitly from a Mapping sheet instead of
' being derived by stripping a prefix.
'
' All settings live on the Config sheet:
'   Expected Sheet Name, Actual Sheet Name, Expected Key Column,
'   Actual Key Column, Mapping Sheet Name, Mapping: Actual Header
'   Column, Mapping: Expected Header Column, Mapping Data Start Row,
'   Numeric Tolerance (Absolute), Numeric Tolerance (Relative),
'   Case-Insensitive Header Match.
'
' Writes one row per (Key, Column) comparison to the Comparison
' sheet, labeled with the actual Expected/Actual sheet names.
' ============================================================

Sub RunComparison()

    Dim wsExpected As Worksheet, wsActual As Worksheet, wsConfig As Worksheet
    Dim wsMap As Worksheet, wsCmp As Worksheet

    Dim expSheetName As String, actSheetName As String
    Dim expKeyHeader As String, actKeyHeader As String
    Dim mapSheetName As String, mapActCol As Long, mapExpCol As Long, mapStartRow As Long
    Dim absTolerance As Double, relTolerance As Double
    Dim caseInsensitive As Boolean

    Dim expKeyCol As Long, actKeyCol As Long
    Dim expLastRow As Long, actLastRow As Long, mapLastRow As Long
    Dim r As Long, outRow As Long
    Dim totalCompared As Long, totalMismatch As Long

    ' ---- read config ----
    Set wsConfig = ThisWorkbook.Sheets("Config")
    expSheetName = Trim(wsConfig.Range("B2").Value)
    actSheetName = Trim(wsConfig.Range("B3").Value)
    expKeyHeader = Trim(wsConfig.Range("B4").Value)
    actKeyHeader = Trim(wsConfig.Range("B5").Value)
    mapSheetName = Trim(wsConfig.Range("B6").Value)
    mapActCol = Range(Trim(wsConfig.Range("B7").Value) & "1").Column
    mapExpCol = Range(Trim(wsConfig.Range("B8").Value) & "1").Column
    mapStartRow = CLng(wsConfig.Range("B9").Value)
    absTolerance = CDbl(wsConfig.Range("B10").Value)
    relTolerance = CDbl(wsConfig.Range("B11").Value)
    caseInsensitive = (UCase(Trim(wsConfig.Range("B12").Value)) = "Y")

    If Not SheetExists(expSheetName) Then
        MsgBox "Expected sheet '" & expSheetName & "' not found. Check Config!B2.", vbCritical
        Exit Sub
    End If
    If Not SheetExists(actSheetName) Then
        MsgBox "Actual sheet '" & actSheetName & "' not found. Check Config!B3.", vbCritical
        Exit Sub
    End If
    If Not SheetExists(mapSheetName) Then
        MsgBox "Mapping sheet '" & mapSheetName & "' not found. Check Config!B6.", vbCritical
        Exit Sub
    End If

    Set wsExpected = ThisWorkbook.Sheets(expSheetName)
    Set wsActual = ThisWorkbook.Sheets(actSheetName)
    Set wsMap = ThisWorkbook.Sheets(mapSheetName)
    Set wsCmp = ThisWorkbook.Sheets("Comparison")

    ' ---- header -> column lookups ----
    Dim expHeaderCol As Object, actHeaderCol As Object
    Set expHeaderCol = CreateObject("Scripting.Dictionary")
    Set actHeaderCol = CreateObject("Scripting.Dictionary")

    Dim c As Long, lastCol As Long, hv As String
    lastCol = wsExpected.Cells(1, wsExpected.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        hv = Trim(CStr(wsExpected.Cells(1, c).Value))
        If Len(hv) > 0 Then expHeaderCol(HKey(hv, caseInsensitive)) = c
    Next c
    lastCol = wsActual.Cells(1, wsActual.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        hv = Trim(CStr(wsActual.Cells(1, c).Value))
        If Len(hv) > 0 Then actHeaderCol(HKey(hv, caseInsensitive)) = c
    Next c

    expKeyCol = 0: actKeyCol = 0
    If expHeaderCol.Exists(HKey(expKeyHeader, caseInsensitive)) Then expKeyCol = expHeaderCol(HKey(expKeyHeader, caseInsensitive))
    If actHeaderCol.Exists(HKey(actKeyHeader, caseInsensitive)) Then actKeyCol = actHeaderCol(HKey(actKeyHeader, caseInsensitive))

    If expKeyCol = 0 Then
        MsgBox "Could not find key column '" & expKeyHeader & "' in sheet '" & expSheetName & "'.", vbCritical
        Exit Sub
    End If
    If actKeyCol = 0 Then
        MsgBox "Could not find key column '" & actKeyHeader & "' in sheet '" & actSheetName & "'.", vbCritical
        Exit Sub
    End If

    ' ---- read mapping pairs ----
    mapLastRow = wsMap.Cells(wsMap.Rows.Count, mapActCol).End(xlUp).Row
    Dim nPairs As Long
    Dim pairActHeader() As String, pairExpHeader() As String
    ReDim pairActHeader(1 To mapLastRow)
    ReDim pairExpHeader(1 To mapLastRow)
    nPairs = 0
    For r = mapStartRow To mapLastRow
        If Len(Trim(CStr(wsMap.Cells(r, mapActCol).Value))) > 0 And _
           Len(Trim(CStr(wsMap.Cells(r, mapExpCol).Value))) > 0 Then
            nPairs = nPairs + 1
            pairActHeader(nPairs) = Trim(CStr(wsMap.Cells(r, mapActCol).Value))
            pairExpHeader(nPairs) = Trim(CStr(wsMap.Cells(r, mapExpCol).Value))
        End If
    Next r

    ' ---- actual key -> row (first match wins) ----
    Dim actKeyRow As Object
    Set actKeyRow = CreateObject("Scripting.Dictionary")
    actLastRow = wsActual.Cells(wsActual.Rows.Count, actKeyCol).End(xlUp).Row
    Dim ar As Long, kv As String
    For ar = 2 To actLastRow
        kv = Trim(CStr(wsActual.Cells(ar, actKeyCol).Value))
        If Len(kv) > 0 Then
            If Not actKeyRow.Exists(kv) Then actKeyRow(kv) = ar
        End If
    Next ar

    ' ---- reset Comparison sheet ----
    wsCmp.Rows("2:" & wsCmp.Rows.Count).ClearContents
    wsCmp.Rows("2:" & wsCmp.Rows.Count).Interior.ColorIndex = 0
    wsCmp.Cells(1, 3).Value = expSheetName
    wsCmp.Cells(1, 4).Value = actSheetName

    Application.ScreenUpdating = False
    Application.StatusBar = "Comparing " & expSheetName & " to " & actSheetName & "..."

    outRow = 1
    totalCompared = 0
    totalMismatch = 0

    Dim er As Long, matchedActualRow As Long
    Dim i As Long
    Dim expVal As Variant, actVal As Variant
    Dim isMatch As Boolean, diffVal As Variant, noteStr As String

    expLastRow = wsExpected.Cells(wsExpected.Rows.Count, expKeyCol).End(xlUp).Row

    For er = 2 To expLastRow
        kv = Trim(CStr(wsExpected.Cells(er, expKeyCol).Value))
        If kv = "" Then GoTo NextRow

        If actKeyRow.Exists(kv) Then
            matchedActualRow = actKeyRow(kv)
        Else
            matchedActualRow = 0
        End If

        If matchedActualRow = 0 Then
            outRow = outRow + 1
            wsCmp.Cells(outRow, 1).Value = kv
            wsCmp.Cells(outRow, 2).Value = "(ALL COLUMNS)"
            wsCmp.Cells(outRow, 6).Value = "NO MATCH"
            wsCmp.Cells(outRow, 7).Value = "No row with this key found in " & actSheetName
            HighlightRow wsCmp, outRow, RGB(255, 199, 206)
            totalMismatch = totalMismatch + 1
            GoTo NextRow
        End If

        For i = 1 To nPairs
            expVal = wsExpected.Cells(er, expHeaderCol(HKey(pairExpHeader(i), caseInsensitive))).Value
            actVal = wsActual.Cells(matchedActualRow, actHeaderCol(HKey(pairActHeader(i), caseInsensitive))).Value

            CompareValues expVal, actVal, absTolerance, relTolerance, isMatch, diffVal, noteStr

            outRow = outRow + 1
            wsCmp.Cells(outRow, 1).Value = kv
            wsCmp.Cells(outRow, 2).Value = pairActHeader(i)
            wsCmp.Cells(outRow, 3).Value = expVal
            wsCmp.Cells(outRow, 4).Value = actVal
            wsCmp.Cells(outRow, 5).Value = diffVal
            wsCmp.Cells(outRow, 6).Value = IIf(isMatch, "MATCH", "MISMATCH")
            wsCmp.Cells(outRow, 7).Value = noteStr

            totalCompared = totalCompared + 1
            If Not isMatch Then
                totalMismatch = totalMismatch + 1
                HighlightRow wsCmp, outRow, RGB(255, 199, 206)
            End If
        Next i
NextRow:
    Next er

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Comparison complete." & vbCrLf & _
           totalCompared & " value(s) compared." & vbCrLf & _
           totalMismatch & " mismatch(es) found.", vbInformation

End Sub

Private Function HKey(s As String, ci As Boolean) As String
    If ci Then HKey = UCase(Trim(s)) Else HKey = Trim(s)
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Sub HighlightRow(ws As Worksheet, r As Long, clr As Long)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = clr
End Sub

' Compares two variant cell values (numeric — including text numbers with
' thousands separators like "2,500" — or text) and returns whether they
' match, the numeric difference (Actual - Expected), and an explanatory note.
' A pair matches if it is within the absolute tolerance OR the relative
' tolerance (relative to the Expected value).
Private Sub CompareValues(expVal As Variant, actVal As Variant, absTol As Double, relTol As Double, _
                           ByRef isMatch As Boolean, ByRef diffVal As Variant, ByRef noteStr As String)

    Dim expBlank As Boolean, actBlank As Boolean
    expBlank = (Trim(CStr(expVal)) = "")
    actBlank = (Trim(CStr(actVal)) = "")

    diffVal = ""
    noteStr = ""

    If expBlank And actBlank Then
        isMatch = True
        Exit Sub
    End If

    If expBlank Or actBlank Then
        isMatch = False
        noteStr = "One side is blank"
        Exit Sub
    End If

    Dim expNum As Double, actNum As Double
    Dim expOk As Boolean, actOk As Boolean
    expOk = TryCleanNumeric(expVal, expNum)
    actOk = TryCleanNumeric(actVal, actNum)

    If expOk And actOk Then
        Dim d As Double
        d = actNum - expNum
        diffVal = d
        Dim relOk As Boolean
        If expNum <> 0 Then
            relOk = (Abs(d) / Abs(expNum) <= relTol)
        Else
            relOk = (actNum = 0)
        End If
        isMatch = (Abs(d) <= absTol) Or relOk
    Else
        If Trim(CStr(expVal)) = Trim(CStr(actVal)) Then
            isMatch = True
        Else
            isMatch = False
            noteStr = "Text values differ"
        End If
    End If

End Sub

' Parses a value that may be a genuine number or text with thousands
' separators (e.g. "2,500") into a Double. Returns False for non-numeric text.
Private Function TryCleanNumeric(v As Variant, ByRef result As Double) As Boolean
    Dim s As String
    If IsNumeric(v) Then
        result = CDbl(v)
        TryCleanNumeric = True
        Exit Function
    End If
    s = Replace(CStr(v), ",", "")
    If IsNumeric(s) Then
        result = CDbl(s)
        TryCleanNumeric = True
    Else
        TryCleanNumeric = False
    End If
End Function
