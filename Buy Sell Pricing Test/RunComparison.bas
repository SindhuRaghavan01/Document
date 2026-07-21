Attribute VB_Name = "modComparison"
Option Explicit

' ============================================================
' Compares an "Expected" sheet against an "Actual" sheet. Rows
' are matched by a key column (names may differ between the two
' sheets). Columns are matched by stripping the configured
' Header Prefix from EITHER side's header (whichever has it)
' and comparing what remains.
'
' Writes one row per (Key, ColumnName) comparison to the
' Comparison sheet. The two value columns there (C and D) are
' labeled with the actual Expected/Actual sheet names from
' Config, so the output is self-describing no matter what the
' sheets are called or which one carries the prefix.
' ============================================================

Sub RunComparison()

    Dim wsExpected As Worksheet, wsActual As Worksheet, wsConfig As Worksheet, wsCmp As Worksheet
    Dim expSheetName As String, actSheetName As String
    Dim expKeyHeader As String, actKeyHeader As String
    Dim prefix As String
    Dim tolerance As Double
    Dim caseInsensitive As Boolean

    Dim expLastRow As Long, expLastCol As Long
    Dim actLastRow As Long, actLastCol As Long
    Dim r As Long, c As Long
    Dim expKeyCol As Long, actKeyCol As Long

    Dim expHeaders() As String, actHeaders() As String
    Dim expBaseHeaders() As String, actBaseHeaders() As String   ' headers with prefix stripped, either side

    Dim outRow As Long
    Dim totalCompared As Long, totalMismatch As Long

    ' ---- read config ----
    Set wsConfig = ThisWorkbook.Sheets("Config")
    expSheetName = Trim(wsConfig.Range("B2").Value)
    actSheetName = Trim(wsConfig.Range("B3").Value)
    expKeyHeader = Trim(wsConfig.Range("B4").Value)
    actKeyHeader = Trim(wsConfig.Range("B5").Value)
    prefix = Trim(wsConfig.Range("B6").Value)
    tolerance = CDbl(wsConfig.Range("B7").Value)
    caseInsensitive = (UCase(Trim(wsConfig.Range("B8").Value)) = "Y")

    If Not SheetExists(expSheetName) Then
        MsgBox "Expected sheet '" & expSheetName & "' not found. Check Config!B2.", vbCritical
        Exit Sub
    End If
    If Not SheetExists(actSheetName) Then
        MsgBox "Actual sheet '" & actSheetName & "' not found. Check Config!B3.", vbCritical
        Exit Sub
    End If

    Set wsExpected = ThisWorkbook.Sheets(expSheetName)
    Set wsActual = ThisWorkbook.Sheets(actSheetName)
    Set wsCmp = ThisWorkbook.Sheets("Comparison")

    expLastRow = wsExpected.Cells(wsExpected.Rows.Count, 1).End(xlUp).Row
    expLastCol = wsExpected.Cells(1, wsExpected.Columns.Count).End(xlToLeft).Column
    actLastRow = wsActual.Cells(wsActual.Rows.Count, 1).End(xlUp).Row
    actLastCol = wsActual.Cells(1, wsActual.Columns.Count).End(xlToLeft).Column

    ' ---- read headers, strip prefix on whichever side has it ----
    ReDim expHeaders(1 To expLastCol)
    ReDim expBaseHeaders(1 To expLastCol)
    For c = 1 To expLastCol
        expHeaders(c) = Trim(CStr(wsExpected.Cells(1, c).Value))
        expBaseHeaders(c) = StripPrefix(expHeaders(c), prefix)
    Next c

    ReDim actHeaders(1 To actLastCol)
    ReDim actBaseHeaders(1 To actLastCol)
    For c = 1 To actLastCol
        actHeaders(c) = Trim(CStr(wsActual.Cells(1, c).Value))
        actBaseHeaders(c) = StripPrefix(actHeaders(c), prefix)
    Next c

    ' ---- find key columns (matched on raw header text, not stripped) ----
    expKeyCol = FindHeaderCol(expHeaders, expLastCol, expKeyHeader, caseInsensitive)
    actKeyCol = FindHeaderCol(actHeaders, actLastCol, actKeyHeader, caseInsensitive)

    If expKeyCol = 0 Then
        MsgBox "Could not find key column '" & expKeyHeader & "' in sheet '" & expSheetName & "'.", vbCritical
        Exit Sub
    End If
    If actKeyCol = 0 Then
        MsgBox "Could not find key column '" & actKeyHeader & "' in sheet '" & actSheetName & "'.", vbCritical
        Exit Sub
    End If

    ' ---- reset Comparison sheet, label value columns with the actual sheet names ----
    wsCmp.Rows("2:" & wsCmp.Rows.Count).ClearContents
    wsCmp.Rows("2:" & wsCmp.Rows.Count).Interior.ColorIndex = 0
    wsCmp.Cells(1, 3).Value = expSheetName
    wsCmp.Cells(1, 4).Value = actSheetName
    outRow = 1
    totalCompared = 0
    totalMismatch = 0

    Application.ScreenUpdating = False
    Application.StatusBar = "Comparing " & expSheetName & " to " & actSheetName & "..."

    Dim keyVal As String
    Dim er As Long, ar As Long
    Dim matchedActualRow As Long

    For er = 2 To expLastRow
        keyVal = Trim(CStr(wsExpected.Cells(er, expKeyCol).Value))
        If keyVal = "" Then GoTo NextExpRow

        matchedActualRow = 0
        For ar = 2 To actLastRow
            If Trim(CStr(wsActual.Cells(ar, actKeyCol).Value)) = keyVal Then
                matchedActualRow = ar
                Exit For
            End If
        Next ar

        If matchedActualRow = 0 Then
            outRow = outRow + 1
            wsCmp.Cells(outRow, 1).Value = keyVal
            wsCmp.Cells(outRow, 2).Value = "(ALL COLUMNS)"
            wsCmp.Cells(outRow, 6).Value = "NO MATCH"
            wsCmp.Cells(outRow, 7).Value = "No row with this key found in " & actSheetName
            HighlightRow wsCmp, outRow, RGB(255, 199, 206)
            totalMismatch = totalMismatch + 1
            GoTo NextExpRow
        End If

        ' compare every expected column (except the key column) against its actual counterpart,
        ' matching on the PREFIX-STRIPPED header text so it works regardless of which sheet
        ' carries the prefix
        For c = 1 To expLastCol
            If c <> expKeyCol And expBaseHeaders(c) <> "" Then
                Dim actCol As Long
                actCol = FindHeaderCol(actBaseHeaders, actLastCol, expBaseHeaders(c), caseInsensitive)

                If actCol > 0 Then
                    Dim expVal As Variant, actVal As Variant
                    Dim isMatch As Boolean
                    Dim diffStr As String, noteStr As String

                    expVal = wsExpected.Cells(er, c).Value
                    actVal = wsActual.Cells(matchedActualRow, actCol).Value

                    CompareValues expVal, actVal, tolerance, isMatch, diffStr, noteStr

                    outRow = outRow + 1
                    wsCmp.Cells(outRow, 1).Value = keyVal
                    wsCmp.Cells(outRow, 2).Value = expBaseHeaders(c)
                    wsCmp.Cells(outRow, 3).Value = expVal
                    wsCmp.Cells(outRow, 4).Value = actVal
                    wsCmp.Cells(outRow, 5).Value = diffStr
                    wsCmp.Cells(outRow, 6).Value = IIf(isMatch, "MATCH", "MISMATCH")
                    wsCmp.Cells(outRow, 7).Value = noteStr

                    totalCompared = totalCompared + 1
                    If Not isMatch Then
                        totalMismatch = totalMismatch + 1
                        HighlightRow wsCmp, outRow, RGB(255, 199, 206)
                    End If
                End If
            End If
        Next c

NextExpRow:
    Next er

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Comparison complete." & vbCrLf & _
           totalCompared & " value(s) compared." & vbCrLf & _
           totalMismatch & " mismatch(es) found.", vbInformation

End Sub

Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Strips the configured prefix from a header if present (case-insensitive),
' otherwise returns the header unchanged. Works whether the prefix lives on
' the Expected sheet's headers or the Actual sheet's headers (or neither).
Private Function StripPrefix(header As String, prefix As String) As String
    If Len(prefix) > 0 And Len(header) >= Len(prefix) Then
        If LCase(Left(header, Len(prefix))) = LCase(prefix) Then
            StripPrefix = Trim(Mid(header, Len(prefix) + 1))
            Exit Function
        End If
    End If
    StripPrefix = header
End Function

Private Function FindHeaderCol(headers() As String, lastCol As Long, target As String, caseInsensitive As Boolean) As Long
    Dim i As Long
    Dim t As String, h As String
    t = Trim(target)
    If caseInsensitive Then t = UCase(t)
    For i = 1 To lastCol
        h = headers(i)
        If caseInsensitive Then h = UCase(h)
        If h = t Then
            FindHeaderCol = i
            Exit Function
        End If
    Next i
    FindHeaderCol = 0
End Function

Private Sub HighlightRow(ws As Worksheet, r As Long, clr As Long)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = clr
End Sub

' Compares two variant cell values (numeric, text, blank) and returns
' whether they match, a difference string, and a note explaining any
' type mismatch or special case.
Private Sub CompareValues(expVal As Variant, actVal As Variant, tolerance As Double, _
                           ByRef isMatch As Boolean, ByRef diffStr As String, ByRef noteStr As String)

    Dim expBlank As Boolean, actBlank As Boolean
    expBlank = (Trim(CStr(expVal)) = "")
    actBlank = (Trim(CStr(actVal)) = "")

    diffStr = ""
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

    If IsNumeric(expVal) And IsNumeric(actVal) Then
        Dim d As Double
        d = CDbl(actVal) - CDbl(expVal)
        diffStr = Format(d, "0.############")
        isMatch = (Abs(d) <= tolerance)
    Else
        ' text-style comparison (e.g. "N/A", "Yes"/"No")
        If Trim(CStr(expVal)) = Trim(CStr(actVal)) Then
            isMatch = True
        Else
            isMatch = False
            noteStr = "Text values differ"
        End If
    End If

End Sub
