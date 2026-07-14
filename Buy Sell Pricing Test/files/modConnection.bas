Attribute VB_Name = "modConnection"
Option Explicit

' Uses late binding (CreateObject) so this works even if the
' "Microsoft ActiveX Data Objects x.x Library" reference isn't
' checked under Tools > References on a colleague's machine.

Public Function GetDBConnection() As Object
    Dim conn As Object
    Dim connStr As String

    connStr = GetConfigValue("ConnectionString")
    If Len(Trim(connStr)) = 0 Then
        MsgBox "ConnectionString is blank on the Config sheet.", vbCritical
        Set GetDBConnection = Nothing
        Exit Function
    End If

    On Error GoTo ConnErr
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = CLng(Val(GetConfigValue("TimeoutSeconds")))
    conn.Open connStr
    Set GetDBConnection = conn
    Exit Function

ConnErr:
    MsgBox "Failed to connect to database:" & vbCrLf & Err.Description, vbCritical
    Set GetDBConnection = Nothing
End Function

' SQL Server can return several messages for one call (PRINT output, warnings, then the
' actual terminating error). VBA's Err.Description only ever shows ONE of them - usually
' the first - which is misleading when a proc PRINTs debug SQL before a real error further
' down (exactly what @DEBUG=1 can trigger). This walks the full Connection.Errors
' collection so the real SQL Server error text isn't hidden behind debug PRINT output.
Public Function GetFullSqlErrorText(conn As Object, ByVal fallbackDescription As String) As String
    Dim msg As String, e As Object
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.Errors.Count > 0 Then
            Dim i As Long
            For i = 0 To conn.Errors.Count - 1
                Set e = conn.Errors(i)
                msg = msg & "(" & e.NativeError & ") " & e.Description & " | "
            Next i
        End If
    End If
    On Error GoTo 0
    If Len(msg) = 0 Then
        GetFullSqlErrorText = fallbackDescription
    Else
        GetFullSqlErrorText = msg
    End If
End Function

Public Function GetConfigValue(ByVal key As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Sheets("Config")
    r = 2
    Do While ws.Cells(r, 1).Value <> ""
        If ws.Cells(r, 1).Value = key Then
            GetConfigValue = ws.Cells(r, 2).Value
            Exit Function
        End If
        r = r + 1
    Loop
    GetConfigValue = ""
End Function
