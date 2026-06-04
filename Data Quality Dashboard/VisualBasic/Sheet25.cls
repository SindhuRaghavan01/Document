VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Function SQL_connection()

    Dim connectionstring As String
    Dim StrSproc As String
    
    connectionstring = Sheets("SQL Connection").range("B7").Value
    
    Set SQL_connection = New ADODB.Connection
    'cnn.Open connectionstring

With SQL_connection
        .CommandTimeout = 900
        .connectionstring = connectionstring
        .Open
    End With
    
    
    'If CBool(SQL_Connection.State And adStateOpen) Then SQL_Connection.Close
    'Set SQL_Connection = Nothing
    

End Function


Sub GetExposureSets()
    Dim exposureDbName As String
    exposureDbName = Application.InputBox("Enter Exposure Database Name (Without brackets) :", "Title", "") ', "", "Left", "Top", "HelpFile", "HelpContextID", "Type")
    'Debug.Print ExposureDBName
    Dim Output As Worksheet
    Dim cnn As ADODB.Connection
    Set cnn = SQL_connection()
    Dim rowOut, colOut As Integer
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set Output = Worksheets("Input")
    cmd.CommandTimeout = 120
    
    sheetName = "Input"
    rowOut = 11 'Worksheets("Input").Cells(Rows.count, 1).End(xlUp).Row
    
    serverName = Worksheets("Input").Cells(9, 5).Value
    colOut = 1
    With cmd
        .CommandType = adCmdStoredProc
        .ActiveConnection = cnn
        'https://www.w3schools.com/asp/met_comm_createparameter.asp
        '.CommandText = "Exec " & storedprocedurename & " " & "'" & Param1 & "'," & Param2
        .CommandText = "[" & "sp_AIRDQ_ExposureSetsFromDB" & "]"
        'Set objparameter=objcommand.CreateParameter (name,type,direction,size,value)
        If exposureDbName <> "" Then .Parameters.Append .CreateParameter("@ExposureDbName", adVarChar, adParamInput, 2048, exposureDbName)
        If serverName <> "" Then .Parameters.Append .CreateParameter("@sqlServer", adVarChar, adParamInput, 255, serverName)
        'If Param2 <> "" Then .Parameters.Append .CreateParameter("@ExposureSetSID", adInteger, adParamInput, 255, Param2)
        'If Param3 <> "" Then .Parameters.Append .CreateParameter("@parameter3", adInteger, adParamInput, , Param3)
        .Execute
    End With
    With rst
        .CursorType = adOpenStatic
        .CursorLocation = adOpenStatic
        .LockType = adLockOptimistic
        .Open cmd
    End With
    
    Call ClearAutoRange(sheetName, rowOut + 1, colOut)

        If Not rst.EOF Then
            intMaxCol = rst.Fields.count
            intMaxRowCount = rst.RecordCount

            'make sure to paste the output from recordset in the correct sections
            'Header Line
            For iCols = 0 To rst.Fields.count - 1
            Output.Cells(rowOut, iCols + colOut).Value = rst.Fields(iCols).Name
            Next
            Output.range(Output.Cells(rowOut, colOut), _
            Output.Cells(rowOut, colOut + intMaxCol)).Font.Bold = True

            Output.range(Output.Cells(rowOut + 1, colOut), _
            Output.Cells(intMaxRowCount + rowOut, colOut + intMaxCol)).CopyFromRecordset rst
            
            Call Create_Table(intMaxCol, intMaxRowCount, rowOut, colOut, sheetName)
            
            rst.Close
            Set rst = Nothing
            Set cmd = Nothing
        End If

    cnn.Close

End Sub


Function CallSP(storedProcedureName, inputVariableList, outputVariableList, outputSheetName, rowOut, colOut, printHeader)

    Dim InputvariableValue As Variant
'    Dim InputvariableName As Variant
    Dim tableSheet As Worksheet
    Dim LookupRange As range
    Dim OutputvariableValue As Variant
    Dim Output, template As Worksheet
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
'    Set the table sheet and range containing variable values
'    Set tableSheet = ThisWorkbook.Worksheets("Inputs")
'    Set tableRange = tableSheet.Range("A1:B10") ' Update range as per your requirement

    If outputSheetName <> "" Then
        Set Output = ThisWorkbook.Worksheets(outputSheetName)
    End If
       
    ' Set up the connection and command objects for the stored procedure call
    Set cnn = SQL_connection()
    
    'Set command type and stored procedure name
    
    cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = storedProcedureName
    cmd.CommandTimeout = 0
    
    inputVariableNames = Split(inputVariableList, ",") ' Replace with your comma-separated string
    outputVariableNames = Split(outputVariableList, ",")
    
    Set tableSheet = ThisWorkbook.Worksheets("Inputs")
'    Set LookupRange = tableRange
    Set LookupRange = tableSheet.range("A:A")
    
     ' Loop through each variable name, extract its value from the table, and build the SQL string
    For Each InputvariableName In inputVariableNames
'        InputvariableValue = Application.VLookup(InputvariableName, tableRange, 2, False)
        
        'Changing the date type value to yyyy-mm-dd format
        a = Application.Match(InputvariableName, LookupRange, 0)
        InputvariableValue = LookupRange.Cells(a, 2)
                
        If IsDate(InputvariableValue) Then
            InputvariableValue = WorksheetFunction.Text(InputvariableValue, "yyyy-mm-dd")
        End If
        
        cmd.Parameters.Append cmd.CreateParameter("@" & InputvariableName, adVarChar, adParamInput, Len(InputvariableValue), InputvariableValue)
        Debug.Print InputvariableValue
        
    Next InputvariableName
    
    If outputVariableList <> "" Then
        ' Loop through each variable name, extract its value from the table, and build the SQL string
        For Each OutputvariableName In outputVariableNames
            'variableValue = Application.VLookup(OutputvariableName, tableRange, 2, False)
        '        sql = sql & " @" & OutputvariableName & " = " & OutputvariableName & " Output ," ' Update SQL string with variable value
        
            'Set output variable
            Set exportID = cmd.CreateParameter("@" & OutputvariableName, adInteger, adParamOutput) ', Len(OutputvariableName), "@" & OutputvariableName)
            cmd.Parameters.Append exportID
        '        Debug.Print sql
        Next OutputvariableName
    End If
   
   ' Print execution command
    sql = cmd.CommandText
    For i = 0 To cmd.Parameters.count - 1
        sql = sql & " " & cmd.Parameters(i).Name & "=" & cmd.Parameters(i).Value
    Next i
    
    Debug.Print sql
       
    'Execute command
    Set rst = cmd.Execute
    
    ' Save the Output from stored procedure
        Do While rst.State = 1 ' If Not rst.EOF Then
        
            intMaxCol = rst.Fields.count
            intMaxRowCount = rst.RecordCount

            'make sure to paste the output from recordset in the correct sections
            
            If printHeader = "True" Then
            'Header Line
                For iCols = 0 To rst.Fields.count - 1
                Output.Cells(rowOut, iCols + colOut).Value = rst.Fields(iCols).Name
                Next
                
                Output.range(Output.Cells(rowOut, colOut), _
                Output.Cells(rowOut, colOut + intMaxCol)).Font.Bold = True
            End If
            
            Output.range(Output.Cells(rowOut + 1, colOut), _
            Output.Cells(intMaxRowCount + rowOut + 2, colOut + intMaxCol)).CopyFromRecordset rst

            rst.Close
            Set rst = Nothing
        Loop
    
    ' Set the Output variable values to Input Tables
    If outputVariableList <> "" Then
        ' set the output variable value
'        tableSheet.Cells(10, 2).Value = exportID.Value
        
        For Each OutputvariableName In outputVariableNames
            OutputvariableValue = cmd.Parameters("@" & OutputvariableName).Value
            
            output_a = Application.Match(OutputvariableName, LookupRange, 0)
            LookupRange.Cells(output_a, 2).Value = OutputvariableValue
                
        Next OutputvariableName
    End If
    
     cnn.Close
    ' Clean up
    Set cmd = Nothing
    
End Function


Function extractDataandCallSP()
    Dim listSheet As Worksheet
    Dim tableSheet As Worksheet
    Dim tableRange As range
    Dim LookupRange As range
    Dim variableSheet As Variant
    Dim variableName As Variant
    Dim sql, SourceString As String
    Dim Database As String
    Dim storedProcedure As String
    Dim Output As Worksheet
    Dim rowOut, colOut As Integer
    Dim printHeader As String
    Dim StartRowNumber, StartColumnNumber, MaxRowCount, currentRow As Integer

    ' Set the table sheet and range containing variable values
    Set tableSheet = ThisWorkbook.Worksheets("Inputs")
    Set tableRange = tableSheet.range("A:B") ' Update range as per your requirement
    
    StartRowNumber = 2
    currentRow = StartRowNumber
    StartColumnNumber = 1
    MaxRowCount = Worksheets("References").Cells(Rows.count, 1).End(xlUp).Row
    
    Worksheets("References").Activate
    Application.ScreenUpdating = False
    
    While currentRow <= MaxRowCount
    
        ' Locate Column Position of each columns in reference table
        Worksheets("References").Activate
        Set LookupRange = ActiveSheet.range("A1:Z1")
        
        InputPos = Application.Match("QueryInputParams", LookupRange, 0)
        OutputPos = Application.Match("QueryOutputParams", LookupRange, 0)
        OutputsheetPos = Application.Match("SheetName", LookupRange, 0)
        RowIdxPos = Application.Match("RowIndex", LookupRange, 0)
        ColumnIdxPos = Application.Match("ColIndex", LookupRange, 0)
        spPos = Application.Match("QueryName", LookupRange, 0)
        PrintHeadersPos = Application.Match("PrintHeaders", LookupRange, 0)
        
        ' Set the Variable sheet and Get the variable names from the input string
        inputVariableList = Worksheets("References").Cells(currentRow, InputPos).Value
        outputVariableList = Worksheets("References").Cells(currentRow, OutputPos).Value
        
        ' Set the Output sheet name and save the result from stroed procedures output
        outputSheetName = Worksheets("References").Cells(currentRow, OutputsheetPos).Value
        
        ' Set the printHeaders
        printHeader = Worksheets("References").Cells(currentRow, PrintHeadersPos).Value
        
        ' Set the Row Output to paste data on SheetName
        rowOut = Worksheets("References").Cells(currentRow, RowIdxPos).Value
        
        ' Set the Column Output to paste data on SheetName
        colOut = Worksheets("References").Cells(currentRow, ColumnIdxPos).Value
        
        If SheetExists(outputSheetName) Then
            Call ClearAutoRange(outputSheetName, rowOut, colOut)
        Else
            ActiveWorkbook.Sheets("Template").Visible = True
            Set template = ThisWorkbook.Worksheets("Template")
            'Copy template sheet into newsheet
            template.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
            'Rename newly created sheet to exposureSetName
            ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count).Name = outputSheetName
        End If
        
        ' Set Database Name
        Database = Worksheets("SQL Connection").Cells(2, 2).Value
        ' Set Stored Procedure Name
        storedProcedure = Worksheets("References").Cells(currentRow, spPos).Value
        storedProcedureName = Database & ".dbo." & storedProcedure
        
        If storedProcedureName <> "" Then
            Debug.Print storedProcedureName
            Call CallSP(storedProcedureName, inputVariableList, outputVariableList, outputSheetName, rowOut, colOut, printHeader)
            
        End If
        currentRow = currentRow + 1
    Wend
    
End Function

Sub FillUpInputTable()

    Dim WS As Worksheet
    Dim SheetNameRange As range
    
    Set WS = ThisWorkbook.Worksheets("Input")
    
    WS.Activate
    
    StartRowNumber = 12
    currentRow = StartRowNumber
    StartColumnNumber = 1
    MaxRowCount = WS.Cells(Rows.count, 2).End(xlUp).Row
    MaxRowRefer = Worksheets("References").Cells(Rows.count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    While currentRow <= MaxRowCount
        
        ' Set the éxposureDBName and exposureSetSID from the input sheet
        exposureDbName = WS.Cells(currentRow, 2).Value
        exposureSetSID = WS.Cells(currentRow, 3).Value
        exposureSetName = WS.Cells(currentRow, 4).Value
        
        Worksheets("Inputs").Cells(2, 2).Value = exposureDbName
        Worksheets("Inputs").Cells(3, 2).Value = exposureSetSID
        
        Set SheetNameRange = Worksheets("References").range("D2:D" & MaxRowRefer)
        SheetNameRange.Value = exposureSetName
        
        If exposureDbName <> "" Then
            Call extractDataandCallSP
        End If
               
        currentRow = currentRow + 1
    Wend
    
    
    ' Call YoY Data from DB
    Call callYoYData(exposureDbName)
    
    Application.StatusBar = "Done"
    Application.ScreenUpdating = True
    
    MsgBox "Data imported Successfully."

End Sub

Function callYoYData(DataSourceName)

    Dim cnn As ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim currentRow, currentColumn As Long
    Dim range As range
    Dim rowOut, colOut As Integer
    Dim rangeString As String
    Dim Output As New Worksheet
    Dim rst As New ADODB.Recordset
    Dim intMaxCol, intMaxRowCount, iCols As Integer
    Dim rngData As range
    
    Set cnn = SQL_connection()
    
    sheetName = "YoY Data"
        
    Set Output = NewSheet(sheetName)
    
    'Clear Validation Rule Summary Tab+
    Sheets("YoY Data").UsedRange.ClearContents
    
    currentRow = 1
    currentColumn = 1
    rowOut = 1
    colOut = 1
    cmd.CommandTimeout = 300
    
    With cmd
        .CommandType = adCmdStoredProc
        .ActiveConnection = cnn
        .CommandText = "[" & "sp_AIRDQ_call_YoY_Data" & "]"
        If DataSourceName <> "" Then .Parameters.Append .CreateParameter("@exposureDbName", adVarChar, adParamInput, 255, DataSourceName)
        .Execute
    End With
       
    With rst
    .CursorType = adOpenStatic
    .CursorLocation = adOpenStatic
    .LockType = adLockOptimistic
    .Open cmd
    End With
    
    Do While rst.State = 1 'Not rst.EOF
        
        intMaxCol = rst.Fields.count
        intMaxRowCount = rst.RecordCount
        
        'make sure to paste the output from recordset in the correct sections
        'Header Line
        For iCols = 0 To rst.Fields.count - 1
            Output.Cells(rowOut, iCols + colOut).Value = rst.Fields(iCols).Name
        Next
        Output.range(Output.Cells(rowOut, colOut + 1), _
        Output.Cells(rowOut, colOut + intMaxCol + 1)).Font.Bold = True
        
        Output.range(Output.Cells(rowOut + 1, colOut), _
        Output.Cells(intMaxRowCount + rowOut, colOut + intMaxCol + 1)).CopyFromRecordset rst
                   
        'range.CopyFromRecordset rst
        
        Call Create_Table(intMaxCol, intMaxRowCount, rowOut, colOut, sheetName)
        
        colOut = colOut + intMaxCol + 1
        
        Set rst = rst.NextRecordset
        
    Loop
    cnn.Close
    Set rst = Nothing
    Set cmd = Nothing
       
End Function


Function ClearAutoRange(sheetName, rowOut, colOut)
    Dim ws1 As Worksheet
    Dim startCell As range
    Dim endCell As range
    
    Set ws1 = ThisWorkbook.Worksheets(sheetName)
    Set startCell = ws1.range(ws1.Cells(rowOut, colOut).Address(0, 0)) ' Change this to the starting cell of your range
    Set endCell = startCell.End(xlDown).End(xlToRight)
    
    ws1.range(startCell, endCell).ClearContents
End Function

Function SheetExists(sheetName) As Boolean
    Dim WS As Worksheet
    SheetExists = False
    For Each WS In ThisWorkbook.Worksheets
        If WS.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next WS
End Function

Function Create_Table(intMaxCol, intMaxRowCount, rowOut, colOut, sheetName)
    Dim tbOb As ListObject
    Dim TblRng As range
    With Sheets(sheetName)
        Set TblRng = .range(.Cells(rowOut, colOut), .Cells(intMaxRowCount + rowOut, intMaxCol + colOut - 1))
        Set tbOb = .ListObjects.Add(xlSrcRange, TblRng, , xlYes)
        tbOb.Name = sheetName & colOut
        tbOb.TableStyle = "TableStyleLight8"
    End With
    Sheets(sheetName).Activate
    ActiveSheet.Cells(rowOut, colOut).CurrentRegion
End Function

Function NewSheet(sheetName) As Worksheet
    Dim Sheet As Object
    For Each Sheet In Sheets:
        If sheetName = Sheet.Name Then
            Set NewSheet = Sheet
            Exit Function
        End If
    Next Sheet
    With ThisWorkbook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.count))
        NewSheet.Name = Right(sheetName, 31)
    End With

End Function

'Sub TestSub()
'
'Dim WS As Worksheet
'Dim rowOut, colOut As Integer
'
'
'sheetName = "Test"
''Set ws = ThisWorkbook.Worksheets(SheetName)
'
'rowOut = 1
'colOut = 1
'
'Call ClearAutoRange(sheetName, rowOut, colOut)
'
'End Sub

'Sub TestSub()
'
'    Dim tableSheet As Worksheet
'    Dim tableRange As Range
'
'     Set the table sheet and range containing variable values
'    Set tableSheet = ThisWorkbook.Worksheets("Inputs")
'    Set tableRange = tableSheet.Range("A:A")
'
'    a = Application.Match("ReportDate", tableRange, 0)
'
'    b = tableRange.Cells(a, 2)
'
'    tableRange.Cells(a, 2).Value = 786
'
'    Debug.Print b
'    b = "2/29/2022"
'    If Not IsDate(b) Then
'        Debug.Print "Not a Date Format"
'    Else
'        Debug.Print "Date"
'    End If
'
'    For Each OutputvariableName In outputVariableNames
'
'
'            OutputvariableValue = 808561 '= cmd.Parameters("@" & OutputvariableName).Value
'
'            output_a = Application.Match("ExportID", tableRange, 0)
'            tableRange.Cells(output_a, 2).Value = OutputvariableValue
'
'        Next OutputvariableName
'
'
'End Sub

