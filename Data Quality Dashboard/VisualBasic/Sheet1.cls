VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub mkdirtest()
Dim strFolderPath As String

strFolderPath = Environ("USERPROFILE") & "\Desktop\TEST\"
CheckDir (strFolderPath)

End Sub

Function CheckDir(path As String)

    If Dir(path, vbDirectory) = "" Then
        MkDir (path)
        MsgBox "Making Directory!"
    'End If
    Else
        MsgBox "Dir Exists!"
    End If

End Function


'requires reference to Microsoft Scripting Runtime
Function MkDir(strDir As String, strPath As String)

Dim fso As New FileSystemObject
Dim path As String

'examples of the input arguments
'strDir = "Folder"
'strPath = "C:\"

path = strPath & strDir

If Not fso.FolderExists(path) Then

' doesn't exist, so create the folder
          fso.CreateFolder path

End If

End Function

Function copyFile(source, destination)
Dim fso As Object

' FileSystemObject.copyFile "c:\mydocuments\letters\sample.doc", "c:\tempfolder\sample_new.doc"

Set fso = VBA.CreateObject("Scripting.FileSystemObject")
Call fso.copyFile(source, destination, True)
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

'In order to access the database you must first include the appropriate library which contains needed objects and methods.
' This is achieved by using ALT + F11 to enter the VBA Editor. Then you need to open the Tools menu, and further select the References option.
' A window will appear that contains a list of available libraries where it is necessary to check the Microsoft ActiveX Data Objects 2.0 library.

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
 

Function RunDataQualityQueries(dbName, exposureSetSID, exposureSetName)
', locCount, conCount, TIV, floorArea, Perils)
' loop thru' all stored procs and pass parameters
    Dim cnn As ADODB.Connection
    Dim sh, Output, template As Worksheet
    
    
    ActiveWorkbook.Sheets("Template").Visible = True
    Set template = ThisWorkbook.Worksheets("Template")
    
    'Delete the Sheet if exists
    For Each Sheet In ActiveWorkbook.Worksheets
            If Sheet.Name = exposureSetName Then
                Sheet.Delete
            End If
    Next Sheet

    'Copy template sheet into newsheet
    template.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    
    'Rename newly created sheet to exposureSetName
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count).Name = exposureSetName
    
    Set sh = ThisWorkbook.Sheets("StoredProcs")
    
    Set Output = ThisWorkbook.Worksheets(exposureSetName)
    'create new sheet naming ExposureSetName
    

    Dim Querycount As Long
    
    Set cnn = SQL_connection()
    
    'exposure_db_name as param 1
     Param1 = dbName
     'ExposureSetSID as Param 2
     Param2 = exposureSetSID
     'ExposureSetName as Param 3
     Param3 = exposureSetName

             
    Querycount = Worksheets("StoredProcs").UsedRange.Rows.count
    'vba_count_rows_with_data
    'Querycount = 1
    i = 2
    
    While i <= Querycount
        Application.ScreenUpdating = True
        Application.StatusBar = dbName & " - " & exposureSetName & " - " & storedProcedureName
        Application.ScreenUpdating = True
     'dbName = Sheets("StoredProcs").Cells(i, 1)
     storedProcedureName = Sheets("StoredProcs").Cells(i, 2)
     
     rowOut = Sheets("StoredProcs").Cells(i, 6)
     colOut = Sheets("StoredProcs").Cells(i, 7)
            
        
        Dim rst As New ADODB.Recordset
        Dim cmd As New ADODB.Command
        cmd.CommandTimeout = 120
        
        With cmd
            .CommandType = adCmdStoredProc
            .ActiveConnection = cnn
            'https://www.w3schools.com/asp/met_comm_createparameter.asp
            '.CommandText = "Exec " & storedprocedurename & " " & "'" & Param1 & "'," & Param2
            .CommandText = "[" & storedProcedureName & "]"
            'Set objparameter=objcommand.CreateParameter (name,type,direction,size,value)
            If Param1 <> "" Then .Parameters.Append .CreateParameter("@ExposureDBName", adVarChar, adParamInput, 255, Param1)
            If Param2 <> "" Then .Parameters.Append .CreateParameter("@ExposureSetSID", adInteger, adParamInput, 255, Param2)
            'If Param3 <> "" Then .Parameters.Append .CreateParameter("@isinsert", adInteger, adParamInput, 255, isinsert)
            'If Param4 <> "" Then .Parameters.Append .CreateParameter("@exposureMetaID", adInteger, adParamInput, 255, exposureMetaID)
            .Execute
        End With
        With rst
            .CursorType = adOpenStatic
            .CursorLocation = adOpenStatic
            .LockType = adLockOptimistic
            .Open cmd
        End With
        
    
        If Not rst.EOF Then
            intMaxCol = rst.Fields.count
            intMaxRowCount = rst.RecordCount
            
            'make sure to paste the output from recordset in the correct sections
            
            For iCols = 0 To rst.Fields.count - 1
            Output.Cells(rowOut, iCols + colOut).Value = rst.Fields(iCols).Name
            Next
            Output.range(Output.Cells(rowOut, colOut), _
            Output.Cells(rowOut, colOut + intMaxCol)).Font.Bold = True
      
    '- MoveLast/First to get an accurate RecordCount
    'Rst.MoveLast
    '
         
            Output.range(Output.Cells(rowOut + 1, colOut), _
            Output.Cells(intMaxRowCount, colOut + intMaxCol)).CopyFromRecordset rst
            
            rst.Close
            Set rst = Nothing
            Set cmd = Nothing
            
        'Else
        '    MsgBox "Error, no data!", vbCritical
        End If
       
        i = i + 1
        Wend
    cnn.Close
    
    ActiveWorkbook.Sheets("Template").Visible = False
    
End Function



Function vba_count_rows_with_data()

Dim counter As Long
Dim iRange As range

With Sheets("StoredProcs").UsedRange

    'loop through each row from the used range
    For Each iRange In .Rows

        'check if the row contains a cell with a value
        If Application.CountA(iRange) > 0 Then

            'counts the number of rows non-empty Cells
            counter = counter + 1

        End If

    Next

End With

End Function




' Function which will loop over all the Exposure Sets of DB

' Take DBName, ExposureSetName, ExposureSetSID as Param

' Call RunDataQualityQueries(DbName, ExposureSetName, ExposureSetSID)

' Close the function


Sub create_DQ()
    Dim DataSourceName, exposureSetName As String
    Dim exposureSetSID As Integer
    Dim locCount As Integer
    Dim conCount As Integer
    Dim TIV As Double
    Dim floorArea As Double
    Dim Perils As String
    Dim StartRowNumber, StartColumnNumber, MaxRowCount, currentRow As Integer
    StartRowNumber = 12
    currentRow = StartRowNumber
    StartColumnNumber = 1
    'MaxRowCount = Worksheets("Input").UsedRange.Rows.Count
    MaxRowCount = Worksheets("Input").Cells(Rows.count, 3).End(xlUp).Row
    Worksheets("Input").Activate
    Application.ScreenUpdating = False
    While currentRow <= MaxRowCount
        DataSourceName = Worksheets("Input").Cells(currentRow, StartColumnNumber + 1).Value
        exposureSetSID = Worksheets("Input").Cells(currentRow, StartColumnNumber + 2).Value
        exposureSetName = Worksheets("Input").Cells(currentRow, StartColumnNumber + 3).Value
        'locCount = Worksheets("Input").Cells(CurrentRow, StartColumnNumber + 6).Value
        'conCount = Worksheets("Input").Cells(CurrentRow, StartColumnNumber + 7).Value
        'TIV = Worksheets("Input").Cells(CurrentRow, StartColumnNumber + 8).Value
        'floorArea = Worksheets("Input").Cells(CurrentRow, StartColumnNumber + 9).Value
        'Perils = Worksheets("Input").Cells(CurrentRow, StartColumnNumber + 10).Value
        If DataSourceName <> "" Then
                        
            Call RunDataQualityQueries(DataSourceName, exposureSetSID, exposureSetName)
            ',exposureSetName, locCount, conCount, TIV, floorArea, Perils)
            
        End If
        
        currentRow = currentRow + 1
    Wend
    
    If DataSourceName <> "" Then
            
            Call callvalidationsummary(DataSourceName)
            Call callYoYData(DataSourceName)
    End If
    

    
    Application.StatusBar = "Done"
    Application.ScreenUpdating = True
End Sub


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
    
    rowOut = Worksheets("Input").Cells(Rows.count, 1).End(xlUp).Row
    
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

            rst.Close
            Set rst = Nothing
            Set cmd = Nothing
        End If

    cnn.Close

End Sub


Function callvalidationsummary(DataSourceName)

    Dim cnn As ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rowOut, colOut As Integer
    'Dim rangeString As String
    Dim Output As New Worksheet
    Dim rst As New ADODB.Recordset
    
    Set cnn = SQL_connection()
        
    Set Output = NewSheet("ValidationRuleSummary")
    
    'Clear Validation Rule Summary Tab
    Sheets("ValidationRuleSummary").UsedRange.ClearContents
    
    rowOut = 1
    colOut = 1
    cmd.CommandTimeout = 120
    
    With cmd
        .CommandType = adCmdStoredProc
        .ActiveConnection = cnn
        .CommandText = "[" & "sp_AIRDQ_call_ValidationSummary" & "]"
        If DataSourceName <> "" Then .Parameters.Append .CreateParameter("@exposureDbName", adVarChar, adParamInput, 255, DataSourceName)
        .Execute
    End With
    
    With rst
    .CursorType = adOpenStatic
    .CursorLocation = adOpenStatic
    .LockType = adLockOptimistic
    .Open cmd
    End With
    
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

            rst.Close
            Set rst = Nothing
            Set cmd = Nothing
    End If
    
    cnn.Close
    

End Function


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
        
    Set Output = NewSheet("YoY Data")
    
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
        
        Call Create_Table(intMaxCol, intMaxRowCount, rowOut, colOut)
        
        colOut = colOut + intMaxCol + 1
        
        Set rst = rst.NextRecordset
        
    Loop
    
    cnn.Close
    
    Set rst = Nothing
    Set cmd = Nothing
       

End Function

Function ColumnNumber_To_ColumnLetter(colNum As Long) As String

    ColumnNumber_To_ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)

End Function

Function Create_Table(intMaxCol, intMaxRowCount, rowOut, colOut)

Dim tbOb As ListObject
Dim TblRng As range
With Sheets("YoY Data")
Set TblRng = .range(.Cells(rowOut, colOut), .Cells(intMaxRowCount + rowOut, intMaxCol + colOut - 1))
Set tbOb = .ListObjects.Add(xlSrcRange, TblRng, , xlYes)
tbOb.Name = "DynamicTable" & colOut
tbOb.TableStyle = "TableStyleLight8"
End With
End Function


Function ExposureMetaTable(extractDQ, exposureDbName, exposureSetName, exposureSetSID, TIV, locCount, conCount, floorArea, peril, serverName, exposureVintage, version, dealYear, dealName, exposureVintageDate)

    Dim cnn As ADODB.Connection
    Set cnn = SQL_connection()
    Dim cmd As New ADODB.Command
    cmd.CommandTimeout = 120
    
    With cmd
        .CommandType = adCmdStoredProc
        .ActiveConnection = cnn
        .CommandText = "[" & "sp_AIRDQ_ExposureMetaTable" & "]"
        If extractDQ <> "" Then .Parameters.Append .CreateParameter("@extractDQ", adVarChar, adParamInput, 255, extractDQ)
        If exposureDbName <> "" Then .Parameters.Append .CreateParameter("@exposureDbName", adVarChar, adParamInput, 255, exposureDbName)
        If exposureSetName <> "" Then .Parameters.Append .CreateParameter("@exposureSetName", adVarChar, adParamInput, 255, exposureSetName)
        If exposureSetSID <> "" Then .Parameters.Append .CreateParameter("@exposureSetSID", adInteger, adParamInput, 255, exposureSetSID)
        If TIV <> "" Then .Parameters.Append .CreateParameter("@TIV", adDouble, adParamInput, 255, TIV)
        If locCount <> "" Then .Parameters.Append .CreateParameter("@locCount", adInteger, adParamInput, 255, locCount)
        If conCount <> "" Then .Parameters.Append .CreateParameter("@conCount", adInteger, adParamInput, 255, conCount)
        If floorArea <> "" Then .Parameters.Append .CreateParameter("@floorArea", adDouble, adParamInput, 255, floorArea)
        If peril <> "" Then .Parameters.Append .CreateParameter("@peril", adVarChar, adParamInput, 255, peril)
        If serverName <> "" Then .Parameters.Append .CreateParameter("@serverName", adVarChar, adParamInput, 255, serverName)
        If exposureVintage <> "" Then .Parameters.Append .CreateParameter("@exposureVintage", adVarChar, adParamInput, 255, exposureVintage)
        If exposureVintageDate <> "" Then .Parameters.Append .CreateParameter("@exposureVintageDate", adVarChar, adParamInput, 255, exposureVintageDate)
        If version <> "" Then .Parameters.Append .CreateParameter("@version", adVarChar, adParamInput, 255, version)
        If dealYear <> "" Then .Parameters.Append .CreateParameter("@dealYear", adInteger, adParamInput, 255, dealYear)
        If dealName <> "" Then .Parameters.Append .CreateParameter("@dealName", adVarChar, adParamInput, 255, dealName)
        .Execute
    End With

    Set cmd = Nothing
    
    cnn.Close

End Function


Sub updateExposureMetaTable()

    
Dim extractDQ As String
Dim dealYear As Integer
Dim dealName  As String
Dim exposureDbName  As String
Dim exposureSetName  As String
Dim exposureSetSID As Long
Dim TIV  As Double
Dim locCount As Long
Dim conCount As Long
Dim floorArea As Double
Dim peril  As String
Dim serverName  As String
Dim exposureVintage As String
Dim exposureVinatgeDate As String
Dim version  As String
Dim Id As Long
Dim StartRowNumber, StartColumnNumber, MaxRowCount, currentRow As Long

StartRowNumber = 12
currentRow = StartRowNumber
StartColumnNumber = 0
    'MaxRowCount = Worksheets("Input").UsedRange.Rows.Count
MaxRowCount = Worksheets("Input").Cells(Rows.count, 1).End(xlUp).Row
Worksheets("Input").Activate
Application.ScreenUpdating = False
While currentRow <= MaxRowCount
        extractDQ = Worksheets("Input").Cells(currentRow, StartColumnNumber + 1).Value
        exposureDbName = Worksheets("Input").Cells(currentRow, StartColumnNumber + 2).Value
        exposureSetSID = Worksheets("Input").Cells(currentRow, StartColumnNumber + 3).Value
        exposureSetName = Worksheets("Input").Cells(currentRow, StartColumnNumber + 4).Value
        locCount = Worksheets("Input").Cells(currentRow, StartColumnNumber + 5).Value
        conCount = Worksheets("Input").Cells(currentRow, StartColumnNumber + 6).Value
        TIV = Worksheets("Input").Cells(currentRow, StartColumnNumber + 7).Value
        floorArea = Worksheets("Input").Cells(currentRow, StartColumnNumber + 8).Value
        peril = Worksheets("Input").Cells(currentRow, StartColumnNumber + 9).Value
        serverName = Worksheets("Input").Cells(9, 5).Value
        exposureVintage = Worksheets("Input").Cells(7, 5).Value
        exposureVintageDate = Worksheets("Input").Cells(8, 5).Value
        version = Worksheets("Input").Cells(6, 5).Value
        dealYear = Worksheets("Input").Cells(5, 5).Value
        dealName = Worksheets("Input").Cells(4, 5).Value
        If exposureDbName <> "" Then
            
            Call ExposureMetaTable(extractDQ, exposureDbName, exposureSetName, exposureSetSID, TIV, locCount, conCount, floorArea, peril, serverName, exposureVintage, version, dealYear, dealName, exposureVintageDate)
            
        End If
        
        
        currentRow = currentRow + 1
    Wend
    Application.StatusBar = "Done"
    Application.ScreenUpdating = True
End Sub



