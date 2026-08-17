Option Explicit

Const adCmdText = 1
Const adParamInput = 1
Const adInteger = 3
Const adDate = 7
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adExecuteNoRecords = 128

Dim fileSystem
Dim projectRoot
Dim databasePath
Dim providerName
Dim connection
Dim providerError
Dim testTitle
Dim command
Dim rows
Dim matchCount
Dim transactionStarted
Dim affectedRows

Set fileSystem = CreateObject("Scripting.FileSystemObject")
projectRoot = fileSystem.GetParentFolderName(fileSystem.GetParentFolderName(WScript.ScriptFullName))
databasePath = fileSystem.BuildPath(fileSystem.BuildPath(projectRoot, "App_Data"), "tasks.mdb")

If Not fileSystem.FileExists(databasePath) Then
    WScript.Echo "FAIL: Database does not exist: " & databasePath
    WScript.Quit 1
End If

providerName = "Microsoft.Jet.OLEDB.4.0"
providerError = ""
Set connection = CreateObject("ADODB.Connection")

On Error Resume Next
connection.Open "Provider=" & providerName & ";Data Source=" & databasePath & ";Persist Security Info=False;"
If Err.Number <> 0 Then
    providerError = Err.Description
    Err.Clear
End If
On Error GoTo 0

If Len(providerError) > 0 Then
    WScript.Echo "FAIL: Jet 4.0 could not open tasks.mdb: " & providerError
    WScript.Echo "Run this script with %WINDIR%\SysWOW64\cscript.exe."
    WScript.Quit 1
End If

testTitle = "Smoke " & UnicodeText(Array(&H30C6, &H30B9, &H30C8)) & " - " & _
    UnicodeText(Array(&H1EA1, &H1EEF)) & " - " & CStr(Timer)
transactionStarted = False

On Error Resume Next
connection.BeginTrans
If Err.Number <> 0 Then
    Call AbortTest("Could not begin transaction: " & Err.Description, connection, False)
End If
transactionStarted = True

Set command = CreateObject("ADODB.Command")
Set command.ActiveConnection = connection
command.CommandType = adCmdText
command.CommandText = _
    "INSERT INTO [Tasks] ([Title], [Description], [Assignee], [Status], [Priority], [DueDate], [CreatedAt], [UpdatedAt]) " & _
    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
command.Parameters.Append command.CreateParameter("p1", adVarWChar, adParamInput, 150, testTitle)
command.Parameters.Append command.CreateParameter("p2", adLongVarWChar, adParamInput, 4000, "Temporary transaction row")
command.Parameters.Append command.CreateParameter("p3", adVarWChar, adParamInput, 100, "SmokeTest")
command.Parameters.Append command.CreateParameter("p4", adVarWChar, adParamInput, 20, "Todo")
command.Parameters.Append command.CreateParameter("p5", adInteger, adParamInput, , 2)
command.Parameters.Append command.CreateParameter("p6", adDate, adParamInput, , DateAdd("d", 1, Date()))
command.Parameters.Append command.CreateParameter("p7", adDate, adParamInput, , Now())
command.Parameters.Append command.CreateParameter("p8", adDate, adParamInput, , Now())
command.Execute , , adExecuteNoRecords
If Err.Number <> 0 Then
    Call AbortTest("Parameterized insert failed: " & Err.Description, connection, transactionStarted)
End If

Set command = CreateObject("ADODB.Command")
Set command.ActiveConnection = connection
command.CommandType = adCmdText
command.CommandText = "SELECT Count(*) AS MatchCount FROM [Tasks] WHERE [Title] = ?"
command.Parameters.Append command.CreateParameter("p1", adVarWChar, adParamInput, 150, testTitle)
Set rows = command.Execute()
If Err.Number <> 0 Then
    Call AbortTest("Parameterized select failed: " & Err.Description, connection, transactionStarted)
End If

matchCount = CLng(rows("MatchCount").Value)
rows.Close
Set rows = Nothing

If matchCount <> 1 Then
    Call AbortTest("Expected one Unicode test row, found " & CStr(matchCount), connection, transactionStarted)
End If

Set command = CreateObject("ADODB.Command")
Set command.ActiveConnection = connection
command.CommandType = adCmdText
command.CommandText = "UPDATE [Tasks] SET [Status] = ?, [UpdatedAt] = ? WHERE [Title] = ?"
command.Parameters.Append command.CreateParameter("p1", adVarWChar, adParamInput, 20, "Done")
command.Parameters.Append command.CreateParameter("p2", adDate, adParamInput, , Now())
command.Parameters.Append command.CreateParameter("p3", adVarWChar, adParamInput, 150, testTitle)
affectedRows = 0
command.Execute affectedRows, , adExecuteNoRecords
If Err.Number <> 0 Then
    Call AbortTest("Parameterized update failed: " & Err.Description, connection, transactionStarted)
End If
If affectedRows <> 1 Then
    Call AbortTest("Expected one updated row, found " & CStr(affectedRows), connection, transactionStarted)
End If

Set command = CreateObject("ADODB.Command")
Set command.ActiveConnection = connection
command.CommandType = adCmdText
command.CommandText = "DELETE FROM [Tasks] WHERE [Title] = ?"
command.Parameters.Append command.CreateParameter("p1", adVarWChar, adParamInput, 150, testTitle)
affectedRows = 0
command.Execute affectedRows, , adExecuteNoRecords
If Err.Number <> 0 Then
    Call AbortTest("Parameterized delete failed: " & Err.Description, connection, transactionStarted)
End If
If affectedRows <> 1 Then
    Call AbortTest("Expected one deleted row, found " & CStr(affectedRows), connection, transactionStarted)
End If

connection.RollbackTrans
transactionStarted = False
If Err.Number <> 0 Then
    Call AbortTest("Rollback failed: " & Err.Description, connection, False)
End If
On Error GoTo 0

connection.Close
Set connection = Nothing

WScript.Echo "PASS: schema, CRUD transaction, parameterized queries and Unicode round-trip"
WScript.Echo "Provider: " & providerName
WScript.Quit 0

Sub AbortTest(ByVal message, ByRef activeConnection, ByVal shouldRollback)
    Dim originalMessage

    originalMessage = message
    Err.Clear
    If shouldRollback Then activeConnection.RollbackTrans
    If Not activeConnection Is Nothing Then activeConnection.Close
    On Error GoTo 0

    WScript.Echo "FAIL: " & originalMessage
    WScript.Quit 1
End Sub

Function UnicodeText(ByVal codePoints)
    Dim codePoint
    Dim result

    result = ""
    For Each codePoint In codePoints
        If codePoint > &H7FFF Then
            result = result & ChrW(codePoint - &H10000)
        Else
            result = result & ChrW(codePoint)
        End If
    Next

    UnicodeText = result
End Function
