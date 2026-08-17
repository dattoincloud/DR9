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
Dim dataFolder
Dim databasePath
Dim buildPath
Dim providerName
Dim operationError
Dim connection

Set fileSystem = CreateObject("Scripting.FileSystemObject")
projectRoot = fileSystem.GetParentFolderName(fileSystem.GetParentFolderName(WScript.ScriptFullName))
dataFolder = fileSystem.BuildPath(projectRoot, "App_Data")
databasePath = fileSystem.BuildPath(dataFolder, "tasks.mdb")
buildPath = fileSystem.BuildPath(dataFolder, "tasks.building.mdb")
providerName = "Microsoft.Jet.OLEDB.4.0"

If Not fileSystem.FolderExists(dataFolder) Then
    fileSystem.CreateFolder dataFolder
End If

If fileSystem.FileExists(databasePath) Then
    WScript.Echo "Database already exists: " & databasePath
    WScript.Echo "Move or delete it explicitly before rebuilding the demo data."
    WScript.Quit 2
End If

If fileSystem.FileExists(buildPath) Then
    fileSystem.DeleteFile buildPath, True
End If

operationError = ""
If Not TryCreateDatabase(providerName, buildPath, operationError) Then
    WScript.Echo "ERROR: Could not create the temporary Access database."
    WScript.Echo operationError
    WScript.Echo "Use the 32-bit host: %WINDIR%\SysWOW64\cscript.exe //nologo scripts\create_database.vbs"
    WScript.Quit 1
End If

Set connection = CreateObject("ADODB.Connection")
On Error Resume Next
connection.Open "Provider=" & providerName & ";Data Source=" & buildPath & ";Persist Security Info=False;"
If Err.Number <> 0 Then
    operationError = "Open: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

If Len(operationError) > 0 Then
    Call AbortBuild(connection, buildPath, operationError)
End If

If Not ExecuteSql(connection, _
    "CREATE TABLE [Tasks] (" & _
    "[Id] COUNTER CONSTRAINT [PK_Tasks] PRIMARY KEY, " & _
    "[Title] TEXT(150) NOT NULL, " & _
    "[Description] MEMO, " & _
    "[Assignee] TEXT(100), " & _
    "[Status] TEXT(20) NOT NULL, " & _
    "[Priority] BYTE NOT NULL, " & _
    "[DueDate] DATETIME, " & _
    "[CreatedAt] DATETIME NOT NULL, " & _
    "[UpdatedAt] DATETIME NOT NULL)", operationError) Then
    Call AbortBuild(connection, buildPath, operationError)
End If

If Not ExecuteSql(connection, "CREATE INDEX [IX_Tasks_Status] ON [Tasks] ([Status])", operationError) Then
    Call AbortBuild(connection, buildPath, operationError)
End If

If Not ExecuteSql(connection, "CREATE INDEX [IX_Tasks_UpdatedAt] ON [Tasks] ([UpdatedAt])", operationError) Then
    Call AbortBuild(connection, buildPath, operationError)
End If

If Not InsertTask(connection, _
    UnicodeText(Array(&H8981, &H4EF6, &H3092, &H78BA, &H8A8D, &H3059, &H308B)), _
    "Confirm the legacy requirements and document the current behavior.", _
    "Sato", "InProgress", 3, DateAdd("d", 3, Date()), operationError) Then
    Call AbortBuild(connection, buildPath, operationError)
End If

If Not InsertTask(connection, _
    "Build Classic ASP CRUD demo", _
    "Create list, search, validation and parameterized ADO commands.", _
    "Nguyen", "Todo", 2, DateAdd("d", 7, Date()), operationError) Then
    Call AbortBuild(connection, buildPath, operationError)
End If

If Not InsertTask(connection, _
    "Verify Access backup", _
    "Keep a recoverable copy before schema or data changes.", _
    "Tanaka", "Done", 1, DateAdd("d", -1, Date()), operationError) Then
    Call AbortBuild(connection, buildPath, operationError)
End If

connection.Close
Set connection = Nothing

On Error Resume Next
fileSystem.MoveFile buildPath, databasePath
If Err.Number <> 0 Then
    operationError = "Publish database: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

If Len(operationError) > 0 Then
    Call AbortBuild(connection, buildPath, operationError)
End If

WScript.Echo "Created: " & databasePath
WScript.Echo "Provider: " & providerName
WScript.Echo "Seed rows: 3"
WScript.Quit 0

Function TryCreateDatabase(ByVal provider, ByVal targetPath, ByRef errorText)
    Dim catalog
    Dim connectionString

    TryCreateDatabase = False
    connectionString = "Provider=" & provider & ";Data Source=" & targetPath & ";Jet OLEDB:Engine Type=5;"

    On Error Resume Next
    Set catalog = CreateObject("ADOX.Catalog")
    catalog.Create connectionString

    If Err.Number = 0 Then
        TryCreateDatabase = True
    Else
        errorText = provider & ": " & Err.Description
        Err.Clear
    End If

    Set catalog = Nothing
    On Error GoTo 0

    If Not TryCreateDatabase And fileSystem.FileExists(targetPath) Then
        fileSystem.DeleteFile targetPath, True
    End If
End Function

Function ExecuteSql(ByRef activeConnection, ByVal sqlText, ByRef errorText)
    ExecuteSql = False

    On Error Resume Next
    activeConnection.Execute sqlText
    If Err.Number = 0 Then
        ExecuteSql = True
    Else
        errorText = "Schema: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function InsertTask(ByRef activeConnection, ByVal title, ByVal description, ByVal assignee, _
    ByVal statusCode, ByVal priorityValue, ByVal dueDateValue, ByRef errorText)

    Dim command
    Dim timestampValue

    InsertTask = False
    timestampValue = Now()

    On Error Resume Next
    Set command = CreateObject("ADODB.Command")
    Set command.ActiveConnection = activeConnection
    command.CommandType = adCmdText
    command.CommandText = _
        "INSERT INTO [Tasks] ([Title], [Description], [Assignee], [Status], [Priority], [DueDate], [CreatedAt], [UpdatedAt]) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"

    command.Parameters.Append command.CreateParameter("p1", adVarWChar, adParamInput, 150, title)
    command.Parameters.Append command.CreateParameter("p2", adLongVarWChar, adParamInput, 4000, description)
    command.Parameters.Append command.CreateParameter("p3", adVarWChar, adParamInput, 100, assignee)
    command.Parameters.Append command.CreateParameter("p4", adVarWChar, adParamInput, 20, statusCode)
    command.Parameters.Append command.CreateParameter("p5", adInteger, adParamInput, , priorityValue)
    command.Parameters.Append command.CreateParameter("p6", adDate, adParamInput, , dueDateValue)
    command.Parameters.Append command.CreateParameter("p7", adDate, adParamInput, , timestampValue)
    command.Parameters.Append command.CreateParameter("p8", adDate, adParamInput, , timestampValue)
    command.Execute , , adExecuteNoRecords

    If Err.Number = 0 Then
        InsertTask = True
    Else
        errorText = "Seed: " & Err.Description
        Err.Clear
    End If

    Set command = Nothing
    On Error GoTo 0
End Function

Sub AbortBuild(ByRef activeConnection, ByVal temporaryPath, ByVal message)
    On Error Resume Next
    activeConnection.Close
    Set activeConnection = Nothing
    If fileSystem.GetFileName(temporaryPath) = "tasks.building.mdb" And fileSystem.FileExists(temporaryPath) Then
        fileSystem.DeleteFile temporaryPath, True
    End If
    Err.Clear
    On Error GoTo 0

    WScript.Echo "ERROR: Database build failed."
    WScript.Echo message
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
