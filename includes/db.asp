<%
Const adCmdText = 1
Const adParamInput = 1
Const adInteger = 3
Const adDate = 7
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adExecuteNoRecords = 128

Function OpenConnection()
    Dim connection
    Dim connectionString
    Dim databasePath
    Dim providerName
    Dim providerError

    databasePath = Server.MapPath("App_Data/tasks.mdb")
    providerName = "Microsoft.Jet.OLEDB.4.0"
    connectionString = "Provider=" & providerName & ";Data Source=" & databasePath & ";Persist Security Info=False;"
    Set connection = Server.CreateObject("ADODB.Connection")

    On Error Resume Next
    connection.Open connectionString
    If Err.Number <> 0 Then
        providerError = Err.Description
        Err.Clear
        On Error GoTo 0
        Set connection = Nothing
        Err.Raise vbObjectError + 2100, "OpenConnection", providerName & ": " & providerError
    End If
    On Error GoTo 0

    Set OpenConnection = connection
End Function

Function OpenConnectionOrFail()
    Dim connection
    Dim errorMessage

    Set connection = Nothing
    errorMessage = ""

    On Error Resume Next
    Set connection = OpenConnection()
    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If connection Is Nothing Then
        Call AppendServerDiagnostic("DB_OPEN", errorMessage)
        Call FailRequest(500, "Database unavailable", _
            "Không thể mở database. Kiểm tra Jet 4.0 32-bit, cấu hình Application Pool và quyền App_Data.")
    End If

    Set OpenConnectionOrFail = connection
End Function

Function NewCommand(ByRef connection, ByVal sqlText)
    Dim command

    Set command = Server.CreateObject("ADODB.Command")
    Set command.ActiveConnection = connection
    command.CommandType = adCmdText
    command.CommandText = sqlText

    Set NewCommand = command
End Function

Sub AddTextParameter(ByRef command, ByVal value, ByVal maxLength)
    Dim parameter

    Set parameter = command.CreateParameter("@p" & CStr(command.Parameters.Count), _
        adVarWChar, adParamInput, maxLength, TextValue(value))
    command.Parameters.Append parameter
End Sub

Sub AddLongTextParameter(ByRef command, ByVal value, ByVal maxLength)
    Dim parameter

    Set parameter = command.CreateParameter("@p" & CStr(command.Parameters.Count), _
        adLongVarWChar, adParamInput, maxLength, TextValue(value))
    command.Parameters.Append parameter
End Sub

Sub AddIntegerParameter(ByRef command, ByVal value)
    Dim parameter

    Set parameter = command.CreateParameter("@p" & CStr(command.Parameters.Count), _
        adInteger, adParamInput, , CLng(value))
    command.Parameters.Append parameter
End Sub

Sub AddDateParameter(ByRef command, ByVal value)
    Dim parameter

    Set parameter = command.CreateParameter("@p" & CStr(command.Parameters.Count), _
        adDate, adParamInput, , CDate(value))
    command.Parameters.Append parameter
End Sub

Sub AddNullableDateParameter(ByRef command, ByVal value)
    Dim parameter

    Set parameter = command.CreateParameter("@p" & CStr(command.Parameters.Count), _
        adDate, adParamInput)

    If IsNull(value) Or IsEmpty(value) Then
        parameter.Value = Null
    Else
        parameter.Value = CDate(value)
    End If

    command.Parameters.Append parameter
End Sub

Sub CloseRecordsetQuietly(ByRef recordset)
    On Error Resume Next
    recordset.Close
    Set recordset = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Sub CloseConnectionQuietly(ByRef connection)
    On Error Resume Next
    connection.Close
    Set connection = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Sub FailDatabaseRequest(ByVal contextName, ByVal detail)
    Call AppendServerDiagnostic(contextName, detail)
    Call FailRequest(500, "Database error", _
        "Không thể hoàn tất thao tác database. Xem IIS log và kiểm tra file Access/ACL.")
End Sub
%>
