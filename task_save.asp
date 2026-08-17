<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/util.asp"-->
<!--#include file="includes/db.asp"-->
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
Session.CodePage = 65001

Dim taskId
Dim taskTitle
Dim taskDescription
Dim taskAssignee
Dim taskStatus
Dim priorityText
Dim taskPriority
Dim dueDateText
Dim taskDueDate
Dim dueDateValid
Dim validationErrors
Dim redirectUrl
Dim connection
Dim command
Dim affectedRows
Dim nowValue
Dim databaseError

Call RequirePost()

If Not IsValidCsrfToken(Request.Form("csrf_token")) Then
    Call FailRequest(403, "Forbidden", "Phiên làm việc hoặc CSRF token không hợp lệ. Hãy quay lại biểu mẫu và thử lại.")
End If

taskId = ParseOptionalId(FormText("id"))
If taskId < 0 Then
    Call FailRequest(400, "Bad Request", "Mã công việc không hợp lệ.")
End If

taskTitle = FormText("title")
taskDescription = FormText("description")
taskAssignee = FormText("assignee")
taskStatus = FormText("status")
priorityText = FormText("priority")
dueDateText = FormText("due_date")
validationErrors = ""

If IsBlankText(taskTitle) Then
    Call AppendValidationError(validationErrors, "Tiêu đề là bắt buộc.")
ElseIf Len(taskTitle) > 150 Then
    Call AppendValidationError(validationErrors, "Tiêu đề không được dài quá 150 ký tự.")
End If

If Len(taskDescription) > 4000 Then
    Call AppendValidationError(validationErrors, "Mô tả không được dài quá 4000 ký tự.")
End If

If Len(taskAssignee) > 100 Then
    Call AppendValidationError(validationErrors, "Người phụ trách không được dài quá 100 ký tự.")
End If

If Not IsAllowedStatus(taskStatus) Then
    Call AppendValidationError(validationErrors, "Trạng thái không hợp lệ.")
End If

taskPriority = 0
If priorityText = "1" Or priorityText = "2" Or priorityText = "3" Then
    taskPriority = CInt(priorityText)
Else
    Call AppendValidationError(validationErrors, "Mức ưu tiên không hợp lệ.")
End If

taskDueDate = ParseIsoDate(dueDateText, dueDateValid)
If Not dueDateValid Then
    Call AppendValidationError(validationErrors, "Hạn hoàn thành phải là một ngày hợp lệ theo yyyy-mm-dd.")
End If

If Len(validationErrors) > 0 Then
    Session("DraftId") = CStr(taskId)
    Session("DraftTitle") = taskTitle
    Session("DraftDescription") = taskDescription
    Session("DraftAssignee") = taskAssignee
    Session("DraftStatus") = taskStatus
    Session("DraftPriority") = taskPriority
    Session("DraftDueDate") = dueDateText
    Session("DraftErrors") = validationErrors

    redirectUrl = "task_form.asp?invalid=1"
    If taskId > 0 Then redirectUrl = redirectUrl & "&id=" & Server.URLEncode(CStr(taskId))
    Response.Redirect redirectUrl
End If

Set connection = OpenConnectionOrFail()
nowValue = Now()
affectedRows = 0
databaseError = ""

If taskId = 0 Then
    Set command = NewCommand(connection, _
        "INSERT INTO [Tasks] ([Title], [Description], [Assignee], [Status], [Priority], [DueDate], [CreatedAt], [UpdatedAt]) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)")
    Call AddTextParameter(command, taskTitle, 150)
    Call AddLongTextParameter(command, taskDescription, 4000)
    Call AddTextParameter(command, taskAssignee, 100)
    Call AddTextParameter(command, taskStatus, 20)
    Call AddIntegerParameter(command, taskPriority)
    Call AddNullableDateParameter(command, taskDueDate)
    Call AddDateParameter(command, nowValue)
    Call AddDateParameter(command, nowValue)
    On Error Resume Next
    command.Execute affectedRows, , adExecuteNoRecords
    If Err.Number <> 0 Then
        databaseError = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
Else
    Set command = NewCommand(connection, _
        "UPDATE [Tasks] SET [Title] = ?, [Description] = ?, [Assignee] = ?, [Status] = ?, " & _
        "[Priority] = ?, [DueDate] = ?, [UpdatedAt] = ? WHERE [Id] = ?")
    Call AddTextParameter(command, taskTitle, 150)
    Call AddLongTextParameter(command, taskDescription, 4000)
    Call AddTextParameter(command, taskAssignee, 100)
    Call AddTextParameter(command, taskStatus, 20)
    Call AddIntegerParameter(command, taskPriority)
    Call AddNullableDateParameter(command, taskDueDate)
    Call AddDateParameter(command, nowValue)
    Call AddIntegerParameter(command, taskId)
    On Error Resume Next
    command.Execute affectedRows, , adExecuteNoRecords
    If Err.Number <> 0 Then
        databaseError = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If Len(databaseError) = 0 And affectedRows = 0 Then
        Call CloseConnectionQuietly(connection)
        Call FailRequest(404, "Not Found", "Công việc cần sửa không còn tồn tại.")
    End If
End If

If Len(databaseError) > 0 Then
    Set command = Nothing
    Call CloseConnectionQuietly(connection)
    Call FailDatabaseRequest("DB_SAVE", databaseError)
End If

Set command = Nothing
Call CloseConnectionQuietly(connection)

Response.Redirect "default.asp?saved=1"
%>
