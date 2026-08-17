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
Dim taskPriority
Dim taskDueDate
Dim validationErrors
Dim isDraft
Dim isEdit
Dim connection
Dim command
Dim rows
Dim csrfToken
Dim pageTitle
Dim errorItems
Dim errorItem
Dim databaseError
Dim taskFound

taskId = ParseOptionalId(QueryText("id"))
If taskId < 0 Then
    Call FailRequest(400, "Bad Request", "Mã công việc không hợp lệ.")
End If

taskTitle = ""
taskDescription = ""
taskAssignee = ""
taskStatus = STATUS_TODO
taskPriority = PRIORITY_NORMAL
taskDueDate = ""
validationErrors = ""
isDraft = (QueryText("invalid") = "1" And Len(TextValue(Session("DraftErrors"))) > 0)

If isDraft Then
    taskId = ParseOptionalId(Session("DraftId"))
    If taskId < 0 Then taskId = 0

    taskTitle = TextValue(Session("DraftTitle"))
    taskDescription = TextValue(Session("DraftDescription"))
    taskAssignee = TextValue(Session("DraftAssignee"))
    taskStatus = TextValue(Session("DraftStatus"))
    taskPriority = Session("DraftPriority")
    taskDueDate = TextValue(Session("DraftDueDate"))
    validationErrors = TextValue(Session("DraftErrors"))

    Session.Contents.Remove "DraftId"
    Session.Contents.Remove "DraftTitle"
    Session.Contents.Remove "DraftDescription"
    Session.Contents.Remove "DraftAssignee"
    Session.Contents.Remove "DraftStatus"
    Session.Contents.Remove "DraftPriority"
    Session.Contents.Remove "DraftDueDate"
    Session.Contents.Remove "DraftErrors"
ElseIf taskId > 0 Then
    Set connection = OpenConnectionOrFail()
    Set command = NewCommand(connection, _
        "SELECT [Title], [Description], [Assignee], [Status], [Priority], [DueDate] FROM [Tasks] WHERE [Id] = ?")
    Call AddIntegerParameter(command, taskId)
    Set rows = Nothing
    databaseError = ""
    taskFound = False
    On Error Resume Next
    Set rows = command.Execute()
    If Err.Number = 0 Then
        If Not rows.EOF Then
            taskFound = True
            taskTitle = TextValue(rows("Title").Value)
            taskDescription = TextValue(rows("Description").Value)
            taskAssignee = TextValue(rows("Assignee").Value)
            taskStatus = TextValue(rows("Status").Value)
            taskPriority = CLng(rows("Priority").Value)
            taskDueDate = FormatIsoDate(rows("DueDate").Value)
        End If
    End If
    If Err.Number <> 0 Then
        databaseError = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If Len(databaseError) > 0 Then
        Call CloseRecordsetQuietly(rows)
        Call CloseConnectionQuietly(connection)
        Call FailDatabaseRequest("DB_EDIT", databaseError)
    End If

    If Not taskFound Then
        Call CloseRecordsetQuietly(rows)
        Call CloseConnectionQuietly(connection)
        Call FailRequest(404, "Not Found", "Không tìm thấy công việc cần sửa.")
    End If

    Call CloseRecordsetQuietly(rows)
    Set command = Nothing
    Call CloseConnectionQuietly(connection)
End If

If Not IsAllowedStatus(taskStatus) Then taskStatus = STATUS_TODO
If Not IsNumeric(taskPriority) Then taskPriority = PRIORITY_NORMAL
taskPriority = CLng(taskPriority)
If Not IsAllowedPriority(taskPriority) Then taskPriority = PRIORITY_NORMAL

isEdit = (taskId > 0)
csrfToken = GetCsrfToken()
If isEdit Then
    pageTitle = "Sửa công việc"
Else
    pageTitle = "Tạo công việc"
End If
%>
<!doctype html>
<html lang="vi">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><%= H(pageTitle) %> · Legacy Task Board</title>
    <link rel="stylesheet" href="assets/styles.css">
</head>
<body>
    <header class="site-header compact-header">
        <div class="shell header-inner">
            <div>
                <a class="back-link" href="default.asp">← Danh sách công việc</a>
                <p class="eyebrow">Classic ASP form</p>
                <h1><%= H(pageTitle) %></h1>
            </div>
        </div>
    </header>

    <main class="shell form-shell">
        <% If Len(validationErrors) > 0 Then
            errorItems = Split(validationErrors, vbLf)
        %>
            <div class="alert alert-error" role="alert">
                <strong>Vui lòng kiểm tra lại:</strong>
                <ul>
                <% For Each errorItem In errorItems %>
                    <li><%= H(errorItem) %></li>
                <% Next %>
                </ul>
            </div>
        <% End If %>

        <form method="post" action="task_save.asp" accept-charset="UTF-8" class="card task-form" data-task-form>
            <input type="hidden" name="id" value="<%= H(taskId) %>">
            <input type="hidden" name="csrf_token" value="<%= H(csrfToken) %>">

            <div class="form-heading">
                <div>
                    <p class="eyebrow"><% If isEdit Then %>ID <%= H(taskId) %><% Else %>Bản ghi mới<% End If %></p>
                    <h2>Thông tin công việc / 作業情報</h2>
                </div>
                <span class="required-note"><span aria-hidden="true">*</span> Bắt buộc</span>
            </div>

            <div class="field">
                <label for="title">Tiêu đề <span aria-hidden="true">*</span></label>
                <input id="title" name="title" type="text" maxlength="150" required autofocus
                    value="<%= H(taskTitle) %>" data-counted-field aria-describedby="title-count">
                <span id="title-count" class="field-hint" data-counter-for="title">0 / 150</span>
            </div>

            <div class="field">
                <label for="description">Mô tả</label>
                <textarea id="description" name="description" rows="7" maxlength="4000"
                    data-counted-field aria-describedby="description-count"><%= H(taskDescription) %></textarea>
                <span id="description-count" class="field-hint" data-counter-for="description">0 / 4000</span>
            </div>

            <div class="form-grid">
                <div class="field">
                    <label for="assignee">Người phụ trách</label>
                    <input id="assignee" name="assignee" type="text" maxlength="100" value="<%= H(taskAssignee) %>">
                </div>

                <div class="field">
                    <label for="due_date">Hạn hoàn thành</label>
                    <input id="due_date" name="due_date" type="date" value="<%= H(taskDueDate) %>">
                    <span class="field-hint">Định dạng gửi lên server: yyyy-mm-dd</span>
                </div>

                <div class="field">
                    <label for="status">Trạng thái <span aria-hidden="true">*</span></label>
                    <select id="status" name="status" required>
                        <option value="Todo"<% If taskStatus = STATUS_TODO Then Response.Write " selected" End If %>>Chưa làm / 未着手</option>
                        <option value="InProgress"<% If taskStatus = STATUS_IN_PROGRESS Then Response.Write " selected" End If %>>Đang làm / 対応中</option>
                        <option value="Done"<% If taskStatus = STATUS_DONE Then Response.Write " selected" End If %>>Hoàn thành / 完了</option>
                    </select>
                </div>

                <div class="field">
                    <label for="priority">Ưu tiên <span aria-hidden="true">*</span></label>
                    <select id="priority" name="priority" required>
                        <option value="1"<% If taskPriority = PRIORITY_LOW Then Response.Write " selected" End If %>>Thấp</option>
                        <option value="2"<% If taskPriority = PRIORITY_NORMAL Then Response.Write " selected" End If %>>Bình thường</option>
                        <option value="3"<% If taskPriority = PRIORITY_HIGH Then Response.Write " selected" End If %>>Cao</option>
                    </select>
                </div>
            </div>

            <div class="form-actions">
                <button class="button button-primary" type="submit" data-submit-button><% If isEdit Then %>Lưu thay đổi<% Else %>Tạo công việc<% End If %></button>
                <a class="button button-ghost" href="default.asp">Hủy</a>
            </div>
        </form>
    </main>

    <script src="assets/app.js"></script>
</body>
</html>
