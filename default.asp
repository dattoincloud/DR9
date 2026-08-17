<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/util.asp"-->
<!--#include file="includes/db.asp"-->
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
Session.CodePage = 65001

Dim searchText
Dim statusFilter
Dim connection
Dim summaryCommand
Dim summaryRows
Dim listCommand
Dim taskRows
Dim taskData
Dim hasTasks
Dim rowIndex
Dim sqlText
Dim totalCount
Dim todoCount
Dim inProgressCount
Dim doneCount
Dim csrfToken
Dim rowId
Dim rowTitle
Dim rowAssignee
Dim rowStatus
Dim rowPriority
Dim rowDueDate
Dim rowUpdatedAt
Dim dueClass
Dim databaseError

searchText = QueryText("q")
statusFilter = QueryText("status")

If Len(searchText) > 100 Then
    Call FailRequest(400, "Bad Request", "Từ khóa tìm kiếm không được dài quá 100 ký tự.")
End If

If Len(statusFilter) > 0 And Not IsAllowedStatus(statusFilter) Then
    Call FailRequest(400, "Bad Request", "Trạng thái lọc không hợp lệ.")
End If

Set connection = OpenConnectionOrFail()
csrfToken = GetCsrfToken()

Set summaryCommand = NewCommand(connection, _
    "SELECT Count(*) AS TotalCount, " & _
    "Sum(IIf([Status]='Todo', 1, 0)) AS TodoCount, " & _
    "Sum(IIf([Status]='InProgress', 1, 0)) AS InProgressCount, " & _
    "Sum(IIf([Status]='Done', 1, 0)) AS DoneCount FROM [Tasks]")
Set summaryRows = Nothing
databaseError = ""
On Error Resume Next
Set summaryRows = summaryCommand.Execute()
If Err.Number = 0 Then
    totalCount = NzLong(summaryRows("TotalCount").Value)
    todoCount = NzLong(summaryRows("TodoCount").Value)
    inProgressCount = NzLong(summaryRows("InProgressCount").Value)
    doneCount = NzLong(summaryRows("DoneCount").Value)
End If
If Err.Number <> 0 Then
    databaseError = Err.Description
    Err.Clear
End If
On Error GoTo 0

If Len(databaseError) > 0 Then
    Call CloseRecordsetQuietly(summaryRows)
    Call CloseConnectionQuietly(connection)
    Call FailDatabaseRequest("DB_SUMMARY", databaseError)
End If

Call CloseRecordsetQuietly(summaryRows)
Set summaryCommand = Nothing

sqlText = "SELECT [Id], [Title], [Assignee], [Status], [Priority], [DueDate], [UpdatedAt] FROM [Tasks] WHERE 1=1"
Set listCommand = NewCommand(connection, "")

If Len(statusFilter) > 0 Then
    sqlText = sqlText & " AND [Status] = ?"
    Call AddTextParameter(listCommand, statusFilter, 20)
End If

If Len(searchText) > 0 Then
    sqlText = sqlText & " AND (InStr(1, [Title], ?, 1) > 0 OR " & _
        "InStr(1, [Description], ?, 1) > 0 OR InStr(1, [Assignee], ?, 1) > 0)"
    Call AddTextParameter(listCommand, searchText, 100)
    Call AddTextParameter(listCommand, searchText, 100)
    Call AddTextParameter(listCommand, searchText, 100)
End If

sqlText = sqlText & " ORDER BY [UpdatedAt] DESC, [Id] DESC"
listCommand.CommandText = sqlText
Set taskRows = Nothing
hasTasks = False
databaseError = ""

On Error Resume Next
Set taskRows = listCommand.Execute()
If Err.Number = 0 Then
    If Not taskRows.EOF Then
        taskData = taskRows.GetRows()
        If Err.Number = 0 Then hasTasks = True
    End If
End If
If Err.Number <> 0 Then
    databaseError = Err.Description
    Err.Clear
End If
On Error GoTo 0

If Len(databaseError) > 0 Then
    Call CloseRecordsetQuietly(taskRows)
    Call CloseConnectionQuietly(connection)
    Call FailDatabaseRequest("DB_LIST", databaseError)
End If

Call CloseRecordsetQuietly(taskRows)
Set listCommand = Nothing
Call CloseConnectionQuietly(connection)
%>
<!doctype html>
<html lang="vi">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Legacy Task Board</title>
    <link rel="stylesheet" href="assets/styles.css">
</head>
<body>
    <header class="site-header">
        <div class="shell header-inner">
            <div>
                <p class="eyebrow">Classic ASP · ADO · MS Access</p>
                <h1>Legacy Task Board</h1>
                <p class="subtitle">Demo quản lý công việc Việt–Nhật / 作業管理デモ</p>
            </div>
            <a class="button button-primary" href="task_form.asp">+ Tạo công việc</a>
        </div>
    </header>

    <main class="shell main-content">
        <% If QueryText("saved") = "1" Then %>
            <div class="alert alert-success" role="status">Đã lưu công việc thành công.</div>
        <% ElseIf QueryText("deleted") = "1" Then %>
            <div class="alert alert-success" role="status">Đã xóa công việc.</div>
        <% End If %>

        <section class="stats-grid" aria-label="Thống kê công việc">
            <article class="stat-card">
                <span>Tổng</span>
                <strong><%= H(totalCount) %></strong>
            </article>
            <article class="stat-card stat-todo">
                <span>Chưa làm</span>
                <strong><%= H(todoCount) %></strong>
            </article>
            <article class="stat-card stat-progress">
                <span>Đang làm</span>
                <strong><%= H(inProgressCount) %></strong>
            </article>
            <article class="stat-card stat-done">
                <span>Hoàn thành</span>
                <strong><%= H(doneCount) %></strong>
            </article>
        </section>

        <section class="card toolbar-card">
            <form method="get" action="default.asp" class="filter-form" data-filter-form>
                <div class="field field-grow">
                    <label for="q">Tìm kiếm</label>
                    <input id="q" name="q" type="search" maxlength="100"
                        value="<%= H(searchText) %>" placeholder="Tiêu đề, mô tả, người phụ trách">
                </div>
                <div class="field">
                    <label for="status">Trạng thái</label>
                    <select id="status" name="status" data-auto-submit>
                        <option value="">Tất cả</option>
                        <option value="Todo"<% If statusFilter = STATUS_TODO Then Response.Write " selected" End If %>>Chưa làm / 未着手</option>
                        <option value="InProgress"<% If statusFilter = STATUS_IN_PROGRESS Then Response.Write " selected" End If %>>Đang làm / 対応中</option>
                        <option value="Done"<% If statusFilter = STATUS_DONE Then Response.Write " selected" End If %>>Hoàn thành / 完了</option>
                    </select>
                </div>
                <div class="toolbar-actions">
                    <button class="button button-secondary" type="submit">Lọc</button>
                    <% If Len(searchText) > 0 Or Len(statusFilter) > 0 Then %>
                        <a class="button button-ghost" href="default.asp">Xóa lọc</a>
                    <% End If %>
                </div>
            </form>
        </section>

        <section class="card table-card" aria-labelledby="task-list-title">
            <div class="section-heading">
                <div>
                    <p class="eyebrow">Danh sách</p>
                    <h2 id="task-list-title">Công việc</h2>
                </div>
                <span class="result-count"><%= H(totalCount) %> bản ghi trong hệ thống</span>
            </div>

            <% If Not hasTasks Then %>
                <div class="empty-state">
                    <h3>Không tìm thấy công việc</h3>
                    <p>Hãy đổi bộ lọc hoặc tạo một công việc mới.</p>
                </div>
            <% Else %>
                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>Công việc</th>
                                <th>Trạng thái</th>
                                <th>Ưu tiên</th>
                                <th>Hạn</th>
                                <th>Cập nhật</th>
                                <th><span class="sr-only">Thao tác</span></th>
                            </tr>
                        </thead>
                        <tbody>
                        <% For rowIndex = 0 To UBound(taskData, 2)
                            rowId = CLng(taskData(0, rowIndex))
                            rowTitle = TextValue(taskData(1, rowIndex))
                            rowAssignee = TextValue(taskData(2, rowIndex))
                            rowStatus = TextValue(taskData(3, rowIndex))
                            rowPriority = CLng(taskData(4, rowIndex))
                            rowDueDate = taskData(5, rowIndex)
                            rowUpdatedAt = taskData(6, rowIndex)
                            dueClass = ""
                            If Not IsNull(rowDueDate) And rowStatus <> STATUS_DONE Then
                                If DateValue(rowDueDate) < Date() Then dueClass = " overdue"
                            End If
                        %>
                            <tr>
                                <td data-label="Công việc">
                                    <a class="task-title" href="task_form.asp?id=<%= H(rowId) %>"><%= H(rowTitle) %></a>
                                    <span class="task-meta">
                                        <% If Len(rowAssignee) > 0 Then %>Phụ trách: <%= H(rowAssignee) %><% Else %>Chưa phân công<% End If %>
                                    </span>
                                </td>
                                <td data-label="Trạng thái">
                                    <span class="badge badge-<%= H(StatusCssClass(rowStatus)) %>"><%= H(StatusLabel(rowStatus)) %></span>
                                </td>
                                <td data-label="Ưu tiên"><%= H(PriorityLabel(rowPriority)) %></td>
                                <td data-label="Hạn" class="date-cell<%= dueClass %>">
                                    <% If IsNull(rowDueDate) Then %>—<% Else %><%= H(FormatIsoDate(rowDueDate)) %><% End If %>
                                </td>
                                <td data-label="Cập nhật" class="date-cell"><%= H(FormatIsoDate(rowUpdatedAt)) %></td>
                                <td class="row-actions">
                                    <a class="button button-small button-ghost" href="task_form.asp?id=<%= H(rowId) %>">Sửa</a>
                                    <form method="post" action="task_delete.asp" class="inline-form" data-delete-form data-task-title="<%= H(rowTitle) %>">
                                        <input type="hidden" name="id" value="<%= H(rowId) %>">
                                        <input type="hidden" name="csrf_token" value="<%= H(csrfToken) %>">
                                        <button class="button button-small button-danger" type="submit">Xóa</button>
                                    </form>
                                </td>
                            </tr>
                        <%
                        Next %>
                        </tbody>
                    </table>
                </div>
            <% End If %>
        </section>

        <p class="footer-note">Server-rendered HTML · VBScript · ES5-friendly JavaScript · Access parameterized queries</p>
    </main>

    <script src="assets/app.js"></script>
</body>
</html>
