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
Dim connection
Dim command
Dim affectedRows
Dim databaseError

Call RequirePost()

If Not IsValidCsrfToken(Request.Form("csrf_token")) Then
    Call FailRequest(403, "Forbidden", "Phiên làm việc hoặc CSRF token không hợp lệ. Hãy tải lại danh sách và thử lại.")
End If

taskId = ParseOptionalId(FormText("id"))
If taskId <= 0 Then
    Call FailRequest(400, "Bad Request", "Mã công việc không hợp lệ.")
End If

Set connection = OpenConnectionOrFail()
Set command = NewCommand(connection, "DELETE FROM [Tasks] WHERE [Id] = ?")
Call AddIntegerParameter(command, taskId)
affectedRows = 0
databaseError = ""

On Error Resume Next
command.Execute affectedRows, , adExecuteNoRecords
If Err.Number <> 0 Then
    databaseError = Err.Description
    Err.Clear
End If
On Error GoTo 0

If Len(databaseError) > 0 Then
    Set command = Nothing
    Call CloseConnectionQuietly(connection)
    Call FailDatabaseRequest("DB_DELETE", databaseError)
End If

Set command = Nothing
Call CloseConnectionQuietly(connection)

If affectedRows = 0 Then
    Call FailRequest(404, "Not Found", "Công việc cần xóa không còn tồn tại.")
End If

Response.Redirect "default.asp?deleted=1"
%>
