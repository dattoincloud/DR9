<%
Const STATUS_TODO = "Todo"
Const STATUS_IN_PROGRESS = "InProgress"
Const STATUS_DONE = "Done"

Const PRIORITY_LOW = 1
Const PRIORITY_NORMAL = 2
Const PRIORITY_HIGH = 3

Function TextValue(ByVal value)
    If IsNull(value) Or IsEmpty(value) Then
        TextValue = ""
    Else
        TextValue = CStr(value)
    End If
End Function

Function H(ByVal value)
    H = Server.HTMLEncode(TextValue(value))
End Function

Function FormText(ByVal fieldName)
    FormText = Trim(TextValue(Request.Form(fieldName)))
End Function

Function QueryText(ByVal fieldName)
    QueryText = Trim(TextValue(Request.QueryString(fieldName)))
End Function

Function IsBlankText(ByVal value)
    Dim normalized

    normalized = TextValue(value)
    normalized = Replace(normalized, vbTab, "")
    normalized = Replace(normalized, vbCr, "")
    normalized = Replace(normalized, vbLf, "")
    normalized = Replace(normalized, ChrW(&H3000), "")
    normalized = Replace(normalized, " ", "")

    IsBlankText = (Len(normalized) = 0)
End Function

Function ParseOptionalId(ByVal rawValue)
    Dim valueText
    Dim matcher

    valueText = Trim(TextValue(rawValue))
    If Len(valueText) = 0 Then
        ParseOptionalId = 0
        Exit Function
    End If

    Set matcher = Server.CreateObject("VBScript.RegExp")
    matcher.Pattern = "^[1-9][0-9]{0,8}$"
    matcher.Global = False

    If Not matcher.Test(valueText) Then
        ParseOptionalId = -1
        Exit Function
    End If

    ParseOptionalId = CLng(valueText)
End Function

Function ParseIsoDate(ByVal rawValue, ByRef isValid)
    Dim valueText
    Dim matcher
    Dim yearPart
    Dim monthPart
    Dim dayPart
    Dim parsedDate

    valueText = Trim(TextValue(rawValue))
    isValid = True

    If Len(valueText) = 0 Then
        ParseIsoDate = Null
        Exit Function
    End If

    Set matcher = Server.CreateObject("VBScript.RegExp")
    matcher.Pattern = "^[0-9]{4}-[0-9]{2}-[0-9]{2}$"
    matcher.Global = False

    If Not matcher.Test(valueText) Then
        isValid = False
        ParseIsoDate = Null
        Exit Function
    End If

    yearPart = CInt(Left(valueText, 4))
    monthPart = CInt(Mid(valueText, 6, 2))
    dayPart = CInt(Right(valueText, 2))

    On Error Resume Next
    parsedDate = DateSerial(yearPart, monthPart, dayPart)
    If Err.Number <> 0 Then
        Err.Clear
        isValid = False
        On Error GoTo 0
        ParseIsoDate = Null
        Exit Function
    End If
    On Error GoTo 0

    If Year(parsedDate) <> yearPart Or Month(parsedDate) <> monthPart Or Day(parsedDate) <> dayPart Then
        isValid = False
        ParseIsoDate = Null
        Exit Function
    End If

    ParseIsoDate = parsedDate
End Function

Function FormatIsoDate(ByVal value)
    If IsNull(value) Or IsEmpty(value) Or Not IsDate(value) Then
        FormatIsoDate = ""
    Else
        FormatIsoDate = Right("0000" & CStr(Year(value)), 4) & "-" & _
            Right("00" & CStr(Month(value)), 2) & "-" & _
            Right("00" & CStr(Day(value)), 2)
    End If
End Function

Function NzLong(ByVal value)
    If IsNull(value) Or IsEmpty(value) Then
        NzLong = 0
    Else
        NzLong = CLng(value)
    End If
End Function

Function IsAllowedStatus(ByVal value)
    IsAllowedStatus = (value = STATUS_TODO Or value = STATUS_IN_PROGRESS Or value = STATUS_DONE)
End Function

Function StatusLabel(ByVal value)
    Select Case value
        Case STATUS_TODO
            StatusLabel = "Chưa làm / 未着手"
        Case STATUS_IN_PROGRESS
            StatusLabel = "Đang làm / 対応中"
        Case STATUS_DONE
            StatusLabel = "Hoàn thành / 完了"
        Case Else
            StatusLabel = TextValue(value)
    End Select
End Function

Function StatusCssClass(ByVal value)
    Select Case value
        Case STATUS_TODO
            StatusCssClass = "todo"
        Case STATUS_IN_PROGRESS
            StatusCssClass = "in-progress"
        Case STATUS_DONE
            StatusCssClass = "done"
        Case Else
            StatusCssClass = "unknown"
    End Select
End Function

Function IsAllowedPriority(ByVal value)
    IsAllowedPriority = (value = PRIORITY_LOW Or value = PRIORITY_NORMAL Or value = PRIORITY_HIGH)
End Function

Function PriorityLabel(ByVal value)
    Select Case CLng(value)
        Case PRIORITY_LOW
            PriorityLabel = "Thấp"
        Case PRIORITY_NORMAL
            PriorityLabel = "Bình thường"
        Case PRIORITY_HIGH
            PriorityLabel = "Cao"
        Case Else
            PriorityLabel = "-"
    End Select
End Function

Sub AppendValidationError(ByRef errors, ByVal message)
    If Len(errors) > 0 Then
        errors = errors & vbLf
    End If
    errors = errors & message
End Sub

Function GetCsrfToken()
    Dim token

    token = TextValue(Session("CsrfToken"))
    If Len(token) = 0 Then
        Randomize
        token = CStr(Session.SessionID) & "-" & Hex(CLng(Timer * 100)) & "-" & Hex(CLng(Rnd() * 2000000000))
        Session("CsrfToken") = token
    End If

    GetCsrfToken = token
End Function

Function IsValidCsrfToken(ByVal submittedToken)
    Dim expectedToken

    expectedToken = TextValue(Session("CsrfToken"))
    submittedToken = TextValue(submittedToken)

    IsValidCsrfToken = (Len(expectedToken) > 0 And Len(submittedToken) > 0 And _
        StrComp(expectedToken, submittedToken, vbBinaryCompare) = 0)
End Function

Sub RequirePost()
    If UCase(TextValue(Request.ServerVariables("REQUEST_METHOD"))) <> "POST" Then
        Response.AddHeader "Allow", "POST"
        Call FailRequest(405, "Method Not Allowed", "Yêu cầu này chỉ chấp nhận phương thức POST.")
    End If
End Sub

Sub AppendServerDiagnostic(ByVal contextName, ByVal detail)
    Dim logText

    logText = Replace(TextValue(detail), vbCr, " ")
    logText = Replace(logText, vbLf, " ")
    logText = Left(contextName & "=" & logText, 70)

    On Error Resume Next
    Response.AppendToLog " " & logText
    Err.Clear
    On Error GoTo 0
End Sub

Sub FailRequest(ByVal statusCode, ByVal statusText, ByVal message)
    Response.Clear
    Response.Status = CStr(statusCode) & " " & statusText
    Response.ContentType = "text/html"
    Response.Charset = "utf-8"
    Response.Write "<!doctype html><html lang=""vi""><head><meta charset=""utf-8"">"
    Response.Write "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
    Response.Write "<title>" & H(statusText) & "</title><link rel=""stylesheet"" href=""assets/styles.css"">"
    Response.Write "</head><body><main class=""error-page card""><p class=""eyebrow"">HTTP " & H(statusCode) & "</p>"
    Response.Write "<h1>" & H(statusText) & "</h1><p>" & H(message) & "</p>"
    Response.Write "<p><a class=""button button-primary"" href=""default.asp"">Về danh sách</a></p></main></body></html>"
    Response.End
End Sub
%>
