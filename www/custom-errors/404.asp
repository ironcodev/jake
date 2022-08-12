<% option explicit %>
<!--#include virtual="/init.asp"-->
<html>
    <head>
        <title>404</title>
    </head>
    <body>
        <h1>404</h1>
        <p>
<%
    dim origin, url, path, i

    origin = Request.ServerVariables("QUERY_STRING")

    i = Instr(origin, ";")

    If i > 0 Then
        url = Mid(origin, i + 1)
    Else
        url = origin
    End If

    Response.Write "origin: " & origin & "<br/>"
    Response.Write "path: " & path & "<br/>"
    Response.Write "url: " & url & "<br/>"
%>
        </p>
    </body>
</html>
