<% option explicit %>
<!--#include virtual="/init.asp"-->
<html>
    <head>
        <title>500</title>
    </head>
    <body>
        <h1>Internal error 500</h1>
        <p>
<%
    on error resume next
    
    dim e

    set e = Server.GetLastError()
    
    Response.Write "ASPCode: " & e.ASPCode & "<br/>"
    Response.Write "ASPDescription: " & e.ASPDescription & "<br/>"
    Response.Write "Category: " & e.Category & "<br/>"
    Response.Write "Column: " & e.Column & "<br/>"
    Response.Write "Number: " & e.Number & "<br/>"
    Response.Write "File: " & e.File & "<br/>"
    Response.Write "Description: " & e.Description & "<br/>"
    Response.Write "Line: " & e.Line & "<br/>"
    Response.Write "Source: " & e.Source & "<br/>"
    Response.Write "<br/>"
    Response.Write "Err.Number: " & Err.Number & "<br/>"
    Response.Write "Err.Description: " & Err.Description & "<br/>"
%>
        </p>
    </body>
</html>
