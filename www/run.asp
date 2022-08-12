<%
    dim c, r
        
    set c = new Controller
    set r = c.Execute()

    execute "r." & RouteValues.Action & "()"

    set r = Nothing
    set c = Nothing
%>