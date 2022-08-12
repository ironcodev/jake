<%
    public function Execute()
        dim t

        set t = new Template

        set t.View = me

        set Execute = t.Render()
    end function
%>