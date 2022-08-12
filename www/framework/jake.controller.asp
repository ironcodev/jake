<%
    public function View(vw, model, requestContext)
        set View = vw
        
        set View.Controller = me
        set View.RequestContext = requestContext
        set View.Model = model
    end function

    public function Json(data)
    end function

    public function Xml(data)
    end function

    public function File(path)
    end function

    public function Bytes(data)
    end function
%>