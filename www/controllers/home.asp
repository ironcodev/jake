<!--#include virtual="/init.asp"-->
<!--#include virtual="/views/home/index.asp"-->
<%
    class Controller
%>
<!--#include virtual="/framework/jake.controller.asp"-->
<%
        public function Index(requestContext)
            set Index = View(new ViewHomeIndex, null, requestContext)
        end function
    end class
%>
<!--#include virtual="/run.asp"-->