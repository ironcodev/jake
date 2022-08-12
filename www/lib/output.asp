<%
    class clsOutput
        public property get DebugMode()
            If DEBUG_MODE And (IsNullOrEmpty(DEBUG_USERS) Or Instr(DEBUG_USERS, AccountService.CurrentUser) > 0) Then
                DebugMode = true
            Else
                DebugMode = false
            End If
        End property
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        public Function GetLog(title, msg)
            If DebugMode Then
                GetLog = vbCrLf & "<!-- " & title & Now() & vbCrLf & msg & vbCrLf & "-->"
            Else
                GetLog = ""
            End If
        End Function
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        public Sub Write(msg)
            Response.Write msg
        End Sub

        public Sub Lug(title, msg)
            If DebugMode Then
                me.Write GetLog(title, msg)
            End If
        End Sub
        
        public Sub Log(msg)
            Call Lug("", msg)
        End Sub

        public Sub Warn(msg)
            Call Lug("Warning", msg)
        End Sub

        public Sub Danger(msg)
            Call Lug("Danger", msg)
        End Sub

        public Sub Debug(msg)
            Call Lug("Debug", msg)
        End Sub

        public Sub Info(msg)
            Call Lug("Info", msg)
        End Sub
    end class

    dim Output

    set Output = new clsOutput
%>