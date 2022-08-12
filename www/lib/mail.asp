<%
    class clsMailCDO
        public function GetResponse(action, errDesc)
            set GetResponse = new ServiceResponse

            GetResponse.Service = "Mail"
            GetResponse.Action = action
            GetResponse.Error = errDesc
        end function

        public function Send(sTo, subject, body, html)
            dim status, myMail, errDesc

            Do
                If IsNullOrEmpty(sTo) Then
                    status = "NoTo"
                    Exit Do
                End If

                If Not isEmail(sTo) Then
                    status = "InvalidTo"
                    Exit Do
                End If

                If IsNullOrEmpty(body) Then
                    status = "NoBody"
                    Exit Do
                End If

                On Error Resume Next
                
                Set myMail = CreateObject("CDO.Message")

                If Err.Number <> 0 Then
                    status = "Faulted"
                    errDesc = Err.Description
                    Exit Do
                End If

                myMail.Subject = subject
                myMail.Sender = MAIL_DEFAULT
                myMail.From = MAIL_DEFAULT
                myMail.To = sTo

                myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = "1" ' SMTP
                myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MAIL_SERVER
                myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = MAIL_PORT
                myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = MAIL_SSL
                myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = MAIL_TIMEOUT

                If MAIL_AUTHENTICATE Then
                    myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = MAIL_USER
                    myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = MAIL_PASS
                    myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                End If

                myMail.Configuration.Fields.Update

                If html Then
                    myMail.HTMLBody = body
                    myMail.HTMLBodyPart.Charset = "UTF-8"
                Else
                    myMail.TextBody = body
                    myMail.BodyPart.Charset = "UTF-8"
                End If

                myMail.Send

                If Err.Number <> 0 Then
                    status = "Failed"
                    errDesc = Err.Description
                    Exit Do
                End If
                
                set myMail = nothing

                status = "Success"
            Loop While False
            
            Set send = me.GetResponse("send", errDesc)
            call send.setStatus(status, 0)

            On Error Goto 0
        end function
    end class

    class clsMailFake
        public function GetResponse(action)
            set GetResponse = new ServiceResponse

            GetResponse.Service = "Mail"
            GetResponse.Action = action
        end function

        public function Send(sTo, subject, body, html)
            dim status, myMail

            Do
                If IsNullOrEmpty(sTo) Then
                    status = "NoTo"
                    Exit Do
                End If

                If Not isEmail(sTo) Then
                    status = "InvalidTo"
                    Exit Do
                End If

                If IsNullOrEmpty(body) Then
                    status = "NoBody"
                    Exit Do
                End If

                Response.Write "<!--" & vbCrLf & _
                                "FakeMailer:" & vbCrLf & _
                                "To: " & sTo & vbCrLf & _
                                "Subject: " & subject & vbCrLf & _
                                "Body: " & vbCrLf & body & vbCrLf & _
                                "-->"
            Loop While False
            
            Set send = me.GetResponse("send")
            call send.setStatus(status, 0)
        end function
    end class

    dim MailUtil

    If IS_DEVELOPMENT Then
        set MailUtil = new clsMailFake
    Else
        set MailUtil = new clsMailCDO
    End If
%>