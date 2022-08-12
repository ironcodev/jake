<%
    class ServiceResponse
        public Success
        public Service
        public Action
        public Status
        public Message
        public Date
        public Data
        public Error
        public Info

        public sub class_initialize()
            me.Date = now()
        end sub

        public function GetKey(lang)
            GetKey = "/" & lang & "/" & Service & "/" & Action & "/" & Status
        end function

        public sub setStatus(s, d)
            me.Status = s

            if IsObject(d) then
                set me.Data = d
            else
                me.Data = d
            end if
            
            me.Success = Instr(1, s, "success", vbTextCompare) > 0
        end sub
        public Sub Failed()
            call me.setStatus("Failed", null)
        end sub
        public Sub Faulted()
            call me.setStatus("Faulted", null)
        end sub
        public Sub notFound()
            call me.setStatus("NotFound", null)
        end sub
        public Sub Succeeded(d)
            call me.setStatus("Success", d)
        end sub
        public sub Copy(sr)
            me.Success = sr.Success 
            me.Service = sr.Service
            me.Action = sr.Action
            me.Status = sr.Status
            me.Message = sr.Message
            me.Date = sr.Date
            me.Data = sr.Data
            me.Info = sr.Info
        end sub
        public function toXml()
            toXml = "<ServiceResponse>" & vbCrLf & _
                    vbTab & "<Success>" & Server.HtmlEncode(Success) & "</Success>" & vbCrLf & _
                    vbTab & "<Status>" & Server.HtmlEncode(Status) & "</Status>" & vbCrLf & _
                    IIf(IsVoid(Message), "", vbTab & "<Message>" & Server.HtmlEncode(Message) & "</Message>" & vbCrLf) & _
                    IIf(IsVoid(Error), "", vbTab & "<Error>" & Server.HtmlEncode(Error) & "</Error>" & vbCrLf) & _
                    vbTab & "<Date>" & Server.HtmlEncode(StrDate(Date)) & "</Date>" & vbCrLf & _
                    IIf(IsVoid(Data), "", vbTab & "<Data>" & Server.HtmlEncode(SafeCStr(Data)) & "</Data>" & vbCrLf) & _
                    "</ServiceResponse>"
        end function

        public function toJson()
            toJson = "{" & vbCrLf & _
                    vbTab & " ""Success"": " & LCase(CStr(Success)) & "," & vbCrLf & _
                    vbTab & " ""Status"": """ & Status & """" & vbCrLf & _
                    vbTab & " ""Message"": """ & Message & """" & vbCrLf & _
                    IIf(IsVoid(Error), "", vbTab & " ""Error"": """ & Error & """" & vbCrLf) & _
                    vbTab & " ""Date"": """ & StrDate(Date) & """" & vbCrLf & _
                    IIf(IsVoid(Data), "", vbTab & " ""Data"": """ & SafeCStr(Data) & """" & vbCrLf) & _
                    "}"
        end function

        public sub Log()
            call Output.Log(Replace(Replace(Server.HtmlEncode(toXml()), vbCrLf, "<br/>"), vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;"))
        end sub
    end class
%>