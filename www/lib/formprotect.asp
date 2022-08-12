<%
    class clsFormProtect
        public function GetToken()
            dim result

            result = CStr(now())
            result = Base64Encode(XorCipher(result, CRYPT_KEY))

            GetToken = result
        end function

        public function CheckToken(token)
            dim x
            dim result

            on Error Resume Next

            result = false

            x = CDate(XorCipher(Base64Decode(token), CRYPT_KEY))

            If Err.Number = 0 Then
                If DateDiff("s", x, now()) < 2 * 60 Then    ' token assumed valid only for 2 minutes
                    result = true
                End If
            End If

            on Error Goto 0

            CheckToken = result
        end function
    end class

    dim FormProtect

    set FormProtect = new clsFormProtect
%>