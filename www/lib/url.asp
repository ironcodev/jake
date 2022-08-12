<%
    class clsUrl
        private mUrl
        private mScheme
        private mHost
        private mPort
        private mPath
        private mQuery
        private mFragment
        private mFileName
        private fileNameSet
        private mExtension
        private extensionSet
        private mAuthority

        public property get Url()
            Url = mUrl
        end property

        public property get Scheme()
            Scheme = mScheme
        end property

        public property get Host()
            Host = mHost
        end property

        public property get Port()
            Port = mPort
        end property

        public property get Path()
            Path = mPath
        end property

        public property get Query()
            Query = mQuery
        end property

        public property get Fragment()
            Fragment = mFragment
        end property

        public property get Authority()
            Authority = mHost & IIf(IsNullOrEmpty(mPort), "", ":" & mPort)
        end property

        public property get FileName()
            FileName = mFileName
        end property

        public property get Extension()
            Extension = mExtension
        end property

        private sub FindFileNameAndExtension()
            dim i, j

            i = InstrRev(mPath, "/")

            If i > 1 Then
                j = InstrRev(mPath, ".")

                If j > 0 And j > i Then
                    mFileName = Mid(mPath, i + 1)
                    mExtension = Mid(mPath, j + 1)
                    mPath = Left(mPath, i - 1)
                End If
            End If
        end sub

        public sub Load(value)
            dim i, j, k, l

            mScheme = ""
            mHost = ""
            mPort = ""
            mPath = ""
            mQuery = ""
            mFragment = ""
            mFileName = ""
            fileNameSet = false
            mExtension = ""
            extensionSet = false
            mAuthority = ""

            mUrl = SafeCStr(value)

            i = Instr(mUrl, "://")

            If i > 0 Then
                mScheme = IIf(i > 1, Left(mUrl, i - 1), "")
            End If

            j = Instr(i + 3, mUrl, ":")

            If j > 0 Then
                mHost = Mid(mUrl, i + 3, j - i - 3)

                k = Instr(j + 1, mUrl, "/")

                If k > 0 Then
                    mPort = Mid(mUrl, j + 1, k - j - 1)

                    j = Instr(k + 1, mUrl, "?")

                    If j > 0 Then
                        mPath = Mid(mUrl, k, j - k)

                        k = Instr(j + 1, mUrl, "#")

                        If k > 0 Then
                            mQuery = Mid(mUrl, j + 1, k - j - 1)
                            mFragment = Mid(mUrl, k + 1)
                        Else
                            mQuery = Mid(mUrl, j + 1)
                        End If
                    Else
                        mPath = Mid(mUrl, k)
                    End If
                Else
                    mPort = Mid(mUrl, j + 1)
                End If
            Else
                k = Instr(i + 3, mUrl, "/")

                If k > 0 Then
                    mHost = Mid(mUrl, i + 3, k - i - 3)

                    j = Instr(k + 1, mUrl, "?")

                    If j > 0 Then
                        mPath = Mid(mUrl, k, j - k)

                        k = Instr(j + 1, mUrl, "#")

                        If k > 0 Then
                            mQuery = Mid(mUrl, j + 1, k - j - 1)
                            mFragment = Mid(mUrl, k + 1)
                        Else
                            mQuery = Mid(mUrl, j + 1)
                        End If
                    Else
                        mPath = Mid(mUrl, k)
                    End If
                Else
                    mHost = Mid(mUrl, i + 3)
                End If
            End If

            call FindFileNameAndExtension()
        end sub

        public function ToXml()
            ToXml = "<Url>" & vbCrLf & _
                    vbTab & "<Value>" & Server.HtmlEncode(Url) & "</Value>" & vbCrLf & _
                    vbTab & "<Scheme>" & Server.HtmlEncode(Scheme) & "</Scheme>" & vbCrLf & _
                    vbTab & "<Host>" & Server.HtmlEncode(Host) & "</Host>" & vbCrLf & _
                    vbTab & "<Port>" & Server.HtmlEncode(Port) & "</Port>" & vbCrLf & _
                    vbTab & "<Path>" & Server.HtmlEncode(Path) & "</Path>" & vbCrLf & _
                    vbTab & "<Query>" & Server.HtmlEncode(Query) & "</Query>" & vbCrLf & _
                    vbTab & "<Fragment>" & Server.HtmlEncode(Fragment) & "</Fragment>" & vbCrLf & _
                    vbTab & "<FileName>" & Server.HtmlEncode(FileName) & "</FileName>" & vbCrLf & _
                    vbTab & "<Extension>" & Server.HtmlEncode(Extension) & "</Extension>" & vbCrLf & _
                    vbTab & "<Authority>" & Server.HtmlEncode(Authority) & "</Authority>" & vbCrLf & _
                    "</Url>"
        end function
    end class

    dim UrlHelper

    set UrlHelper = new clsUrl
%>