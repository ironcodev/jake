<%
    class clsPathHelper
        public function GetExtension(filenameAndPath)
            dim result, i

            i = InstrRev(filenameAndPath, ".")

            if i > 0 then
                result = Mid(filenameAndPath, i + 1)
            else
                result = filenameAndPath
            end if

            GetExtension = result
        end function

        public function GetFileName(filenameAndPath)
            dim result, i

            i = InstrRev(filenameAndPath, "/")

            if i <= 0 then
                i = InstrRev(filenameAndPath, "\")
            end if

            if i > 0 then
                result = Mid(filenameAndPath, i + 1)
            else
                result = filenameAndPath
            end if

            GetFileName = result
        end function

        public function GetFileNameWithoutExtension(filenameAndPath)
            dim result, i

            result = GetFileName(filenameAndPath)

            i = InstrRev(result, ".")

            if i > 1 then
                result = Mid(result, 1, i - 1)
            end if

            GetFileNameWithoutExtension = result
        end function
    end class

    dim PathHelper

    set PathHelper = new clsPathHelper
%>