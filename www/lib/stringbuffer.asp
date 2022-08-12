<%
    class StringBuffer
        public arr()
        public Growth
        private current
        private size
        private iLength

        public sub class_initialize
            Growth = 4
            current = 0
            size = 0
            iLength = 0
        end sub

        public sub Append(s)
            if current = size then
                size = size + Growth
                redim preserve arr(size)
            end if

            arr(current) = SafeCStr(s)
            
            iLength = iLength + len(arr(current))
            current = current + 1
        end sub

        public function Flush()
            Flush = join(arr, "")
            current = 0
            iLength = 0
            size = 0
            redim arr(0)
        end function

        public property get Length
            Length = iLength
        end property

        public function Merge(delimiter)
            Merge = join(arr, IIf(IsNullOrEmpty(delimiter), "", delimiter))
        end function
    end class
%>