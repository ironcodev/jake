<%

    class clsFormatter
        private FORMATTER_STATE_START
        private FORMATTER_STATE_SLASH
        private FORMATTER_STATE_BRACE
        
        public sub class_initialize
            FORMATTER_STATE_START = 1
            FORMATTER_STATE_SLASH = 2
            FORMATTER_STATE_BRACE = 3
        end sub

        public function Format(template, args)
            dim state, ch, i, result, key

            state = FORMATTER_STATE_START

            set result = new StringBuffer

            for i = 1 to len(template)
                ch = Mid(template, i, 1)

                select case state
                    case FORMATTER_STATE_START:
                        if ch = "\" then
                            state = FORMATTER_STATE_SLASH
                        elseif ch = "{" then
                            state = FORMATTER_STATE_BRACE
                        else
                            result.Append ch
                        end if
                    case FORMATTER_STATE_SLASH:
                        if ch = "\" then
                            result.Append "\"
                        elseif ch = "{" then
                            result.Append "{"
                        else
                            result.Append "\" & ch
                        end if

                        state = FORMATTER_STATE_START
                    case FORMATTER_STATE_BRACE:
                        if ch = "}" then
                            if not IsNullOrEmpty(key) then
                                if args.Exists(key) then
                                    result.Append args.Item(key)
                                end if

                                key = ""
                            end if

                            state = FORMATTER_STATE_START
                        else
                            key = key & ch
                        end if
                end select
            next

            Format = result.Flush()

            set result = Nothing
        end function

        public function Argify(keys, values)
            dim arr, i, value

            set Argify = Server.CreateObject("Scripting.Dictionary")

            arr = Split(SafeCStr(keys), ",")

            if ubound(arr) > 0 and IsArray(values) then
                for i = lbound(arr) to ubound(arr)
                    if i >= lbound(values) and i <= ubound(values) then
                        if IsObject(values(i)) then
                            set value = values(i)
                        else
                            value = values(i)
                        end if
                    end if

                    Argify.add arr(i), value
                next
            end if
        end function
    end class

    dim Formatter

    set Formatter = new clsFormatter
%>