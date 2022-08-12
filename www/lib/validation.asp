<%
    Function IsEmail(e)
		Dim regx
    
		Set regx = New RegExp
		
		regx.pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$"
		regx.ignoreCase = True
		regx.multiLine = False
		
		IsEmail = regx.test(e)
		
		Set regx = Nothing
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsUserName(e)
		Dim regx
    
		Set regx = New RegExp
		
		regx.pattern = "^\w{3,15}$"
		regx.ignoreCase = True
		regx.multiLine = False
		
		IsUserName = regx.test(e)
		
		Set regx = Nothing
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	public function IsAlpha(ch)
		dim s, n, result

		result = false

		s = SafeCStr(ch)

		If len(s) > 0 then
			n = Asc(Mid(s, 1, 1))

			result = IIf((n >= 65 and n <= 65 + 26) or (n >= 97 and n <= 97 + 26), true, false)
		end if

		IsAlpha = result
	end function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	public function IsLower(ch)
		dim s, n, result

		result = false

		s = SafeCStr(ch)

		If len(s) > 0 then
			n = Asc(Mid(s, 1, 1))

			result = IIf(n >= 97 and n <= 97 + 26, true, false)
		end if

		IsLower = result
	end function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	public function IsUpper(ch)
		dim s, n, result

		result = false

		s = SafeCStr(ch)

		If len(s) > 0 then
			n = Asc(Mid(s, 1, 1))

			result = IIf(n >= 65 and n <= 65 + 26, true, false)
		end if

		IsUpper = result
	end function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	public function IsDigit(ch)
		dim s, n, result

		result = false

		s = SafeCStr(ch)

		If len(s) > 0 then
			n = Asc(Mid(s, 1, 1))

			result = IIf(n >= 48 and n <= 48 + 10, true, false)
		end if

		IsAlpha = result
	end function
%>