<%
    Function IIf(condition, true_part, false_part)
		Dim result
		
		If Convert.ToBool(condition) Then
			result = true_part
		Else
			result = false_part
		End If
		
		If IsObject(result) Then
			set IIf = result
		Else
			IIf = result
		End If
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsInstance(o)
		Dim result
		
		result = False

		If IsObject(o) Then
			If Not o Is Nothing Then
				result = True
			End If
		End If
		
		IsInstance = result
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsNothing(o)
		Dim result
		
		result = False

		If IsObject(o) Then
			result = o Is Nothing
		End If
		
		IsNothing = result
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsVoid(o)
		IsVoid = IsNothing(o) or IsNullOrEmpty(o)
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsBasicType(x)
		IsBasicType = Not (IsNull(x) Or IsEmpty(x) Or IsInstance(x))
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsNumber(x)
		dim result

		result = false

		If IsBasicType(x) Then
			result = IsNumeric(x) And Not VarType(x) = vbBoolean
		End if

		IsNumber = result
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsPostBack()
		If StrComp(Request.ServerVariables("REQUEST_METHOD"), "POST", vbTextCompare) = 0 Then
			IsPostBack = true
		Else
			IsPostBack = false
		End If
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function IsNullOrEmpty(x)
		If IsEmpty(x) or IsNull(x) Then
			IsNullOrEmpty = true
		ElseIf VarType(x) = vbString Then
			IsNullOrEmpty = IIf(Len(Trim(x)) = 0, true, false)
		Else
			IsNullOrEmpty = false
		End If
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function PadLeft(x, size, ch)
		dim result

		result = IIf(IsVoid(x), "", IIf(IsObject(x), TypeName(x), CStr(x)))
		size = IIf(IsNumber(size), size, Len(result))

		If Len(result) < size Then
			result = RepeatStr(ch, size - Len(result)) & result
		End If

		PadLeft = result
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function PadRight(x, size, ch)
		dim result

		result = IIf(IsVoid(x), "", IIf(IsObject(x), TypeName(x), CStr(x)))
		size = IIf(IsNumber(size), size, Len(result))
		
		If Len(result) < size Then
			result = result & RepeatStr(ch, size - Len(result))
		End If

		PadLeft = result
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function StrDate(d)
		dim result

		If VarType(d) = vbDate Then
        	result = PadLeft(CStr(Year(d)), 4, "0") & _
					"/" & _
					PadLeft(CStr(Month(d)), 2, "0") & _
					"/" & _
					PadLeft(CStr(Day(d)), 2, "0") & _
					" " & _
					PadLeft(CStr(Hour(d)), 2, "0") & _
					":" & _
					PadLeft(CStr(Minute(d)), 2, "0") & _
					":" & _
					PadLeft(CStr(Second(d)), 2, "0")
		Else
			result = ""
		End If

		StrDate = result
    End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function read(name)
		Dim result
		
		result = CStr(Request.Querystring.Item(name))

		If Len(result) = 0 Then
			result = CStr(Request.Form.Item(name))
		End If
		
		read = result
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function SafeReplace(str, sWhat, sWith)
		Dim s
		
		s = Convert.ToString(str)

		If Len(s) > 0 Then
			s = Replace(s, sWhat, sWith)
		End If

		SafeReplace = s
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function prepareDate(d)
		Dim s

		s = Convert.ToString(d)

		If Len(s) = 0 Then
			s = "0000/00/00"
		ElseIf Len(s) > 10 Then
			s = Left(s,10)

			If s <> "0000/00/00" Then
				If Not IsDate(s) Then
					s = "0000/00/00"
				End If
			End If
		End If

		prepareDate = s
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function prepareTime(d)
		Dim s

		s = Trim(Convert.ToString(d))

		If Len(s) = 0 Or Len(s) > 5 Or Instr(s,":") = 0 Then
			s = "00:00"
		End If

		prepareTime = s
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Private Function Ceil(number)
		Ceil = -Sgn(number) * Int(-Abs(number))
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function selectOf(value, values, delimiter, default)
		If Instr(delimiter & values & delimiter, delimiter & value & delimiter) = 0 Then
			selectOf = default
		Else
			selectOf = value
		End If
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function JsEncode(x)
		JsEncode = Replace(x, """", """""")
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function XorCipher(str, pass)
		Dim i
		Dim RetStr
		Dim Charuse,CharPwd

		For i = 1 to Len(str)
			charuse = Mid(str,i,1)
			charpwd = Mid(pass,(i mod len(pass))+1,1)
			retstr = retstr + chr(asc(charuse) xor asc(charpwd))
		Next
		
		XorCipher = retstr
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Private Function Stream_StringToBinary(Text)
		Const adTypeText = 2
		Const adTypeBinary = 1
		Dim BinaryStream 'As New Stream
		Set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = adTypeText
		BinaryStream.CharSet = "us-ascii"
		BinaryStream.Open
		BinaryStream.WriteText Text
		BinaryStream.Position = 0
		BinaryStream.Type = adTypeBinary
		BinaryStream.Position = 0
		Stream_StringToBinary = BinaryStream.Read
		Set BinaryStream = Nothing
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Private Function Stream_BinaryToString(Binary)
		Const adTypeText = 2
		Const adTypeBinary = 1
		Dim BinaryStream 'As New Stream
		Set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = adTypeBinary
		BinaryStream.Open
		BinaryStream.Write Binary
		BinaryStream.Position = 0
		BinaryStream.Type = adTypeText
		BinaryStream.CharSet = "us-ascii"
		Stream_BinaryToString = BinaryStream.ReadText
		Set BinaryStream = Nothing
	End Function
	Function Base64Encode(sText)
		Dim oXML, oNode

		Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
		Set oNode = oXML.CreateElement("base64")

		oNode.dataType = "bin.base64"
		oNode.nodeTypedValue = Stream_StringToBinary(sText)
		Base64Encode = oNode.text
		Set oNode = Nothing
		Set oXML = Nothing
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function Base64Decode(ByVal vCode)
		Dim oXML, oNode
		Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
		Set oNode = oXML.CreateElement("base64")
		oNode.dataType = "bin.base64"
		oNode.text = vCode
		Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
		Set oNode = Nothing
		Set oXML = Nothing
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function RepeatStr(s, count)
		dim result(), i
		redim result(count - 1)

		For i = 0 To count - 1
			result(i) = s
		Next

		RepeatStr = Join(result, "")
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function SplitStr(str, separator)
		If IsNullOrEmpty(separator) Then
			dim result(), i

			redim result(len(str) - 1)

			for i = 0 to len(str) - 1
				result(i) = Mid(str, i + 1, 1)
			next

			SplitStr = result
		Else
			SplitStr = Split(str, separator)
		End If
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Function AscSplit(str, separator)
		dim result, i

		If IsNullOrEmpty(separator) Then
			redim result(len(str) - 1)

			for i = 0 to len(str) - 1
				result(i) = Mid(str, i + 1, 1)
			next
		Else
			result = Split(str, separator)
		End If

		for i = lbound(result) to ubound(result)
			result(i) = asc(result(i))
		next

		AscSplit = result
	End Function
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
%>