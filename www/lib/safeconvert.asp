<%
    Function SafeCSng(sngNumber)
		Dim sngResult
		On Error Resume Next

		If IsNullOrEmpty(sngNumber) Or Not IsNumeric(sngNumber) Then
			sngResult = 0
		Else
			If Err.Number <> 0 Then
				sngResult = 0
			Else
				sngResult = CSng(sngNumber)

				If Err.Number <> 0 Then
					If Err.Number = 6 Then	' overflow
						If Left(Trim(CStr(sngNumber)),1) = "-" Then
							sngResult = -3.402823 * (10 ^ 38)
						Else
							sngResult =  3.402823 * (10 ^ 38)
						End If
					Else
						sngResult = 0
					End If
				End If
			End If
		End If

		On Error Goto 0
		
		SafeCSng = sngResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCDbl(dblNumber)
		Dim dblResult
		On Error Resume Next

		If IsNullOrEmpty(dblNumber) Or Not IsNumeric(dblNumber) Then
			dblResult = 0
		Else
			If Err.Number <> 0 Then
				dblResult = 0
			Else
				dblResult = CDbl(dblNumber)

				If Err.Number <> 0 Then
					If Err.Number = 6 Then	' overflow
						If Left(Trim(CStr(dblNumber)),1) = "-" Then
							dblResult = -1.79769313486232 * (10 ^ 308)
						Else
							dblResult =  1.79769313486232 * (10 ^ 308)
						End If
					Else
						dblResult = 0
					End If
				End If
			End If
		End If
		
		On Error Goto 0

		SafeCDbl = dblResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCLng(lNumber)
		Dim lResult
		On Error Resume Next

		If IsNullOrEmpty(lNumber) Or Not IsNumeric(lNumber) Then
			lResult = 0
		Else
			If Err.Number <> 0 Then
				lResult = 0
			Else
				lResult = CLng(lNumber)

				If Err.Number <> 0 Then
					If Err.Number = 6 Then	' overflow
						If Left(Trim(CStr(lNumber)),1) = "-" Then
							lResult = -2147483648
						Else
							lResult = 2147483647
						End If
					Else
						lResult = 0
					End If
				End If
			End If
		End If

		On Error Goto 0
		
		SafeCLng = lResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCInt(iNumber)
		Dim iResult
		On Error Resume Next

		If IsNullOrEmpty(iNumber) Or Not IsNumeric(iNumber) Then
			iResult = 0
		Else
			If Err.Number <> 0 Then
				iResult = 0
			Else
				iResult = CInt(iNumber)

				If Err.Number <> 0 Then
					If Err.Number = 6 Then	' overflow
						If Left(Trim(CStr(iNumber)),1) = "-" Then
							iResult = -32768
						Else
							iResult = 32767
						End If
					Else
						iResult = 0
					End If
				End If
			End If
		End If

		On Error Goto 0
		
		SafeCInt = iResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCByte(btNumber)
		Dim btResult
		Dim sNumber
		
		On Error Resume Next

		If IsNullOrEmpty(btNumber) Or Not IsNumeric(btNumber) Then
			btResult = 0
		Else
			If Err.Number <> 0 Then
				btResult = 0
				Err.Clear
			Else
				btResult = CByte(btNumber)

				If Err.Number <> 0 Then
					If Err.Number = 6 Then	' overflow
						sNumber = CStr(btNumber)
						If Left(Trim(sNumber),1) = "-" Then
							btResult = 0
						Else
							btResult = 255
						End If
					Else
						btResult = 0
					End If

					Err.Clear
				End If
			End If
		End If

		On Error Goto 0
		
		SafeCByte = btResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCStr(sValue)
		Dim sResult
		
		On Error Resume Next
		
		If IsVoid(sValue) Then
			sResult = ""
		Else
			sResult = CStr(sValue)

			If Err.Number <> 0 Then
				sResult = ""
			End If
		End If

		On Error Goto 0
		
		SafeCStr = sResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCBool(bValue)
		Dim bResult
		
		If IsVoid(bValue) Then
			bResult = False
		ElseIf IsNumber(bValue) Then
			bResult = bValue <> 0
		ElseIf LCase(CStr(bValue)) <> "true" And LCase(CStr(bValue)) <> "false" Then
			bResult = False
		Else
			bResult = CBool(bValue)
		End If
		
		SafeCBool = bResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCDate(dValue)
		Dim dResult
		On Error Resume Next
		
		dResult = CDate(dValue)

		If Err.Number <> 0 Then
			dResult = NULL
			Err.Clear
		End If
		
		On Error Goto 0

		SafeCDate = dResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SafeCCur(curValue)
		Dim curResult
		On Error Resume Next
		
		curResult = CDate(curValue)
		If Err.Number <> 0 Then
			curResult = NULL
		End If
		
		SafeCDate = curResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SCLng(iNumber,initValue)
		Dim iResult
		On Error Resume Next
		
		If IsNull(iNumber) Or IsEmpty(iNumber) Or (Not IsNumeric(iNumber)) Then
			iResult = initValue
		Else
			iResult = SafeCLng(iNumber)
		End If
		SCLng = iResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SCInt(iNumber,initValue)
		Dim iResult
		On Error Resume Next
		
		If IsNull(iNumber) Or IsEmpty(iNumber) Or (Not IsNumeric(iNumber)) Then
			iResult = initValue
		Else
			iResult = SafeCInt(iNumber)
		End If
		SCInt = iResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SCStr(s, initValue)
		Dim sResult
		On Error Resume Next
		
		If IsNull(s) Or IsEmpty(s) Then
			sResult = initValue
		Else
			sResult = SafeCStr(s)
		End If
		SCStr = sResult
	End Function
	' --------------------------------------------------------------------------------------------
	Function SCBool(b,initValue)
		Dim bResult
		Dim s
		On Error Resume Next
		
		s = LCase(Trim(CStr(b)))
		If Err.Number <> 0 Then
			b = NULL
		End If
		If IsNull(s) Or IsEmpty(s) Or (s <> "true" And s <> "false") Then
			bResult = initValue
		Else
			bResult = SafeCBool(b)
		End If
		SCBool = bResult
	End Function
%>