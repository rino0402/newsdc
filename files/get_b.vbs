' -------------------------------------------------------------
' get_b.vbs
' -------------------------------------------------------------
Function Get_LeftB(a_Str, a_int)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
		If iLenCount > Cint(a_int) Then
			Exit For
		Else
			iLeftStr = iLeftStr + Mid(a_Str, iCount, 1)
		End If
	Next
	Get_LeftB = iLeftStr
End Function

Function Get_MidB(a_Str,s_int, a_int)
	Dim iCount, iAscCode, iLenCount, iMidStr
	iLenCount = 0
	iMidStr = ""
	If Len(a_Str) = 0 Then
		Get_MidB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_MidB = ""
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
		if iLenCount >= s_int then
			If iLenCount > Cint(s_int) + Cint(a_int) - 1 Then
				Exit For
			Else
				iMidStr = iMidStr + Mid(a_Str, iCount, 1)
			End If
		end if
	Next
	Get_MidB = iMidStr
End Function

Function Get_LenB(a_Str)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LenB = 0
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
	Next
	Get_LenB = iLenCount
End Function
