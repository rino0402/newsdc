'-----------------------------------------------------------------------
'デバッグメッセージ
'-----------------------------------------------------------------------
Function Debug(byval strMsg)
	if WScript.Arguments.Named.Exists("debug") then
		Wscript.Echo strMsg
	end if
End Function
Function isDebug()
	isDebug = WScript.Arguments.Named.Exists("debug")
End Function
'-----------------------------------------------------------------------
'Err表示
'-----------------------------------------------------------------------
Private Function DispErr(objErr)
	dim	strMsg
	dim	intErrNumber
	intErrNumber = objErr.Number
	if intErrNumber <> 0 then
		Call DispMsg("Error.Number:0x" & Hex(intErrNumber))
		Call DispMsg("Error.Description:" & objErr.Description)
	end if
	DispErr = intErrNumber
End Function
'-----------------------------------------------------------------------
'(Debug用)全Propertiesを文字列で取得
'-----------------------------------------------------------------------
Function GetProperties(obj)
	dim	strP
	strP = ""
	dim	p
	for each p in (obj.Properties)
		if strP <> "" Then
			strP = strP & vbCrLf
		end if
		strP = strP & " " & p.Name & ":" & p
	next
	GetProperties = strP
End Function
'-----------------------------------------------------------------------
'select Where条件
'-----------------------------------------------------------------------
Function makeWhere(byval strWhere _
				  ,byval strField _
				  ,byval strValue1 _
				  ,byval strValue2 _
				  )
	dim	strAnd
	dim	strNot
	dim	strCmp
	
	if len(strValue1) > 0 then
		if inStr(strWhere,"where") > 0 then
			strAnd = " and "
		else
			strAnd = " where "
		end if
		if len(strValue2) > 0 then
			strCmp = "between"
			strWhere = strWhere & strAnd & " " & strField & " " & strCmp & " '" & strValue1 & "' and '" & strValue2 & "'"
		else
			select case left(strValue1,1)
			case "<"
				strValue1 = right(strValue1,len(strValue1)-1)
				strCmp = "<"
			case else
				strValue1 = "'" & strValue1 & "'"
				if instr(1,strValue1,"%") > 0 then
					strCmp = strNot & "like"
				elseif instr(strValue1,",") > 0 then
					strCmp = strNot & "in "
					strValue1 = "(" & replace(strValue1,",","','") & ")"
				else
					if strNot = "" Then
						strCmp = "="
					else
						strCmp = "<>"
					end if
				end if
			end select
			strWhere = strWhere & strAnd & " " & strField & " " & strCmp & " " & strValue1 & ""
		end if
	end if
	makeWhere = strWhere
End Function
