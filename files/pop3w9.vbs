Option Explicit
'
'cscript pop3w9.vbs
' /debug
' /save : メールを残す
' /savd : メールを削除(default)
'
dim	objPop3
Set objPop3 = New Pop3
Set objPop3 = Nothing
Class Pop3
	Private	objBasp
	Private Sub Debug(strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	Private Sub Disp(strMsg)
		Wscript.Echo strMsg
	End Sub
	Private	strSvName
	Private	strUser
	Private	strPass
	Private	strCommand
	Private	strDirname
    Private Sub Class_Initialize
		Debug ".Class_Initialize"
		strSvName	= GetOption("s","ns")
		strUser		= GetOption("u","w9")
		strPass		= GetOption("p","123daaa")
		strDirname	= "pop3w9\new"
		strCommand	= "SAVD 1"	' メールを削除
		if WScript.Arguments.Named.Exists("save") then
			strCommand	= "SAVE 1"	' メールを残す
		end if
		if WScript.Arguments.Named.Exists("savd") then
			strCommand	= "SAVD 1"	' メールを削除
		end if
		Set objBasp = CreateObject("Basp21")
		Call RcvMail()
    End Sub
    Private Sub Class_Terminate
		Debug ".Class_Terminate"
		Set objBasp = Nothing
    End Sub
	Private	Function isTrue(byVal b)
		Debug ".isTrue:" & b
		isTrue = False
		if b then
			isTrue = True
		end if
	End Function
	Private Sub RcvMail
		Debug ".RcvMail"
		Disp "受信中：" & strSvName & "," & strUser & "," & strPass & "," & strCommand & "," & strDirname
		dim	aryRcv
		aryRcv = objBasp.RcvMail(strSvName,strUser,strPass,strCommand,strDirname)
		if isTrue(IsArray(aryRcv)) = False then
			Disp "受信メールなし：" & aryRcv
			Exit Sub
		end if
		Debug ".RcvMail:IsArray"
		dim	strRcv
		for each strRcv in aryRcv
			dim	strSubject
			dim	strFrom
			dim	aryFile()
			redim	aryFile(0)
			strSubject	= ""
			strFrom		= ""
			objBasp.CodePage = 65001	'utf-8
			dim	aryMail
			aryMail = objBasp.ReadMail(strRcv,"subject:from:date:",strDirname)
			if IsArray(aryMail) then	' OK ?
				dim	strMail
				for each strMail in aryMail
					'1行目を表示
'					Disp Split(data,vbCrLf)(0)
					Disp strRcv & ">" & strMail
					dim	strHead
					strHead = Split(strMail,":")(0)
					dim	strBody
					strBody = Split(strMail,":")(1)
					select case strHead
					case "Subject"
'									strSubject = Right(data,Len(data) - 9)
									strSubject = strBody
					case "From"
'									strFrom	= lcase(Right(data,Len(data) - 6))
									strFrom	= lcase(strBody)
					case "Body"
					case "File"
									if UBound(aryFile) > 0 then
										ReDim Preserve aryFile(UBound(aryFile) + 1)
									end if
'									aryFile(UBound(aryFile)) = Right(data,Len(data) - 6)
									aryFile(UBound(aryFile)) = strBody
					end select
				next
			end if
			dim	strFile
			dim	i
			i = 0
			strBody = "" & vbCrLf
			for each strFile in aryFile
				i = i + 1
				Debug "File(" & i & ")" & strFile
				strBody = strBody & strFile & vbCrLf
			next
			strBody = strBody & "" & vbCrLf
			strBody = strBody & "添付ファイルを" & i & "件 受信しました。" & vbCrLf
			strBody = strBody & "変換処理を開始します。" & vbCrLf
			strBody = strBody & "終了後に再度メールを送信します。" & vbCrLf
			strSubject = "変換開始 Re:" & strSubject
			'メールを返信
			dim	strMailTo
			strMailTo = strFrom
			Disp "返信中：" & strSvName & "," & strMailTo & "," & strSubject & "," & strBody
			dim	strMsg
			objBasp.CodePage = 932	'JIS
			strMsg = objBasp.SendMail(strSvName,strMailTo,"pop3w9", strSubject,strBody,"")
			Disp strMsg
		next
    End Sub
	'-----------------------------------------------------------------------
	'オプション取得
	'-----------------------------------------------------------------------
	Private Function GetOption(byval strName ,byval strDefault)
		dim	strValue

		if strName = "" then
			strValue = ""
			if strDefault < WScript.Arguments.UnNamed.Count then
				strValue = WScript.Arguments.UnNamed(strDefault)
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
End Class
