Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Call Main()
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objTable
	Set objTable = New Table
	objTable.Run
	Set objTable = Nothing
End Function
'-----------------------------------------------------------------------
'Table
'2017.05.17 新規
'-----------------------------------------------------------------------
Class Table
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "Duct200.vbs [option]"
		Echo "/db:newsdc5"
		Echo "/debug"
		Echo "Ex."
		Echo "cscript//nologo Duct200.vbs /db:newsdc4"
		Echo ""
		Echo "strDBName=" & strDBName
		Echo "    strDt=" & strDt
	End Sub
	'-----------------------------------------------------------------------
	'Private 変数
	'-----------------------------------------------------------------------
	Private	strDBName
	Private	objDB
	Private	strDt
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName		= GetOption("db","newsdc5")
		set objDB		= nothing
		strDt			= Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2)
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB		= nothing
    End Sub
	'-----------------------------------------------------------------------
	'Quit() 強制終了
	'-----------------------------------------------------------------------
	Private Function Quit()
		Debug ".Quit()"
		Wscript.Quit
	End Function
	'-----------------------------------------------------------------------
	'Echo()
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		Wscript.Echo s
	End Sub
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
	Private Function Init()
		Debug ".Init()"
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			strDt = strArg
			exit for
			Echo "Error:オプション:" & strArg
			Disp Init
			Usage
			Quit
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Echo "Error:オプション:" & strArg
				Usage
				Quit
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Init
		OpenDb
		Load
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load()"
		WriteLine strDt
		AddSql ""
		AddSql "select * from y_syuka_tei"
		AddSql "where SND_YMD = '" & strDt & "'"
		AddSql "  and HINB_CD like 'AD-5%'"
		AddSql "order by"
		AddSql " SND_YMD"
		AddSql ",SND_HMS"
		AddSql ",TEI_LABELID"
		AddSql ",L_UCHI_NO"
		CallSql strSql
		dim	strId
		dim	strPn0
		dim	strPn1
		strId	= ""
		strPn0	= ""
		strPn1	= ""
		dim	prvId
		dim	aryMsg(10)
		dim	i
		for i = 1 to 10
			aryMsg(i) = ""
		next
		dim	strCall
		strCall = ""
		do while True
			strId = GetField("TEI_LABELID")
			if prvId <> strId then
				WriteLine strCall
				if strCall <> "" then
					
				end if
				prvId = strId
				strPn0	= ""
				strCall = ""
			end if
			if objRs.Eof = True then
				exit do
			end if

			strPn1 = GetField("HINB_CD")
			dim	strAdd
			strAdd = ""
			if Left(strPn1,9) = "AD-5008SH" then
				if strPn0 = "" then
					strPn0 = strPn1
				elseif strPn0 < strPn1 then
					'分割
					strAdd = " 分離"
					strCall = "分離処理実行"
				elseif strPn0 > strPn1 then
					'分割
					strAdd = strPn0 & " 分離"
					strPn0 = strPn1
					strCall = "分離処理実行"
				end if
			end if
			Write T(GetField("SND_YMD"),-9)
			Write T(GetField("SND_HMS"),-7)
			Write T(GetField("SEQ_NO"),-6)
'			Write T(GetField("L_SERIES1"),-21)
			Write T(GetField("L_SERIES2"),-21)
			Write T(GetField("TEI_LABELID"),-10)
			Write T(GetField("L_UCHI_NO"),2) & " "
			Write T(GetField("HINB_CD"),-15)
			Write strAdd
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-------------------------------------------------------------------
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		GetField = ""
		if objRs.Eof = True then
			exit function
		end if
		dim	strField
		on error resume next
		strField = RTrim("" & objRs.Fields(strName))
		if Err.Number <> 0 then
			WScript.Echo "GetField(" & strName & "):Error:0x" & Hex(Err.Number)
			WScript.Quit
		end if
		on error goto 0
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
	End Function
	'-----------------------------------------------------------------------
	'T() 文字列
	'-----------------------------------------------------------------------
	Private Function T(byVal v,byVal i)
		if i > 0 then
			T = right(space(i) & v,i)
		else
			i = i * -1
			T = LeftB(v & space(i),i)
		end if
	End Function
	'-----------------------------------------------------------------------
	'LeftB() 文字列
	'-----------------------------------------------------------------------
	Private Function LeftB(byVal a_Str, byVal a_int)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			LeftB = ""
			Exit Function
		End If
		If a_int = 0 Then
			LeftB = ""
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
		if LenB(iLeftStr) < a_int then
			iLeftStr = iLeftStr & space(a_int - LenB(iLeftStr))
		end if
		LeftB = iLeftStr
	End Function
	'-----------------------------------------------------------------------
	'LenB() 文字列
	'-----------------------------------------------------------------------
	Function LenB(byVal a_Str)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			LenB = 0
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
		LenB = iLenCount
	End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private	objRs
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		select case Err.Number
		case -2147467259	'重複
		case 0,500
		case else
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine Err.Number & "(0x" & Hex(Err.Number) & "):" & Err.Description
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end select
		on error goto 0
'		on error resume next
'		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
		objDB.Open strDbName
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		objDB.Close
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'文字列追加 strSql
	'-------------------------------------------------------------------
	dim	strSql
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
	End Function
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-------------------------------------------------------------------
	'Write
	'-------------------------------------------------------------------
	Private	Sub Write(byVal s)
		Wscript.StdOut.Write s
	End Sub
	'-------------------------------------------------------------------
	'WriteLine
	'-------------------------------------------------------------------
	Private	Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine s
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
	'-----------------------------------------------------------------------
	Function GetOption(byval strName _
					  ,byval strDefault _
					  )
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
