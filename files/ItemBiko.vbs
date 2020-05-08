Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objItemBiko
	Set objItemBiko = New ItemBiko
	objItemBiko.Run
	Set objItemBiko = nothing
End Function
'-----------------------------------------------------------------------
'ItemBiko
'-----------------------------------------------------------------------
Class ItemBiko
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "ItemBiko.vbs [option]"
		Echo "Ex."
		Echo "cscript//nologo ItemBiko.vbs /db:newsdc4 /J:R /T:R101"
		Echo "Option."
		Echo "   DBName:" & strDBName
		Echo "   JGyobu:" & strJGyobu
		Echo "    Tanto:" & strCsTanto
	End Sub
	Private	strDBName
	Private	strJGyobu
	Private	strCsTanto
	Private	objDB
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		strJGyobu = ""
		strCsTanto = ""
		set objDB = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		if Init() = True then
			OpenDb
			List
			CloseDb
		end if
	End Function
	'-----------------------------------------------------------------------
	'List() リスト
	'-----------------------------------------------------------------------
    Private Function List()
		Debug ".List()"
		AddSql ""
		AddSql "select"
		AddSql "*"
		AddSql "from item"
		AddSql "where (CS_TANTO_CD <> '' or BIKOU20 <> '')"
		AddWhere "JGYOBU",strJGyobu
		AddWhere "CS_TANTO_CD",strCsTanto
		CallSql strSql
		do while objRs.Eof = False
			Line
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line() 1行
	'-----------------------------------------------------------------------
    Private Function Line()
		Debug ".Line()"
		Write objRs.Fields("JGyobu")		,2
		Write objRs.Fields("NAIGAI")		,2
		Write objRs.Fields("HIN_GAI")		,20
		Write RTrim(objRs.Fields("CS_TANTO_CD"))	,8
		Write RTrim(objRs.Fields("BIKOU20"))		,20
		ItemUpdate
	End Function
	'-----------------------------------------------------------------------
	'GetTantoNm() 担当者名
	'-----------------------------------------------------------------------
    Private Function GetTantoNm(byVal strTanto)
		GetTantoNm = ""
		select case strTanto
		case "R101"
				GetTantoNm = "砂畠"
		case "R102"
				GetTantoNm = "今井"
		case "R103"
				GetTantoNm = "遠山"
		case "R104"
				GetTantoNm = "今井"
		case "R105"
				GetTantoNm = "川村"
		case "R106"
				GetTantoNm = "砂畠"
		end select
	End Function
	'-----------------------------------------------------------------------
	'GetTanto() 担当者コード
	'-----------------------------------------------------------------------
    Private Function GetTanto(byVal strTantoNm)
		GetTanto = ""
		if inStr(strTantoNm,"砂畠") > 0 then
			GetTanto = "砂畠"
		elseif inStr(strTantoNm,"砂旗") > 0 then
			GetTanto = "砂旗"
		elseif inStr(strTantoNm,"砂") > 0 then
			GetTanto = "砂"
		elseif inStr(strTantoNm,"白田") > 0 then
			GetTanto = "白田"
		elseif inStr(strTantoNm,"矢藤") > 0 then
			GetTanto = "矢藤"
		elseif inStr(strTantoNm,"今井") > 0 then
			GetTanto = "今井"
		elseif inStr(strTantoNm,"遠山") > 0 then
			GetTanto = "遠山"
		elseif inStr(strTantoNm,"川村") > 0 then
			GetTanto = "川村"
		elseif inStr(strTantoNm,"田中") > 0 then
			GetTanto = "田中"
		elseif inStr(strTantoNm,"千葉") > 0 then
			GetTanto = "千葉"
		elseif inStr(strTantoNm,"坂上") > 0 then
			GetTanto = "坂上"
		end if
	End Function
	'-----------------------------------------------------------------------
	'ItemUpdate() Item更新
	'-----------------------------------------------------------------------
    Private Function ItemUpdate()
		Debug ".ItemUpdate()"
		dim	strTantoNm
		strTantoNm = GetTantoNm(RTrim(objRs.Fields("CS_TANTO_CD")))
		if strTantoNm = "" then
			exit function
		end if
		dim	strBiko20
		strBiko20 = RTrim(objRs.Fields("BIKOU20"))
		if strBiko20 = strTantoNm then
			exit function
		end if
		if strBiko20 = "" then
			strBiko20 = strTantoNm
		else
			dim	strTantoNmOld
			strTantoNmOld = GetTanto(strBiko20)
			if strTantoNmOld = "" then
				exit function
			end if
			strBiko20 = Replace(strBiko20,strTantoNmOld,strTantoNm)
			if strBiko20 = RTrim(objRs.Fields("BIKOU20")) then
				exit function
			end if
		end if
		Write "→",0
		Write strBiko20,0
'		exit function
		AddSql ""
		AddSql "update Item"
		AddSql "set BIKOU20 = '" & strBiko20 & "'"
		AddSql ",UPD_TANTO='ItemB'"
		AddSql ",UPD_DATETIME = left(replace(replace(replace(convert(CURRENT_TIMESTAMP(),sql_char),'-',''),':',''),' ',''),14)"
		AddSql "where JGyobu = '" & RTrim(objRs.Fields("JGyobu")) & "'"
		AddSql "  and NAIGAI = '" & RTrim(objRs.Fields("NAIGAI")) & "'"
		AddSql "  and HIN_GAI = '" & RTrim(objRs.Fields("HIN_GAI")) & "'"
		Write ":" & Execute(strSql) ,0
	End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private Function Execute(byVal strSql)
		Debug ".Execute():" & strSql
		on error resume next
		Call objDb.Execute(strSql)
		Execute = Err.Number
		select case Execute
		case 0
		case -2147467259	'0x80004005 重複キー
		case else
			Wscript.StdErr.WriteLine
			Wscript.StdErr.WriteLine Err.Description
			Wscript.StdErr.WriteLine strSql
		end select
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private	objRs
	Private Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		if Err.Number <> 0 then
			Wscript.StdOut.WriteLine "0x" & Hex(Err.Number) & ":" & Err.Description
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end if
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'Where strSql
	'-------------------------------------------------------------------
	Private	Function AddWhere(byVal strF,byVal strV)
		if strV = "" then
			exit function
		end if
		if inStr(strSql,"where") > 0 then
			AddSql " and "
		else
			AddSql " where "
		end if
		AddSql strF
		if inStr(strV,"%") > 0 then
			AddSql " like "
		else
			AddSql " = "
		end if
		AddSql " '" & strV & "'"
	End Function
	'-------------------------------------------------------------------
	'文字列追加 strSql
	'-------------------------------------------------------------------
	Private	strSql
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
	Private Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
	'-----------------------------------------------------------------------
	Private Function GetOption(byval strName ,byval strDefault)
		dim	strValue

		if strName = "" then
			strValue = ""
			if WScript.Arguments.Named.Exists(strDefault) then
				strValue = strDefault
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
	'-----------------------------------------------------------------------
	'WriteLine
	'-----------------------------------------------------------------------
	Private Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine s
	End Sub
	'-----------------------------------------------------------------------
	'Write
	'-----------------------------------------------------------------------
	Private Sub Write(byVal s,byVal i)
		if i > 0 then
			s = left(RTrim(s) & space(i),i)
		elseif i < 0 then
			s = right(space(-i) & LTrim(s),-i)
		end if
		Wscript.StdOut.Write "" & s
	End Sub
	'-----------------------------------------------------------------------
	'Echo
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
		Init = False
		For Each strArg In WScript.Arguments.UnNamed
			Echo "オプションエラー:" & strArg
			Usage
			Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "j","jgyobu"
				strJGyobu = GetOption(strArg,"")
			case "t","tanto"
				strCsTanto = GetOption(strArg,"")
			case "?"
				Usage
				Exit Function
			case else
				Echo "オプションエラー:" & strArg
				Usage
				Exit Function
			end select
		Next
		Init = True
	End Function
End Class
