Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objP_SAGYO_LOG
	Set objP_SAGYO_LOG = New P_SAGYO_LOG
	objP_SAGYO_LOG.Run
	Set objP_SAGYO_LOG = nothing
End Function
'-----------------------------------------------------------------------
'P_SAGYO_LOG
'-----------------------------------------------------------------------
Class P_SAGYO_LOG
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "p_sagyo_log.vbs [option]"
		Echo "Ex."
		Echo "cscript//nologo p_sagyo_log.vbs /db:newsdc4"
		Echo "Option."
		Echo "   DBName:" & strDBName
		Echo "     List:" & strList
		Echo "      Top:" & strTop
		Echo "       Dt:" & strDt
	End Sub
	Private	objDB
	Private	strDBName
	Private	strList
	Private	strTop
	Private	strDt
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName	= GetOption("db"	,"newsdc")
		strList		= GetOption("list"	,"")
		strTop		= GetOption("top"	,"")
		strDt		= GetOption("dt"	,"")
		set objDB	= nothing
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
			case "debug"
			case "db"
				strDBName	= GetOption(strArg,strDBName)
			case "list"
				strList		= GetOption(strArg,strList)
			case "top"
				strTop		= GetOption(strArg,strDt)
			case "dt"
				strDt		= GetOption(strArg,strDt)
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
			select case strList
			case "1"
				List1
			case "0"
				List0
			case else
				List0
			end select
			CloseDb
		end if
	End Function
	'-----------------------------------------------------------------------
	'List0() 処理日別
	'-----------------------------------------------------------------------
    Private Function List0()
		Debug ".List0()"
		AddSql ""
		AddSql "select"
		AddTop
		AddSql " p.JITU_DT"
		AddSql ",count(*) cnt"
		AddSql "from p_sagyo_log p"
		AddWhere "p.JITU_DT",strDt
		AddSql "group by"
		AddSql " p.JITU_DT"
		AddSql "order by"
		AddSql " p.JITU_DT desc"
		CallSql strSql
		do while objRs.Eof = False
			Line0
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line0() 1行表示
	'-----------------------------------------------------------------------
    Private Function Line0()
		Debug ".Line0()"
		Write objRs.Fields("JITU_DT")	,0
		Write objRs.Fields("cnt")		,-5
	End Function
	'-----------------------------------------------------------------------
	'List1() 担当者別
	'-----------------------------------------------------------------------
    Private Function List1()
		Debug ".List1()"
		AddSql ""
		AddSql "select"
		AddTop
		AddSql " p.JITU_DT"
		AddSql ",case p.JGYOBU"
		AddSql " when '4' then '*'"
		AddSql " when '5' then '*'"
		AddSql " when 'D' then '*'"
		AddSql " else p.JGYOBU"
		AddSql " end JGYOBU"
		AddSql ",p.TANTO_CODE"
		AddSql ",ifnull(t.TANTO_NAME,'') TANTO_NAME"
		AddSql ",count(*) cnt"
		AddSql ",round(sum(convert(WORK_TM,sql_decimal))/60/60,2) wtm"
		AddSql ",min(p.JITU_TM) min_tm"
		AddSql ",max(p.JITU_TM) max_tm"
		AddSql ",min(p.WEL_ID) min_wel"
		AddSql ",max(p.WEL_ID) max_wel"
		AddSql "from p_sagyo_log p"
		AddSql "left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
		AddWhere "p.JITU_DT",strDt
		AddSql "group by"
		AddSql " p.JITU_DT"
		AddSql ",JGYOBU"
		AddSql ",p.TANTO_CODE"
		AddSql ",TANTO_NAME"
		AddSql "order by"
		AddSql " p.JITU_DT desc"
		AddSql ",JGYOBU"
		AddSql ",p.TANTO_CODE"
		CallSql strSql
		do while objRs.Eof = False
			Line1
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line1() 入荷リスト １行表示
	'-----------------------------------------------------------------------
    Private Function Line1()
		Debug ".Line1()"
		Debug ".Line0()"
		Write objRs.Fields("JITU_DT")	,9
		Write objRs.Fields("JGYOBU")	,2
		Write objRs.Fields("TANTO_CODE"),6
		Write objRs.Fields("TANTO_NAME"),14
		Write objRs.Fields("cnt")		,-5
		Write "",1
		Write objRs.Fields("wtm")		,-5
		Write "",1
		Write objRs.Fields("min_tm")	,4
		Write "",1
		Write objRs.Fields("max_tm")	,4
		Write "",1
		dim	strWel1
		strWel1 = RTrim(objRs.Fields("min_wel"))
		Write strWel1	,4
		dim	strWel2
		strWel2 = RTrim(objRs.Fields("max_wel"))
		if strWel1 = strWel2 then
			strWel2 = ""
		end if
		Write strWel2	,4
	End Function
	'-------------------------------------------------------------------
	'GroupHead() グループヘッダー
	'	True:グループヘッダー
	'  Flase:継続行
	'-------------------------------------------------------------------
	Private	curHead
	Private	newHead
	Private	Function GroupHead(byVal intHead)
		if intHead < 0 then
			curHead = ""
			exit function
		end if
		dim	i
		newHead = ""
		for i = 0 to intHead
			newHead = newHead + objRs.Fields(i)
		next
		if curHead = newHead then
			GroupHead = False
			exit function
		end if
		curHead = newHead
		GroupHead = True
	End Function
	'-----------------------------------------------------------------------
	'List2()
	'-----------------------------------------------------------------------
    Private Function List2()
		Debug ".List2()"
		AddSql ""
		AddSql "select"
		AddSql " h.Filename Filename"
		AddSql ",h.JGyobu JGyobu"
		AddSql ",h.DenDt"
		AddSql ",h.IoKbn"
		AddSql ",h.AkaKuro"
		AddSql ",h.SyukoCd"
		AddSql ",h.NyukoCd"
		AddSql ",h.SyushiCd"
		AddSql ",count(*) cnt"
		AddSql ",if(n.JGyobu is null,'','入荷 ' + n.JGyobu + ' ' + n.NYUKO_TANABAN) Nyuka"
		AddSql2 "from ",strTable & " h"
		AddSql "left outer join y_nyuka n"
		AddSql " on (h.JGyobu = n.JGyobu"
		AddSql " and h.DenDt = n.SYUKA_YMD"
		AddSql " and (h.SyoriMD + h.Bin + h.SeqNo) = n.Text_No"
		AddSql "	)"
		AddWhere "h.Filename",strFileName
		AddWhere "h.DenDt",strDt
		AddSql "group by"
		AddSql " Filename"
		AddSql ",h.JGyobu"
		AddSql ",h.DenDt"
		AddSql ",h.IoKbn"
		AddSql ",h.AkaKuro"
		AddSql ",h.SyukoCd"
		AddSql ",h.NyukoCd"
		AddSql ",h.SyushiCd"
		AddSql ",Nyuka"
		AddSql "order by"
		AddSql " Filename"
		AddSql ",h.JGyobu"
		AddSql ",h.DenDt"
		AddSql ",h.IoKbn"
		AddSql ",h.SyukoCd"
		AddSql ",h.NyukoCd"
		AddSql ",h.SyushiCd"
		AddSql ",h.AkaKuro"
		CallSql strSql
		do while objRs.Eof = False
			Line2
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line2() 1行表示
	'-----------------------------------------------------------------------
    Private Function Line2()
		Debug ".Line2()"
		Write objRs.Fields("JGyobu")	,2
		Write objRs.Fields("DenDt")		,9
		Write objRs.Fields("IoKbn")		,1
		Write objRs.Fields("AkaKuro")	,2
		Write objRs.Fields("SyukoCd")	,6
		Write objRs.Fields("NyukoCd")	,6
		Write objRs.Fields("SyushiCd")	,3
		Write objRs.Fields("cnt")		,-5
		Write " " & objRs.Fields("Nyuka")		,0
'		Write "" & objRs.Fields("NYUKO_TANABAN")		,0
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
	'-----------------------------------------------------------------------
	'SetSql() 
	'-----------------------------------------------------------------------
    Private Function SetSql(byVal strSet,byVal strTitle,byVal strName,byVal strSrc,byVal strDst)
		Debug ".SetSql()"
		Write strTitle,0
		Write strSrc,0
		if strDst <> "" then
			select case strName
			case "KANKYO_KBN_SURYO"
				if strDst = "0" then
					strDst = strSrc
				end if
			case "INSP_MESSAGE"
				if strSrc = "単価改訂 リチウム電池搭載" then
					strDst = strSrc
				end if
			end select
			if strDst <> strSrc then
				Write "→",0
				Write strDst,0
				if strSet = "" then
					strSet = " set "
				else
					strSet = strSet & " ,"
				end if
				strSet = strSet & strName & " = '" & strDst & "'"
			end if
		end if
		SetSql = strSet
	End Function
	'-------------------------------------------------------------------
	'AddSql2
	'-------------------------------------------------------------------
	Private	Function AddSql2(byVal str1,byVal str2)
		if Right(str1,1) = "'" then
			'Char
			str2 = Replace(RTrim(str2),"'","''") & "'"
		end if
		AddSql str1 & str2
	End Function
	'-------------------------------------------------------------------
	'top strTop
	'-------------------------------------------------------------------
	Private	Function AddTop()
		if strTop = "" then
			exit function
		end if
		AddSql " top " & strTop
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
		dim	strCmp
		strCmp = "="
		if left(strV,1) = "-" then
			strV = Right(strV,len(strV)-1)
			strCmp = "<>"
		end if
		if inStr(strV,"%") > 0 then
			if strCmp = "=" then
				strCmp = " like "
			else
				strCmp = " not like "
			end if
		end if
		AddSql strF & " " & strCmp & " '" & strV & "'"
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
			s = LeftB(RTrim(s) & space(i),i)
		elseif i < 0 then
			s = right(space(-i) & LTrim(s),-i)
		end if
		Wscript.StdOut.Write s
	End Sub
	'-------------------------------------------------------------------
	'LeftB()
	'-------------------------------------------------------------------
	Private Function LeftB(byVal a_Str,byVal a_int)
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
		LeftB = iLeftStr
	End Function
	'-----------------------------------------------------------------------
	'Echo
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		Wscript.Echo s
	End Sub
End Class
