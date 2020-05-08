Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objHMEM700
	Set objHMEM700 = New HMEM700
	objHMEM700.Run
	Set objHMEM700 = nothing
End Function
'-----------------------------------------------------------------------
'HMEM700
'-----------------------------------------------------------------------
Class HMEM700
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "HMEM700.vbs [option] [filename]"
		Echo "Ex."
		Echo "cscript//nologo HMEM700.vbs /db:newsdc1 hmem704szz.dat.20170807-101929.29213"
		Echo "cscript//nologo HMEM700.vbs /db:newsdc4"
		Echo "cscript//nologo HMEM700.vbs /db:newsdc4 /list:0"
		Echo "cscript//nologo HMEM700.vbs /db:newsdc4 /list:1"
	End Sub
	Private	optDebug
	Private	objDB
	Private	strDBName
	Private	strFileName
	Private	strDt
	Private	optList
	Private	optSaki
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		optDebug	= False
		set objDB	= nothing
		strDBName	= GetOption("db","newsdc")
		strFileName = ""
		strDt		= GetOption("dt","")
		optList		= GetOption("list","1")
		optSaki		= GetOption("saki","")
	End Sub
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Private Function Init()
		Debug ".Init()"
		dim	strArg
		Init = False
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Echo "オプションエラー:" & strArg
				Usage
				Exit Function
			end if
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
				strDBName	= GetOption(strArg,strDBName)
			case "dt"
				strDt		= GetOption(strArg,strDt)
			case "debug"
				optDebug	= True
			case "list"
				optList		= GetOption(strArg,optList)
			case "saki"
				optSaki		= GetOption(strArg,optSaki)
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
			Look
			CloseDb
		end if
	End Function
	'-----------------------------------------------------------------------
	'Look() 読込
	'-----------------------------------------------------------------------
    Private Function Look()
		Debug ".Look()"
		AddSql ""
		select case optList
		case "0"
			AddSql "select"
			AddSql " Filename"
			AddSql ",count(*) cnt"
			AddSql "from HMEM700"
			AddWhere "Filename",strFileName
			AddWhere "SyukaDt",strDt
			AddWhere "AiteCode",optSaki
			AddSql "group by"
			AddSql " Filename"
			AddSql "order by"
			AddSql " Filename"
		case "1"
			AddSql "select"
			AddSql " h.JCode"
			AddSql ",h.UkeKbn"
			AddSql ",h.IOKbn"
			AddSql ",h.SyukaDt"
			AddSql ",h.Syushi"
			AddSql ",h.ChuKbn"
			AddSql ",h.ChuName"
			AddSql ",h.AiteCode"
			AddSql ",h.AiteName"
			AddSql ",count(*) cnt"
			AddSql ",if(y.JGYOBU is null,if(n.JGYOBU is null,'','入荷 ' + n.JGYOBU + ' ' + n.NYUKO_TANABAN ),'出荷 ' + y.JGYOBU) NyuSyuka"
			AddSql "from HMEM700 h"
			AddSql "left outer join y_syuka y"
			AddSql "on (h.IDNo = y.KEY_ID_NO)"
			AddSql "left outer join y_nyuka n"
			AddSql " on (h.SyukaDt = n.SYUKA_YMD and right(h.IDNo,9) = n.TEXT_NO and h.IDNo = n.ID_NO2)"
			AddWhere "h.Filename",strFileName
			AddWhere "h.SyukaDt",strDt
			AddWhere "h.AiteCode",optSaki
			AddSql "group by"
			AddSql " h.JCode"
			AddSql ",h.UkeKbn"
			AddSql ",h.IOKbn"
			AddSql ",h.SyukaDt"
			AddSql ",h.Syushi"
			AddSql ",h.ChuKbn"
			AddSql ",h.ChuName"
			AddSql ",h.AiteCode"
			AddSql ",h.AiteName"
			AddSql ",NyuSyuka"
			AddSql "order by"
			AddSql " h.JCode"
			AddSql ",h.UkeKbn"
			AddSql ",h.IOKbn"
			AddSql ",h.SyukaDt"
			AddSql ",h.Syushi"
			AddSql ",h.ChuKbn"
'			AddSql ",h.ChuName"
			AddSql ",h.AiteCode"
'			AddSql ",h.AiteName"
		case "2","3"	'冷蔵庫メール用
			AddSql "select"
			AddSql " h.JCode"
			AddSql ",h.UkeKbn"
			AddSql ",h.IOKbn"
			AddSql ",h.SyukaDt"
			AddSql ",h.ChuKbn"
			AddSql ",h.ChuName"
			AddSql ",h.AiteCode"
			AddSql ",h.AiteName"
			AddSql ",h.IDNo"
			AddSql ",h.Pn"
			AddSql ",h.Qty"
			AddSql ",h.Syushi"
			AddSql "from HMEM700 h"
			AddWhere "h.Filename",strFileName
			AddWhere "h.SyukaDt",strDt
			AddWhere "h.AiteCode",optSaki
			AddSql "order by 1,2,3,4,5,6,7,8,9"
		end select
		CallSql strSql
		Call GroupHead(-1)
		do while objRs.Eof = False
			Line
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line() 読込
	'-----------------------------------------------------------------------
    Private Function Line()
		Debug ".Line():" & optList
		select case optList
		case "0"
			Write objRs.Fields("Filename")	,0
			Write objRs.Fields("cnt")		,-5
		case "1"
			Write objRs.Fields("JCode")		,9
			Write objRs.Fields("UkeKbn")	,2
			Write objRs.Fields("IOKbn")		,3
			Write objRs.Fields("SyukaDt")	,9
			Write objRs.Fields("Syushi")	,3
			Write objRs.Fields("ChuKbn")	,2
			Write objRs.Fields("ChuName")	,10
			Write objRs.Fields("AiteCode")	,9
			Write objRs.Fields("AiteName")	,20
			Write objRs.Fields("cnt")		,-5
			Write " " & objRs.Fields("NyuSyuka")	,0
		case "2"
			Write objRs.Fields("JCode")		,9
			Write objRs.Fields("UkeKbn")	,2
			Write objRs.Fields("IOKbn")		,3
			Write objRs.Fields("SyukaDt")	,9
			Write objRs.Fields("ChuKbn")	,2
			Write objRs.Fields("ChuName")	,10
			Write objRs.Fields("AiteCode")	,9
			Write objRs.Fields("AiteName")	,20
			Write objRs.Fields("IDNo")		,13
			Write objRs.Fields("Pn")		,14
			Write objRs.Fields("Qty")		,8
			Write objRs.Fields("Syushi")	,0
		case "3"
			if GroupHead(7) = True then
				Write "■",0
				Write objRs.Fields("SyukaDt")	,0
				WriteLine ""
				Write "■",0
				Write RTrim(objRs.Fields("AiteCode"))	,9
				Write RTrim(objRs.Fields("AiteName"))	,0
				WriteLine ""
			end if
			Write objRs.Fields("Pn")		,14
			Write CLng(objRs.Fields("Qty"))	,-5
			Write " " & RTrim(objRs.Fields("Syushi"))	,0
		end select
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
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private Function Execute(byVal strSql)
		Debug ".Execute():" & strSql
		on error resume next
		Call objDb.Execute(strSql)
		Execute = Err.Number & "(0x" & Hex(Err.Number) & ")" & Err.Description
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
	Private Function GetOption(byval strName _
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
