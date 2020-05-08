Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "SaDlv.vbs [option]"
	Wscript.Echo " /db:newsdc3	データベース"
	Wscript.Echo " /j:7			事業部"
	Wscript.Echo " /s:10000		開始行"
	Wscript.Echo " /l:100		読み込む行数"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript SaDlv.vbs /db:newsdc3 /j:7"
End Sub

'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objSaDlv
	Set objSaDlv = New SaDlv
	if objSaDlv.Init() <> "" then
		call usage()
		exit function
	end if
	call objSaDlv.Run()
End Function
'-----------------------------------------------------------------------
'SaDlv
'-----------------------------------------------------------------------
Class SaDlv
	Private	strDBName
	Private	objDB
	Private	objRs
	Public	strJGYOBU
	Private	strAction
	Private	strFileName
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
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
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		strAction = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "j"
			case "s"
			case "l"
			case "debug"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		strJGYOBU = GetOption("j"	,"7")
		set objDB = nothing
		set objRs = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		Call Load()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
	Private	strSql
    Private Function Load()
		Debug ".Load()"
		strSql = ""
		strSql = strSql & "select"
		strSql = strSql & " * "
		strSql = strSql & " from SaDelv"
		strSql = strSql & " order by"
		strSql = strSql & "  Pn"
		strSql = strSql & " ,PlnDt"
		strSql = strSql & " ,AMPM"
		strSql = strSql & " ,PlnTm"
		set objRs = objDB.Execute(strSql)
		dim	strCur
		dim	strPrv
		strPrv = ""
		dim	i
		do while objRs.Eof = false
			strCur = objRS.Fields("Pn")
			Disp objRS.Fields("Pn") _
			 	& " " & objRS.Fields("PlnDt") _
				& " " & objRS.Fields("AMPM") _
				& " " & objRS.Fields("PlnTm") _
				& " " & objRS.Fields("DlvQty")
			if strCur <> strPrv then
				i = 1
				strSql = ""
				strSql = strSql & "insert into SaDelvSum"
				strSql = strSql & " (Pn,Dt_1,Qt_1)"
				strSql = strSql & " values "
				strSql = strSql & " ('" & objRS.Fields("Pn") & "'"
				strSql = strSql & " ,'" & objRS.Fields("PlnDt") & "'"
				strSql = strSql & " ," & objRS.Fields("DlvQty") & ""
				strSql = strSql & " )"
			else
				i = i + 1
				strSql = ""
				strSql = strSql & "update SaDelvSum"
				strSql = strSql & " set Dt_" & i & " = '" & objRS.Fields("PlnDt") & "'"
				strSql = strSql & "   , Qt_" & i & " = " & objRS.Fields("DlvQty") & ""
				strSql = strSql & " where Pn = '" & strCur & "'"
			end if
			strPrv = strCur
			Call objDB.Execute(strSql)
			objRs.MoveNext
		loop
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
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
'		on error resume next
		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field名
	'-------------------------------------------------------------------
	Public Function GetFields(byVal strTable)
		Debug ".GetFields():" & strTable
		dim	strFields
		strFields = ""
		dim	objRs
		set objRS = objDB.Execute("select top 1 * from " & strTable)
		dim	objF
		for each objF in objRS.Fields
			if strFields <> "" then
				strFields = strFields & ","
			end if
			strFields = strFields & objF.Name
		next
		set objRs = nothing
		GetFields = strFields
	End Function
End Class
