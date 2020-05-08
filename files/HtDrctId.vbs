Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "HtDrctId.vbs [option]"
	Wscript.Echo " /db:newsdc1 データベース"
	Wscript.Echo " /new        新規のみ(default)"
	Wscript.Echo " /all        全て"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript HtDrctId.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'HtDirct
'2016.10.15 新規 産機直送先登録
'2016.10.27 EntTm をセットするように変更(トリガー削除)
'-----------------------------------------------------------------------
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CommandTypeEnum Values ----
Const adCmdUnspecified	= -1	' Unspecified type of command 
Const adCmdText			= 1		' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2		' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4		' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8		' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

Class HtDrctId
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strAction	' make/csv
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		optNew = "new"
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
		Call Make()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Make() 
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		if optNew = "all" then
			SetSql ""
			SetSql "delete from HtDrctId where IDNo in (select distinct IDNo from HMTAH015)"
			Disp strSql
			Call CallSql(strSql)
		end if
		SetSql ""
		SetSql "select"
		SetSql "distinct"
		SetSql " IDNo"
		SetSql ",ChoCode"
		SetSql ",ChoName"
		SetSql ",ChoZip"
		SetSql ",ChoTel"
		SetSql ",ChoAddress"
		SetSql ",ChoMemo"
		SetSql ",TMark"
		SetSql "from HMTAH015"
		SetSql "where ChoCode<>''"
		if optNew = "new" then
			SetSql " and IDNo not in (select distinct IDNo from HtDrctId)"
		end if
		Debug ".Make():" & strSql
		set objRs = objDB.Execute(strSql)
		do while objRs.Eof = False
			Call MakeData()
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData()"
		Call DispLine()
		SetSql ""
		SetSql "insert into HtDrctId"
		SetSql "("
		SetSql " IDNo"
		SetSql ",ChoCode"
		SetSql ",ChoName"
		SetSql ",ChoZip"
		SetSql ",ChoTel"
		SetSql ",ChoAddress"
		SetSql ",ChoMemo"
		SetSql ",TMark"
		SetSql ",EntID"
		SetSql ",EntTm"
		SetSql ") values ("
		SetSql " '" & GetField("IDNo") & "'"
		SetSql ",'" & GetField("ChoCode") & "'"
		SetSql ",'" & GetField("ChoName") & "'"
		SetSql ",'" & GetField("ChoZip") & "'"
		SetSql ",'" & GetField("ChoTel") & "'"
		SetSql ",'" & GetField("ChoAddress") & "'"
		SetSql ",'" & GetField("ChoMemo") & "'"
		SetSql ",'" & GetField("TMark") & "'"
		SetSql ",'HtDrctId.vbs'"
		SetSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
		SetSql ")"
		Call CallSql(strSql)
	End Function
	'-------------------------------------------------------------------
	'1行表示
	'-------------------------------------------------------------------
	Private objF
	Private Function DispLine()
		Debug ".DispLine()"
		for each objF in objRs.Fields
			WScript.StdOut.Write objF
			WScript.StdOut.Write " "
		next
		WScript.StdOut.WriteLine 
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
	'SQL文字列追加
	'-----------------------------------------------------------------------
	Private	strSql
	Public Function SetSql(byVal s)
		if s = "" then
			strSql = ""
		else
			if strSql <> "" then
				strSql = strSql & " "
			end if
			strSql = strSql & s
		end if
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
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		strField = RTrim("" & objRs.Fields(strName))
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
	End Function
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
	Private	optNew
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			Init = "オプションエラー:" & strArg
			Disp Init
			Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "new"
				optNew = "new"
			case "all"
				optNew = "all"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objHtDrctId
	Set objHtDrctId = New HtDrctId
	if objHtDrctId.Init() <> "" then
		call usage()
		exit function
	end if
	call objHtDrctId.Run()
End Function
