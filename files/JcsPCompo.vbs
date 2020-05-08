Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "JcsPCompo.vbs [option]"
	Wscript.Echo " /db:newsdc1 データベース"
	Wscript.Echo " /make			登録(default)"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo JcsPCompo.vbs /db:newsdc7"
End Sub
'-----------------------------------------------------------------------
'JcsPCompo.vbs
'2016.10.20 商品化構成マスター登録
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

Class PsShiji
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
	Private	optAction
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "make"
				optAction = "make"
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
	Private	strDBName
	Private	objDB
	Private	objRs
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		optAction = "make"
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
		if optAction = "make" then
			Call Make()
		end if
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Make() 
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		SetSql ""
		SetSql "select"
		SetSql "*"
		SetSql "from JcsOrder"
		set objRs = objDB.Execute(strSql)
		do while objRs.Eof = False
			DispLine
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
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
	'1行表示
	'-------------------------------------------------------------------
	Private objF
	Private Function DispLine()
		Debug ".DispLine()"
		WScript.StdOut.Write 	   GetField("XlsRow")
		WScript.StdOut.Write " " & GetField("OrderDt")
		WScript.StdOut.Write " " & GetField("NohinNo")
		WScript.StdOut.Write " " & GetField("NohinNo2")
		WScript.StdOut.Write " " & GetField("PartWH")
		WScript.StdOut.Write " " & GetField("DlvDt")
		WScript.StdOut.Write " " & Left(GetField("MazdaPn") & Space(20),20)
		WScript.StdOut.Write " " & Left(GetField("Pn") & Space(20),20)
		WScript.StdOut.Write " " & Right(Space(5) & GetField("Qty"),5)
		WScript.StdOut.Write " " & GetField("Location")
		WScript.StdOut.Write " " & GetField("MPn")
		WScript.StdOut.Write " " & GetField("InfoDlvDt")
		WScript.StdOut.Write " " & GetField("InfoTotal")
		WScript.StdOut.Write " " & GetField("InfoZan")
		WScript.StdOut.Write " " & GetField("NG")
		WScript.StdOut.Write " " & GetField("Biko")
		WScript.StdOut.Write " " & GetField("No")
		WScript.StdOut.WriteLine
'		for each objF in objRs.Fields
'			WScript.StdOut.Write RTrim("" & objF)
'			WScript.StdOut.Write " "
'		next
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
	Function GetOption(byval strName ,byval strDefault)
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
End Class	' PsShiji
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objJcsPCompo
	Set objJcsPCompo = New JcsPCompo
	if objJcsPCompo.Init() <> "" then
		call usage()
		exit function
	end if
	call objJcsPCompo.Run()
End Function
