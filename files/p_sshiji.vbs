Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "p_sshiji.vbs [option]"
	Wscript.Echo " /db:newsdc1 データベース"
	Wscript.Echo " /make			登録(default)"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo p_sshiji.vbs /db:newsdc7"
End Sub
'-----------------------------------------------------------------------
'p_sshiji.vbs
'2016.10.20 商品化指示データ登録
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
			case "check"
				optAction = "check"
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
		if optAction = "check" then
			Call Check()
		end if
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Make() 
	'-----------------------------------------------------------------------
    Public Function Check()
		Debug ".Check()"
		SetSql ""
		if GetOption("check","") = "child" then
			SetSql "select"
			SetSql "*"
			SetSql "from p_sshiji_k k"
			SetSql "inner join p_sshiji_o o on(k.SHIJI_NO = o.SHIJI_NO)"
		else
			SetSql "select"
			SetSql " o.SHIJI_NO SHIJI_NO"
			SetSql ",o.ORDER_DT ORDER_DT"
			SetSql ",o.HAKKO_DT HAKKO_DT"
			SetSql ",o.HIN_GAI HIN_GAI"
			SetSql ",o.KAN_DT KAN_DT"
			SetSql ",o.SHIJI_QTY SHIJI_QTY"
			SetSql ",o.BIKOU oBIKOU"
			SetSql ",p.BIKOU pBIKOU"
			SetSql "from p_sshiji_o o"
			SetSql "inner join p_compo p on ("
			SetSql "p.SHIMUKE_CODE=o.SHIMUKE_CODE and"
			SetSql "p.JGYOBU=o.JGYOBU and"
			SetSql "p.NAIGAI=o.NAIGAI and"
			SetSql "p.HIN_GAI=o.HIN_GAI and"
			SetSql "p.DATA_KBN='0')"
		end if
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
    '切り上げ
	'-------------------------------------------------------------------
	Private Function RoundUp(byVal curNum)
		RoundUp = Int(Abs(curNum) * -1) * (Sgn(curNum) * -1)
	End Function
	'-------------------------------------------------------------------
    '切り捨て
	'-------------------------------------------------------------------
	Private Function RoundDown(byVal curNum)
	    RoundDown = Fix(curNum)
	End Function
	'-------------------------------------------------------------------
    '四捨五入
	'-------------------------------------------------------------------
	Private Function RoundOff(byVal curNum)
	    RoundOff = Fix(curNum + (0.5 * Sgn(curNum)))
	End Function
	'-------------------------------------------------------------------
    '使用数訂正
	'-------------------------------------------------------------------
	Private Function UpdateQty(byVal newQty)
		Debug ".UpdateQty():" & newQty
		SetSql ""
		SetSql "update p_sshiji_k"
		SetSql "set"
		SetSql " KO_SHIJI_QTY = '" & newQty & "'"
		SetSql ",UPD_DATETIME = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"		'
		SetSql "where SHIJI_NO = '" & GetField("SHIJI_NO") & "'"
		SetSql "and DATA_KBN = '" & GetField("DATA_KBN") & "'"
		SetSql "and SEQNO = '" & GetField("SEQNO") & "'"
		objDB.Execute strSql
		WScript.StdOut.Write " update:ok"
	End Function
	'-------------------------------------------------------------------
    '備考訂正
	'-------------------------------------------------------------------
	Private Function UpdateBikou()
		Debug ".UpdateBikou()"
		if GetField("oBIKOU") = GetField("pBIKOU") then
			exit function
		end if
		SetSql ""
		SetSql "update p_sshiji_o"
		SetSql "set"
		SetSql " BIKOU = '" & GetField("pBIKOU") & "'"
		SetSql ",UPD_DATETIME = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"		'
		SetSql "where SHIJI_NO = '" & GetField("SHIJI_NO") & "'"
		objDB.Execute strSql
		WScript.StdOut.Write " update:" & GetField("SHIJI_NO") & ":" & GetField("pBIKOU") & ":ok"
	End Function
	'-------------------------------------------------------------------
	'1行表示
	'-------------------------------------------------------------------
	Private objF
	Private Function DispLine()
		Debug ".DispLine()"
		if GetOption("check","") = "child" then
			WScript.StdOut.Write 	   GetField("SHIJI_NO")
			WScript.StdOut.Write " " & GetField("HIN_GAI")
			WScript.StdOut.Write " " & Right(Space(5) & GetField("SHIJI_QTY"),5)
			WScript.StdOut.Write " " & GetField("DATA_KBN")
			WScript.StdOut.Write " " & GetField("SEQNO")
			WScript.StdOut.Write " " & GetField("KO_SYUBETSU")
			WScript.StdOut.Write " " & GetField("KO_JGYOBU")
			WScript.StdOut.Write " " & GetField("KO_NAIGAI")
			WScript.StdOut.Write " " & GetField("KO_HIN_GAI")
			WScript.StdOut.Write " " & Right(Space(5) & GetField("KO_QTY"),5)
			WScript.StdOut.Write " " & Right(Space(5) & GetField("KO_SHIJI_QTY"),5)
			dim	curQty
			dim	newQty
			curQty = CCur(GetField("KO_SHIJI_QTY"))
			if GetField("DATA_KBN") = "2" then
				newQty = RoundUp(CCur(GetField("SHIJI_QTY")) / CCur(GetField("KO_QTY")))
			else
				newQty = CCur(GetField("SHIJI_QTY")) * CCur(GetField("KO_QTY"))
			end if
			if curQty <> newQty then
				WScript.StdOut.Write "(ng)" & newQty
				UpdateQty newQty
			end if
		else
			WScript.StdOut.Write 	   GetField("SHIJI_NO")
			WScript.StdOut.Write " " & GetField("ORDER_DT")
			WScript.StdOut.Write " " & GetField("HAKKO_DT")
			WScript.StdOut.Write " " & GetField("HIN_GAI")
			WScript.StdOut.Write " " & GetField("KAN_DT")
			WScript.StdOut.Write " " & Right(Space(5) & GetField("SHIJI_QTY"),5)
			WScript.StdOut.Write ":" & GetField("oBIKOU")
			WScript.StdOut.Write ":" & GetField("pBIKOU")
			UpdateBikou
		end if
		WScript.StdOut.WriteLine
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
	dim	objPsShiji
	Set objPsShiji = New PsShiji
	if objPsShiji.Init() <> "" then
		call usage()
		exit function
	end if
	call objPsShiji.Run()
End Function
