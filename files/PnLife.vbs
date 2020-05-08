Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "PnLife.vbs [option]"
	Wscript.Echo " /db:newsdc1"
	Wscript.Echo " /Zaiko"
	Wscript.Echo " /Syuka"
	Wscript.Echo " /YmNo"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo PnLife.vbs /db:newsdc1 5"
	Wscript.Echo "cscript//nologo PnLife.vbs /db:newsdc1 5 /Zaiko"
End Sub

'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objPnLife
	Set objPnLife = New PnLife
	if objPnLife.Init() <> "" then
		call usage()
		exit function
	end if
	call objPnLife.Run()
End Function
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset		= 1
Const adOpenDynamic		= 2
Const adOpenStatic		= 3

'---- LockTypeEnum Values ----
Const adLockReadOnly 		= 1
Const adLockPessimistic 	= 2
Const adLockOptimistic 		= 3
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

Const adStateClosed		= 0 ' オブジェクトが閉じている

'-----------------------------------------------------------------------
'PnLife
'-----------------------------------------------------------------------
Class PnLife
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strJGyobu
	Private	strAction
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Private Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Private Sub Disp(byVal strMsg)
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
		bError = 0
		dim	strArg
		Init = ""
		strAction = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strJGyobu = "" then
				strJGyobu = strArg
'			else
'				Init = "オプションエラー:" & strArg
'				Disp Init
'				Exit Function
			end if
		Next
		if strJGyobu = "" then
			Init = "." & strArg
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "zaiko"
				strAction = "Zaiko"
			case "syuka"
				strAction = "Syuka"
			case "ymno"
				strAction = "YmNo"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'CheckFunction
	'-----------------------------------------------------------------------
	Private Function CheckFunction(byval strA)
		Debug ".CheckFunction():" & strA
		CheckFunction = False
		if strAction = "" then
			exit function
		end if
		if WScript.Arguments.Named.Exists(strA) then
			exit function
		end if
		CheckFunction = True
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db","newsdc")
		set objDB = nothing
		set objRs = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			strJGyobu = strArg
			Call MonthlyQty()
			Call DailyZaiko()
			Call YmNo()
		Next
		Call CloseDB()
	End Function
	'-------------------------------------------------------------------
	'YmNo
	'-------------------------------------------------------------------
	Private	startYM
	Private Function YmNo()
		if CheckFunction("YmNo") then
			exit function
		end if
		Debug ".YmNo()"
		strSql = ""
		strSql = strSql & "select"
		strSql = strSql & " distinct"
		strSql = strSql & " *"
		strSql = strSql & " from PnLife"
		strSql = strSql & " where JGYOBU='" & strJGyobu & "'"
		strSql = strSql & " order by JGYOBU,HIN_GAI,YM"
		Debug ".YmNo():" & strSql
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strSql, objDb , adOpenDynamic, adLockOptimistic
		dim	intYmNo
		intYmNo = 0
		prvJGYOBU	= ""
		prvHIN_GAI	= ""
		prvYM		= ""
		startYM		= ""
		do while objRs.Eof = False
			curJGYOBU	= RTrim(objRs.Fields("JGYOBU"))
			curHIN_GAI	= RTrim(objRs.Fields("HIN_GAI"))
			curYM		= RTrim(objRs.Fields("YM"))
			if curJGYOBU <> prvJGYOBU or curHIN_GAI	<> prvHIN_GAI then
				startYm = curYM
			end if
			intYmNo = DateDiff("m", startYm  , curYM) + 1
			prvJGYOBU	= curJGYOBU
			prvHIN_GAI	= curHIN_GAI
			prvYM		= curYM
			objRs.Fields("YM_No") = intYmNo
			objRs.Fields("UpdID") = "YmNo"
			objRs.Update
			Call YmNoDisp()
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = nothing
	End Function
	'-------------------------------------------------------------------
	'YmNoDisp
	'-------------------------------------------------------------------
	Private Function YmNoDisp()
		Debug ".YmNoDisp()"
		dim	strMsg
		strMsg = objRs.Fields("JGYOBU")
		strMsg = strMsg & " " & objRs.Fields("HIN_GAI")
		strMsg = strMsg & " " & startYM
		strMsg = strMsg & " " & objRs.Fields("YM")
		strMsg = strMsg & " " & objRs.Fields("YM_No")
		Disp strMsg
	End Function
	'-------------------------------------------------------------------
	'DailyZaiko
	'-------------------------------------------------------------------
	private	strYM
	private	objYM
	Private Function DailyZaiko()
		if CheckFunction("Zaiko") then
			exit function
		end if
		Debug ".DailyZaiko()"
		strSql = ""
		strSql = strSql & "select"
		strSql = strSql & " distinct"
		strSql = strSql & " left(replace(convert(ym,sql_char),'-',''),6) ym"
		strSql = strSql & " from PnLife"
		strSql = strSql & " where JGYOBU='" & strJGyobu & "'"
		strSql = strSql & " order by ym"
		Debug ".DailyZaiko():" & strSql
		set objYM = objDB.Execute(strSql)
		do while objYM.Eof = False
			strYM = objYM.Fields("ym")
			Call DailyZaikoSub()
			objYM.MoveNext
		loop
		set objYM = nothing
	End Function
	'-------------------------------------------------------------------
	'DailyZaiko
	'-------------------------------------------------------------------
	private	bError
	Private Function DailyZaikoSub()
		Debug ".DailyZaikoSub()"
		strSql = ""
		strSql = strSql & "select"
		strSql = strSql & " distinct"
'		strSql = strSql & " top 100"
		strSql = strSql & " JGYOBU"
		strSql = strSql & ",HIN_GAI"
		strSql = strSql & ",Left(DT,6)	DT"
		strSql = strSql & ",Max(ZaikoQty) ZaikoQty"
		strSql = strSql & " from DailyZaiko"
		strSql = strSql & " where JGYOBU='" & strJGyobu & "'"
		strSql = strSql & " and NAIGAI='1'"
		strSql = strSql & " and DT like '" & strYM & "%'"
		strSql = strSql & " group by JGYOBU,HIN_GAI,DT"
		Debug ".DailyZaikoSub():" & strSql
'		Debug ".commandTimeout:" & objDB.commandTimeout
'		objDB.CursorLocation = adUseServer
'		Set objRead = Wscript.CreateObject("ADODB.Recordset")
'		objRead.Open strSql, objDB , adOpenForwardOnly, adLockReadOnly
		set objRead = objDB.Execute(strSql)
		prvJGYOBU	= ""
		prvHIN_GAI	= ""
		prvYM		= ""
		do while objRead.Eof = False
			Call DailyZaikoDisp()
			Call DailyZaikoAdd()
			if bError <> 0 then
				exit do
			end if
			objRead.MoveNext
		loop
		objRead.Close
		set objRead = Nothing
	End Function
	'-------------------------------------------------------------------
	'DailyZaikoAdd
	'-------------------------------------------------------------------
	dim	curJGYOBU
	dim	curHIN_GAI
	dim	curYM
	dim	prvJGYOBU
	dim	prvHIN_GAI
	dim	prvYM
	Private Function DailyZaikoAdd()
		Debug ".DailyZaikoAdd()"
		curJGYOBU	= RTrim(objRead.Fields("JGYOBU"))
		curHIN_GAI	= RTrim(objRead.Fields("HIN_GAI"))
		curYM		= Left(objRead.Fields("DT"),4) & "-" & Mid(objRead.Fields("DT"),5,2) & "-01"
		if curJGYOBU	= prvJGYOBU		then
		if curHIN_GAI	= prvHIN_GAI	then
		if curYM		= prvYM			then
			exit function
		end if
		end if
		end if
		prvJGYOBU	= curJGYOBU
		prvHIN_GAI	= curHIN_GAI
		prvYM		= curYM

		strSql = "insert into PnLife"
		strSql = strSql & "(JGYOBU"
		strSql = strSql & ",HIN_GAI"
		strSql = strSql & ",YM"
		strSql = strSql & ",Z_Qty"
		strSql = strSql & ",EntID"
		strSql = strSql & ") values ("
		strSql = strSql & " '" & curJGYOBU & "'"
		strSql = strSql & ",'" & curHIN_GAI & "'"
		strSql = strSql & ",'" & curYM & "'"
		strSql = strSql & "," & objRead.Fields("ZaikoQty") & ""
'							 1234567890
		strSql = strSql & ",'DailyZaiko'"
		strSql = strSql & ")"
		Debug ".DailyZaikoAdd():" & strSql
		dim	bUpd
		bUpd = False
		on error resume next
			Call objDB.Execute(strSql)
			select case Err.Number
			case 0
				strMsg = ""
			case &h80004005
				strMsg = "■更新■"
				bUpd = True
			case else
				strMsg = "0x" & Hex(Err.Number) & " " & Err.Description
				bError = Err.Number
			end select
			if strMsg <> "" then
				Disp strMsg
			end if
		on error goto 0
		if bUpd then
			Call DailyZaikoUpd()
		end if
	End Function
	'-------------------------------------------------------------------
	'DailyZaikoUpd
	'-------------------------------------------------------------------
	Private Function DailyZaikoUpd()
		Debug ".DailyZaikoUpd()"
		strSql = "update PnLife"
		strSql = strSql & " set Z_Qty=" & objRead.Fields("ZaikoQty")
		strSql = strSql & " , UpdID='DailyZaiko'"
		strSql = strSql & " where JGYOBU='" & curJGYOBU & "'"
		strSql = strSql & " and HIN_GAI='" & curHIN_GAI & "'"
		strSql = strSql & " and YM='" & curYM & "'"
		Debug ".DailyZaikoUpd():" & strSql
'		on error resume next
			Call objDB.Execute(strSql)
			select case Err.Number
			case 0
				strMsg = ""
			case else
				strMsg = "0x" & Hex(Err.Number) & " " & Err.Description
				bError = Err.Number
			end select
			if strMsg <> "" then
				Disp strMsg
			end if
'		on error goto 0
	End Function
	'-------------------------------------------------------------------
	'MonthlyQtyDisp
	'-------------------------------------------------------------------
	Private Function DailyZaikoDisp()
		Debug ".DailyZaikoDisp()"
		dim	strMsg
		strMsg = objRead.Fields("JGYOBU")
		strMsg = strMsg & " " & objRead.Fields("HIN_GAI")
		strMsg = strMsg & " " & objRead.Fields("DT")
		strMsg = strMsg & " " & objRead.Fields("ZaikoQty")
		Disp strMsg
	End Function
	'-------------------------------------------------------------------
	'MonthlyQty
	'-------------------------------------------------------------------
	dim	objRead
	Private	strSql
	Private	strMsg
	Private Function MonthlyQty()
		if CheckFunction("Syuka") then
			exit function
		end if
		Debug ".MonthlyQty()"
		strSql = ""
		strSql = strSql & "select"
'		strSql = strSql & " top 100"
		strSql = strSql & " JGYOBU"
		strSql = strSql & ",HIN_GAI"
		strSql = strSql & ",DT"
		strSql = strSql & ",SyukaCnt"
		strSql = strSql & ",SyukaQty"
		strSql = strSql & " from MonthlyQty"
		strSql = strSql & " where JGYOBU='" & strJGyobu & "'"
		strSql = strSql & " and NAIGAI='1'"
		Debug ".MonthlyQty():" & strSql
		set objRead = objDB.Execute(strSql)
		do while objRead.Eof = False
			Call MonthlyQtyDisp()
			Call MonthlyQtyAdd()
			objRead.MoveNext
		loop
		objRead.Close
		set objRead = Nothing
	End Function
	'-------------------------------------------------------------------
	'MonthlyQtyAdd
	'-------------------------------------------------------------------
	Private Function MonthlyQtyAdd()
		Debug ".MonthlyQtyAdd()"
		strSql = "insert into PnLife"
		strSql = strSql & "(JGYOBU"
		strSql = strSql & ",HIN_GAI"
		strSql = strSql & ",YM"
		strSql = strSql & ",S_Cnt"
		strSql = strSql & ",S_Qty"
		strSql = strSql & ",EntID"
		strSql = strSql & ") values ("
		strSql = strSql & " '" & objRead.Fields("JGYOBU") & "'"
		strSql = strSql & ",'" & RTrim(objRead.Fields("HIN_GAI")) & "'"
		strSql = strSql & ",'" & Left(objRead.Fields("DT"),4) & "-" & Mid(objRead.Fields("DT"),5,2) & "-01'"
		strSql = strSql & "," & objRead.Fields("SyukaCnt") & ""
		strSql = strSql & "," & objRead.Fields("SyukaQty") & ""
'							 1234567890
		strSql = strSql & ",'MonthlyQty'"
		strSql = strSql & ")"
		Debug ".MonthlyQtyAdd():" & strSql
		dim	bUpd
		bUpd = False
		on error resume next
			Call objDB.Execute(strSql)
			select case Err.Number
			case 0
				strMsg = ""
			case &h80004005
				strMsg = "■更新■"
				bUpd = True
			case else
				strMsg = "0x" & Hex(Err.Number) & " " & Err.Description
				bError = Err.Number
			end select
			if strMsg <> "" then
				Disp strMsg
			end if
		on error goto 0
		if bUpd then
			Call MonthlyQtyUpd()
		end if
	End Function
	'-------------------------------------------------------------------
	'MonthlyQtyUpd
	'-------------------------------------------------------------------
	Private Function MonthlyQtyUpd()
		Debug ".MonthlyQtyUpd()"
		strSql = "update PnLife"
		strSql = strSql & " set S_Cnt=" & objRead.Fields("SyukaCnt")
		strSql = strSql & " , S_Qty=" & objRead.Fields("SyukaQty")
		strSql = strSql & " , UpdID='MonthlyQty'"
		strSql = strSql & " where JGYOBU='" & objRead.Fields("JGYOBU") & "'"
		strSql = strSql & " and HIN_GAI='" & RTrim(objRead.Fields("HIN_GAI")) & "'"
		strSql = strSql & " and YM='" & Left(objRead.Fields("DT"),4) & "-" & Mid(objRead.Fields("DT"),5,2) & "-01'"
		Debug ".MonthlyQtyUpd():" & strSql
'		on error resume next
			Call objDB.Execute(strSql)
			select case Err.Number
			case 0
				strMsg = ""
			case else
				strMsg = "0x" & Hex(Err.Number) & " " & Err.Description
				bError = Err.Number
			end select
			if strMsg <> "" then
				Disp strMsg
			end if
'		on error goto 0
	End Function
	'-------------------------------------------------------------------
	'MonthlyQtyDisp
	'-------------------------------------------------------------------
	Private Function MonthlyQtyDisp()
		Debug ".MonthlyQtyDisp()"
		dim	strMsg
		strMsg = objRead.Fields("JGYOBU")
		strMsg = strMsg & " " & objRead.Fields("HIN_GAI")
		strMsg = strMsg & " " & objRead.Fields("DT")
		strMsg = strMsg & " " & objRead.Fields("SyukaCnt")
		strMsg = strMsg & " " & objRead.Fields("SyukaQty")
		Disp strMsg
	End Function
	'-------------------------------------------------------------------
	'OpenRs
	'-------------------------------------------------------------------
	Private	strTable
	Private Function OpenRs()
		strTable = "PnLife"
		Debug ".OpenRs():" & strTable
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strTable, objDb , adOpenDynamic, adLockOptimistic , adCmdTableDirect
'		Set objRs = Wscript.CreateObject("ADODB.Recordset")
'		objRs.Open strTable, objDb , adOpenKeyset, adLockOptimistic , adCmdTableDirect
	End Function
	'-------------------------------------------------------------------
	'CloseRs
	'-------------------------------------------------------------------
	Private Function CloseRs()
		Debug ".CloseRs():" & strTable
		if not objRs is nothing then
			Call objRs.Close()
			set objRs = nothing
		end if
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
		Set objRs = nothing
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		if not objRs is nothing then
			Call objRs.Close()
		end if
		Call objDB.Close()
		set objDB = Nothing
    End Function
End Class
