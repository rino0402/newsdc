Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "ItemSize.vbs [option] [PN]"
	Wscript.Echo " /db:newsdc1	データベース"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript ItemSize.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'ItemSize(才数テーブル)更新
'2016.10.27 新規
'2016.10.29 Insert/Update後 @@rowcount 表示
'2016.10.31 外装才数／入数の場合 小数４桁四捨五入
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

Class ItemSize
	Private	strDBName
	Private	strPn
	Private	objDB
	Private	objRs
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		strPn = GetOption("pn"	,"")
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
		Call Make()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Make() 才数データ更新
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		SetSql	""
		SetSql	"select"
		SetSql	"distinct"
		SetSql	"ifnull(z.JGYOBU,'*')										zJ"
		SetSql	",i.JGYOBU													JGYOBU"
		SetSql	",i.HIN_GAI													HIN_GAI"
		SetSql	",ifNull(z.Size,0)		Size"
		SetSql	",convert(i.SAI_SU,sql_decimal)								iSize"
		SetSql	",k.SJGyobu				SJGyobu"
		SetSql	",k.SPn					SPn"
		SetSql	",ifNull(k.SQty,0)		SQty"
		SetSql	",ifNull(k.SSize,0)		SSize"
		SetSql	",k.GJGyobu				GJGyobu"
		SetSql	",k.GPn					GPn"
		SetSql	",ifNull(k.GQty,0)		GQty"
		SetSql	",ifNull(k.GSize,0)		GSize"
'		SetSql	"from p_compo_k k"
'		SetSql	"inner join item i"
		SetSql	"from item i"
		SetSql	"left outer join ("
		SetSql	"select"
		SetSql	"k.JGYOBU													JGYOBU"
		SetSql	",k.HIN_GAI													HIN_GAI"
		SetSql	",Max(if(k.DATA_KBN='1',k.KO_JGYOBU,''))					SJGyobu"
		SetSql	",Max(if(k.DATA_KBN='1',k.KO_HIN_GAI,''))					SPn"
		SetSql	",Max(if(k.DATA_KBN='1',convert(KO_QTY,sql_decimal),0))		SQty"
		SetSql	",Max(if(k.DATA_KBN='1',convert(s.SAI_SU,sql_decimal),0))	SSize"
		SetSql	",Max(if(k.DATA_KBN='2',k.KO_JGYOBU,''))					GJGyobu"
		SetSql	",Max(if(k.DATA_KBN='2',k.KO_HIN_GAI,''))					GPn"
		SetSql	",Max(if(k.DATA_KBN='2',convert(k.KO_QTY,sql_decimal),0))	GQty"
		SetSql	",Max(if(k.DATA_KBN='2',convert(s.SAI_SU,sql_decimal),0))	GSize"
		SetSql	"from p_compo_k k"
		SetSql	"left outer join Item s"
		SetSql	"on (k.KO_JGYOBU=s.JGYOBU and k.KO_NAIGAI=s.NAIGAI and k.KO_HIN_GAI=s.HIN_GAI)"
		SetSql	"where k.SHIMUKE_CODE in (select Min(C_Code) from p_code where data_kbn = '04' group by OPTION1)"
		SetSql	"and ((k.DATA_KBN in ('1','2') and k.SEQNO = '010')"
		SetSql	"   or(k.DATA_KBN = '0' and k.SEQNO = '000'))"
		SetSql	"group by"
		SetSql	"JGYOBU,HIN_GAI"
		SetSql	") k"
		SetSql	"on (i.JGYOBU=k.JGYOBU and i.HIN_GAI=k.HIN_GAI)"
		SetSql	"left outer join ItemSize z"
		SetSql	"on (i.JGYOBU=z.JGYOBU and i.HIN_GAI=z.HIN_GAI)"
		SetSql	"where JGYOBU <> 'S' and NAIGAI='1'"
		if strPn <> "" then
			SetSql	"and HIN_GAI='" & strPn & "'"
		end if
		SetSql	"and (iSize <> 0 or SSize <> 0 or GSize <> 0"
		SetSql	"or SPn <> '' or GPn <> '')"
		SetSql	"order by ifnull(z.Size,-1)"
		Debug ".Make():" & strSql
		WScript.StdErr.Write "検索中..." & strPn
		set objRs = objDB.Execute(strSql)
		WScript.StdErr.WriteLine "...Eof:" & objRs.Eof
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
		WScript.StdOut.Write Left(GetField("zJ") & Space(2),2)
		WScript.StdOut.Write Left(GetField("JGYOBU") & Space(2),2)
		WScript.StdOut.Write Left(GetField("HIN_GAI") & Space(20),16)
		WScript.StdOut.Write Right(Space(6) & GetField("Size"),6) & " "
		WScript.StdOut.Write Right(Space(6) & GetField("iSize"),6) & " "
		WScript.StdOut.Write Left(GetField("SJGyobu") & Space(2),2)
		WScript.StdOut.Write Left(GetField("SPn") & Space(10),8)
		WScript.StdOut.Write Right(Space(3) & GetField("SQty"),3)
		WScript.StdOut.Write Right(Space(6) & GetField("SSize"),6) & " "
		WScript.StdOut.Write Left(GetField("GJGyobu") & Space(2),2)
		WScript.StdOut.Write Left(GetField("GPn") & Space(10),8)
		WScript.StdOut.Write Right(Space(3) & GetField("GQty"),3)
		WScript.StdOut.Write Right(Space(6) & GetField("GSize"),6)
		Call Insert()
		WScript.StdOut.WriteLine
	End Function
	'-------------------------------------------------------------------
	'Update
	'-------------------------------------------------------------------
	Private	Function Update()
		Debug ".Update()"
		SetSql ""
		SetSql "update ItemSize "
		SetSql "set"
		SetSql " Size = " & dSize & ""
		SetSql ",SJGYOBU = '" & GetField("SJGYOBU") & "'"
		SetSql ",SPn = '" & GetField("SPn") & "'"
		SetSql ",SQty = " & GetField("SQty") & ""
		SetSql ",SSize = " & GetField("SSize") & ""
		SetSql ",GJGYOBU = '" & GetField("GJGYOBU") & "'"
		SetSql ",GPn = '" & GetField("GPn") & "'"
		SetSql ",GQty = " & GetField("GQty") & ""
		SetSql ",GSize = " & GetField("GSize") & ""
		SetSql ",UpdID = 'ItemSize.vbs'"
		SetSql ",UpdTM = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
		SetSql "where JGYOBU = '" & GetField("JGYOBU") & "'"
		SetSql "and HIN_GAI = '" & GetField("HIN_GAI") & "'"
		SetSql "and ("
		SetSql "Size <> " & dSize
		SetSql "or SJGYOBU <> '" & GetField("SJGYOBU") & "'"
		SetSql "or SPn <> '" & GetField("SPn") & "'"
		SetSql "or SQty <> " & GetField("SQty") & ""
		SetSql "or SSize <> " & GetField("SSize") & ""
		SetSql "or GJGYOBU <> '" & GetField("GJGYOBU") & "'"
		SetSql "or GPn <> '" & GetField("GPn") & "'"
		SetSql "or GQty <> " & GetField("GQty") & ""
		SetSql "or GSize <> " & GetField("GSize") & ""
		SetSql ")"
		WScript.StdOut.Write " Upd:"
		CallSql strSql
	End Function
	'-------------------------------------------------------------------
	'GetDbl
	'-------------------------------------------------------------------
	Private Function GetDbl(byVal v)
		GetDbl = 0
		if v = "" then exit function
		GetDbl = CDbl(v)
	End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
	Private	dSize
	Private Function Insert()
		Debug ".Insert()"
		dSize = 0
		if GetDbl(GetField("GQty")) > 0 then
			dSize = Round(GetDbl(GetField("GSize")) / GetDbl(GetField("GQty")),4)
		end if
		if dSize = 0 then
			dSize = GetDbl(GetField("SSize"))
		end if
		if dSize = 0 then
			dSize = GetDbl(GetField("iSize"))
		end if
		if GetField("zJ") <> "*" then
			Update
			exit function
		end if
		SetSql ""
		SetSql "insert into ItemSize ("
		SetSql " JGYOBU"	
		SetSql ",HIN_GAI"
		SetSql ",Size"
		SetSql ",SJGYOBU"
		SetSql ",SPn"
		SetSql ",SQty"
		SetSql ",SSize"
		SetSql ",GJGYOBU"
		SetSql ",GPn"
		SetSql ",GQty"
		SetSql ",GSize"
		SetSql ",EntID"
		SetSql ",EntTm"
		SetSql ") values ("
		SetSql "'" & GetField("JGYOBU") & "'"
		SetSql ",'" & GetField("HIN_GAI") & "'"
		SetSql "," & dSize & ""
		SetSql ",'" & GetField("SJGYOBU") & "'"
		SetSql ",'" & GetField("SPn") & "'"
		SetSql "," & GetField("SQty") & ""
		SetSql "," & GetField("SSize") & ""
		SetSql ",'" & GetField("GJGYOBU") & "'"
		SetSql ",'" & GetField("GPn") & "'"
		SetSql "," & GetField("GQty") & ""
		SetSql "," & GetField("GSize") & ""
		SetSql ",'ItemSize.vbs'"
		SetSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
		SetSql ")"
		Debug strSql
		WScript.StdOut.Write " Ins:"
		CallSql strSql
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
		on error resume next
		objDb.Execute strSql
		on error goto 0
		dim	intNumver
		dim	strDescription
		intNumver = Err.Number
		strDescription	= Err.Description
		if intNumver = 0 then
			dim	objRc
			set objRc = objDb.Execute("select @@rowcount")
			WScript.StdOut.Write objRc.Fields(0)
		else
			WScript.StdOut.Write RTrim("0x" & Hex(intNumver) & " " & strDescription)
		end if
    End Function
	'-------------------------------------------------------------------
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
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
		For Each strArg In WScript.Arguments.UnNamed
			Init = "オプションエラー:" & strArg
			Disp Init
			Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "pn"
			case "debug"
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
	dim	objItemSize
	Set objItemSize = New ItemSize
	if objItemSize.Init() <> "" then
		call usage()
		exit function
	end if
	call objItemSize.Run()
End Function
