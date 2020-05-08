Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "JcsPsShiji.vbs [option]"
	Wscript.Echo " /db:newsdc7 データベース"
	Wscript.Echo " /make			P_SSHIJI_O 登録(default)"
	Wscript.Echo " /child			P_SSHIJI_K 登録"
	Wscript.Echo "Ex."
	Wscript.Echo "sc32 //nologo JcsPsShiji.vbs /db:newsdc7"
End Sub
'-----------------------------------------------------------------------
'JcsPsShiji.vbs
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

Class JcsPsShiji
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
			case "child"
				optAction = "child"
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
		select case optAction
		case "make"
			Call Make()
		case "child"
			Call Child()
		end select
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Child() 
	'-----------------------------------------------------------------------
    Public Function Child()
		Debug ".Child()"
		SetSql ""
		SetSql "select"
		SetSql "*"
		SetSql "from p_sshiji_o o"
		SetSql "inner join p_compo_k k"
		SetSql " on (o.SHIMUKE_CODE = k.SHIMUKE_CODE and o.JGYOBU = k.JGYOBU and o.NAIGAI = k.NAIGAI and o.HIN_GAI = k.HIN_GAI and k.DATA_KBN <> '0')"
'		SetSql "inner join JcsItem i on (i.MazdaPn = o.HIN_GAI)"
'		SetSql "inner join JcsType t on (t.SType = i.SType)"
		SetSql "where SHIJI_NO not in (select distinct SHIJI_NO from p_sshiji_k)"
		SetSql "order by"
		SetSql " o.SHIJI_NO"
		SetSql ",k.DATA_KBN"
		SetSql ",k.SEQNO"
		set objRs = objDB.Execute(strSql)
		do while objRs.Eof = False
			CDispLine
			CMakeLine
'			CDispLine
'			CMakeLine "1","010","","S","1",GetField("SPn1"),GetField("SQty1")
'			CMakeLine "1","020","","S","1",GetField("SPn2"),GetField("SQty2")
'			CMakeLine "1","030","","S","1",GetField("SPn3"),GetField("SQty3")
'			CMakeLine "1","040","","S","1",GetField("SPn4"),GetField("SQty4")
'			CMakeLine "2","010","","S","1",GetField("GPn"),GetField("GQty")
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'子 1行登録
	'-------------------------------------------------------------------
	Private Function CMakeLine()
		Debug ".CMakeLine()"
		dim	strKO_SHIJI_QTY
		strKO_SHIJI_QTY = ""
		if GetField("DATA_KBN") = "2" then
			strKO_SHIJI_QTY = Round(CDbl(GetField("SHIJI_QTY")) / CDbl(GetField("KO_QTY")) + 0.5,0)
		else
			strKO_SHIJI_QTY = CDbl(GetField("SHIJI_QTY")) * CDbl(GetField("KO_QTY"))
		end if
		SetSql ""
		SetSql "insert into p_sshiji_k ("
		SetSql "DATA_KBN"			'
		SetSql ",SEQNO"				'
		SetSql ",KO_SYUBETSU"		'
		SetSql ",KO_JGYOBU"			'
		SetSql ",KO_NAIGAI"			'
		SetSql ",KO_HIN_GAI"		'
		SetSql ",KO_QTY"			'
		SetSql ",KO_SHIJI_QTY"		'
		SetSql ",KO_BIKOU"			'
		SetSql ",KO_ID_NO"			'
		SetSql ",CALCEL_F"			'
		SetSql ",CANCEL_DATETIME"	'
		SetSql ",SHIJI_NO"			'
		SetSql ",HIKIATE_QTY"		'在庫引当数 2012.03.09
		SetSql ",IDO_SUMI"			'移動済み 空白:未　9:済み 2012.03.09
		SetSql ",ST_TANABAN"		'標準棚番 2012.03.18
		SetSql ",IDO_SUMI_QTY"		'移動済み数量 2012.04.13
		SetSql ",COMPO_TANTO"		'構成ﾁｪｯｸ   担当者          2012.04.20
		SetSql ",COMPO_YMDHS"		'           日時            2012.04.20
		SetSql ",COMPO_Sumi_Cnt"	'           ﾁｪｯｸ済み数      2012.04.20
		SetSql ",COMPO_ALL_Cnt"		'           構成数          2012.04.20
		SetSql ",UPD_DATETIME"		'
		SetSql ") values ("
		SetSql " '" & GetField("DATA_KBN") & "'"
		SetSql ",'" & GetField("SEQNO") & "'"
		SetSql ",'" & GetField("KO_SYUBETSU") & "'"
		SetSql ",'" & GetField("KO_JGYOBU") & "'"
		SetSql ",'" & GetField("KO_NAIGAI") & "'"
		SetSql ",'" & GetField("KO_HIN_GAI") & "'"
		SetSql ",'" & GetField("KO_QTY") & "'"
		SetSql ",'" & strKO_SHIJI_QTY & "'"
		SetSql ",'" & GetField("KO_BIKOU") & "'"
		SetSql ",''"	'KO_ID_NO"		'
		SetSql ",''"	'CALCEL_F"		'
		SetSql ",''"	'CANCEL_DATETIME"		'
		SetSql ",'" & GetField("SHIJI_NO") & "'"	'SHIJI_NO"		'
		SetSql ",''"	'HIKIATE_QTY"		'在庫引当数 2012.03.09
		SetSql ",''"	'IDO_SUMI"		'移動済み 空白:未　9:済み 2012.03.09
		SetSql ",''"	'ST_TANABAN"		'標準棚番 2012.03.18
		SetSql ",''"	'IDO_SUMI_QTY"		'移動済み数量 2012.04.13
		SetSql ",''"	'COMPO_TANTO"		'構成ﾁｪｯｸ   担当者          2012.04.20
		SetSql ",''"	'COMPO_YMDHS"		'           日時            2012.04.20
		SetSql ",''"	'COMPO_Sumi_Cnt"		''           ﾁｪｯｸ済み数      2012.04.20
		SetSql ",''"	'COMPO_ALL_Cnt"		''           構成数          2012.04.20
		SetSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
												'UPD_DATETIME"	'	20161020102638
		SetSql ")"
		Debug strSql
		on error resume next
		Call objDB.Execute(strSql)
		if Err.Number <> 0 then
			WScript.StdOut.Write ":" & Hex(Err.Number) & ":" & Err.Description
			WScript.Quit
		end if
		on error goto 0
		WScript.StdOut.WriteLine
	End Function
	'-------------------------------------------------------------------
    '切り上げ
	'-------------------------------------------------------------------
	Private Function RoundUp(byVal curNum)
		RoundUp = Int(Abs(curNum) * -1) * (Sgn(curNum) * -1)
	End Function
	'-------------------------------------------------------------------
	'子 1行登録
	'-------------------------------------------------------------------
	Private Function CMakeLine000(	byVal strDATA_KBN _
							,byVal strSEQNO _
							,byVal strKO_SYUBETSU _
							,byVal strKO_JGYOBU _
							,byVal strKO_NAIGAI _
							,byVal strKO_HIN_GAI _
							,byVal strKO_QTY _
								)
		Debug ".CMakeLine()"
		if strKO_HIN_GAI = "" then
			exit function
		end if
		dim	curQty
		if strDATA_KBN = "2" then
			curQty = RoundUp(CCur(GetField("SHIJI_QTY")) / CCur(strKO_QTY))
		else
			curQty = CCur(GetField("SHIJI_QTY")) * CCur(strKO_QTY)
		end if
		SetSql ""
		SetSql "insert into p_sshiji_k ("
		SetSql "DATA_KBN"			'
		SetSql ",SEQNO"				'
		SetSql ",KO_SYUBETSU"		'
		SetSql ",KO_JGYOBU"			'
		SetSql ",KO_NAIGAI"			'
		SetSql ",KO_HIN_GAI"		'
		SetSql ",KO_QTY"			'
		SetSql ",KO_SHIJI_QTY"		'
		SetSql ",KO_BIKOU"			'
		SetSql ",KO_ID_NO"			'
		SetSql ",CALCEL_F"			'
		SetSql ",CANCEL_DATETIME"	'
		SetSql ",SHIJI_NO"			'
		SetSql ",HIKIATE_QTY"		'在庫引当数 2012.03.09
		SetSql ",IDO_SUMI"			'移動済み 空白:未　9:済み 2012.03.09
		SetSql ",ST_TANABAN"		'標準棚番 2012.03.18
		SetSql ",IDO_SUMI_QTY"		'移動済み数量 2012.04.13
		SetSql ",COMPO_TANTO"		'構成ﾁｪｯｸ   担当者          2012.04.20
		SetSql ",COMPO_YMDHS"		'           日時            2012.04.20
		SetSql ",COMPO_Sumi_Cnt"	'           ﾁｪｯｸ済み数      2012.04.20
		SetSql ",COMPO_ALL_Cnt"		'           構成数          2012.04.20
		SetSql ",UPD_DATETIME"		'
		SetSql ") values ("
		SetSql " '" & strDATA_KBN & "'"
		SetSql ",'" & strSEQNO & "'"
		SetSql ",'" & strKO_SYUBETSU & "'"
		SetSql ",'" & strKO_JGYOBU & "'"
		SetSql ",'" & strKO_NAIGAI & "'"
		SetSql ",'" & strKO_HIN_GAI & "'"
		SetSql ",'" & strKO_QTY & "'"
		SetSql ",'" & curQty & "'"
		SetSql ",''"	'KO_BIKOU"		'
		SetSql ",''"	'KO_ID_NO"		'
		SetSql ",''"	'CALCEL_F"		'
		SetSql ",''"	'CANCEL_DATETIME"		'
		SetSql ",'" & GetField("SHIJI_NO") & "'"	'SHIJI_NO"		'
		SetSql ",''"	'HIKIATE_QTY"		'在庫引当数 2012.03.09
		SetSql ",''"	'IDO_SUMI"		'移動済み 空白:未　9:済み 2012.03.09
		SetSql ",''"	'ST_TANABAN"		'標準棚番 2012.03.18
		SetSql ",''"	'IDO_SUMI_QTY"		'移動済み数量 2012.04.13
		SetSql ",''"	'COMPO_TANTO"		'構成ﾁｪｯｸ   担当者          2012.04.20
		SetSql ",''"	'COMPO_YMDHS"		'           日時            2012.04.20
		SetSql ",''"	'COMPO_Sumi_Cnt"		''           ﾁｪｯｸ済み数      2012.04.20
		SetSql ",''"	'COMPO_ALL_Cnt"		''           構成数          2012.04.20
		SetSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
												'UPD_DATETIME"	'	20161020102638
		SetSql ")"
		Debug strSql
		on error resume next
		Call objDB.Execute(strSql)
		on error goto 0
	End Function
	'-------------------------------------------------------------------
	'子 1行表示
	'-------------------------------------------------------------------
	Private Function CDispLine()
		Debug ".CDispLine()"
		WScript.StdOut.Write "" & GetField("SHIJI_NO")
		WScript.StdOut.Write " " & GetField("HAKKO_DT")
		WScript.StdOut.Write " " & GetField("SHIMUKE_CODE")
		WScript.StdOut.Write " " & GetField("JGYOBU")
		WScript.StdOut.Write " " & GetField("NAIGAI")
		WScript.StdOut.Write " " & Left(GetField("HIN_GAI") & Space(20),20)
		WScript.StdOut.Write " " & GetField("DATA_KBN")
		WScript.StdOut.Write " " & GetField("SEQNO")
		WScript.StdOut.Write " " & GetField("KO_SYUBETSU")
		WScript.StdOut.Write " " & GetField("KO_JGYOBU")
		WScript.StdOut.Write " " & GetField("KO_NAIGAI")
		WScript.StdOut.Write " " & Left(GetField("KO_HIN_GAI") & Space(20),10)
		WScript.StdOut.Write " " & Right(Space(5) & GetField("KO_QTY"),5)
		WScript.StdOut.Write " " & Left(GetField("KO_BIKOU") & Space(20),20)
		WScript.StdOut.Write " " & GetField("CLASS_CODE")
'		WScript.StdOut.WriteLine
	End Function
	Private Function CDispLine000()
		Debug ".CDispLine()"
		WScript.StdOut.Write "" & GetField("SHIJI_NO")
		WScript.StdOut.Write " " & Left(GetField("S_CLASS_CODE") & Space(10),8)
		WScript.StdOut.Write " " & Left(GetField("SPn1") & Space(10),8)
		WScript.StdOut.Write " " & Right(Space(4) & GetField("SQty1"),4)
		WScript.StdOut.Write " " & Left(GetField("SPn2") & Space(10),8)
		WScript.StdOut.Write " " & Right(Space(4) & GetField("SQty2"),4)
		WScript.StdOut.Write " " & Left(GetField("SPn3") & Space(10),8)
		WScript.StdOut.Write " " & Right(Space(4) & GetField("SQty3"),4)
		WScript.StdOut.Write " " & Left(GetField("SPn4") & Space(10),8)
		WScript.StdOut.Write " " & Right(Space(4) & GetField("SQty4"),4)
		WScript.StdOut.Write " " & GetField("PUnit")
		WScript.StdOut.Write " " & Left(GetField("GPn") & Space(10),8)
		WScript.StdOut.Write " " & Right(Space(4) & GetField("GQty"),4)
		WScript.StdOut.WriteLine
	End Function
	'-----------------------------------------------------------------------
	'Make() 
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		if GetOption("make","order") = "order" then
			SetSql ""
			SetSql "select"
			SetSql "*"
			SetSql "from JcsOrder j"
			SetSql "left outer join p_sshiji_o o on (o.SHIJI_NO = (RTrim(j.NohinNo) + RTrim(j.NohinNo2)))"
			SetSql "where (RTrim(j.NohinNo) + RTrim(j.NohinNo2)) not in (select distinct SHIJI_NO from p_sshiji_o)"
		else
			SetSql ""
			SetSql "select"
			SetSql "j.ID XlsRow"
			SetSql ",o.SHIJI_NO SHIJI_NO"
			SetSql ",'' ""No"""
			SetSql ",'' ""PartWH"""
			SetSql ",'' ""Biko"""
			SetSql ",j.*"
			SetSql "from JcsIdo j"
			SetSql "left outer join p_sshiji_o o on (o.SHIJI_NO = (RTrim(j.NohinNo) + RTrim(j.NohinNo2)))"
			SetSql "where j.ID >= 10"
			SetSql "and j.NohinNo<>''"
			SetSql "and (RTrim(j.NohinNo) + RTrim(j.NohinNo2)) not in (select distinct SHIJI_NO from p_sshiji_o)"
		end if
		set objRs = objDB.Execute(strSql)
		do while objRs.Eof = False
			DispLine
			MakeLine
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
	'1行登録
	'-------------------------------------------------------------------
	Private Function MakeLine()
		Debug ".MakeLine()"
		dim	strNohinNo
		strNohinNo = GetField("NohinNo")
		if strNohinNo = "" then
			exit function
		end if
		if GetField("SHIJI_NO") <> "" then
			exit function
		end if
		SetSql ""
		SetSql "insert into p_sshiji_o ("
		SetSql "HAKKO_DT"	'	20161020
		SetSql ",PRINT_DATETIME"	'	20161020091701
		SetSql ",TANTO_CODE"	'	00010
		SetSql ",SHONIN_CODE"	'	02430
		SetSql ",SHIMUKE_CODE"	'	01
		SetSql ",JGYOBU"	'	A
		SetSql ",NAIGAI"	'	1
		SetSql ",HIN_GAI"	'	ACRA50C00320
		SetSql ",SHIJI_QTY"	'	00000001.00
		SetSql ",UKEHARAI_CODE"	'	ZD3
		SetSql ",S_CLASS_CODE"	'	ZZZ
		SetSql ",F_CLASS_CODE"	'	
		SetSql ",N_CLASS_CODE"	'	
		SetSql ",S_TANTO"	'	
		SetSql ",SAMPLE_F"	'	0
		SetSql ",SHIJI_F"	'	0
		SetSql ",TORI_KBN"	'	3
		SetSql ",PRI_SHIJI"	'	1
		SetSql ",PRI_PARTS"	'	1
		SetSql ",PRI_GAISOU"	'	1
		SetSql ",PRI_KISHU"	'	0
		SetSql ",BIKOU"	'	重要安全部品 20160216資材設定
		SetSql ",KAN_F"	'	1
		SetSql ",KAN_DT"	'	20161020
		SetSql ",BUNNOU_CNT"	'	00
		SetSql ",UKEIRE_QTY"	'	00000001.00
		SetSql ",CANCEL_F"	'0
		SetSql ",CANCEL_DATETIME"
		SetSql ",ORDER_DT"	'
		SetSql ",SHIJI_NO"	'00197218
		SetSql ",ORDER_DT_SEQ"
		SetSql ",COMPO_END_F"	
		SetSql ",UPD_DATETIME"	'	20161020102638
		SetSql ") values ("
		SetSql "'" & Replace(GetField("OrderDt"),"/","") & "'"	'HAKKO_DT"	'	20161020
		SetSql ",''"							''PRINT_DATETIME"	'	20161020091701
		SetSql ",'99999'"						'TANTO_CODE"	'	00010
		SetSql ",'99999'"						'SHONIN_CODE"	'	02430
		SetSql ",'01'"							'SHIMUKE_CODE"	'	01
		SetSql ",'J'"							'JGYOBU"	'	A
		SetSql ",'1'"							'NAIGAI"	'	1
		SetSql ",'" & GetField("MazdaPn") & "'"	'HIN_GAI"	'	ACRA50C00320
		SetSql ",'" & GetField("Qty") & "'"		'SHIJI_QTY"	'	00000001.00
		SetSql ",'ZI0'"							'UKEHARAI_CODE"	'	ZD3
		SetSql ",'" & GetField("Pn") & "'"		'S_CLASS_CODE"	'	ZZZ
		SetSql ",'" & GetField("Location") & "'"	'F_CLASS_CODE"	'	
		SetSql ",'" & GetField("No") & "'"		'N_CLASS_CODE"	'	
		SetSql ",'" & GetField("PartWH") & "'"	'S_TANTO"	'	
		SetSql ",'0'"							'SAMPLE_F"	'	0
		SetSql ",'0'"							'SHIJI_F"	'	0
		SetSql ",'3'"							'TORI_KBN"	'	3
		SetSql ",'1'"							'PRI_SHIJI"	'	1
		SetSql ",'0'"							'PRI_PARTS"	'	1
		SetSql ",'0'"							'PRI_GAISOU"	'	1
		SetSql ",'0'"							'PRI_KISHU"	'	0
		SetSql ",'" & GetField("Biko") & "'"	'BIKOU"	'	重要安全部品 20160216資材設定
		SetSql ",'0'"							'KAN_F"	'	1
		SetSql ",''"							'KAN_DT"	'	20161020
		SetSql ",''"							'BUNNOU_CNT"	'	00
		SetSql ",''"							'UKEIRE_QTY"	'	00000001.00
		SetSql ",'0'"							'CANCEL_F"	'0
		SetSql ",''"							'CANCEL_DATETIME"
		SetSql ",'" & Replace(GetField("DlvDt"),"/","") & "'"	'ORDER_DT"	'
		SetSql ",'" & GetField("NohinNo") & GetField("NohinNo2") & "'"	'SHIJI_NO"	'00197218
		SetSql ",''"							'ORDER_DT_SEQ"
		SetSql ",''"							'COMPO_END_F"	
		SetSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
												'UPD_DATETIME"	'	20161020102638
		SetSql ")"
		Call objDB.Execute(strSql)
	End Function
	'-------------------------------------------------------------------
	'1行表示
	'-------------------------------------------------------------------
'	Private objF
	Private Function DispLine()
		Debug ".DispLine()"
		WScript.StdOut.Write 	   GetField("XlsRow")
		WScript.StdOut.Write " " & GetField("OrderDt")
		WScript.StdOut.Write " " & GetField("NohinNo")
		WScript.StdOut.Write " " & GetField("NohinNo2")
		WScript.StdOut.Write " " & GetField("Location")
		WScript.StdOut.Write " " & GetField("DlvDt")
		WScript.StdOut.Write " " & Left(GetField("MazdaPn") & Space(20),20)
		WScript.StdOut.Write " " & Left(GetField("Pn") & Space(20),20)
		WScript.StdOut.Write " " & Right(Space(5) & GetField("Qty"),5)
'		WScript.StdOut.Write " " & GetField("Dt")
'		WScript.StdOut.Write " " & GetField("DestCode")
		WScript.StdOut.Write "(" & GetField("SHIJI_NO") & ")"
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
		Debug ".GetField():" & strName
		on error resume next
		strField = RTrim("" & objRs.Fields(strName))
		If Err.Number <> 0 then
			WScript.StdOut.WriteLine strName
			WScript.StdOut.WriteLine "0x" & Hex(Err.Number)
			WScript.StdOut.WriteLine Err.Description
			WScript.Quit
		end if
		on error goto 0
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
	dim	objJcsPsShiji
	Set objJcsPsShiji = New JcsPsShiji
	if objJcsPsShiji.Init() <> "" then
		call usage()
		exit function
	end if
	call objJcsPsShiji.Run()
End Function
