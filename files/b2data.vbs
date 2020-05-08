Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "b2data.vbs [option]"
	Wscript.Echo " /db:newsdc1	データベース"
	Wscript.Echo " /make 送り状データ作成(default)"
	Wscript.Echo " /csv  送り状データ出力"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript b2data.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'B2Data
'2016.10.06 追加：配達時間帯区分
'           0812 午前中
'　　　　　 1416 S5YN30		産機群馬ﾊﾟｰﾂ受注管理
'　　　　　　　　5YE2S7000	テクノシステム郡山
'2016.10.07 住所の建物名を分割 <1770棟2F> <ﾃｸﾉWING510> <ﾃｸﾉWING503>
'           会社名を分割 <ＰＥＳ産機システム（株）ディライト　神戸>
'           ３件以上で品名２に<他>が入るように修正
'           顧客管理Noを <ファイル日時>-<便> に変更
'2016.10.25 R-smile(SSX)対応
'2017.06.29 伝票枚数／才数セット処理速度Up
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

Class B2Data
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
		strAction = "make"
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
		select case strAction
		case "csv"
			Call Csv()
		case "make"
			Call Make()
		end select
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Csv() 送り状データ作成
	'-----------------------------------------------------------------------
    Public Function Csv()
		Debug ".Csv()"
		dim	dtToday
		dtToday = Date()
		strSyukaDt = Year(dtToday) & Right("0" & Month(dtToday),2) & Right("0" & Day(dtToday),2)
		Call SetSql("")
		Call SetSql("select")
		Call SetSql("*")
		Call SetSql("from b2data")
		Call SetSql("where SyukaDt = '" & strSyukaDt & "'")
'		Call SetSql("where SyukaDt in ('20180904','20180905')") '
		Call SetSql("order by")
		Call SetSql(" SyukaDt")
		Call SetSql(",ClientNo")
		Call SetSql(",EntTm")
		Call SetSql(",SCode")
		Debug ".Csv():" & strSql
		set objRs = objDB.Execute(strSql)
		intCnt = 0
		Call CsvLine()
		do while objRs.Eof = False
			intCnt = intCnt + 1
			Call CsvLine()
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function

	'-------------------------------------------------------------------
	'YoteiTime() 配達時間帯区分
	'-------------------------------------------------------------------
	'■配達時間帯区分
	'配達時間帯を指定します。
	'半角4文字
	'タイム、ＤＭ便、ネコポス以外
	' 0812 : 午前中
	' 1214 : 12～14時
	' 1416 : 14～16時
	' 1618 : 16～18時
	' 1820 : 18～20時
	' 2021 : 20～21時
	'タイムのみ
	' 0010 : 午前10時まで
	' 0017 : 午後5時まで
	Private	Function YoteiTime()
		select case GetField("SCode")
		'S5YN30		産機群馬ﾊﾟｰﾂ受注管理
		'5YE2S7000	テクノシステム郡山
		case "S5YN30" _
			,"5YE2S7000"
			YoteiTime = "1416"
		case else
			YoteiTime = "0812"
		end select
	End Function
	'-------------------------------------------------------------------
	'YoteiDt() 予定日
	'-------------------------------------------------------------------
	Private Function YoteiDt()
		dim	strDt
		strDt = GetField("SyukaDt")
		dim	dtTmp
		strDt = Left(strDt,4) & "/" & Mid(strDt,5,2) & "/" & Right(strDt,2)
		Debug ".YoteiDt:" & strDt
		dtTmp = CDate(strDt)
		dtTmp = DateAdd("d",1,dtTmp)
		YoteiDt = CStr(dtTmp)
	End Function
	'-------------------------------------------------------------------
	'CsvLine() 1行出力
	'-------------------------------------------------------------------
	Private Function CsvLine()
		Debug ".CsvLine()"
		dim	objF
		'1行目：項目名
		if intCnt = 0 then
			for each objF in objRs.Fields
				WScript.StdOut.Write objF.Name
				WScript.StdOut.Write ","
			next
			WScript.StdOut.Write "予定日"
			WScript.StdOut.Write ",配達時間帯"
			WScript.StdOut.Write ",RsTel"
			WScript.StdOut.Write ",Rs町域"
			WScript.StdOut.Write ",Rs番地"
			WScript.StdOut.Write ",Rs名称1"
			WScript.StdOut.Write ",Rs名称2"
			WScript.StdOut.Write ",Rs荷送人"
			WScript.StdOut.Write ",Rs配達指定日"
			WScript.StdOut.Write ",Rs航空区分"
			WScript.StdOut.Write ",Rs記事1"
			WScript.StdOut.Write ",Rs記事2"
			WScript.StdOut.Write ",Rs記事3"
			WScript.StdOut.Write ",Rs記事4"
			WScript.StdOut.Write ",Rs記事5"
			WScript.StdOut.Write ",Rs出荷番号"
			WScript.StdOut.WriteLine
			exit function
		end if
		'明細
		for each objF in objRs.Fields
			WScript.StdOut.Write Replace(GetField(objF.Name),",",".")
			WScript.StdOut.Write ","
		next
		WScript.StdOut.Write YoteiDt()
		WScript.StdOut.Write "," & YoteiTime()
		WScript.StdOut.Write "," & RsCsv("RsTel")
		WScript.StdOut.Write "," & RsCsv("Rs町域")
		WScript.StdOut.Write "," & RsCsv("Rs番地")
		WScript.StdOut.Write "," & RsCsv("Rs名称1")
		WScript.StdOut.Write "," & RsCsv("Rs名称2")
		WScript.StdOut.Write "," & RsCsv("Rs荷送人")
		WScript.StdOut.Write "," & RsCsv("Rs配達指定日")
		WScript.StdOut.Write "," & RsCsv("Rs航空区分")
		WScript.StdOut.Write "," & RsCsv("Rs記事1")
		WScript.StdOut.Write "," & RsCsv("Rs記事2")
		WScript.StdOut.Write "," & RsCsv("Rs記事3")
		WScript.StdOut.Write "," & RsCsv("Rs記事4")
		WScript.StdOut.Write "," & RsCsv("Rs記事5")
		WScript.StdOut.Write "," & RsCsv("Rs出荷番号")
		WScript.StdOut.WriteLine
	End Function
	'-----------------------------------------------------------------------
	'R-smile用(CSV)
	'-----------------------------------------------------------------------
	Private Function RsCsv(byVal strName)
		dim	strValue
		strValue = ""
		select case strName
		case "RsTel"
			strValue = GetField("STel")
		case "Rs町域"
		case "Rs番地"
		case "Rs名称1"	' 60
			strValue = GetField("SCampany1") & GetField("SCampany2") & GetField("SName")
		case "Rs名称2"
		case "Rs荷送人"
			strValue = RsSender()
		case "Rs配達指定日"
			strValue = YoteiDt()
		case "Rs航空区分"
			select case RsSender()
			case 7,8
				strValue = "ＡＩＲ"
			end select
		case "Rs記事1"
			strValue = GetField("HinName1")
		case "Rs記事2"
			strValue = GetField("HinName2")
		case "Rs記事3"
			strValue = GetField("Kiji")
		case "Rs記事4"
		case "Rs記事5"
		case "Rs出荷番号"	'15桁
							'20
'			strValue = Replace(GetField("ClientNo"),"-","")
			strValue = GetField("SCode")
		end select
		RsCsv = strValue
	End Function
	'-----------------------------------------------------------------------
	'RsSender Rs荷送人
	'	SDC小野⑥Ｐ産機(陸送)
	'	SDC小野⑦Ｐ産機(沖縄)
	'	SDC小野⑧Ｐ産機(エアー)
	'-----------------------------------------------------------------------
	Private	Function RsSender()
		RsSender = 6
		dim	strAddress
		strAddress = GetField("SAddress")
		if Left(strAddress,2) = "沖縄" then
			RsSender = 7
			exit function
		end if
		if Left(strAddress,3) = "北海道" then
			RsSender = 8
			exit function
		end if
	End Function

	'-----------------------------------------------------------------------
	'strClientNo
	'-----------------------------------------------------------------------
	Private	intBin
    Public Function B2ClientNo()
		Debug ".B2ClientNo()"
		dim	strToday
		strToday = Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2)
		Call SetSql("")
		Call SetSql("select")
		Call SetSql("distinct")
		Call SetSql("Filename")
		Call SetSql("from HMTAH015_t")
		Call SetSql("where Filename like 'HMTAH015SZZ.dat." & strToday & "-%'")
		Call SetSql("order by")
		Call SetSql(" Filename")
		Debug ".B2ClientNo():" & strSql
		set objRs = objDB.Execute(strSql)
		strClientNo = ""
		intBin = 0
		do while objRs.Eof = false
			intBin = intBin + 1
			strClientNo = GetField("Filename")
			'HMTAH015SZZ.dat.20161007-062300
			strClientNo = Split(strClientNo,".")(2)
			objRs.MoveNext
		loop
		strClientNo = strClientNo & "-" & intBin
		objRs.Close
	End Function

	'-----------------------------------------------------------------------
	'Make() 送り状データ作成
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		Call B2ClientNo()

		Call SetSql("")
		Call SetSql("select")
		Call SetSql("y.KEY_SYUKA_YMD SyukaDt")
		Call SetSql(",y.KEY_HIN_NO Pn")
		Call SetSql(",convert(y.SURYO,SQL_DECIMAL) Qty")
		Call SetSql(",y.bikou1 Biko1")
		Call SetSql(",d.ChoCode ChoCode")
		Call SetSql(",d.ChoName ChoName")
		Call SetSql(",d.ChoAddress ChoAddress")
		Call SetSql(",d.ChoTel ChoTel")
		Call SetSql(",d.ChoZip ChoZip")
		Call SetSql("from y_syuka y")
		Call SetSql("inner join HtDrctId d on (d.IDNo = y.KEY_ID_NO)")
'		Call SetSql("where Aitesaki	=	'00027768'")
'		Call SetSql(  "and ChoCode	<>	''")
'		Call SetSql(  "and Stts		=	'4'")
'		Call SetSql(  "and TMark	<>	'T'")
'		Call SetSql(  "and SyukaDt	=	'20160913'")
		Call SetSql("order by")
		Call SetSql(" SyukaDt")
'		Call SetSql(",Aitesaki")
		Call SetSql(",ChoCode")
		Debug ".Make():" & strSql
		set objRs = objDB.Execute(strSql)
		prvSyukaDt = ""
		prvChoCode = ""
		do while objRs.Eof = False
			Call MakeData()
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'送り先ごとの件数カウント
	'-------------------------------------------------------------------
	Private	strSyukaDt
	Private	strChoCode
	Private	strChoName
	Private	strClientNo
	Private	strSCode
	Private	prvSyukaDt
	Private	prvChoCode
	Private	intCnt
	Private	Function Count()
		intCnt = intCnt + 1
		if strSyukaDt <> prvSyukaDt _
		or strChoCode <> prvChoCode then
			intCnt = 1
		end if
		prvSyukaDt = strSyukaDt
		prvChoCode = strChoCode
	End Function
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData():" & GetField("SyukaDt") & " " & GetField("ChoCode") & " " & GetField("ChoName")
		strSyukaDt = GetField("SyukaDt")
		strChoCode = GetField("ChoCode")
		strChoName = GetField("ChoName")
'		strClientNo	= GetField("Aitesaki")
		strSCode	= strChoCode
		Call Count()
		Call GetB2Data()
		Call SetB2Data()
	End Function
	'-------------------------------------------------------------------
	'B2Dataフィールドセット
	'-------------------------------------------------------------------
	Private Function SetB2DataField(byVal strName,byVal strValue)
		dim	intLen
		intLen = objB2DataRs.Fields(strName).DefinedSize
		Debug ".SetB2DataField():" & strName & "(" & intLen & "):" & strValue
		strValue = Get_LeftB(strValue,intLen)
		Debug ".SetB2DataField():" & strName & "(" & intLen & "):" & strValue
		if objB2DataRs.Fields(strName) = strValue then
			exit function
		end if
		objB2DataRs.Fields(strName) = strValue
		objB2DataRs.Fields("UpdID")	= "b2data." & intCnt
	End Function
	'-------------------------------------------------------------------
	'お届け先名（漢字）半角32
	'お届け先会社・部門名１	半角50
	'お届け先会社・部門名２	半角50
	'-------------------------------------------------------------------
	Private	strSName
	Private	strSCampany1
	Private	strSCampany2
	Private Function SetB2DataSName()
		strSName = GetField("ChoName")
		strSCampany1 = ""
		strSCampany2 = ""
		dim	intLen
		intLen = LenB(strSName)
		if intLen <= 32 then
			Exit Function
		end if

		select case strSName
		case "ＰＥＳ産機システム（株）ディライト　神戸"
			strSCampany2	= "ＰＥＳ産機システム（株）"
			strSName		= "ディライト　神戸"
			Exit Function
		end select

		dim	aryWord
		aryWord = ""
		strSName = Replace(strSName,"　"," ")
		if inStr(strSName," ") > 0 then
			aryWord = Split(strSName," ")
		end if
		if isArray(aryWord) then
			dim	strWord
			for each strWord in aryWord
				strWord = Trim(strWord)
				if strSCampany2 = "" then
					strSCampany2 = strWord
					strSName = ""
				else
					if strSName <> "" then
						strSName = strSName & " "
					end if
					strSName = strSName & strWord
				end if
			next
		end if
		if LenB(strSName) <= 32 then
			Exit Function
		end if
		strSName = zen2han(strSName)
	End Function
	'-------------------------------------------------------------------
	'B2Data住所分割
	'-------------------------------------------------------------------
	Private	strAddress
	Private	strBillding
	Private	Function SetB2DataAddress()
		strAddress	= GetField("ChoAddress")
		strBillding	= ""
		select case strAddress
		case "群馬県邑楽郡大泉町坂田1-1-1_1770棟2F"
			strAddress	= "群馬県邑楽郡大泉町坂田1-1-1"
			strBillding	= "1770棟2F"
		case "東京都大田区本羽田2丁目12―1ﾃｸﾉWING510"
			strAddress	= "東京都大田区本羽田2丁目12-1"
			strBillding	= "ﾃｸﾉWING510"
		case "東京都大田区本羽田2丁目12―1ﾃｸﾉWING503"
			strAddress	= "東京都大田区本羽田2丁目12-1"
			strBillding	= "ﾃｸﾉWING503"
		end select
	End Function
	'-------------------------------------------------------------------
	'伝票枚数、才数
	'-------------------------------------------------------------------
	Private	intDenCnt
	Private	dblSaisu
	Private Function SetB2DataSaisu()
		'遅い
		'Call SetSql("")
		'Call SetSql("select")
		'Call SetSql("Count(*) c")
		'Call SetSql(",Sum(convert(y.SURYO,SQL_DECIMAL) * convert(i.SAI_SU,sql_decimal)) s")
		'Call SetSql("from HtDrctId d")
		'Call SetSql("inner join y_syuka y on (d.IDNo = y.KEY_ID_NO)")
		'Call SetSql("inner join Item i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)")
		'Call SetSql("where y.KEY_SYUKA_YMD='" & GetField("SyukaDt") & "'")
		'Call SetSql("and d.ChoCode='" & GetField("ChoCode") & "'")
		'遅くない
		SetSql ""
		SetSql "select"
		SetSql "Count(*) c"
		SetSql ",Sum(convert(y.SURYO,SQL_DECIMAL) * convert(i.SAI_SU,sql_decimal)) s"
		SetSql "from y_syuka y"
		SetSql "inner join Item i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
		SetSql "where y.KEY_SYUKA_YMD='" & GetField("SyukaDt") & "'"
		SetSql "and y.KEY_ID_NO in"
		SetSql " (select distinct IDNo from HtDrctId where ChoCode = '" & GetField("ChoCode") & "')"

		dim	objSaisu
		Debug ".SetB2DataSaisu().call"
		set objSaisu = objDB.Execute(strSql)
		Debug ".SetB2DataSaisu().done"
		intDenCnt	= 0
		dblSaisu	= 0
		if objSaisu.EOF = False then
			intDenCnt	= objSaisu.Fields("c")
			dblSaisu	= objSaisu.Fields("s")
		end if
		objSaisu.Close
		set objSaisu = Nothing
	End Function
	'-------------------------------------------------------------------
	'B2Dataレコードセット
	'-------------------------------------------------------------------
	Private Function SetB2Data()
		Debug ".SetB2Data():" & intCnt
		if intCnt = 1 then
'			Call SetB2DataField("ClientNo"	,GetField("Aitesaki"))
			Call SetB2DataField("STel"		,GetField("ChoTel"))
			Call SetB2DataSName()
			Call SetB2DataField("SName"		,strSName)
			Call SetB2DataField("SCampany1"	,strSCampany1)
			Call SetB2DataField("SCampany2"	,strSCampany2)
			Call SetB2DataAddress()
			Call SetB2DataField("SZip"		,GetField("ChoZip"))
			Call SetB2DataField("SAddress"	,strAddress)
			Call SetB2DataField("SBillding"	,strBillding)
			Call SetB2DataField("HinCode1"	,GetField("Pn"))
			Call SetB2DataField("HinName1"	,GetField("Pn") & " " & GetField("Qty") & "個")
			Call SetB2DataField("HinCode2"	,"")
			Call SetB2DataField("HinName2"	,"")
			Call SetB2DataSaisu()
'			Call SetB2DataField("Kiji"		,GetField("Biko1") & " " & intDenCnt & "枚(" & dblSaisu & "才)")
			Call SetB2DataField("Kiji"		,GetField("Biko1") & " 伝票：" & intDenCnt & "枚")
		elseif intCnt = 2 then
			Call SetB2DataField("HinCode2"	,GetField("Pn"))
			Call SetB2DataField("HinName2"	,GetField("Pn") & " " & GetField("Qty") & "個")
		elseif intCnt = 3 then
			Call SetB2DataField("HinName2"	,RTrim(objB2DataRs.Fields("HinName2")) & " 他")
			Debug ".SetB2Data()他:" & objB2DataRs.Fields("HinName2")
		end if
'		Call SetB2DataField("Kiji"		,GetField("ChoCode") & " 伝票：" & intCnt & "枚")
'		Call SetB2DataField("Kiji"		,GetField("Biko1") & " 伝票：" & intCnt & "枚")

		Call DispB2Data()
		Call objB2DataRs.Update()
	End Function
	'-------------------------------------------------------------------
	'B2Data表示
	'-------------------------------------------------------------------
	dim	strDisp
	Private Function DispB2Data()
		strDisp = ""
		strDisp = strDisp & strSyukaDt
		strDisp = strDisp & " " & strClientNo
		strDisp = strDisp & " " & strChoCode
		strDisp = strDisp & " " & strChoName
		strDisp = strDisp & " " & strSCode
		strDisp = strDisp & ":" & intCnt
		Disp strDisp
	End Function
	'-------------------------------------------------------------------
	'B2Dataレコード
	'-------------------------------------------------------------------
	dim	objB2DataRs
	Private Function GetB2Data()
		Debug ".GetB2Data()"
		Call SetSql("")
		Call SetSql("select")
		Call SetSql("*")
		Call SetSql("from B2Data")
		Call SetSql("where SyukaDt	=	'" & strSyukaDt & "'")
'		Call SetSql(  "and ClientNo	=	'" & strClientNo & "'")
		Call SetSql(  "and SCode	=	'" & strSCode & "'")
		Debug ".GetB2Data():" & strSql
		set objB2DataRs = Nothing
		Set objB2DataRs = Wscript.CreateObject("ADODB.Recordset")
		Call objB2DataRs.Open(strSql, objDb, adOpenKeyset, adLockOptimistic)
		if objB2DataRs.Eof = False then
			Exit Function
		end if
		Call objB2DataRs.Close
		Call objB2DataRs.Open("B2Data", objDb, adOpenKeyset, adLockOptimistic,adCmdTableDirect)
		Call objB2DataRs.AddNew
		objB2DataRs.Fields("SyukaDt")	= strSyukaDt
		objB2DataRs.Fields("ClientNo")	= strClientNo
		objB2DataRs.Fields("SCode")		= strSCode
		objB2DataRs.Fields("EntID")		= "b2data.vbs"
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
		on error resume next
		strField = RTrim("" & objRs.Fields(strName))
		if Err.Number <> 0 then
			WScript.StdErr.WriteLine "GetField(" & strName & "):Error:0x" & Hex(Err.Number)
'			WScript.Quit
		end if
		on error goto 0
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
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
	'-------------------------------------------------------------------
	'全角→半角
	'-------------------------------------------------------------------
	Private Function zen2han( byVal strVal )
		dim	objBasp
		Set objBasp = CreateObject("Basp21")
		zen2han = objBasp.StrConv( strVal, 8 )
		Set objBasp = Nothing
	End Function
	'-------------------------------------------------------------------
	'LenB()
	'-------------------------------------------------------------------
	Private Function LenB(byVal strVal)
	    Dim i, strChr
	    LenB = 0
	    If Trim(strVal) <> "" Then
	        For i = 1 To Len(strVal)
	            strChr = Mid(strVal, i, 1)
	            '２バイト文字は＋２
	            If (Asc(strChr) And &HFF00) <> 0 Then
	                LenB = LenB + 2
	            Else
	                LenB = LenB + 1
	            End If
	        Next
	    End If
	End Function
	'-------------------------------------------------------------------
	'Get_LeftB()
	'-------------------------------------------------------------------
	Private Function Get_LeftB(byVal a_Str,byVal a_int)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			Get_LeftB = ""
			Exit Function
		End If
		If a_int = 0 Then
			Get_LeftB = ""
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
		Get_LeftB = iLeftStr
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
			case "make"
				strAction = "make"
			case "csv"
				strAction = "csv"
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
	dim	objB2Data
	Set objB2Data = New B2Data
	if objB2Data.Init() <> "" then
		call usage()
		exit function
	end if
	call objB2Data.Run()
End Function
