Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "RSmile.vbs [option]"
	Wscript.Echo " /db:newsdc1	データベース"
	Wscript.Echo " /make      送り状データ作成(default)"
	Wscript.Echo " /make:test 送り状データ作成:Test用(全件)"
	Wscript.Echo " /csv       送り状データ出力"
	Wscript.Echo "Ex."
	Wscript.Echo "sc32//nologo RSmile.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'RSmile CSVファイル作成
'2016.10.25 R-smile(SSX)
'2016.10.28 /make:test 送り状データ作成:Test用(全件)
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

Class Rsmile
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
		strPrgId = GetOption("make"	,"RSmile.vbs")
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
		Call SetSql("from RSmile")
		Call SetSql("where SyukaDt = '" & strSyukaDt & "'")
		Call SetSql("order by")
		Call SetSql(" Id")
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
			WScript.StdOut.Write "届先"		'1	お届先コード		半角数字	15	15	
			WScript.StdOut.Write ",電話番号"	'2	電話番号	○	半角数字	15	15	ハイフン付き
			WScript.StdOut.Write ",郵便番号"	'3	郵便番号		半角数字	7	7	ハイフンなし
			WScript.StdOut.Write ",都道府県"	'4	都道府県	○	全角文字	10	20	
			WScript.StdOut.Write ",市区町村"	'5	市区町村	○	全角文字	20	40	
			WScript.StdOut.Write ",町域"		'6	町域	○	全角文字	30	60	
			WScript.StdOut.Write ",番地"		'7	番地・ビル名		全角文字	30	60	番地・ビル名など
			WScript.StdOut.Write ",名称１"	'8	名称１	○	全角文字	30	60	
			WScript.StdOut.Write ",名称２"	'9	名称２		全角文字	30	60	
'			WScript.StdOut.Write ",敬称"		'10	お届け先敬称		全角文字	2	4	「なし」 、「様」、「御中」、「殿」、ブランク
'			WScript.StdOut.Write ",個数"		'11	個数	○	半角数字	4	4	インポート可能な桁数は3桁です
'			WScript.StdOut.Write ",重量"		'12	重量	○	半角数字	7	7	"小数点入力可（小数点以下3桁まで）         ０も可"
			WScript.StdOut.Write ",荷送人"	'13	荷送人ＣＤ		半角数字	6	6	"未入力時はデフォルトの荷送人ＣＤを使用         ※デフォルトの荷送人ＣＤが未設定の場合、必須入力"
											'14	荷送人電話番号	○	半角数字	15	15	ハイフン付き
											'15	荷送人郵便番号		半角数字	7	7	ハイフンなし
											'16	荷送人都道府県	○	全角文字	10	20	
											'17	荷送人市区町村	○	全角文字	20	40	
											'18	荷送人町域	○	全角文字	30	60	
											'19	荷送人番地		全角文字	30	60	荷送人の番地・ビル名など
											'20	会社名	○	全角文字	30	60	
											'21	荷送人担当者名		全角文字	30	60	未入力時はマスタの内容を使用
											'22	お客さまＮｏ	○	半角文字	6	6	
											'23	請求単位		半角文字	20	20	99：共通レイアウトを選択している場合は、請求単位フリー入力無許可ユーザーは紐付不可
											'24	伝票区分	○	全角文字	2	4	"元払：""元払""、着払：""着払""、      代引：""代引"""
			WScript.StdOut.Write ",航空区分"	'25	航空区分	△	全角文字	3	6	"輸送区分が航空の時のみ有効         急便：空白、ＡＩＲ：""ＡＩＲ"""
			WScript.StdOut.Write ",配達日"	'26	配達指定日		半角数字	8	8	年月日（年は西暦4桁）
'			WScript.StdOut.Write ",配達時間"	'27	配達指定時間帯		半角英字	2	2	
			WScript.StdOut.Write ",記事1"	'28	品名	○	全角文字	30	60	
			WScript.StdOut.Write ",記事2"	'29	記事２		全角文字	30	60	
			WScript.StdOut.Write ",記事3"	'30	記事３		全角文字	30	60	
			WScript.StdOut.Write ",記事4"	'31	記事４		全角文字	30	60	
			WScript.StdOut.Write ",記事5"	'32	記事５		全角文字	30	60	
			WScript.StdOut.Write ",出荷日"	'33	出荷日		半角数字	8	8	"年月日（年は西暦4桁）      30日先まで指定可能。      過去日の場合は当日の日付に変更される"
											'34			半角数字	10	10	現在使用していません。
			WScript.StdOut.Write ",出荷番号"	'35	出荷番号		半角数字	15	15	
			WScript.StdOut.Write ",伝票番号"	'36	伝票番号		半角数字	1	1	入力されている場合、印刷時に採番しない
			WScript.StdOut.WriteLine
			exit function
		end if
		'明細
		RsCsv "届先"
		RsCsv "電話番号"	
		RsCsv "郵便番号"	
		RsCsv "都道府県"	
		RsCsv "市区町村"	
		RsCsv "町域"		
		RsCsv "番地"		
		RsCsv "名称１"	
		RsCsv "名称２"	
'		RsCsv "敬称"		
'		RsCsv "個数"		
'		RsCsv "重量"		
		RsCsv "荷送人"	
		RsCsv "航空区分"	
		RsCsv "配達日"	
'		RsCsv "配達時間"	
		RsCsv "記事1"	
		RsCsv "記事2"	
		RsCsv "記事3"	
		RsCsv "記事4"	
		RsCsv "記事5"	
		RsCsv "出荷日"	
		RsCsv "出荷番号"	
		RsCsv "伝票番号"	
		WScript.StdOut.WriteLine
	End Function
	'-----------------------------------------------------------------------
	'電話番号 ハイフン追加
	'-----------------------------------------------------------------------
	Private	Function Tel(byVal strTel)
		Debug ".Tel():" & strTel
		Tel = strTel
		if inStr(strTel,"-") > 0 then
			exit function
		end if
		dim	strTel1
		dim	strTel2
		dim	strTel3
		select case len(strTel)
		case 11
			Debug ".Tel():11:" & strTel
			strTel1 = left(strTel,3)
			strTel2 = mid(strTel,4,4)
			strTel3 = right(strTel,4)
			strTel = strTel1 & "-" & strTel2 & "-" & strTel3
			Debug ".Tel():11:" & strTel
		case 10
			Debug ".Tel():10:" & strTel
			select case left(strTel,2)
			case "03","06"
				Debug ".Tel():10-2:" & strTel
				' 03-3456-7890
				strTel1 = left(strTel,2)
				strTel2 = mid(strTel,3,4)
				strTel3 = right(strTel,4)
				strTel = strTel1 & "-" & strTel2 & "-" & strTel3
				Debug ".Tel():10-2:" & strTel
			case else
				select case left(strTel,3)
				case "028","045","046","048","078","080"
					'028-688-8168
					Debug ".Tel():10-3:" & strTel
					strTel1 = left(strTel,3)
					strTel2 = mid(strTel,5,3)
					strTel3 = right(strTel,4)
					strTel = strTel1 & "-" & strTel2 & "-" & strTel3
					Debug ".Tel():10-3:" & strTel
				case else
					select case left(strTel,4)
					case "0258"
						'0258-42-2211
						Debug ".Tel():10-4:" & strTel
						strTel1 = left(strTel,4)
						strTel2 = mid(strTel,5,2)
						strTel3 = right(strTel,4)
						strTel = strTel1 & "-" & strTel2 & "-" & strTel3
						Debug ".Tel():10-4:" & strTel
					case else
					end select
				end select
			end select
		end select
		Tel = strTel
	End Function
	'-----------------------------------------------------------------------
	'R-smile用(CSV)
	'-----------------------------------------------------------------------
	Private Function RsCsv(byVal strName)
		dim	strComma
		strComma = ","
		dim	strValue
		strValue = ""
		select case strName
		case "届先"
			strValue = GetField("SCode")
			strComma = ""
		case "電話番号"	
			strValue = Tel(GetField("STel"))
		case "郵便番号"	
			strValue = GetField("SZip")
		case "都道府県"	
		case "市区町村"	
		case "町域"		
			strValue = GetField("SAddress")
		case "番地"		
			strValue = GetField("SBillding")
		case "名称１"	
			strValue = GetField("SName1")
		case "名称２"	
			strValue = GetField("SName2")
'		case "敬称"		
'		case "個数"		
'		case "重量"		
		case "荷送人"	
			strValue = RsSender()
		case "航空区分"	
			select case RsSender()
			case 7,8
				strValue = "ＡＩＲ"
			end select
		case "配達日"	
			strValue = YoteiDt()
'		case "配達時間"	
		case "記事1"	
			strValue = GetField("SKiji1")
		case "記事2"	
			strValue = GetField("SKiji2")
		case "記事3"	
			strValue = GetField("SKiji3")
		case "記事4"	
			strValue = GetField("SKiji4")
		case "記事5"	
			strValue = GetField("SKiji5")
		case "出荷日"	
			strValue = GetField("SyukaDt")
		case "出荷番号"	
			strValue = GetField("ID") & "0"
		case "伝票番号"	
		end select
		WScript.StdOut.Write strComma
		WScript.StdOut.Write Replace(strValue,",",".")
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
	'Make() 送り状データ作成
	'-----------------------------------------------------------------------
	Private	strPrgId
    Public Function Make()
		Debug ".Make()"
		SetSql	""
		if strPrgId = "test" then
			Disp "テストデータ作成"
			Call objDb.Execute("delete from RSmile where EntID = 'test'")
			SetSql	"select"
			SetSql	"distinct"
			SetSql	"Min(IdNo) IdNo"
			SetSql	",Left(Replace(Convert(Now(),sql_char),'-',''),8) SyukaDt"
			SetSql	",d.ChoCode ChoCode"
			SetSql	",d.ChoName ChoName"
			SetSql	",d.ChoAddress ChoAddress"
			SetSql	",d.ChoTel ChoTel"
			SetSql	",d.ChoZip ChoZip"
			SetSql	",'' Id"
			SetSql	"from HtDrctId d"
			SetSql	"where RTrim(ChoCode) <> ''"
			SetSql	"group by"
			SetSql	"SyukaDt"
			SetSql	",ChoCode"
			SetSql	",ChoName"
			SetSql	",ChoCode"
			SetSql	",ChoAddress"
			SetSql	",ChoTel"
			SetSql	",ChoZip"
			SetSql	"order by"
			SetSql	"ChoCode"
		else
			SetSql	"select"
			SetSql	"distinct"
			SetSql	"y.KEY_SYUKA_YMD SyukaDt"
			SetSql	",d.ChoCode ChoCode"
			SetSql	",d.ChoName ChoName"
			SetSql	",d.ChoAddress ChoAddress"
			SetSql	",d.ChoTel ChoTel"
			SetSql	",d.ChoZip ChoZip"
			SetSql	",r.Id Id"
			SetSql	"from y_syuka y"
			SetSql	"inner join HtDrctId d on (d.IDNo = y.KEY_ID_NO)"
			SetSql	"left outer join RSmile r on (y.KEY_SYUKA_YMD = r.SyukaDt and d.ChoCode = r.SCode)"
			SetSql	"order by"
			SetSql	" SyukaDt"
			SetSql	",ChoCode"
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
	'送り先ごとの件数カウント
	'-------------------------------------------------------------------
	Private	strSyukaDt
	Private	strChoCode
	Private	strChoName
	Private	strClientNo
	Private	strSCode
	Private	intCnt
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData()"
		WScript.StdOut.Write GetField("SyukaDt")
		if strPrgId = "test" then
			WScript.StdOut.Write " " & GetField("IdNo")
		end if
'		WScript.StdOut.Write " " & GetField("Biko1")
		WScript.StdOut.Write " " & Left(GetField("ChoCode") & Space(9),9)
		WScript.StdOut.Write Get_LeftB(GetField("ChoName") & Space(40),40)
'		WScript.StdOut.Write " " & GetField("ChoAddress")
'		WScript.StdOut.Write " " & GetField("ChoTel")
'		WScript.StdOut.Write " " & GetField("ChoZip")
		Call GetKiji()
		Call InsertData()
		WScript.StdOut.WriteLine
	End Function
	'-------------------------------------------------------------------
	'Kiji1-5
	'-------------------------------------------------------------------
	Private	intDenCnt
	Private	strId
	Private	strKiji1
	Private	strKiji2
	Private	strKiji3
	Private	strKiji4
	Private	strKiji5
	Private Function GetKiji()
		Debug ".GetKiji()"
		if strPrgId = "test" then
			strId = GetField("IdNo")
			strKiji1 = String(30,"テ")
			strKiji2 = String(30,"ス")
			strKiji3 = String(30,"ト")
			strKiji4 = String(30,"デ")
			strKiji5 = String(30,"ス")
			exit function
		end if
		SetSql	""
		SetSql	"select"
		SetSql	"distinct"
		SetSql	"y.KEY_SYUKA_YMD SyukaDt"
		SetSql	",y.KEY_ID_NO Id"
		SetSql	",y.KEY_HIN_NO HIN_GAI"
		SetSql	",convert(y.SURYO,sql_decimal) Qty"
		SetSql	",y.Bikou1 Biko1"
		SetSql	"from y_syuka y"
		SetSql	"inner join HtDrctId d on (d.IDNo = y.KEY_ID_NO)"
		SetSql	"where y.KEY_SYUKA_YMD='" & GetField("SyukaDt") & "'"
		SetSql	"and d.ChoCode='" & GetField("ChoCode") & "'"
		SetSql	"order by y.KEY_ID_NO"
		Debug strSql
		dim	objKiji
		set objKiji = objDB.Execute(strSql)
		strId		= ""
		strKiji1	= ""
		strKiji2	= ""
		strKiji3	= ""
		strKiji4	= ""
		strKiji5	= ""
		intDenCnt	= 0
		do while objKiji.EOF = False
			intDenCnt = intDenCnt + 1
			Debug ".GetKiji():" & intDenCnt & " " & objKiji.Fields("SyukaDt") & " " & objKiji.Fields("Id") & " " & objKiji.Fields("HIN_GAI") & " " & objKiji.Fields("Qty")
			select case intDenCnt
			case 1:
				strId = RTrim(objKiji.Fields("Id"))
				strKiji1 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "個"
				strKiji5 = RTrim(objKiji.Fields("Biko1"))
			case 2:
				strKiji2 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "個"
			case 3:
				strKiji3 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "個"
			case 4:
				strKiji4 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "個"
			case 5:
				strKiji4 = strKiji4 & " 他"
			case else
			end select
			objKiji.MoveNext
		loop
		strKiji5 = strKiji5 & Space(5) & "伝票：" & intDenCnt & "枚"
		objKiji.Close
		set objKiji = Nothing
	End Function
	'-------------------------------------------------------------------
	'RSmile Update
	'-------------------------------------------------------------------
	Private	Function Update()
		Debug ".Update()"
		SetSql ""
		SetSql "update RSmile "
		SetSql "set SyukaDt = '" & GetField("SyukaDt") & "'"
		SetSql ",SCode = '" & GetField("ChoCode") & "'"
		SetSql ",STel = '" & GetField("ChoTel") & "'"
		SetSql ",SZip = '" & GetField("ChoZip") & "'"
		SetSql ",SAddress = '" & GetField("ChoAddress") & "'"
'		SetSql ",SBillding"
		SetSql ",SName1 = '" & GetField("ChoName") & "'"
'		SetSql ",SName2"
		SetSql ",SKiji1 = '" & strKiji1 & "'"
		SetSql ",SKiji2 = '" & strKiji2 & "'"
		SetSql ",SKiji3 = '" & strKiji3 & "'"
		SetSql ",SKiji4 = '" & strKiji4 & "'"
		SetSql ",SKiji5 = '" & strKiji5 & "'"
		SetSql ",UpdID = '" & strPrgId & "'"
		SetSql "where Id = '" & GetField("Id") & "'"
		SetSql "and ( SyukaDt <> '" & GetField("SyukaDt") & "'"
		SetSql "or SCode <> '" & GetField("ChoCode") & "'"
		SetSql "or STel <> '" & GetField("ChoTel") & "'"
		SetSql "or SZip <> '" & GetField("ChoZip") & "'"
		SetSql "or SAddress <> '" & GetField("ChoAddress") & "'"
		SetSql "or SName1 <> '" & GetField("ChoName") & "'"
		SetSql "or SKiji1 <> '" & strKiji1 & "'"
		SetSql "or SKiji2 <> '" & strKiji2 & "'"
		SetSql "or SKiji3 <> '" & strKiji3 & "'"
		SetSql "or SKiji4 <> '" & strKiji4 & "'"
		SetSql "or SKiji5 <> '" & strKiji5 & "'"
		SetSql ")"
		on error resume next
		objDb.Execute strSql
		WScript.StdOut.Write ":0x" & Hex(Err.Number) & " " & Err.Description
		on error goto 0
	End Function
	'-------------------------------------------------------------------
	'RSmile insert
	'-------------------------------------------------------------------
	Private Function InsertData()
		Debug ".InsertData()"
		if GetField("Id") <> "" then
			WScript.StdOut.Write ":" & GetField("Id")
			Update
			exit function
		end if
		SetSql ""
		SetSql "insert into RSmile ("
		SetSql "Id"
		SetSql ",SyukaDt"
		SetSql ",SCode"
		SetSql ",STel"
		SetSql ",SZip"
		SetSql ",SAddress"
		SetSql ",SBillding"
		SetSql ",SName1"
		SetSql ",SName2"
		SetSql ",SKiji1"
		SetSql ",SKiji2"
		SetSql ",SKiji3"
		SetSql ",SKiji4"
		SetSql ",SKiji5"
		SetSql ",EntID"
		SetSql ") values ("
		SetSql "'" & strId & "'"
		SetSql ",'" & GetField("SyukaDt") & "'"
		SetSql ",'" & GetField("ChoCode") & "'"
		SetSql ",'" & GetField("ChoTel") & "'"
		SetSql ",'" & GetField("ChoZip") & "'"
		SetSql ",'" & GetField("ChoAddress") & "'"
		SetSql ",''"		'SBillding"
		SetSql ",'" & GetField("ChoName") & "'"
		SetSql ",''"		'SName2"
		SetSql ",'" & strKiji1 & "'"	    'SKiji1"
		SetSql ",'" & strKiji2 & "'"	    'SKiji2"
		SetSql ",'" & strKiji3 & "'"	    'SKiji3"
		SetSql ",'" & strKiji4 & "'"	    'SKiji4"
		SetSql ",'" & strKiji5 & "'"	    'SKiji5"
		SetSql ",'" & strPrgId & "'"
		SetSql ")"
		Debug strSql
		on error resume next
		objDb.Execute strSql
		WScript.StdOut.Write ":0x" & Hex(Err.Number) & " " & Err.Description
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
			WScript.Echo "GetField(" & strName & "):Error:0x" & Hex(Err.Number)
			WScript.Quit
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
	dim	objRSmile
	Set objRSmile = New RSmile
	if objRSmile.Init() <> "" then
		call usage()
		exit function
	end if
	call objRSmile.Run()
End Function
