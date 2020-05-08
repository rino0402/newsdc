Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "csvconv.vbs [option]"
	Wscript.Echo " /db:newsdc9	データベース"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript csvconv.vbs /db:newsdc9 pop3w9\tmp\在庫_棚番1.csv"
	Wscript.Echo "cscript csvconv.vbs /db:newsdc9 BoSyukaDet.csv"
End Sub
'-----------------------------------------------------------------------
'BoCnv
'-----------------------------------------------------------------------
Class Csv
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strPathName
	Private	strFileName
	Private	strDT
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
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
			case "i"
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
		strDBName	= GetOption("db"	,"newsdc")
		strDT		= year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2)
		set objDB	= nothing
		set objRs	= nothing
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
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			Debug ".Run():" & strArg
			strPathName = strArg
			strFileName = GetFileName(strPathName)
			Call Conv()
		Next
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Private Function Conv()
		Debug ".Conv():" & strPathName
		select case CsvType()
		case "BoZaiko"
			Call BoZaiko()
		case "BoZaikoName"
			Call BoZaiko()
		case "BoSyukaDet"
			Call BoSyukaDet()
		end select
	End Function
	'-----------------------------------------------------------------------
	'BoSyukaDet()
	'-----------------------------------------------------------------------
	Private	strYm1
	Private	strYm2
    Private Function BoSyukaDet()
		Debug ".BoSyukaDet()"
		Wscript.StdOut.WriteLine "ファイル名:" & strFileName
		Wscript.StdOut.WriteLine "      形式:" & strCsvType
		BoSyukaDet_Btwn
		Wscript.StdOut.WriteLine "      年月:" & strYm1 & "-" & strYm2
		BoSyukaDet_Del
		Wscript.StdOut.WriteLine "      削除:" & RowCount()
		BoSyukaDet_Ins
		Wscript.StdOut.WriteLine "      追加:" & RowCount()
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_Del()
	'-----------------------------------------------------------------------
    Private Function BoSyukaDet_Btwn()
		Debug ".BoSyukaDet_Btwn()"
		AddSql ""
		AddSql "select"
		AddSql " Min(Left(if(Col=11,Col08,Col06),6)) Ym1"
		AddSql ",Max(Left(if(Col=11,Col08,Col06),6)) Ym2"
		AddSql "from CsvTemp"
		AddSql "where FileName = '" & strFileName & "'"
		AddSql "and Row>1"
		AddSql "and Col01 not like '%管理番号'"
'		AddSql "and Col=11"
		CallSql
		strYm1 = ""
		strYm2 = ""
		do while objRs.Eof = False
			strYm1 = RTrim(objRs.Fields("Ym1"))
			strYm2 = RTrim(objRs.Fields("Ym2"))
			exit do
		loop
	End Function
	'-----------------------------------------------------------------------
	'BoSyukaDet_Del()
	'-----------------------------------------------------------------------
    Private Function BoSyukaDet_Del()
		Debug ".BoSyukaDet_Del()"
		AddSql ""
		AddSql "delete from BoSyukaDet where Left(JisekiDt,6) between '" & strYm1 & "' and '" & strYm2 & "'"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoSyukaDet_Ins()
	'-----------------------------------------------------------------------
    Private Function BoSyukaDet_Ins()
		Debug ".BoSyukaDet_Ins()"
		AddSql ""
		AddSql "insert into BoSyukaDet"
		AddSql "("
'		AddSql "No"			'//NO 1
		AddSql " IDNo"		'//受注出荷管理番号 700089663241	
		AddSql ",JCode"		'//事業場CD 00021184
		AddSql ",Syushi"	'//在庫収支 34"
		AddSql ",DenNo"		'//伝票番号 032293
		AddSql ",SyukaCd"	'//出荷先CD 113A
		AddSql ",SyukaNm"	'//出荷先名 MALAYSIA(PM)
		AddSql ",AiteCd"	'//相手先CD 
		AddSql ",AiteNm"	'//相手先名 その他
		AddSql ",ChuKb"		'//注文区分 6:AIR週切（13日）
		AddSql ",JisekiDt"	'//売上実績年月日 20150401"
		AddSql ",Pn"		'//品目番号 A390C7R30WT
		AddSql ",Qty"		'//出荷実績数 2
		AddSql ") select distinct "
		AddSql " RTrim(Col01)"						'受注出荷過実_受注出荷管理番号	700093657489
		AddSql ",RTrim(Col02)"						'受注出荷過実_資産管理事業場コード	00021529
		AddSql ",RTrim(if(Col=11,Col03,Col04))"		'受注出荷過実_在庫収支コード	11D
		AddSql ",RTrim(if(Col=11,Col04,Col05))"		'受注出荷過実_伝票番号	027243
		AddSql ",RTrim(if(Col=11,Col05,Col07))"		'受注出荷過実_得意先コード(相手先CD)	00020162
		AddSql ",RTrim(if(Col=11,Col06,Col08))"		'受注出荷過実_得意先略称(相手先名)	アプライアンス社　本社
		AddSql ",RTrim(if(Col=11,Col11,Col13))"		'受注出荷過実_直送相手先コード	00020162
		AddSql ",''"
		AddSql ",RTrim(if(Col=11,Col07,Col12))"		'受注出荷過実_注文区分	2
		AddSql ",RTrim(if(Col=11,Col08,Col06))"		'受注出荷過実_売上実績年月日	20170119
		AddSql ",RTrim(if(Col=11,Col09,Col03))"		'受注出荷過実_品目番号	ANP300-1530
		AddSql ",Convert(RTrim(if(Col=11,Col10,Col09)),sql_decimal)"	'受注出荷過実_出荷実績数	4
		AddSql "from CsvTemp"
		AddSql "where FileName = '" & strFileName & "'"
		AddSql "and Row > 1"
		AddSql "and Col01 not like '%管理番号'"
'		AddSql "and Col = 11"
'1受注出実_受注出荷管理番号	700093755635
'2受注出実_資産管理事業場コード	00023100
'3受注出実_品目番号	AXW22B-7EM0
'4受注出実_在庫収支コード	11D
'5受注出実_伝票番号	004781
'6受注出実_売上実績年月日	20170207
'7受注出実_得意先コード(相手先コード)	00020162
'8受注出実_得意先略称(相手先名)	アプライアンス社　本社
'9受注出実_出荷実績数	2
'10受注出実_倉庫コード	NAR
'11受注出実_注文区分	2
'12受注出実_入出庫取引区分	20
'13受注出実_直送相手先コード	00020162
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko()
	'-----------------------------------------------------------------------
	Private Function BoZaiko()
		Debug ".BoZaiko()"
		Wscript.StdOut.WriteLine "ファイル名:" & strFileName
		Wscript.StdOut.WriteLine "      形式:" & strCsvType

		Wscript.StdOut.Write "        BoZaiko_Del:"
		BoZaiko_Del
		Wscript.StdOut.WriteLine RowCount()

		select case strCsvType
		case "BoZaiko"
			Wscript.StdOut.Write "        BoZaiko_Ins:"
			BoZaiko_Ins
			Wscript.StdOut.WriteLine RowCount()
		case "BoZaikoName"
			Wscript.StdOut.Write "    BoZaikoName_Ins:"
			BoZaikoName_Ins
			Wscript.StdOut.WriteLine RowCount()
		end select

'		Wscript.StdOut.Write "ZaikoH " & strDT & " Del:"
'		BoZaiko_ZaikoH_Del
'		Wscript.StdOut.WriteLine RowCount()

'		Wscript.StdOut.Write "ZaikoH " & strDT & " Ins:"
'		BoZaiko_ZaikoH
'		Wscript.StdOut.WriteLine RowCount()

'		Wscript.StdOut.Write "            NarTana:"
'		BoZaiko_NarTana()
'		Wscript.StdOut.WriteLine RowCount()
	End Function
	'-----------------------------------------------------------------------
	'RowCount()
	'-----------------------------------------------------------------------
    Private Function RowCount()
		Debug ".RowCount()"
		dim	objRow
		set	objRow = objDb.Execute("select @@rowcount")
		RowCount = objRow.Fields(0)
		set	objRow = Nothing
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_Del()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_Del()
		Debug ".BoZaiko_Del()"
'		Disp "BoZaiko:delete all"
		AddSql ""
		AddSql "delete from BoZaiko"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_Ins()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_Ins()
		Debug ".BoZaiko_Ins()"
'		Disp "BoZaiko:Insert " & strFileName
		AddSql	""
		AddSql	"insert into BoZaiko"
		AddSql	"(Soko"
		AddSql	",JCode"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",SyuShi"
		AddSql	",TanaQty"
		AddSql	",HikiQty"
		AddSql	",SyuShiR"
		AddSql	",SyuShiName"
		AddSql	",Loc1"
		AddSql	") select "
		AddSql	"distinct"
		AddSql	" RTrim(Col01)"				'// 在庫収支_倉庫コード
		AddSql	",RTrim(Col02)"				'// ＰＮ倉庫在庫_事業場コード
		AddSql	",RTrim(Col03)"				'// ＰＮ倉庫在庫_資産管理事業場コード
		AddSql	",RTrim(Col04)"				'// ＰＮ倉庫在庫_品目番号
		AddSql	",RTrim(Col05)"				'// ＰＮ倉庫在庫_在庫収支コード
		AddSql	",Convert(Col06,Sql_Decimal)"	'// ＰＮ倉庫在庫_棚在庫数
		AddSql	",Convert(Col07,Sql_Decimal)"	'// ＰＮ倉庫在庫_正味引当可能在庫数
		AddSql	",RTrim(Col08)"				'// 在庫収支_在庫収支略式名
		AddSql	",Max(RTrim(Col10))"			'// 在庫収支_在庫収支名
		AddSql	",RTrim(Col09)"				'// 棚番_１
		AddSql	"from CsvTemp"  
		AddSql	"where FileName = '" & strFileName & "'"
		AddSql	"and Row > 1"
		AddSql	"and RTrim(Col01) = 'NAR'"
		AddSql	"group by"
		AddSql	"Col01"
		AddSql	",Col02"
		AddSql	",Col03"
		AddSql	",Col04"
		AddSql	",Col05"
		AddSql	",Col06"
		AddSql	",Col07"
		AddSql	",Col08"
		AddSql	",Col09"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaikoName_Ins()
	'-----------------------------------------------------------------------
    Private Function BoZaikoName_Ins()
		Debug ".BoZaikoName_Ins()"
		AddSql	""
		AddSql	"insert into BoZaiko"
		AddSql	"(Soko"
		AddSql	",JCode"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",SyuShi"
		AddSql	",TanaQty"
		AddSql	",HikiQty"
		AddSql	",SyuShiR"
		AddSql	",Loc1"
		AddSql	",SyuShiName"
		AddSql	",PName"
		AddSql	",PNameEng"
		AddSql	") select "
		AddSql	"distinct"
		AddSql	" RTrim(Col01)	Soko"			'PN倉庫（収支)_倉庫コード
		AddSql	",RTrim(Col02)	JCode"			'PN倉庫（PN棚）_事業場コード
		AddSql	",RTrim(Col03)	ShisanJCode"	'PN倉庫（PN棚）_資産管理事業場コード
		AddSql	",RTrim(Col04)	Pn"				'PN倉庫（PN棚）_品目番号
		AddSql	",RTrim(Col06)	SyuShi"			'PN倉庫（PN棚)_在庫収支コード
		AddSql	",Convert(Col07,Sql_Decimal)	TanaQty"	'PN倉庫（PN棚)_棚在庫数　※分析
		AddSql	",Convert(Col08,Sql_Decimal)	HikiQty"'	'PN倉庫（PN棚)_正味引当可能在庫数
		AddSql	",RTrim(Col09)	SyuShiR"		'PN倉庫（PN棚）_在庫収支略式名
		AddSql	",RTrim(Col10)	Loc1"			'PN倉庫（PN棚）_ロケーション番号１
		AddSql	",Max(RTrim(Col11))	SyuShiName"	'在庫収支_在庫収支名'
		AddSql	",Max(RTrim(Col05))	PName"		'ＰＮ名称(JPN)_品目名
		AddSql	",Max(RTrim(Col12)) PNameEng"	'ＰＮ名称_品目別名(ENG)'
		AddSql	"from CsvTemp"
		AddSql	"where FileName = '" & strFileName & "'"
		AddSql	"and Row > 1"
		AddSql	"group by"
		AddSql	" Soko"
		AddSql	",JCode"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",SyuShi"
		AddSql	",TanaQty"
		AddSql	",HikiQty"
		AddSql	",SyuShiR"
		AddSql	",Loc1"
'		AddSql	",SyuShiName"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_ZaikoH()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_ZaikoH_Del()
		Debug ".BoZaiko_ZaikoH_Del()"
'		Disp "BoZaiko:ZaikoH delete " & strDT
		AddSql	""
		AddSql	"delete from ZaikoH"
		AddSql	"where Kubun = 'Bo'"
'		AddSql	"and DT = left(replace(convert(now(),sql_char),'-',''),8)"
		AddSql	"and DT = '" & strDT & "'"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_ZaikoH()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_ZaikoH()
		Debug ".BoZaiko_ZaikoH()"
'		Disp "BoZaiko:ZaikoH insert " & strDT
		AddSql	""
		AddSql	"insert into ZaikoH"
		AddSql	"(Kubun"
		AddSql	",DT"
		AddSql	",JCode"
		AddSql	",Pn"
		AddSql	",Syushi"
		AddSql	",Qty"
		AddSql	",QtyHiki"
		AddSql	",Loc1"
		AddSql	") select"
		AddSql	"'Bo'"
		AddSql	",'" & strDT & "'"
'		AddSql	",left(replace(convert(now(),sql_char),'-',''),8)"
		AddSql	",ShisanJCode"
		AddSql	",Pn"	'					"品番"
		AddSql	",SyuShi"	'				"収支"
		AddSql	",TanaQty"	'				"棚在庫数"
		AddSql	",HikiQty"	'				"引当可能在庫数"
		AddSql	",Loc1"	'					"棚番_１"
		AddSql	"from BoZaiko"
'		AddSql	"where Soko='NAR'"
		AddSql 	"where (Soko = 'NAR' or (Soko = 'NA2' and Left(Loc1,1) = 'E'))"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_NarTana()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_NarTana()
		Debug ".BoZaiko_NarTana()"

'		Disp "BoZaiko:NarTana delete in BoZaiko"
		AddSql	""
		AddSql	"delete from NarTana"
		AddSql	"where RTrim(Soko)+RTrim(ShisanJCode)+RTrim(Pn)"
		AddSql	"in (select distinct RTrim(Soko)+RTrim(ShisanJCode)+RTrim(Pn)"
		AddSql	"from BoZaiko"
'		AddSql	"where Soko='NAR'"
		AddSql 	"where (Soko = 'NAR' or (Soko = 'NA2' and Left(Loc1,1) = 'E'))"
		AddSql	"and left(SyuShi,2) in ('11','12','10','41','71','99','15')"
		AddSql	"and Loc1<>''"
		AddSql	")"
		CallSql

'		Disp "BoZaiko:NarTana insert in BoZaiko"
		AddSql	""
		AddSql	"insert into NarTana"
		AddSql	"(Soko"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",Loc1"
		AddSql	",Loc10"
		AddSql	",Loc11"
		AddSql	",Loc12"
		AddSql	",Loc41"
		AddSql	",Loc71"
		AddSql	",Loc99"
		AddSql	",Loc15"
		AddSql	",EntID"
		AddSql	")"
		AddSql	"select"
		AddSql	"distinct"
		AddSql	" Soko"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",Max(Loc1)"
		AddSql	",Max(if(left(SyuShi,2)='10',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='11',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='12',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='41',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='71',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='99',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='15',Loc1,''))"
		AddSql	",'BoZaiko'"
		AddSql	"from BoZaiko"
		AddSql 	"where (Soko = 'NAR' or (Soko = 'NA2' and Left(Loc1,1) = 'E'))"
		AddSql	"and left(SyuShi,2) in ('11','12','10','41','71','99','15')"
		AddSql	"and Loc1<>''"
		AddSql	"group by"
		AddSql	" Soko"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		CallSql

'		Disp "BoZaiko:NarTana update Loc11"
		AddSql	""
		AddSql	"update NarTana"
		AddSql	"set Loc1=Loc11"
		AddSql	",	UpdID='Loc11'"
		AddSql	",	UpdTm=Now()"
'		AddSql	"where Soko='NAR'"
		AddSql	"where Loc11<>''"
		AddSql	"and Loc1<>Loc11"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'CsvType()
	'-----------------------------------------------------------------------
	Private	strCsvType
    Private Function CsvType()
		Debug ".CsvType()"
		strCsvType = ""
		AddSql	""
		AddSql	"select"
		AddSql	"y.CsvType cType"
'		AddSql	"from CsvType y"
'		AddSql	"inner join CsvTemp t"
		AddSql	"from CsvTemp t"
		AddSql	"inner join CsvType y"
		AddSql	"on (y.Col = t.Col"
		AddSql	"and y.Col01 = t.Col01"
		AddSql	"and y.Col02 = t.Col02"
		AddSql	"and y.Col03 = t.Col03"
		AddSql	"and y.Col04 = t.Col04"
		AddSql	"and y.Col05 = t.Col05"
		AddSql	"and y.Col06 = t.Col06"
		AddSql	"and y.Col07 = t.Col07"
		AddSql	"and y.Col08 = t.Col08"
		AddSql	"and y.Col09 = t.Col09"
		AddSql	"and y.Col10 = t.Col10"
		AddSql	"and y.Col11 = t.Col11"
		AddSql	"and y.Col12 = t.Col12"
		AddSql	")"
		AddSql	"where t.FileName = '" & strFileName & "'"
		AddSql	"and t.Row = 1"
		Wscript.StdErr.Write strFileName & ":"
		CallSql
		if objRs.Eof = False then
			Debug ".CsvType():" & objRs.Fields("cType")
			strCsvType = RTrim(objRs.Fields("cType"))
		end if
		Wscript.StdErr.WriteLine strCsvType
		CsvType = strCsvType
	End Function
	'-------------------------------------------------------------------
	'ファイル名(パスを除く)
	'-------------------------------------------------------------------
	Private Function GetFileName(byVal f)
		dim	objFileSys
		Set objFileSys	= WScript.CreateObject("Scripting.FileSystemObject")

		dim	strFName
		strFName = objFileSys.GetBaseName(f)
		strFName = strFName & "."
		strFName = strFName & objFileSys.GetextensionName(f)
		GetFileName	= strFName

		Set objFileSys	= Nothing
	End Function
	'-------------------------------------------------------------------
	'文字列追加 strSql
	'-------------------------------------------------------------------
	dim	strSql
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
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
	Private Function CallSql()
		Debug ".CallSql():" & strSql
'		on error resume next
		Set objRs = objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field名
	'-------------------------------------------------------------------
	Private Function GetFields(byVal strTable)
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
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Private Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Private Sub Disp(byVal strMsg)
		Wscript.StdErr.WriteLine strMsg
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
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objCsv
	Set objCsv = New Csv
	if objCsv.Init() <> "" then
		call usage()
		exit function
	end if
	call objCsv.Run()
End Function
