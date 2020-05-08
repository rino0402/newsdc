Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objHMEM500
	Set objHMEM500 = New HMEM500
	objHMEM500.Run
	Set objHMEM500 = nothing
End Function
'-----------------------------------------------------------------------
'HMEM500
'-----------------------------------------------------------------------
Class HMEM500
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "hmem500.vbs [option] [filename]"
		Echo "Ex."
		Echo "cscript//nologo hmem500.vbs /db:newsdc1 hmem508szz.dat.20170804-133832.3397"
		Echo "cscript//nologo hmem500.vbs /db:newsdc4 hmem507szz.dat.20170922-185045.19583"
		Echo "cscript//nologo hmem500.vbs /db:newsdc4 hmem507szz.dat.20170922-185045.19583 /list:1 ※入荷リスト"
		Echo "cscript//nologo hmem500.vbs /db:newsdc4 hmem507szz.dat.20170922-185045.19583 /Z:90010101"
		Echo "cscript//nologo hmem500.vbs /db:newsdc4 hmem507szz.dat.20170922-185045.19583 /item:n"
		Echo "cscript//nologo hmem500.vbs /db:newsdc4 hmem507szz.dat.20170922-185045.19583 /y_nyuka"
'hmem507szz.dat.20171017-093734.25457
'hmem507szz.dat.20171017-110115.14558
'hmem507szz.dat.20171017-123752.24415
		Echo "Option."
		Echo "   DBName:" & strDBName
		Echo "    Table:" & strTable
		Echo " FileName:" & strFileName
		Echo "       Dt:" & strDt
		Echo "    Zaiko:" & strZaiko
		Echo "     Item:" & strItem
	End Sub
	Private	objDB
	Private	strDBName
	Private	strTable
	Private	strFileName
	Private	strDt
	Private	strZaiko
	Private	strItem
	Private	strList
	Private	strYNyuka
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName	= GetOption("db"	,"newsdc")
		strFileName = ""
		strDt		= GetOption("dt"	,"")
		strTable	= GetOption("table"	,"hmem500")
		strList		= GetOption("list"	,"")
		strItem		= GetOption("item"	,"")
		strZaiko	= GetOption("z"		,"")
		strYNyuka	= GetOption("y_nyuka"		,"")
		set objDB	= nothing
	End Sub
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Private Function Init()
		Debug ".Init()"
		dim	strArg
		Init = False
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Echo "オプションエラー:" & strArg
				Usage
				Exit Function
			end if
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
				strDBName	= GetOption(strArg,strDBName)
			case "table"
				strTable	= GetOption(strArg,strList)
			case "dt"
				strDt		= GetOption(strArg,strDt)
			case "debug"
			case "list"
				strList		= GetOption(strArg,strList)
			case "item"
				strItem		= GetOption(strArg,strItem)
			case "y_nyuka"
				strYNyuka	= strArg
			case "z"
				strZaiko	= GetOption(strArg,strZaiko)
			case "?"
				Usage
				Exit Function
			case else
				Echo "オプションエラー:" & strArg
				Usage
				Exit Function
			end select
		Next
		Init = True
	End Function
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		if Init() = True then
			OpenDb
			if strItem <> "" then
				Item
			elseif strYNyuka <> "" then
				YNyuka
			elseif strZaiko <> "" then
				Zaiko
			else
				select case strList
				case "1"
					List1
				case "0"
					List0
				case else
					List2
				end select
			end if
			CloseDb
		end if
	End Function
	'-----------------------------------------------------------------------
	'YNyuka() 入荷データ登録
	'-----------------------------------------------------------------------
    Private Function YNyuka()
		Debug ".YNyuka()"
		AddSql ""
		AddSql "select"
		AddSql "*"
		AddSql "from hmem500"
		AddSql "where Right(RTrim(SyukoCd),2) <> RTrim(SyushiCd)"
		AddSql "and convert(Qty,sql_decimal) > 0"
		AddWhere "Filename",strFileName
		AddWhere "DenDt",strDt
		AddWhere "IoKbn","1"
		if strZaiko = "" then
			strZaiko = "90010101"
		else
			'SJ
			AddWhere "SyushiCd","SJ"
		end if
		CallSql strSql
		Call GroupHead(-1)
		do while objRs.Eof = False
			YNyukaInsert
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'YNyukaInsert() 入荷データInsert
	'-----------------------------------------------------------------------
    Private Function YNyukaInsert()
		Debug ".YNyukaInsert()"
		Write objRs.Fields("JGyobu")	,2
		Write objRs.Fields("DenDt")		,9
		Write objRs.Fields("IoKbn")		,0
		Write objRs.Fields("AkaKuro")	,0
		Write objRs.Fields("SyoriMD")	,0
		Write objRs.Fields("Bin")		,0
		Write objRs.Fields("SeqNo")		,5
		Write objRs.Fields("Pn")		,15
		Write objRs.Fields("Qty")		,7
		Write objRs.Fields("SyukoCd")	,6
		Write objRs.Fields("NyukoCd")	,4
		Write objRs.Fields("SyushiCd")	,3
		Write objRs.Fields("SHIIRE_WORK_CENTER")	,8
		Write objRs.Fields("EOR")	,1
		AddSql ""
		AddSql "insert into Y_NYUKA"
		AddSql "("
		AddSql " KAN_KBN			"'  1)	//完了区分
		AddSql ",DT_SYU				"'  1)	//データ種別
		AddSql ",JGYOBU				"'  1)	//事業部区分	key0
		AddSql ",NAIGAI				"'  1)	//国内外
		AddSql ",TEXT_NO			"'  9)	//テキスト№	key0
		AddSql ",JGYOBA				"'  8)	//事業場ｺｰﾄﾞ
		AddSql ",DATA_KBN			"'  1)	//データ区分
		AddSql ",TORI_KBN			"'  2)	//取引区分
		AddSql ",ID_NO				"' 12)	//ID-NO
		AddSql ",KAIKEI_JGYOBA		"'  8)	//会計用事業場ｺｰﾄﾞ
		AddSql ",SHISAN_JGYOBA		"'  8)	//資産管理用事業場ｺｰﾄﾞ
		AddSql ",HIN_NO				"' 20)	//品目番号
		AddSql ",DEN_NO				"' 10)	//伝票番号
		AddSql ",SURYO				"'  7)	//出荷数量
		AddSql ",MUKE_CODE			"'  8)	//得意先コード
		AddSql ",SYUKO_SYUSI		"'  8)	//在庫収支
		AddSql ",SHISAN_SYUSI		"'  8)	//資産管理用在庫収支ｺｰﾄﾞ
		AddSql ",HOJYO_SYUSI		"'  8)	//補助在庫収支ｺｰﾄﾞ
		AddSql ",SYUKO_YMD			"'  8)	//出庫日付
		AddSql ",TANKA				"' 10)	//実際単価
		AddSql ",ODER_NO			"' 12)	//オーダー番号
		AddSql ",ITEM_NO			"'  5)	//アイテム番号
		AddSql ",ODER_NO_R			"'  5)	//注文管理番号略号
		AddSql ",KOSO_KEITAI		"' 10)	//個装形態ｺｰﾄﾞ
		AddSql ",SYUKA_YMD			"'  8)	//出荷予定日	key0
		AddSql ",TANABAN1			"' 10)	//ﾛｹｰｼｮﾝ1
		AddSql ",TANABAN2			"' 10)	//ﾛｹｰｼｮﾝ2
		AddSql ",TANABAN3			"' 10)	//ﾛｹｰｼｮﾝ3
		AddSql ",MUKE_NAME			"' 24)	//得意先名称
		AddSql ",CYU_KBN			"'  1)	//注文区分
		AddSql ",CYU_KBN_NAME		"' 10)	//注文区分名称
		AddSql ",ORIGIN1			"' 10)	//原産国1
		AddSql ",ORIGIN2			"' 10)	//原産国2
		AddSql ",BIKOU2				"' 40)	//備考2
		AddSql ",HAN_KBN			"'  1)	//販売区分
		AddSql ",CHOKU_KBN			"'  1)	//直送指示区分
		AddSql ",UNIT_ID_NO			"' 12)	//ﾕﾆｯﾄ修正管理番号
		AddSql ",ZAIKO_HIKIATE		"'  3)	//在庫引当順序
		AddSql ",GOKON_KANRI_NO		"'  8)	//合梱管理番号
		AddSql ",JYUCHU_ZAN			"'  7)	//受注残数量
		AddSql ",KYOKYU_KBN			"'  1)	//供給区分
		AddSql ",SHOHIN_SYUSI		"'  8)	//商品化納品在庫収支ｺｰﾄﾞ
		AddSql ",S_SHISAN_SYUSI		"'  8)	//商品化納品資産管理収支ｺｰﾄﾞ
		AddSql ",S_HOJYO_SYUSI		"'  8)	//商品化納品補助収支ｺｰﾄﾞ
		AddSql ",BIKOU1				"' 40)	//備考1
		AddSql ",CHOHA_KBN			"'  1)	//帳端区分
		AddSql ",JYU_HIN_NO			"' 20)	//受付品目番号
		AddSql ",HIN_NAME			"' 20)	//品名
		AddSql ",HIN_CHANGE_KBN		"'  1)	//品目番号変更区分
		AddSql ",MODULE_EXCHANGE	"'  1)	//ﾓｼﾞｭｰﾙ交換区分
		AddSql ",ZAIKO_SYUSI		"'  8)	//残在庫まとめ在庫収支ｺｰﾄﾞ
		AddSql ",ZAN_SHISAN_SYUSI	"'  8)	//残在庫まとめ資産管理収支ｺｰﾄﾞ
		AddSql ",ZAN_HOJYO_SYUSI	"'  8)	//残在庫まとめ補助収支ｺｰﾄﾞ
		AddSql ",NOUKI_YMD			"'  8)	//指定納期
		AddSql ",SERVICE_KANRI_NO	"'  9)	//ｻｰﾋﾞｽ会社管理番号
		AddSql ",KI_HIN_NO			"'  3)	//機種品目ｺｰﾄﾞ
		AddSql ",ENVIRONMENT_KBN	"'  1)	//環境企画部品区分
		AddSql ",SS_CODE			"'  8)	//直送相手先ｺｰﾄﾞ
		AddSql ",KEPIN_KAIJYO		"'  1)	//欠品解消区分
		AddSql ",KAN_DT				"'  8)	//完了日付
		AddSql ",BEF_NYU_QTY		"'  8)	//先行入荷数
		AddSql ",YOSAN_FROM			"'  5)	//予算単位（元）
		AddSql ",YOSAN_TO			"'  5)	//予算単位（先）
		AddSql ",HTANABAN			"'  8)	//標準棚番
		AddSql ",HIN_NAI			"' 13)	//品番（内部）
		AddSql ",H_SOKO				"'  2)	//ﾎｽﾄ倉庫 2006.10.17
		AddSql ",NYU_LIST_OUT		"'  1)	//入庫予定出力ﾌﾗｸﾞ 2007.06.12    現在未使用 0:データ出力対象 9:出力済(もしくは出力対象外)
		AddSql ",GENSANKOKU			"' 20)	//原産国名
		AddSql ",GEN_GENSANKOKU		"' 20)	//現物表示原産国名
		AddSql ",SHIIRE_WORK_CENTER	"'  8)	//資材仕入先ﾜｰｸｾﾝﾀｰ
		AddSql ",KANKYO_KBN			"'  3)	//環境種類区分
		AddSql ",KANKYO_KBN_ST		"'  8)	//環境種類区分適用開始
		AddSql ",KANKYO_KBN_SURYO	"' 10)	//環境種類区分数量
		AddSql ",ID_NO2				"' 12)	//ID_NO
		AddSql ",AITESAKI_CODE		"' 16)	//相手先ｺｰﾄﾞ
		AddSql ",JYUCHU_YMD			"'  8)	//受注年月日
		AddSql ",SHITEI_NOUKI_YMD	"'  8)	//指定納期年月日
		AddSql ",LIST_OUT_END_F		"'  1)	//入庫関連ﾘｽﾄ出力F 0:複数原産国部品入庫管理ﾘｽﾄまたは入庫／棚番ﾁｪｯｸﾘｽﾄが未 '9:複数原産国部品入庫管理ﾘｽﾄかつ入庫／棚番ﾁｪｯｸﾘｽﾄが処理済
		AddSql ",LIST_NYU_KANRI_F	"'  1)	//入庫管理ﾘｽﾄ出力F「複数原産国部品入庫管理ﾘｽﾄ用」 0:印刷対象(未印刷) 8:印刷対象外　9:印刷済	(0→9)
		AddSql ",LIST_NYU_CHECK_F	"'  1)	//入庫ﾁｪｯｸﾘｽﾄ出力F「入庫／棚番ﾁｪｯｸﾘｽﾄ用」　0:未印刷 9:印刷済
		AddSql ",NYUKO_TANABAN		"'  8)	//入庫棚番
		AddSql ",MAEGARI_SURYO		"'  8)	//前借相殺数
		AddSql ",INS_TANTO			"'  5)	//追加　担当者
		AddSql ",Ins_DateTime		"' 14)	//追加　日時  
		AddSql ",UPD_TANTO			"'  5)	//更新　担当者
		AddSql ",UPD_DATETIME		"' 14)	//更新　日時  
		AddSql ",MOTO_PROG_ID		"'  8)	//発生元プログラム
		AddSql ",MOTO_TEXT_NO		"'  9)	//元テキスト№
		AddSql ",JITU_SURYO			"'  7)	//実績数量
		AddSql ") values ("
		AddSql " '0'"	'  1)	//完了区分
		AddSql ",'0'"	'  1)	//データ種別
		AddSql ",'" & RTrim(objRs.Fields("JGyobu")) & "'"'	  1)	//事業部区分	key0
		AddSql ",'1'"	'  1)	//国内外
		AddSql ",'" & RTrim(objRs.Fields("SyoriMD")) & RTrim(objRs.Fields("Bin")) & RTrim(objRs.Fields("SeqNo")) & "'"	'  9)	//テキスト№	key0
		AddSql ",''"	'  8)	//事業場ｺｰﾄﾞ
		AddSql ",''"	'  1)	//データ区分
		AddSql ",''"	'  2)	//取引区分
		AddSql ",'" & RTrim(objRs.Fields("ID_NO")) & "'"	' 12)	//ID-NO
		AddSql ",''"	'  8)	//会計用事業場ｺｰﾄﾞ
		AddSql ",''"	'  8)	//資産管理用事業場ｺｰﾄﾞ
		AddSql ",'" & RTrim(objRs.Fields("Pn")) & "'" 	' 20)	//品目番号
		AddSql ",'" & RTrim(objRs.Fields("DenNo")) & "'"	' 10)	//伝票番号
		AddSql ",'" & RTrim(objRs.Fields("Qty")) & "'"	'  7)	//出荷数量
		AddSql ",''"	'  8)	//得意先コード
		AddSql ",''"	'  8)	//在庫収支
		AddSql ",''"	'  8)	//資産管理用在庫収支ｺｰﾄﾞ
		AddSql ",''"	'  8)	//補助在庫収支ｺｰﾄﾞ
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	'  8)	//出庫日付
		AddSql ",''"	' 10)	//実際単価
		AddSql ",''"	' 12)	//オーダー番号
		AddSql ",''"	'  5)	//アイテム番号
		AddSql ",''"	'  5)	//注文管理番号略号
		AddSql ",''"	' 10)	//個装形態ｺｰﾄﾞ
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	'  8)	//出荷予定日	key0
		AddSql ",''"	' 10)	//ﾛｹｰｼｮﾝ1
		AddSql ",''"	' 10)	//ﾛｹｰｼｮﾝ2
		AddSql ",''"	' 10)	//ﾛｹｰｼｮﾝ3
		AddSql ",''"	' 24)	//得意先名称
		AddSql ",''"	'  1)	//注文区分
		AddSql ",''"	' 10)	//注文区分名称
		AddSql ",''"	' 10)	//原産国1
		AddSql ",''"	' 10)	//原産国2
		AddSql ",''"	' 40)	//備考2
		AddSql ",''"	'  1)	//販売区分
		AddSql ",''"	'  1)	//直送指示区分
		AddSql ",''"	' 12)	//ﾕﾆｯﾄ修正管理番号
		AddSql ",''"	'  3)	//在庫引当順序
		AddSql ",''"	'  8)	//合梱管理番号
		AddSql ",''"	'  7)	//受注残数量
		AddSql ",''"	'  1)	//供給区分
		AddSql ",''"	'  8)	//商品化納品在庫収支ｺｰﾄﾞ
		AddSql ",''"	'  8)	//商品化納品資産管理収支ｺｰﾄﾞ
		AddSql ",''"	'  8)	//商品化納品補助収支ｺｰﾄﾞ
		AddSql ",''"	' 40)	//備考1
		AddSql ",''"	'  1)	//帳端区分
		AddSql ",''"	' 20)	//受付品目番号
		AddSql2 ",'",RTrim(objRs.Fields("PName"))	' 20)	//品名
		AddSql ",''"	'  1)	//品目番号変更区分
		AddSql ",''"	'  1)	//ﾓｼﾞｭｰﾙ交換区分
		AddSql ",''"	'  8)	//残在庫まとめ在庫収支ｺｰﾄﾞ
		AddSql ",''"	'  8)	//残在庫まとめ資産管理収支ｺｰﾄﾞ
		AddSql ",''"	'  8)	//残在庫まとめ補助収支ｺｰﾄﾞ
		AddSql ",'" & RTrim(objRs.Fields("SHITEI_NOUKI_YMD")) & "'"	'  8)	//指定納期
		AddSql ",''"	'  9)	//ｻｰﾋﾞｽ会社管理番号
		AddSql ",''"	'  3)	//機種品目ｺｰﾄﾞ
		AddSql ",''"	'  1)	//環境企画部品区分
		AddSql ",''"	'  8)	//直送相手先ｺｰﾄﾞ
		AddSql ",''"	'  1)	//欠品解消区分
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	'  8)	//完了日付
		AddSql ",''"	'  8)	//先行入荷数
		AddSql ",'" & RTrim(objRs.Fields("SyukoCd")) & "'"	'  5)	//予算単位（元）
		AddSql ",'" & RTrim(objRs.Fields("NyukoCd")) & "'"	'  5)	//予算単位（先）
		AddSql ",'" & RTrim(objRs.Fields("Loc1")) & "'"		'  8)	//標準棚番
		AddSql ",'" & RTrim(objRs.Fields("PnNai")) & "'"		' 13)	//品番（内部）
		AddSql ",'" & RTrim(objRs.Fields("SyushiCd")) & "'"	'  2)	//ﾎｽﾄ倉庫 2006.10.17
		AddSql ",''"	'  1)	//入庫予定出力ﾌﾗｸﾞ 現在未使用 0:データ出力対象 9:出力済(もしくは出力対象外)
		AddSql ",'" & RTrim(objRs.Fields("GENSANKOKU")) & "'"	' 20)	//原産国名
		AddSql ",'" & RTrim(objRs.Fields("GEN_GENSANKOKU")) & "'"	' 20)	//現物表示原産国名
		AddSql ",'" & RTrim(objRs.Fields("SHIIRE_WORK_CENTER")) & "'"	'  8)	//資材仕入先ﾜｰｸｾﾝﾀｰ
		AddSql ",'" & RTrim(objRs.Fields("KANKYO_KBN")) & "'"	'  3)	//環境種類区分
		AddSql ",'" & RTrim(objRs.Fields("KANKYO_KBN_ST")) & "'"	'  8)	//環境種類区分適用開始
		AddSql ",'" & RTrim(objRs.Fields("KANKYO_KBN_SURYO")) & "'"	' 10)	//環境種類区分数量
		AddSql ",'" & RTrim(objRs.Fields("ID_NO")) & "'"			' 12)	//ID_NO
		AddSql ",'" & RTrim(objRs.Fields("AITESAKI_CODE")) & "'"	' 16)	//相手先ｺｰﾄﾞ
		AddSql ",'" & RTrim(objRs.Fields("JYUCHU_YMD")) & "'"	'  8)	//受注年月日
		AddSql ",'" & RTrim(objRs.Fields("SHITEI_NOUKI_YMD")) & "'"	'  8)	//指定納期年月日
		AddSql ",'9'"	'  1)	//入庫関連ﾘｽﾄ出力F 0:複数原産国部品入庫管理ﾘｽﾄまたは入庫／棚番ﾁｪｯｸﾘｽﾄが未 '9:複数原産国部品入庫管理ﾘｽﾄかつ入庫／棚番ﾁｪｯｸﾘｽﾄが処理済
		AddSql ",'9'"	'  1)	//入庫管理ﾘｽﾄ出力F「複数原産国部品入庫管理ﾘｽﾄ用」 0:印刷対象(未印刷) 8:印刷対象外　9:印刷済	(0→9)
		AddSql ",'9'"	'  1)	//入庫ﾁｪｯｸﾘｽﾄ出力F「入庫／棚番ﾁｪｯｸﾘｽﾄ用」　0:未印刷 9:印刷済
		AddSql ",'" & strZaiko & "'"	'  8)	//入庫棚番
		AddSql ",''"	'  8)	//前借相殺数
		AddSql ",'HM500'"		'  5)	//追加担当者
		AddSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"	' 14)	//追加 日時
		AddSql ",''"	'  5)	//更新担当者
		AddSql ",''"	' 14)	//更新日時  
		AddSql ",''"	'  8)	//発生元プログラム
		AddSql ",''"	'  9)	//元テキスト№
		AddSql ",''"	'  7)	//実績数量
		AddSql ")"
		Debug strSql
		dim	vRet
		vRet = Execute(strSql)
		select case vRet
		case 0
			Write ":ok:",0
			ZaikoInsert
		case -2147467259	'0x80004005 重複キー
			Write ":dup",0
		case else
			Write ":0x" & Hex(vRet),0
		end select
	End Function
	'-----------------------------------------------------------------------
	'List1() 入荷リスト
	'-----------------------------------------------------------------------
    Private Function List1()
		Debug ".List1()"
		AddSql ""
		AddSql "select"
		AddSql " h.JGyobu"
		AddSql ",h.DenDt"
'		AddSql ",h.IoKbn"
'		AddSql ",h.AkaKuro"
		AddSql ",h.SyukoCd"
'		AddSql ",h.NyukoCd"
		AddSql ",y.YName"
		AddSql ",h.SyushiCd"
		AddSql ",h.Pn"
		AddSql ",h.Qty"
		AddSql "from hmem500 h"
		AddSql "left outer join Yosan y on (y.JGyobu = h.JGyobu and y.YCode = h.SyukoCd)"
		AddWhere "h.Filename",strFileName
		AddWhere "h.DenDt",strDt
		AddWhere "h.IoKbn","1"
		AddSql "order by 1,2,3,4,5"
'		AddSql " DenDt"
'		AddSql ",SyukoCd"
'		AddSql ",SyushiCd"
'		AddSql ",Pn"
		CallSql strSql
'		curDenDt	= ""
'		curSyukoCd	= ""
		Call GroupHead(-1)
		do while objRs.Eof = False
			Line1
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line1() 入荷リスト １行表示
	'-----------------------------------------------------------------------
	Private	curDenDt
 	Private	curSyukoCd
    Private Function Line1()
		Debug ".Line1()"
		if Len(RTrim(objRs.Fields("SyukoCd"))) = 4 then
			if Right(RTrim(objRs.Fields("SyukoCd")),2) = RTrim(objRs.Fields("SyushiCd")) then
				exit function
			end if
		end if
'		dim	curDiff
'		curDiff = False
'		if curDenDt <> RTrim(objRs.Fields("DenDt")) then
'			curDenDt = RTrim(objRs.Fields("DenDt"))
'			curDiff = True
'		end if
'		if curSyukoCd <> RTrim(objRs.Fields("SyukoCd")) then
'			curSyukoCd = RTrim(objRs.Fields("SyukoCd"))
'			curDiff = True
'		end if
'		if curDiff = True then
		if GroupHead(2) = True then
			Write "■",0
			Write objRs.Fields("DenDt"),9
			WriteLine ""
			Write "■",0
			Write objRs.Fields("SyukoCd"),0
			Write "",1
			Write RTrim(objRs.Fields("YName"))	,0
			WriteLine ""
		end if
		Write objRs.Fields("Pn")		,13
		Write CLng(objRs.Fields("Qty"))	,-4
		Write ""			,1
		Write RTrim(objRs.Fields("SyushiCd"))	,0
		WriteLine ""
	End Function
	'-------------------------------------------------------------------
	'GroupHead() グループヘッダー
	'	True:グループヘッダー
	'  Flase:継続行
	'-------------------------------------------------------------------
	Private	curHead
	Private	newHead
	Private	Function GroupHead(byVal intHead)
		if intHead < 0 then
			curHead = ""
			exit function
		end if
		dim	i
		newHead = ""
		for i = 0 to intHead
			newHead = newHead + objRs.Fields(i)
		next
		if curHead = newHead then
			GroupHead = False
			exit function
		end if
		curHead = newHead
		GroupHead = True
	End Function
	'-----------------------------------------------------------------------
	'Item() 品目マスター登録
	'-----------------------------------------------------------------------
    Private Function Item()
		Debug ".Item()"
		AddSql ""
		AddSql "select distinct"
		AddSql " h.Filename"
		AddSql ",h.JGyobu"
		AddSql ",h.Pn"
		AddSql ",h.PnNai"
		AddSql ",h.PName"
		AddSql ",h.SHIIRE_WORK_CENTER"
		AddSql ",i.TORI_SHIIRE_WORK_CTR"
		AddSql ",h.KANKYO_KBN hKANKYO_KBN"
		AddSql ",i.KANKYO_KBN iKANKYO_KBN"
		AddSql ",h.KANKYO_KBN_ST hKANKYO_KBN_ST"
		AddSql ",i.KANKYO_KBN_ST iKANKYO_KBN_ST"
		AddSql ",h.KANKYO_KBN_SURYO hKANKYO_KBN_SURYO"
		AddSql ",i.KANKYO_KBN_SURYO iKANKYO_KBN_SURYO"
		AddSql ",i.INSP_MESSAGE INSP_MESSAGE"
'		AddSql ",ifnull(i.Hin_Name,'*未登録*') Hin_Name"
		AddSql "from hmem500 h"
		AddSql "left outer join item i on (i.JGyobu = h.JGyobu and i.NAIGAI='1' and i.HIN_GAI = h.Pn)"
		select case strItem
		case "n","y"	' 追加
			AddSql "where i.HIN_GAI is null"
		case "u"		' 更新
		end select
		AddWhere "h.Filename",strFileName
		AddWhere "h.DenDt",strDt
		CallSql strSql
		do while objRs.Eof = False
'			Write "item:" & strItem & ":",0
			if inStr(strFileName,"%") > 0 then
				Write RTrim(objRs.Fields("Filename")) & " " ,0
			end if
			Write objRs.Fields("JGyobu")	,2
			Write Rtrim(objRs.Fields("Pn")) & "" 		,0
'			Write Rtrim(objRs.Fields("PnNai")) & " "	,0
'			Write Rtrim(objRs.Fields("PName"))	,0
			Write ":" & strItem & ":",0
			select case strItem
			case "y"
				ItemInsert
			case "u"
				ItemUpdate
			end select
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'ItemInsert() Item追加
	'-----------------------------------------------------------------------
    Private Function ItemInsert()
		Debug ".ItemInsert()"
		AddSql ""
		AddSql "insert into Item"
		AddSql "("
		AddSql " JGYOBU"				'Char(  1) //事業部区分
		AddSql ",NAIGAI"				'Char(  1) //国内外
		AddSql ",HIN_GAI"				'Char( 20) //品番（外部）
		AddSql ",HIN_NAME"				'Char( 40) //品名
		AddSql ",HIN_NAI"				'Char( 20) //品番（内部）
		AddSql ",GLICS1_TANA"			'Char( 10) //グリックス棚番１   2005.05
		AddSql ",GLICS2_TANA"			'Char( 10) //グリックス棚番２   2005.05
		AddSql ",GLICS3_TANA"			'Char( 10) //グリックス棚番３   2005.05
		AddSql ",L_HIN_NAME_E"			'Char( 30) //商品ﾗﾍﾞﾙ   品名
		AddSql ",L_KISHU1"				'Char( 25) //           機種(1)
		AddSql ",L_KISHU3"				'Char(150) //           機種(3)(→適用機種備
		AddSql ",L_URIKIN1"				'Char( 10) //           価格(1)	//NUMERICSA(10,0)
		AddSql ",L_URIKIN2"				'Char( 10) //           価格(2)	//NUMERICSA(10,0)
		AddSql ",L_URIKIN3"				'Char( 10) //           価格(3)	//NUMERICSA(10,0)
		AddSql ",UNIT_BUHIN"			'Char(  1) //ﾕﾆｯﾄ部品区分       2006.07.28
		AddSql ",NAI_BUHIN"				'Char(  1) //国内供給部品区分   2006.07.28
		AddSql ",GAI_BUHIN"				'Char(  1) //海外供給部品区分   2006.07.28
		AddSql ",HYO_TANKA"				'Char( 10) //標準単価   2006.07.28
		AddSql ",KANKYO_KBN"			'Char(  3) //環境種類区分       2010.07.27
		AddSql ",KANKYO_KBN_ST"			'Char(  8) //環境種類区分適用開始 2010.07.
		AddSql ",KANKYO_KBN_SURYO"		'Char( 10) //環境種類区分数量   2010.07.27
		AddSql ",CS_TANTO_CD"			'Char(  8) //CS担当ｺｰﾄﾞ
		AddSql ",D_MODEL"				'Char(  8) //代表機種品目コード PN連携でセット 2011.12.28
		AddSql ",HINMOKU"				'Char(  8) //品目コード         PN連携でセット 2011.12.28
		AddSql ",K_KEITAI"				'Char( 14) //個装形態(14桁)     2012.03.13
		AddSql ",INS_TANTO"				'Char(  5) //追加　担当者
		AddSql ",Ins_DateTime"			'Char( 14) //追加　日時  
		AddSql ",BIKOU20"
		AddSql ",L_PAPER"				' not null   //           紙
		AddSql ",L_PLASTIC"             ' not null   //           プラスチック
		AddSql ",L_LABEL"               ' not null   //           適用機種ﾗﾍﾞﾙ
		AddSql ")"
		AddSql "select top 1"
		AddSql " h.JGyobu"				'//事業部区分
		AddSql ",'1'"					'//国内外
		AddSql ",p.Pn"					'Char( 20) //品番（外部）
		AddSql ",p.PnBetsu"				'Char( 40) //品名
		AddSql ",p.SPn"					'Char( 20) //品番（内部）
		AddSql ",p.Loc1"				'Char( 10) //グリックス棚番１   2005.05
		AddSql ",p.Loc2"				'Char( 10) //グリックス棚番２   2005.05
		AddSql ",p.Loc3"				'Char( 10) //グリックス棚番３   2005.05
		AddSql ",RTrim(p.PNameEngA)"	'Char( 30) //商品ﾗﾍﾞﾙ   品名
		AddSql ",p.NaiModel"			'Char( 25) //           機種(1)
		AddSql ",p.GaiModel"			'Char(150) //           機種(3)(→適用機種備
		AddSql ",p.Tanka2"				'Char( 10) //           価格(1)	//NUMERICSA(10,0)
		AddSql ",p.Tanka3"				'Char( 10) //           価格(2)	//NUMERICSA(10,0)
		AddSql ",p.Tanka4"				'Char( 10) //           価格(3)	//NUMERICSA(10,0)
		AddSql ",p.UnitKbn"				'Char(  1) //ﾕﾆｯﾄ部品区分       2006.07.28
		AddSql ",p.NaiKbn"				'Char(  1) //国内供給部品区分   2006.07.28
		AddSql ",p.GaiKbn"				'Char(  1) //海外供給部品区分   2006.07.28
		AddSql ",p.HyoTan"				'Char( 10) //標準単価   2006.07.28
		AddSql ",h.KANKYO_KBN"			'Char(  3) //環境種類区分       2010.07.27
		AddSql ",h.KANKYO_KBN_ST"		'Char(  8) //環境種類区分適用開始 2010.07.
		AddSql ",h.KANKYO_KBN_SURYO"	'Char( 10) //環境種類区分数量   2010.07.27
		AddSql ",p.KobaiTanto"			'Char(  8) //CS担当ｺｰﾄﾞ
		AddSql ",p.DModel"				'Char(  8) //代表機種品目コード PN連携でセット 2011.12.28
		AddSql ",p.Hinmoku"				'Char(  8) //品目コード         PN連携でセット 2011.12.28
		AddSql ",p.KKeitai"				'Char( 14) //個装形態(14桁)     2012.03.13
		AddSql ",'HM500'"				'Char(  5) //追加　担当者
		AddSql ",left(replace(replace(replace(convert(Now(),sql_char),'-',''),':',''),' ',''),14)"	'Char( 14) //追加　日時  
		AddSql ",case p.KobaiTanto"
		AddSql " when 'R101' then '砂畠'"
		AddSql " when 'R102' then '今井'"
		AddSql " when 'R103' then '遠山'"
		AddSql " when 'R104' then '今井'"
		AddSql " when 'R105' then '川村'"
		AddSql " when 'R106' then '砂畠'"
		AddSql " else ''"
		AddSql " end"
		AddSql ",'0'"	'//           紙
		AddSql ",'0'"	'//           プラスチック
		AddSql ",'0'"	'//           適用機種ﾗﾍﾞﾙ
		AddSql "from hmem500 h"
		AddSql "inner join Pn5 p on (h.Pn = p.Pn)"
		AddWhere "h.Filename",RTrim(objRs.Fields("Filename"))
		AddWhere "h.Pn",RTrim(objRs.Fields("Pn"))
		Write ":" & Execute(strSql) ,0
	End Function
	'-----------------------------------------------------------------------
	'ItemUpdate() Item更新
	'	TORI_SHIIRE_WORK_CTR"	' //仕入ﾜｰｸセンター    
	'	KANKYO_KBN"				' //環境種類区分       
	'	KANKYO_KBN_ST"			' //環境種類区分適用開始
	'	KANKYO_KBN_SURYO"		' //環境種類区分数量   
	'-----------------------------------------------------------------------
    Private Function ItemUpdate()
		Debug ".ItemUpdate()"
		dim	strSet
		strSet = ""
		strSet = SetSql(strSet," 仕WC:","TORI_SHIIRE_WORK_CTR",RTrim(objRs.Fields("TORI_SHIIRE_WORK_CTR")),RTrim(objRs.Fields("SHIIRE_WORK_CENTER")))
		strSet = SetSql(strSet," 環境:","KANKYO_KBN",RTrim(objRs.Fields("iKANKYO_KBN")),RTrim(objRs.Fields("hKANKYO_KBN")))
		strSet = SetSql(strSet," 環境始:","KANKYO_KBN_ST",RTrim(objRs.Fields("iKANKYO_KBN_ST")),RTrim(objRs.Fields("hKANKYO_KBN_ST")))
		strSet = SetSql(strSet," 環境数:","KANKYO_KBN_SURYO",Trim(objRs.Fields("iKANKYO_KBN_SURYO")),Trim(objRs.Fields("hKANKYO_KBN_SURYO")))
		dim	strMsg
		if RTrim(objRs.Fields("hKANKYO_KBN")) = "LIT" then
'			strMsg = "リチウム電池搭載(" & Trim(objRs.Fields("hKANKYO_KBN_SURYO")) & ")"
'			strSet = SetSql(strSet," 検品Msg:","INSP_MESSAGE",Trim(objRs.Fields("INSP_MESSAGE")),strMsg)
		end if
		if strSet = "" then
			exit function
		end if
		AddSql ""
		AddSql "update Item"
		AddSql strSet
		AddSql ",UPD_TANTO='HM500'"
		AddSql ",UPD_DATETIME = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
		AddSql "where JGyobu = '" & RTrim(objRs.Fields("JGyobu")) & "'"
		AddSql "  and NAIGAI = '1'"
		AddSql "  and HIN_GAI = '" & RTrim(objRs.Fields("Pn")) & "'"
		Write ":" & Execute(strSql) ,0
	End Function
	'-----------------------------------------------------------------------
	'SetSql() 
	'-----------------------------------------------------------------------
    Private Function SetSql(byVal strSet,byVal strTitle,byVal strName,byVal strSrc,byVal strDst)
		Debug ".SetSql()"
		Write strTitle,0
		Write strSrc,0
		if strDst <> "" then
			select case strName
			case "KANKYO_KBN_SURYO"
				if strDst = "0" then
					strDst = strSrc
				end if
			case "INSP_MESSAGE"
				if strSrc = "単価改訂 リチウム電池搭載" then
					strDst = strSrc
				end if
			end select
			if strDst <> strSrc then
				Write "→",0
				Write strDst,0
				if strSet = "" then
					strSet = " set "
				else
					strSet = strSet & " ,"
				end if
				strSet = strSet & strName & " = '" & strDst & "'"
			end if
		end if
		SetSql = strSet
	End Function
	'-----------------------------------------------------------------------
	'Zaiko() 在庫登録
	'-----------------------------------------------------------------------
    Private Function Zaiko()
		Debug ".Zaiko():" & strZaiko
		AddSql ""
		AddSql "select"
		AddSql " *"
		AddSql "from hmem500 h"
'		AddSql "where h.IoKbn = '1'"
		AddWhere "h.Filename",strFileName
		AddWhere "h.DenDt",strDt
		AddSql "order by Filename,Row"
		CallSql strSql
		do while objRs.Eof = False
			Write objRs.Fields("JGyobu")	,2
			Write objRs.Fields("DenDt")		,9
			Write objRs.Fields("IoKbn")		,0
			Write objRs.Fields("AkaKuro")	,0
			Write objRs.Fields("SyoriMD")	,0
			Write objRs.Fields("Bin")		,0
			Write objRs.Fields("SeqNo")		,5
			Write objRs.Fields("Pn")		,15
			Write objRs.Fields("Qty")		,7
			Write objRs.Fields("SyukoCd")	,6
			Write objRs.Fields("NyukoCd")	,4
			Write objRs.Fields("SyushiCd")	,3
'			Write objRs.Fields("Loc1")		,9
			Write objRs.Fields("SHIIRE_WORK_CENTER")	,8
			Write objRs.Fields("EOR")	,1
			Write ":",0
			if ZaikoInsert() = 1 then
				AddSql ""
				AddSql "update hmem500"
				AddSql " set EOR = '1'"
				AddSql " where Filename = '" & RTrim(objRs.Fields("Filename")) & "'"
				AddSql "   and Row = " & objRs.Fields("Row")
				Write ":" &  Execute(strSql),0
			end if
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'ZaikoInsert() 在庫データ登録
	'-----------------------------------------------------------------------
    Private Function ZaikoInsert()
		Debug ".ZaikoInsert()"
		ZaikoInsert = 0
'		if objRs.Fields("IoKbn") <> "1" then
'			exit function
'		end if
'		if Len(RTrim(objRs.Fields("SyukoCd"))) = 4 then
'			if Right(RTrim(objRs.Fields("SyukoCd")),2) = RTrim(objRs.Fields("SyushiCd")) then
'				exit function
'			end if
'		end if
		Write strZaiko,0
'		if objRs.Fields("EOR") <> "@" then
'			if GetOption("","u") <> "u" then
'				exit function
'			end if
'		end if
'		select case RTrim(objRs.Fields("SyukoCd"))
'		case "KARI"	'"HIFU"
'			exit function
'		end select
		AddSql ""
		AddSql "insert into Zaiko"
		AddSql "("
		AddSql "	Soko_No				"' //倉庫№
		AddSql ",	Retu				"' //棚番列
		AddSql ",	Ren					"' //棚番連
		AddSql ",	Dan					"' //棚番段
		AddSql ",	JGYOBU				"' //事業部区分
		AddSql ",	NAIGAI				"' //国内外
		AddSql ",	HIN_GAI				"' //品番（外部）
		AddSql ",	GOODS_ON			"' //0:商品化済 1:未商品
		AddSql ",	NYUKA_DT			"' //入荷日付
		AddSql ",	NYUKO_DT			"' //入庫日付
		AddSql ",	HIN_NAI				"' //品番（内部）
		AddSql ",	YUKO_Z_QTY			"' //有効在庫数
		AddSql ",	LOCK_F				"' //排他フラグ
		AddSql ",	WEL_ID				"' //使用子機ID
		AddSql ",	PRG_ID				"' //使用中プログラム
		AddSql ",	GOODS_YMD			"' //商品化日付
		AddSql ",	SHIIRE_CODE			"' //仕入先ｺｰﾄﾞ
		AddSql ",	SHIIRE_TANKA		"' //仕入単価(9(8)V99)
		AddSql ",	KEIJYO_YM			"' //計上年月
		AddSql ",	GENSANKOKU			"' //原産国名
		AddSql ",	SHIIRE_WORK_CENTER	"' //資材仕入先ﾜｰｸｾﾝﾀｰ
		AddSql ",	ID_NO2				"' //ID_NO
		AddSql ",	YOSAN_FROM			"' //予算単位（元）
		AddSql ",	YOSAN_TO			"' //予算単位（先）
		AddSql ") values ("
		AddSql " '" & mid(strZaiko,1,2) & "'"	' //倉庫№
		AddSql ",'" & mid(strZaiko,3,2) & "'"' //棚番列
		AddSql ",'" & mid(strZaiko,5,2) & "'"' //棚番連
		AddSql ",'" & mid(strZaiko,7,2) & "'"' //棚番段
		AddSql ",'" & RTrim(objRs.Fields("JGyobu")) & "'"	' //事業部区分
		AddSql ",'1'"										' //国内外
		AddSql ",'" & RTrim(objRs.Fields("Pn")) & "'"		' //品番（外部）
		AddSql ",'1'"										' //0:商品化済 1:未商品
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	' //入荷日付
		AddSql ",''"										' //入庫日付
		AddSql ",'" & RTrim(objRs.Fields("PnNai")) & "'"	' //品番（内部）
		AddSql ",'" & RTrim(objRs.Fields("Qty")) & "'"		' //有効在庫数
		AddSql ",'0'"										' //排他フラグ
		AddSql ",''"										' //使用子機ID
		AddSql ",''"										' //使用中プログラム
		AddSql ",''"										' //商品化日付
		AddSql ",''"										' //仕入先ｺｰﾄﾞ
		AddSql ",''"										' //仕入単価(9(8)V99)
		AddSql ",''"										' //計上年月
		AddSql ",'" & RTrim(objRs.Fields("GENSANKOKU")) & "'"			' //原産国名
		AddSql ",'" & RTrim(objRs.Fields("SHIIRE_WORK_CENTER")) & "'"	' //資材仕入先ﾜｰｸｾﾝﾀｰ
		AddSql ",'" & RTrim(objRs.Fields("ID_NO")) & "'"	' //ID_NO
		AddSql ",'" & RTrim(objRs.Fields("SyukoCd")) & "'"	' //振替元(予算単位)
		AddSql ",'" & RTrim(objRs.Fields("SyushiCd")) & "'"	' //振替先(在庫収支)
		AddSql ")"
		Debug strSql
		dim	vRet
		vRet = Execute(strSql)
		select case vRet
		case 0
			Write ":ok",0
			ZaikoInsert = 1
		case -2147467259	'0x80004005 重複キー
			AddSql ""
			AddSql "update Zaiko"
			if GetOption("","u") = "u" then
				Write ":u",0
				AddSql " set YUKO_Z_QTY = '" & RTrim(objRs.Fields("Qty")) & "'"
				AddSql "   , YOSAN_TO = '" & RTrim(objRs.Fields("SyushiCd")) & "'"
				AddSql " where JGYOBU = '" & RTrim(objRs.Fields("JGyobu")) & "'"
				AddSql "   and NAIGAI = '1'"
				AddSql "   and HIN_GAI = '" & RTrim(objRs.Fields("Pn")) & "'"
				AddSql "   and GOODS_ON = '1'"	' //0:商品化済 1:未商品
				AddSql "   and Soko_No = '" & mid(strZaiko,1,2) & "'"
				AddSql "   and Retu	   = '" & mid(strZaiko,3,2) & "'"
				AddSql "   and Ren	   = '" & mid(strZaiko,5,2) & "'"
				AddSql "   and Dan	   = '" & mid(strZaiko,7,2) & "'"
				AddSql "   and NYUKA_DT = '" & RTrim(objRs.Fields("DenDt")) & "'"
				AddSql "   and ID_NO2 = '" & RTrim(objRs.Fields("ID_NO")) & "'"
			else
				Write ":w",0
				AddSql " set YUKO_Z_QTY = convert(convert(YUKO_Z_QTY,sql_decimal) + " & RTrim(objRs.Fields("Qty")) & ",sql_char)"
				AddSql "   , YOSAN_TO = RTrim(YOSAN_TO) + '" & RTrim(objRs.Fields("SyushiCd")) & "'"
				AddSql " where JGYOBU = '" & RTrim(objRs.Fields("JGyobu")) & "'"
				AddSql "   and NAIGAI = '1'"
				AddSql "   and HIN_GAI = '" & RTrim(objRs.Fields("Pn")) & "'"
				AddSql "   and GOODS_ON = '1'"	' //0:商品化済 1:未商品
				AddSql "   and Soko_No = '" & mid(strZaiko,1,2) & "'"
				AddSql "   and Retu	   = '" & mid(strZaiko,3,2) & "'"
				AddSql "   and Ren	   = '" & mid(strZaiko,5,2) & "'"
				AddSql "   and Dan	   = '" & mid(strZaiko,7,2) & "'"
				AddSql "   and NYUKA_DT = '" & RTrim(objRs.Fields("DenDt")) & "'"
			end if
			vRet = Execute(strSql)
			select case vRet
			case 0
				Write ":ok",0
				ZaikoInsert = 1
			case else
				Write ":0x",Hex(vRet)
			end select
		case else
			Write ":0x",Hex(vRet)
		end select
	End Function
	'-----------------------------------------------------------------------
	'List0()
	'-----------------------------------------------------------------------
    Private Function List0()
		Debug ".List0()"
		AddSql ""
		AddSql "select"
		AddSql " Filename"
		AddSql ",count(*) cnt"
		AddSql2 "from ",strTable & " h"
		AddWhere "h.Filename",strFileName
		AddWhere "h.DenDt",strDt
		AddSql "group by"
		AddSql " Filename"
		AddSql "order by"
		AddSql " Filename"
		CallSql strSql
		do while objRs.Eof = False
			Line0
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line0() 1行表示
	'-----------------------------------------------------------------------
    Private Function Line0()
		Debug ".Line0()"
		Write objRs.Fields("Filename")	,0
		Write objRs.Fields("cnt")		,-5
	End Function
	'-----------------------------------------------------------------------
	'List2()
	'-----------------------------------------------------------------------
    Private Function List2()
		Debug ".List2()"
		AddSql ""
		AddSql "select"
		AddSql " h.Filename Filename"
		AddSql ",h.JGyobu JGyobu"
		AddSql ",h.DenDt"
		AddSql ",h.IoKbn"
		AddSql ",h.AkaKuro"
		AddSql ",h.SyukoCd"
		AddSql ",h.NyukoCd"
		AddSql ",h.SyushiCd"
		AddSql ",count(*) cnt"
		AddSql ",if(n.JGyobu is null,'','入荷 ' + n.JGyobu + ' ' + n.NYUKO_TANABAN) Nyuka"
		AddSql2 "from ",strTable & " h"
		AddSql "left outer join y_nyuka n"
		AddSql " on (h.JGyobu = n.JGyobu"
		AddSql " and h.DenDt = n.SYUKA_YMD"
		AddSql " and (h.SyoriMD + h.Bin + h.SeqNo) = n.Text_No"
		AddSql "	)"
		AddWhere "h.Filename",strFileName
		AddWhere "h.DenDt",strDt
		AddSql "group by"
		AddSql " Filename"
		AddSql ",h.JGyobu"
		AddSql ",h.DenDt"
		AddSql ",h.IoKbn"
		AddSql ",h.AkaKuro"
		AddSql ",h.SyukoCd"
		AddSql ",h.NyukoCd"
		AddSql ",h.SyushiCd"
		AddSql ",Nyuka"
		AddSql "order by"
		AddSql " Filename"
		AddSql ",h.JGyobu"
		AddSql ",h.DenDt"
		AddSql ",h.IoKbn"
		AddSql ",h.SyukoCd"
		AddSql ",h.NyukoCd"
		AddSql ",h.SyushiCd"
		AddSql ",h.AkaKuro"
		CallSql strSql
		do while objRs.Eof = False
			Line2
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line2() 1行表示
	'-----------------------------------------------------------------------
    Private Function Line2()
		Debug ".Line2()"
		Write objRs.Fields("JGyobu")	,2
		Write objRs.Fields("DenDt")		,9
		Write objRs.Fields("IoKbn")		,1
		Write objRs.Fields("AkaKuro")	,2
		Write objRs.Fields("SyukoCd")	,6
		Write objRs.Fields("NyukoCd")	,6
		Write objRs.Fields("SyushiCd")	,3
		Write objRs.Fields("cnt")		,-5
		Write " " & objRs.Fields("Nyuka")		,0
'		Write "" & objRs.Fields("NYUKO_TANABAN")		,0
	End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private Function Execute(byVal strSql)
		Debug ".Execute():" & strSql
		on error resume next
		Call objDb.Execute(strSql)
		Execute = Err.Number
		select case Execute
		case 0
		case -2147467259	'0x80004005 重複キー
		case else
			Wscript.StdErr.WriteLine
			Wscript.StdErr.WriteLine Err.Description
			Wscript.StdErr.WriteLine strSql
		end select
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private	objRs
	Private Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		if Err.Number <> 0 then
			Wscript.StdOut.WriteLine "0x" & Hex(Err.Number) & ":" & Err.Description
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end if
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
	'-------------------------------------------------------------------
	'AddSql2
	'-------------------------------------------------------------------
	Private	Function AddSql2(byVal str1,byVal str2)
		if Right(str1,1) = "'" then
			'Char
			str2 = Replace(RTrim(str2),"'","''") & "'"
		end if
		AddSql str1 & str2
	End Function
	'-------------------------------------------------------------------
	'Where strSql
	'-------------------------------------------------------------------
	Private	Function AddWhere(byVal strF,byVal strV)
		if strV = "" then
			exit function
		end if
		if inStr(strSql,"where") > 0 then
			AddSql " and "
		else
			AddSql " where "
		end if
		dim	strCmp
		strCmp = "="
		if left(strV,1) = "-" then
			strV = Right(strV,len(strV)-1)
			strCmp = "<>"
		end if
		if inStr(strV,"%") > 0 then
			if strCmp = "=" then
				strCmp = " like "
			else
				strCmp = " not like "
			end if
		end if
		AddSql strF & " " & strCmp & " '" & strV & "'"
	End Function
	'-------------------------------------------------------------------
	'文字列追加 strSql
	'-------------------------------------------------------------------
	Private	strSql
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
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
	'オプション取得
	'-----------------------------------------------------------------------
	Private Function GetOption(byval strName ,byval strDefault)
		dim	strValue

		if strName = "" then
			strValue = ""
			if WScript.Arguments.Named.Exists(strDefault) then
				strValue = strDefault
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
	'WriteLine
	'-----------------------------------------------------------------------
	Private Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine s
	End Sub
	'-----------------------------------------------------------------------
	'Write
	'-----------------------------------------------------------------------
	Private Sub Write(byVal s,byVal i)
		if i > 0 then
			s = left(RTrim(s) & space(i),i)
		elseif i < 0 then
			s = right(space(-i) & LTrim(s),-i)
		end if
		Wscript.StdOut.Write "" & s
	End Sub
	'-----------------------------------------------------------------------
	'Echo
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		Wscript.Echo s
	End Sub
End Class
