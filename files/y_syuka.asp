<% Option Explicit	%>
<% Response.Buffer = false %>
<% Response.Expires = -1 %>
<%
Function GetVersion()
' 2008.06.23 検索条件：請求区分 を追加/集計表(商品化請求内訳)"
' 2008.06.19 出庫OK解除 を追加 (完了区分=9 → 0)"
' 2008.04.12 検索条件：取引区分 追加"
' 2008.02.07 集計表(倉庫) 倉庫2桁で集計するように変更"
' 2008.01.11 G事業場(JGYOBA)対応(一覧表、検索条件)"
' 2007.12.26 削除済→予定の戻し対応
' 2007.10.04 検索条件：検品区分追加(中部デポ出荷明細対応)"
' 2007.08.08 集計表(品番/在庫)・複数ロケーションの出荷数量、件数がおかしい不具合修正"
' 2007.08.09 集計表(品番/在庫)・要商品化数の計算式を(数量－商品化済)に変更"
' 2007.08.06 集計表(品番/在庫) 作成
' ver1.17 2007.07.07 時間別件数個数集計表の対応
' ver1.16 2007.05.31 検品OK解除の対応
' ver1.15 2006.08.07 時間別集計表：直送区分追加/7時(7～9時)10時(10～11時)追加
' ver1.13 2006.07.08 出庫済の検索方法(-0)表示"
' ver1.12 2006.07.08 出庫済の検索方法(0)表示
' ver1.11 2006.07.03 Response.Buffer = false に変更
' ver1.10 2006.06.28 集計表(時間別)の取引区分追加,注文区分の名称表示/検品済の表記追加
' ver1.09 2006.06.24 集計表(時間別) の 12時～18時は伝票日付と同じ場合のみカウントするように変更
' ver1.08 2006.06.24 12時を分けて集計
' ver1.07 2006.06.20 伝票No/ID-NO 検索対応
' ver1.06 2006.06.15 集計表(注文区分別/出荷日) の対応
' ver1.05 2006.06.15 group by , ourder by を修正
' ver1.04 2006.06.13 項目名をテーブルの最終行にも表示
' ver1.03 2006.06.13 向け先マスター参照対応
	GetVersion = "2008.07.02 集計表(品番/在庫) エラーになる問題の対処"
	GetVersion = "2008.07.08 検品OK エラーになる問題の対処( as y 追記)"
	GetVersion = "2008.07.17 単価設定(0:単価未登録/1:単価登録済)の対応"
	GetVersion = "<font color='red'>2008.07.18 単価設定(0:単価未登録/1:単価登録済)の検索不具合修正</font>"
	GetVersion = "<font color='red'>2008.07.19 単価設定(0:単価未登録/1:単価登録済)の検索不具合修正(商品化請求内訳 以外の出力形式対応)</font>"
	GetVersion = "2008.10.28 集計表(資材)の対応"
	GetVersion = "2008.11.04 dbName 変数化"
	GetVersion = "2008.12.12 dbName 変数化"
	GetVersion = "2008.12.24 検索条件 欠品解除 追加"
	GetVersion = "2009.02.06 出力形式のデータ操作関連を非表示"
	GetVersion = "2009.02.24 一覧表：オーダーNo 追加"
	GetVersion = "2009.02.24 出力形式：集計表(オーダーNo) 追加"
	GetVersion = "2009.04.21 集計表(倉庫)：倉庫名を表示"
	GetVersion = "2009.05.08 一覧表：移管状況 削除/検品メッセージ 追加"
	GetVersion = "2009.06.17 【重要】検索条件 単価設定(0:単価未登録/1:単価登録済)の単価設定日で判断するように変更"
	GetVersion = "2009.08.12 印刷時、検索条件を非表示にするよう改善(検索結果を大きく表示する為)"
	GetVersion = "2009.10.07 一覧表：出庫表印刷 追加"
	GetVersion = "2009.10.22 実行SQLを非表示に変更 リンクで表示/非表示切替"
	GetVersion = "2009.10.26 集計表(品番/在庫集計) 対応・・・廃棄データとBU/PPSC在庫数を照合"
	GetVersion = "2009.11.05 検索条件 事業部 の複数指定対応 例：4,D"
	GetVersion = "2009.11.12 集計表(品番 出荷急増) 作成"
	GetVersion = "2009.11.12 集計表(品番 出荷急増) 削除済(DEL_SYUKA) を指定して動作するように変更"
	GetVersion = "2010.01.15 集計表(月別) 【アイテム】追加"
	GetVersion = "2010.03.23 集計表(注文区分:1500前後) 対応(奈良センター倉庫発送の為)"
	GetVersion = "2010.04.07 一覧表：アイテムNo 追加"
	GetVersion = "2010.06.09 集計表(注区別)/集計表(注区別/15:00前後)：才数 追加"
	GetVersion = "2010.06.10 才数 の表示形式を小数2桁に変更"
	GetVersion = "2010.06.10 集計表(品番/向け先) に 品名 を追加"
	GetVersion = "2010.06.25 集計表(時間別) に 才数 を追加"
	GetVersion = "2010.08.20 ポップアップメニュー対応"
	GetVersion = "2010.09.01 集計表(品番/在庫)  集計表(品番/在庫集計) 在庫数の不具合修正"
	GetVersion = "2011.04.10 一覧表 更新日時,完了時刻 追加"
	GetVersion = "2011.07.01 向け先がマスター未登録の場合でも、出荷先コードを表示するように修正"
	GetVersion = "2011.08.30 集計表(品番/在庫) 供給区分(国内・海外)／ユニット区分を追加"
	GetVersion = "2012.04.06 検索条件 品番(対外) の前方一致検索を解除"
	GetVersion = "<font color=red>2012.10.17 集計表(品番/在庫) の棚番を標準棚番に変更</font>"
	GetVersion = "2013.02.26 集計表(商品化請求内訳) に 事業部 を追加"
	GetVersion = "2015.06.29 名称変更:向け先→出荷先、出荷先で直送先も検索するように変更、一覧表の出荷先に海外直送先を追加"
	GetVersion = "2015.06.29 追加：一覧表(外装入数),集計表(外装入数)"
	GetVersion = "2016.10.03 検索条件(備考)追加"
	GetVersion = "2016.10.13 出荷先名：向け先マスターを参照しないように変更"
	GetVersion = "2016.10.14 出荷先：産機直送の出荷先名を表示／直送：産機直送を「1 産機」表示ｊ"
	GetVersion = "2016.10.28 才数 の精度Up：才数テーブル(ItemSize)参照"
	GetVersion = "2016.10.31 集計表(品番/在庫)：ソート順変更(事/棚番/品番)"
	GetVersion = "2017.01.27 集計表(注区別/直送件数)：産機と直送の合計を集計"
	GetVersion = "2017.04.11 産機 直送先の検索対応"
	GetVersion = "2017.09.24 出庫表"
	GetVersion = "2019.06.05 検索条件：更新日時"
	GetVersion = "2019.11.01 検品OK解除：JITU_SURYOに0をセットするように修正"
	GetVersion = "2020.03.23 クリップボードコピー対応(IE以外)"
	GetVersion = "2020.07.21 事業部の検索不具合修正"
	GetVersion = "2020.07.21 クリップボードコピー範囲変更(検索日時を除外)"
	GetVersion = "2020.08.18 出庫テスト用"
End Function

'----------------------------------------------------------
'select from テーブル
'----------------------------------------------------------
Function GetFrom(byVal strName)
	GetFrom = ""
	select case strName
	case "Item"
		GetFrom = " left outer join Item i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
	case "ItemSize"
		GetFrom = " left outer join ItemSize iSize on (y.jgyobu = iSize.jgyobu and y.key_hin_no = iSize.hin_gai)"
	case "Zaiko"
		GetFrom = GetFrom & " left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(if(GOODS_ON = '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) sumi_qty,sum(if(GOODS_ON <> '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) mi_qty from zaiko group by JGYOBU,NAIGAI,HIN_GAI)"
		GetFrom = GetFrom & " z on ("
		GetFrom = GetFrom & " z.JGYOBU = y.JGYOBU and"
		GetFrom = GetFrom & " z.NAIGAI = y.NAIGAI and"
		GetFrom = GetFrom & " z.HIN_GAI = y.KEY_HIN_NO)"
	case "Zaiko.Tana"
		GetFrom = GetFrom & " left outer join (select JGYOBU,NAIGAI,HIN_GAI,Soko_No+Retu+Ren+Dan Tana,sum(if(GOODS_ON = '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) sumi_qty,sum(if(GOODS_ON <> '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) mi_qty from zaiko group by JGYOBU,NAIGAI,HIN_GAI,Tana)"
		GetFrom = GetFrom & " z on ("
		GetFrom = GetFrom & " z.JGYOBU = y.JGYOBU and"
		GetFrom = GetFrom & " z.NAIGAI = y.NAIGAI and"
		GetFrom = GetFrom & " z.HIN_GAI = y.KEY_HIN_NO)"
	case "HtDrctId"
		GetFrom = " left outer join HtDrctId d on (d.IDNo = y.KEY_ID_NO)"
	case "HMTAH015"
		GetFrom = " left outer join HMTAH015 g on (g.IDNo = y.KEY_ID_NO)"
	case "Mts"
		GetFrom = " left outer join Mts m on (m.MUKE_CODE = y.KEY_MUKE_CODE)"
	case else
		GetFrom = " from " & strName & " y"
	end select
End Function
'----------------------------------------------------------
'select フィールド
'----------------------------------------------------------
Function GetSelect(byVal strName)
	GetSelect = ""
	select case strName
	case "出荷日"
		GetSelect = GetSelect & "KEY_SYUKA_YMD"
	case "事業場"
		GetSelect = GetSelect & "JGYOBA"
	case "事業部","事"
		GetSelect = GetSelect & "y.JGYOBU"
	case "在庫<br>収支"
		GetSelect = GetSelect & "y.SYUKO_SYUSI"
	case "ID","ID_BC"
		GetSelect = GetSelect & "y.KEY_ID_NO"
	case "伝票No"
		GetSelect = GetSelect & "y.DEN_NO"
	case "オーダーNo"
		GetSelect = GetSelect & "y.ODER_NO"
	case "アイテムNo"
		GetSelect = GetSelect & "y.ITEM_NO"
	case "品番","品番_BC"
		GetSelect = GetSelect & "y.KEY_HIN_NO"
	case "品名"
		GetSelect = GetSelect & "y.HIN_NAME"
	case "数量"
		GetSelect = GetSelect & "convert(y.SURYO,SQL_DECIMAL)"
	case "(出庫済)"
		GetSelect = GetSelect & "convert(y.JITU_SURYO,SQL_INTEGER)"
	case "直送"
		GetSelect = GetSelect & "CHOKU_KBN + if(CHOKU_KBN='1',if(y.LK_MUKE_CODE = '00027768',' 産機',' 直送'),'')"
	case "出荷先"
		GetSelect = GetSelect & "if(ifnull(d.ChoCode,'')=''"
		GetSelect = GetSelect & " ,LK_MUKE_CODE + ' ' + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME)"
		GetSelect = GetSelect & " ,d.ChoCode + ' ' + d.ChoName)"
	case "出荷先LK"
		GetSelect = GetSelect & " y.KEY_MUKE_CODE + ' ' + y.MUKE_NAME + '<br>' + y.LK_MUKE_CODE"
	case "出荷先_BC"
		GetSelect = GetSelect & " y.LK_MUKE_CODE"
	case "出荷先数<br>"
		GetSelect = GetSelect & "count(distinct if(ifnull(d.ChoCode,'')='',LK_MUKE_CODE + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME),d.ChoCode + d.ChoName))"
	case "出荷先 直送集計"
'		GetSelect = GetSelect & "if(CHOKU_KBN<>'1' or ifnull(d.ChoCode,'')<>'',LK_MUKE_CODE + ' ' + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME),'その他')"
		GetSelect = GetSelect & "if(CHOKU_KBN = '1',if(LK_MUKE_CODE = '00027768','P産機','その他'),LK_MUKE_CODE + ' ' + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME))"
	case "15時以降","データ<br>受信時刻"
		GetSelect = GetSelect & "if(left(INS_NOW,10) < KEY_SYUKA_YMD + '15','','15時以降')"
	case "取引<br>区分"
		GetSelect = GetSelect & "TORI_KBN"
		GetSelect = GetSelect & "+case TORI_KBN"
		GetSelect = GetSelect & " when '25' then ' 売上'"
		GetSelect = GetSelect & " when '29' then ' 振替出庫'"
		GetSelect = GetSelect & " when '19' then ' 振替入庫'"
		GetSelect = GetSelect & " end"
	case "注文<br>区分","注文区分"
		GetSelect = GetSelect & "KEY_CYU_KBN + if(KEY_CYU_KBN = '1',' 月切',if(KEY_CYU_KBN = '2',' 緊急',if(KEY_CYU_KBN = '3',' 補充',if(KEY_CYU_KBN = 'E',' 貿易',''))))"
	case "欠品解除"
		GetSelect = GetSelect & "KEPIN_KAIJYO"
	case "出荷件数"
		GetSelect = GetSelect & "count(distinct y.KEY_ID_NO)"
	case "出荷数"
		GetSelect = GetSelect & "sum(convert(y.SURYO,SQL_DECIMAL))"
	case "商品化(済)"
		GetSelect = GetSelect & "z.sumi_qty"
	case "商品化(未)"
		GetSelect = GetSelect & "z.mi_qty"
	case "要商品化数"
		GetSelect = GetSelect & "if(sum(convert(y.SURYO,SQL_DECIMAL)) >= z.sumi_qty,sum(convert(y.SURYO,SQL_DECIMAL)) - z.sumi_qty,0)"
	case "国内供給区分"
		GetSelect = GetSelect & "i.NAI_BUHIN + ' ' + case i.NAI_BUHIN when '1' then '対象' when '2' then '打切案内中'  when '3' then '打切' when '3' then '単品ユニット' else ''  end"
	case "海外供給区分"
		GetSelect = GetSelect & "i.GAI_BUHIN + ' ' + case i.GAI_BUHIN when '1' then '対象' when '2' then '打切案内中'  when '3' then '打切' when '3' then '単品ユニット' else ''  end"
	case "ユニット区分"
		GetSelect = GetSelect & "i.UNIT_BUHIN + ' ' + case i.UNIT_BUHIN when '0' then '単品' when '1' then 'ユニット親'  when '2' then 'ユニット子' when '3' then '単品ユニット' else '' end"
	case "件数","合計<br>"
		GetSelect = GetSelect & "count(*)"
	case "月切"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '1',1,0))"
	case "月切<br>検品済"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '1' and KENPIN_YMD <> '',1,0))"
		strName = "<br>検品済"
	case "緊急"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '2',1,0))"
	case "緊急<br>検品済"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '2' and KENPIN_YMD <> '',1,0))"
		strName = "<br>検品済"
	case "補充"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '3',1,0))"
	case "補充<br>検品済"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '3' and KENPIN_YMD <> '',1,0))"
		strName = "<br>検品済"
	case "その他"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN not in ('1','2','3'),1,0))"
	case "その他<br>検品済"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN not in ('1','2','3') and KENPIN_YMD <> '',1,0))"
		strName = "<br>検品済"
	case "span1"
		GetSelect = GetSelect & "' '"
	case "出庫<br>済"
		GetSelect = GetSelect & "sum(if(KAN_YMD <> '',1,0))"
	case "出庫<br>残"
		GetSelect = GetSelect & "sum(if(KAN_YMD <> '',0,1))"
	case "検品<br>済","<br>検品済"
		GetSelect = GetSelect & "sum(if(KENPIN_YMD <> '',1,0))"
	case "検品<br>残"
		GetSelect = GetSelect & "sum(if(KENPIN_YMD <> '',0,1))"
	case "送信<br>残"
		GetSelect = GetSelect & "sum(if(g.IDno is null or KEY_CYU_KBN = 'E' or RTrim(LK_SEQ_NO)<>'',0,1))"
	case "実績<br>残"
		GetSelect = GetSelect & "sum(if(g.IDno is null,0,1))"
	case "前日"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8),1,0))"
	case "09時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) <= '09',1,0))"
	case "10時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('10','11'),1,0))"
	case "12時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('12'),1,0))"
	case "13時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('13'),1,0))"
	case "14時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('14'),1,0))"
	case "15時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('15'),1,0))"
	case "16時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('16'),1,0))"
	case "17時"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) >= '17',1,0))"
	case "個数"
		GetSelect = GetSelect & "sum(convert(SURYO,SQL_DECIMAL))"
	case "才数"
		GetSelect = GetSelect & "sum(iSize.Size * convert(y.SURYO,SQL_DECIMAL))"
	case ".才数"
		GetSelect = GetSelect & "iSize.Size * convert(y.SURYO,SQL_DECIMAL)"
	case "xx才数"
		GetSelect = GetSelect & "sum(convert(i.SAI_SU,SQL_DECIMAL)*convert(SURYO,SQL_DECIMAL))"
	case ".xx才数"
		GetSelect = GetSelect & "convert(i.SAI_SU,SQL_DECIMAL)*convert(y.SURYO,SQL_DECIMAL)"
		strName = "才数"
	case "備考1"
		GetSelect = GetSelect & "y.BIKOU1"
	case "備考2"
		GetSelect = GetSelect & "y.BIKOU2"
	case "完了区分"
		GetSelect = GetSelect & "y.KAN_KBN"
	case "出庫日"
		GetSelect = GetSelect & "y.KAN_YMD"
	case "出庫時刻"
		GetSelect = GetSelect & "y.KAN_HMS"
	case "出庫日時"
		GetSelect = GetSelect & "y.KAN_YMD + '-' + left(y.KAN_HMS,4)"
	case "検品日時"
		GetSelect = GetSelect & "y.KENPIN_YMD + '-' + left(y.KENPIN_HMS,4)"
	case "検品担当者"
		GetSelect = GetSelect & "y.KENPIN_TANTO_CODE"
	case "事前チェック"
		GetSelect = GetSelect & "y.LK_SEQ_NO"
	case "移管状況"
		GetSelect = GetSelect & "i.BIKOU_TANA"
	case "検品メッセージ"
		GetSelect = GetSelect & "i.INSP_MESSAGE"
	case "標準棚番"
		GetSelect = GetSelect & "rtrim(i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN)"
	case "棚番"
		GetSelect = GetSelect & "y.HTANABAN"
	case "標準棚番"
		GetSelect = GetSelect & "rtrim(i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN)"
	case "棚番_BC"
		GetSelect = GetSelect & "if(rtrim(i.ST_SOKO)<>'',rtrim(i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN),y.HTANABAN)"
	case "z棚番_BC"
		GetSelect = GetSelect & "'.' + z.Tana"
	case "登録日時"
		GetSelect = GetSelect & "left(y.INS_NOW,8) + '-' + substring(y.INS_NOW,9,4)"
	case "更新日時"
		GetSelect = GetSelect & "left(y.UPD_NOW,8) + '-' + substring(y.UPD_NOW,9,4)"
	case "データ区分"
		GetSelect = GetSelect & "y.DATA_KBN"
	case "販売区分"
		GetSelect = GetSelect & "y.HAN_KBN"
	case "直送先"
		GetSelect = GetSelect & "y.LK_MUKE_CODE"
	case "出庫表印刷"
		GetSelect = GetSelect & "y.PRINT_YMD"
	end select

	GetSelect = GetSelect & " """ & strName & """"
End Function

server.ScriptTimeout = 900
%>
<!--#include file="info.txt" -->
<!--#include file="makeWhere.asp" -->
<HTML>
<HEAD>
<%
	dim	fname
	dim	tblStr
	dim	ptypeStr
	dim	delStr
	dim	dtStr
	dim	dtToStr
	dim	KEY_MUKE_CODEStr
	dim	KEY_CYU_KBNStr
	dim	DATA_KBNStr
	dim	HAN_KBNStr
	dim	JGYOBUStr
	dim	JGYOBAStr
	dim	TOK_KBNStr
	dim	KEY_ID_NOStr
	dim	DEN_NOStr
	dim	pnStr
	dim	KAN_KBNStr
	dim	KEN_KBNStr
	dim	SEI_KBNStr
	dim	SYUKO_SYUSIStr
	dim	SS_CODEStr
	dim	whereStr
	dim	andStr
	dim	sqlStr
	dim	strTh
	dim	cnt
	dim	i
	dim	db
	dim	rsList
	dim	centerStr
	dim	submitStr
	dim	autoStr
	dim	fValue
	dim	tdTag
	dim	TORI_KBNStr
	dim sTankaStr
	dim	KEPIN_KAIJYOStr
	dim	adminStr
	dim	strShimuke
	dim	INSP_MESSAGEStr
	dim	cmpStr
	dim insStr
	dim	lngMax
	dim	maxStr

	insStr = ""

	INSP_MESSAGEStr		= Request.QueryString("INSP_MESSAGE")
	adminStr			= Request.QueryString("admin")
	submitStr			= Request.QueryString("submit1")
	autoStr = 0
	
	tblStr = Request.QueryString("tbl")
	if len(tblStr) = 0 then
		tblStr = "Y_Syuka"
	end if
	if tblStr = "Y_Syuka" then
		delStr = ""
	else
		delStr = "削除済"
	end if
	ptypeStr = Request.QueryString("ptype")
	if len(ptypeStr) = 0 then
		ptypeStr = "pTable"
	end if
	dtStr			= Request.QueryString("dt")
	dtToStr 		= Request.QueryString("dtTo")
	sTankaStr		= ucase(Request.QueryString("sTanka"))
	TORI_KBNStr		= ucase(Request.QueryString("TORI_KBN"))
	KEY_MUKE_CODEStr	= ucase(Request.QueryString("KEY_MUKE_CODE"))
	KEY_CYU_KBNStr		= ucase(Request.QueryString("KEY_CYU_KBN"))
	DATA_KBNStr		= ucase(Request.QueryString("DATA_KBN"))
	HAN_KBNStr		= ucase(Request.QueryString("HAN_KBN"))
	JGYOBUStr		= ucase(Request.QueryString("JGYOBU"))
	JGYOBAStr		= ucase(Request.QueryString("JGYOBA"))
	TOK_KBNStr		= ucase(Request.QueryString("TOK_KBN"))
	KEY_ID_NOStr		= ucase(Request.QueryString("KEY_ID_NO"))
	DEN_NOStr		= ucase(Request.QueryString("DEN_NO"))
	pnStr			= ucase(Request.QueryString("pn"))
	KAN_KBNStr		= ucase(Request.QueryString("KAN_KBN"))
	KEN_KBNStr		= ucase(Request.QueryString("KEN_KBN"))
	SEI_KBNStr		= ucase(Request.QueryString("SEI_KBN"))
	SYUKO_SYUSIStr		= ucase(Request.QueryString("SYUKO_SYUSI"))
	KEPIN_KAIJYOStr		= ucase(Request.QueryString("KEPIN_KAIJYO"))
	SS_CODEStr		= ucase(Request.QueryString("SS_CODE"))

	submitStr		= Request.QueryString("submit1")

	autoStr = 0
	if len(submitStr) > 0 and ptypeStr = "pTable" and tblStr = "Y_Syuka" then
		autoStr = Request.QueryString("auto")
		if len(autoStr) = 0 then
			autoStr = 5
		end if
	end if
	maxStr			= ucase(Request.QueryString("max"))
	if maxStr = "" then
		maxStr = 100
	end if
	lngMax = clng(maxStr)
%>
<%if autoStr > 0 then%>
	<!--meta http-equiv="refresh" content="<%=autoStr * 60%>"-->
<%end if %>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL="STYLESHEET" TYPE="text/css" HREF="result.css" TITLE="CSS" media="all">
<LINK REL="STYLESHEET" TYPE="text/css" HREF="print.css" TITLE="CSS" media="print">
<TITLE><%=centerStr%> 出荷予定</TITLE>
<!-- jdMenu head用 include 開始 -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="https://cdn.jsdelivr.net/clipboard.js/1.5.3/clipboard.min.js"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<script src="./clipboard.js" type="text/javascript"></script>
<!-- jdMenu head用 include 終了 -->
<SCRIPT LANGUAGE="JavaScript"><!--
navi = navigator.userAgent;

function DoCopy(arg){
	var doc = document.body.createTextRange();
	doc.moveToElementText(document.all(arg));
	doc.execCommand("copy");
	window.alert("クリップボードへコピーしました。\n貼り付けできます。" );
}

function DW(n,bname){
    if(navi.indexOf('MSIE') <=0 && navi.indexOf('Windows')>=0){
        return document.write("<INPUT TYPE='button' onClick=\"DoCopy('COPYAREA" + n + "')\" VALUE='" + bname + "'>");
    }
}
	function ptypeChange(typ) {
//		sqlForm.ptype[typ].checked = "true";
		for(var	i = 0;i < document.sqlForm.elements.length;i++) {
			if ( document.sqlForm.elements[i].id == typ ) {
				document.sqlForm.elements[i].checked = "true"
				break;
			}
		}
	}
	function tblChange(typ) {
//		sqlForm.tbl[typ].checked = "true";
		for(var	i = 0;i < document.sqlForm.elements.length;i++) {
			if ( document.sqlForm.elements[i].id == typ ) {
				document.sqlForm.elements[i].checked = "true"
				break;
			}
		}
	}

function autoChange() {
	if(autoBtn.value == "On") {
		window.alert("自動更新" + autoBtn.value + "→Off");
//		DispMsg("自動更新" + autoBtn.value + "→Off");
		autoBtn.value = "Off";
		autoValue.value = "On";
	} else {
		window.alert("自動更新" + autoBtn.value + "→On");
//		DispMsg("自動更新" + autoBtn.value + "→On");
		autoBtn.value = "On";
		autoValue.value = "Off";
	}
}

	function uSyukoClick() {
		if ( window.confirm("出庫OKにします") == false ) {
			ptypeChange("pTable");
		}
	}
	function uSyukoClickCancel() {
		if ( window.confirm("出庫OKを解除にします") == false ) {
			ptypeChange("pTable");
		}
	}
	function uKenpinClick() {
		if ( window.confirm("検品OKにします") == false ) {
			ptypeChange("pTable");
		}
	}
	function uKenpinCancelClick() {
		if ( window.confirm("検品OKを解除します") == false ) {
			ptypeChange("pTable");
		}
	}
	function DeleteClick() {
		if ( window.confirm("データを削除します\n＊元に戻せませんが＊よろしいですか？") == false ) {
			ptypeChange("pTable");
		}
	}
	function DelToYClick() {
		if ( window.confirm("削除済データを予定に戻します\n＊元に戻せませんが＊よろしいですか？") == false ) {
			ptypeChange("pTable");
		}
	}
	function uClick(msg) {
		if ( window.confirm(msg) == false ) {
			ptypeChange("pTable");
		}
	}

  function showhide(id){
    if(document.getElementById){
      if(document.getElementById(id).style.display != "none"){
        document.getElementById(id).style.display = "none";
      }else{
        document.getElementById(id).style.display = "block";
      }
    }
  }
--></SCRIPT>
</HEAD>
<BODY>
<!-- jdMenu body用 include 開始 -->
<!--#include file="jdmenu-sdc.asp" -->
<!-- jdMenu body用 include 終了 -->
<%
	if len(submitStr) + len(dtStr) = 0 then
		dtStr = "today"
		select case centerStr
		case "小野PC","滋賀PC"
			HAN_KBNStr = "1"
			DATA_KBNStr = "1"
		end select
	end if
	if dtStr = "today" then
		dtStr = year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2)
	end if
%>
  <FORM name="sqlForm">
  <div id="sqlDiv">
	<div><%=centerStr%> 出荷予定検索 <%=delStr%></div>
	<table id="sqlTbl">
		<tr>
			<th>伝票日付</th>
			<th title="事業場コード">事業場</th>
			<th>取引区分</th>
			<th>DATA区分</th>
			<th>販売区分</th>
			<th>注文区分</th>
			<th>事業部</th>
			<th>直送</th>
			<th>出荷先</th>
			<th>ID-No</th>
			<th>伝票No</th>
			<th>品番(対外)</th>
			<th>完了区分</th>
			<th>検品区分</th>
			<th>請求対象</th><!-- 2009.06.17 名称変更 -->
			<th>単価設定</th>
			<th>欠品解除</th>
			<th>在庫収支</th>
			<!--th>SS<br>(トラックNo)</th-->
			<th>検品メッセージ</th>
			<th>備考</th>
			<th>登録日時</th>
			<th>更新日時</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="dt" id="dt" VALUE="<%=dtStr%>" size="10" maxlength="8"><br>
				～<br>
				<INPUT TYPE="text" NAME="dtTo" id="dtTo" VALUE="<%=dtToStr%>" size="10" maxlength="8">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="JGYOBA" id="JGYOBA" VALUE="<%=JGYOBAStr%>" size="10" maxlength="8" style="text-align:left;">
				<div style="text-align:left;">
					<font size="-2">00036003：AP社CS</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="TORI_KBN" id="TORI_KBN" VALUE="<%=TORI_KBNStr%>" size="5" maxlength="2" style="text-align:left;">
				<div style="text-align:left;">
					<font size="-2">25:売上</font><br>
					<font size="-2">29:振替出庫</font>
					<font size="-2">19:振替入庫</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="DATA_KBN" id="DATA_KBN" VALUE="<%=DATA_KBNStr%>" size="4" maxlength="" style="text-align:center;">
				<div style="text-align:left;">
				<font size="-2">1：売上<br>3：振替<br>7：科目振替</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="HAN_KBN" id="HAN_KBN" VALUE="<%=HAN_KBNStr%>" size="4" maxlength="3" style="text-align:center;">
				<div style="text-align:left;">
				<font size="-2">1：国内<br>2：輸出</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEY_CYU_KBN" id="KEY_CYU_KBN" VALUE="<%=KEY_CYU_KBNStr%>" size="8" maxl3ength="8" style="text-align:center;">
				<div style="text-align:left;">
				<font size="-2">1：月切<br>2：緊急<br>3：補充<br>E：貿易<br>1,2,3：貿易除く</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="JGYOBU" id="JGYOBU" VALUE="<%=JGYOBUStr%>" size="4" style="text-align:center;">
				<!--div style="text-align:left;">
				1:ﾗﾝﾄﾞﾘｰBU<br>
				4:CABU<br>
				D:IHBU<br>
				7:ｸﾘｰﾅｰBU<br>
				A:ｴｱｺﾝBU
				</div-->
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="TOK_KBN" id="TOK_KBN" VALUE="<%=TOK_KBNStr%>" size="1" maxlength="1" style="text-align:center;"><br>
				1：直送
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEY_MUKE_CODE" id="KEY_MUKE_CODE" VALUE="<%=KEY_MUKE_CODEStr%>" size="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEY_ID_NO" id="KEY_ID_NO" VALUE="<%=KEY_ID_NOStr%>" size="15" maxlength="12">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="DEN_NO" id="DEN_NO" VALUE="<%=DEN_NOStr%>" size="10" maxlength="8">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="pn" id="pn" VALUE="<%=pnStr%>" size="20">
				<div style="text-align:left;"><font size="-2">
				%：あいまい検索<br>
				　例 AZC81%
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KAN_KBN" id="KAN_KBN" VALUE="<%=KAN_KBNStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0：未出庫<br>
				9：出庫済<br>
				=：
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEN_KBN" id="KEN_KBN" VALUE="<%=KEN_KBNStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0：未検品<br>
				9：検品済
				</div>
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="SEI_KBN" id="SEI_KBN" VALUE="<%=SEI_KBNStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0：請求対象外<br>
				1：請求対象
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="sTanka" id="sTanka" VALUE="<%=sTankaStr%>" size="1" maxlength="1"><br>
				<div style="text-align:left;"><font size="-2">
					0:単価未登録<br>
					1:単価登録済<br>
					2:単価0以上<br>
				</font></div>
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="KEPIN_KAIJYO" id="KEPIN_KAIJYO" VALUE="<%=KEPIN_KAIJYOStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0：通常引当<br>
				1：欠品解除
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="SYUKO_SYUSI" id="SYUKO_SYUSI" VALUE="<%=SYUKO_SYUSIStr%>" size="3" maxlength="3" style="text-align:left;"><br>
			</td>
			<!--td align="center">
				<INPUT TYPE="text" NAME="SS_CODE" id="SS_CODE" VALUE="<%=SS_CODEStr%>" size="10" maxlength="5">
			</td-->
			<td align="center">
				<INPUT TYPE="text" NAME="INSP_MESSAGE" id="INSP_MESSAGE" VALUE="<%=INSP_MESSAGEStr%>" size="20">
				<div style="text-align:left;"><font size="-2">
				-：検品メッセージあり
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="BIKOU1" VALUE="<%=GetRequest("BIKOU1","")%>" size="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="INS_NOW" id="INS_NOW" VALUE="<%=Request.QueryString("INS_NOW")%>" size="14" style="text-align:left;">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="UPD_NOW" id="UPD_NOW" VALUE="<%=Request.QueryString("UPD_NOW")%>" size="14" style="text-align:left;">
			</td>
		</tr>
		<tr>
			<td align="right">
				<b>出力形式：</b>
			</td>
			<td colspan="21" nowrap>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable" id="pTable">
					<label for="pTable">集計表(注区別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableCnt" id="pTableCnt">
					<label for="pTableCnt">集計表(注区別/直送集計)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable1500" id="pTable1500">
					<label for="pTable1500">集計表(注区別/15:00前後)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable2" id="pTable2">
					<label for="pTable2">集計表(時間別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable3" id="pTable3">
					<label for="pTable3">集計表(出荷日別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnMonth" id="pTablePnMonth">
					<label for="pTablePnMonth">集計表(品番/月別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSMonth" id="pTableSMonth">
					<label for="pTableSMonth">集計表(出荷先/月別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMonth" id="pTableMonth">
					<label for="pTableMonth">集計表(月別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableDay" id="pTableDay">
					<label for="pTableDay">集計表(時間別件数個数)</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSoko" id="pTableSoko">
					<label for="pTableSoko">集計表(倉庫)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable4" id="pTable4">
					<label for="pTable4">集計表(品番/出荷先)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnZaiko" id="pTablePnZaiko">
					<label for="pTablePnZaiko">集計表(品番/在庫)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnSumzai" id="pTablePnSumzai">
					<label for="pTablePnSumzai">集計表(品番/在庫集計)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable5" id="pTable5">
					<label for="pTable5">集計表(日別／出荷先別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableDaySaki" id="pTableDaySaki">
					<label for="pTableDaySaki">集計表(出荷日別／出荷先別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableDayTime" id="pTableDayTime">
					<label for="pTableDayTime">集計表(出荷日・受信日・時刻別)</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
					<label for="pList">一覧表</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableIri" id="pTableIri">
					<label for="pTableIri">集計表(外装入数)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListIri" id="pListIri">
					<label for="pListIri">一覧表(外装入数)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pPicking" id="pPicking">
					<label for="pPicking">出庫表</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
					<label for="pListAll">一覧表(全項目)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSeikyu" id="pTableSeikyu">
					<label for="pTableSeikyu">集計表(商品化請求内訳)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSeikyuKaigai" id="pTableSeikyuKaigai">
					<label for="pTableSeikyuKaigai">集計表(商品化請求内訳/国内外単価)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableShizai" id="pTableShizai">
					<label for="pTableShizai">集計表(資材)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableOrderNo" id="pTableOrderNo">
					<label for="pTableOrderNo">集計表(オーダーNo)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnJumpUp" id="pTablePnJumpUp">
					<label for="pTablePnJumpUp">集計表(品番 出荷増)</label>
<% if adminStr = "admin" then %>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTest" id="pTest">
					<label for="pTest">出庫テスト</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pPickingTest" id="pPickingTest">
					<label for="pPickingTest">検品テスト</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uSyuko" id="uSyuko" onclick="uSyukoClick();">
					<label for="uSyuko">出庫OK</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uSyukoCancel" id="uSyukoCancel" onclick="uSyukoClickCancel();">
					<label for="uSyukoCancel">出庫OK解除</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uKenpin" id="uKenpin" onclick="uKenpinClick();">
					<label for="uKenpin">検品OK</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uKenpinCancel" id="uKenpinCancel" onclick="uKenpinCancelClick();">
					<label for="uKenpinCancel">検品OK解除</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uJizenCancel" id="uJizenCancel" onclick="uClick('事前チェック済を解除します');">
					<label for="uJizenCancel">事前チェック解除</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uPrintCancel" id="uPrintCancel">
					<label for="uPrintCancel">出庫表OK解除</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="dData" id="dData" onclick="DeleteClick();">
					<label for="dData">削除</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="DelToY" id="DelToY" onclick="DelToYClick();">
					<label for="DelToY">削除済→予定戻し</label>
<% end if	%>
			</td>
		</tr>
		<tr>
			<td align="right">
				<b>対象：</b>
			</td>
			<td colspan="21">
				<INPUT TYPE="radio" NAME="tbl" VALUE="Y_Syuka" id="Y_Syuka">
					<label for="Y_Syuka"><b>出荷予定(Y_Syuka)</b></label>
				<INPUT TYPE="radio" NAME="tbl" VALUE="DEL_SYUKA" id="DEL_SYUKA">
					<label for="DEL_SYUKA"><b>削除済(DEL_SYUKA)</b></label>
			</td>
		</tr>
		<tr bordercolor="White">
			<td colspan="22">
			<INPUT TYPE="submit" value="検索" id=submit1 name=submit1>
			<INPUT TYPE="reset" value="リセット" id=reset1 name=reset1 onClick="location.href='y_syuka.asp?tbl=<%=tblStr%>';">
				最大件数：<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=maxStr%>" size="8">
			<%=GetVersion()%>
	<%		if len(submitStr) > 0 and ptypeStr = "pTable" and tblStr = "Y_Syuka" then	%>
				<!--span>　　　　　　　　　　　　　自動更新 
				<INPUT TYPE="text" NAME="auto" id="auto" VALUE="<%=autoStr%>" size="2" style="text-align : right;">
				分</span-->
	<%		end if	%>
			<INPUT TYPE="hidden" NAME="admin" id="admin" VALUE="<%=adminStr%>">
			</td>
		</tr>
	</table>
	</div>
  </FORM>
<SCRIPT LANGUAGE='JavaScript'>ptypeChange('<%=ptypeStr%>');</SCRIPT>
<SCRIPT LANGUAGE='JavaScript'>ptypeChange('<%=tblStr%>');</SCRIPT>
<%	if len(submitStr) > 0 then %>
	<SCRIPT LANGUAGE=javascript><!--
		sqlForm.disabled = true;
	//--></SCRIPT>
	<div>
		<!--INPUT TYPE="button" onClick="DoCopy('resultDiv')"
			value="検索中...ScriptTimeout=<%=Server.ScriptTimeout%>"
			id="cpTblBtn" disabled-->
		<button id="btnClip" class="btn" data-clipboard-target="#resultTbl" disabled onClick="DoCopy('resultDiv');">
			検索中...ScriptTimeout=<%=Server.ScriptTimeout%>
		</button>
	</div>
	<div id='resultDiv'>
	<div><%=now%> 現在</div>
	<table id="resultTbl" class="<%=ptypeStr%>">
	<%
		dim	strYMD
		dim strHMS

		Set db = Server.CreateObject("ADODB.Connection")
		db.open GetDbName()
		sqlStr = ""
		whereStr = ""
		andStr = " where"
		if ptypeStr = "pTablePnJumpUp" then
			whereStr = whereStr & andStr & " KEY_SYUKA_YMD = '" & dtToStr & "'"
			dtToStr = clng(dtToStr) - 1 & ""
			andStr = " and"
		else
			if len(dtStr) + len(dtToStr) > 0 then
				if len(dtStr) > 0 and len(dtToStr) > 0 then
					whereStr = whereStr & andStr & " (KEY_SYUKA_YMD between '" & dtStr & "' and '" & dtToStr & "')"
				else
					whereStr = whereStr & andStr & " KEY_SYUKA_YMD like '" & dtStr & "%'"
				end if
				andStr = " and"
			end if
		end if
		if len(sTankaStr) > 0 then
			select case left(sTankaStr,1)
			case "0"
'				whereStr = whereStr & andStr & " rtrim(i.S_KOUSU_BAIKA) = ''"
				whereStr = whereStr & andStr & " rtrim(i.S_KOUSU_SET_DATE) = ''"
			case "1"
'				whereStr = whereStr & andStr & " rtrim(i.S_KOUSU_BAIKA) <> ''"
				whereStr = whereStr & andStr & " rtrim(i.S_KOUSU_SET_DATE) <> ''"
			case "2"
'				whereStr = whereStr & andStr & " (convert(ifnull(rtrim(i.S_KOUSU_BAIKA),'0'),SQL_DECIMAL) > 0 or convert(ifnull(rtrim(i.S_SHIZAI_BAIKA),'0'),SQL_DECIMAL) > 0)"
				whereStr = whereStr & andStr & " (convert(i.S_KOUSU_BAIKA,SQL_DECIMAL) > 0 or convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL) > 0)"
			end select
			andStr = " and"
		end if
   		whereStr = makeWhere(whereStr,"JGYOBA",JGYOBAStr,"")
   		whereStr = makeWhere(whereStr,"TORI_KBN",TORI_KBNStr,"")
   		whereStr = makeWhere(whereStr,"KEY_CYU_KBN",KEY_CYU_KBNStr,"")
   		whereStr = makeWhere(whereStr,"y.DATA_KBN",DATA_KBNStr,"")
   		whereStr = makeWhere(whereStr,"HAN_KBN",HAN_KBNStr,"")
   		whereStr = makeWhere(whereStr,"CHOKU_KBN",TOK_KBNStr,"")
		whereStr = makeWhere(whereStr,"BIKOU1",GetRequest("BIKOU1",""),"")
		whereStr = makeWhere(whereStr,"y.JGYOBU",GetRequest("JGYOBU",""),"")

		if len(JGYOBUStr) > 0 and False then	'2020.07.21 Falseコメント
			cmpStr = " = "
			if left(JGYOBUStr,1) = "-" then
				JGYOBUStr = right(JGYOBUStr,len(JGYOBUStr)-1)
				cmpStr = " <> "
			end if
			if instr(JGYOBUStr,",") > 0 then
				JGYOBUStr = trim(JGYOBUStr)
				JGYOBUStr = "('" & replace(JGYOBUStr,",","','") & "')"
				if trim(cmpStr) = "=" then
					cmpStr = " in "
				else
					cmpStr = " not in "
				end if
			else
				JGYOBUStr = "'" & JGYOBUStr & "'"
				select case JGYOBUStr
				case "1"
							strShimuke = "04"
				case "D"
							strShimuke = "01"
				case "4"
							strShimuke = "02"
				case "7"
							strShimuke = "01"
				case else
							strShimuke = "01"
				end select
			end if
			whereStr = whereStr & andStr & " y.JGYOBU " & cmpStr & JGYOBUStr
			andStr = " and"
		end if

		if len(KAN_KBNStr) > 0 then
			if left(KAN_KBNStr,1) = "-" then
				whereStr = andWhere(whereStr) & " KAN_KBN <> '" & right(KAN_KBNStr,1) & "'"
			elseif KAN_KBNStr = "=" then
				whereStr = whereStr & andStr & " KAN_KBN = '0'"
				whereStr = whereStr &          " and convert(SURYO,SQL_INTEGER) = convert(JITU_SURYO,SQL_INTEGER)"
			else
				whereStr = whereStr & andStr & " KAN_KBN = '" & KAN_KBNStr & "'"
			end if
			andStr = " and"
		end if
		if len(KEN_KBNStr) > 0 then
			if KEN_KBNStr = "0" then
				whereStr = whereStr & andStr & " KENPIN_YMD = ''"
			else
				whereStr = whereStr & andStr & " KENPIN_YMD <> ''"
			end if
			andStr = " and"
		end if
		if len(SEI_KBNStr) > 0 then
			select case SEI_KBNStr
			case "0"	' 請求対象外
				whereStr = whereStr & andStr & " ((HAN_KBN = '1' and LK_SEQ_NO = '') or (HAN_KBN = '2' and KAN_KBN <> '9'))"
			case "1"	' 請求対象
				whereStr = whereStr & andStr & " ((HAN_KBN = '1' and LK_SEQ_NO <> '') or (HAN_KBN = '2' and KAN_KBN = '9'))"
			end select
			andStr = " and"
		end if
		if len(KEPIN_KAIJYOStr) > 0 then
			select case KEPIN_KAIJYOStr
			case "0"	' 通常引当
				whereStr = whereStr & andStr & " KEPIN_KAIJYO <> '1'"
			case "1"	' 欠品解除
				whereStr = whereStr & andStr & " KEPIN_KAIJYO = '1'"
			end select
			andStr = " and"
		end if
   		whereStr = makeWhere(whereStr,"SYUKO_SYUSI",SYUKO_SYUSIStr,"")
   		whereStr = makeWhere(whereStr,"KEY_ID_NO",KEY_ID_NOStr,"")
   		whereStr = makeWhere(whereStr,"DEN_NO",DEN_NOStr,"")
   		whereStr = makeWhere(whereStr,"y.SS_CODE",SS_CODEStr,"")
   		whereStr = makeWhere(whereStr,"KEY_HIN_NO",pnStr,"")
   		whereStr = makeWhere(whereStr,"i.INSP_MESSAGE",INSP_MESSAGEStr,"")
		select case ptypeStr
		case "pTable","pTableCnt","pTable1500","pTable3","pTable2","pList"
	   		whereStr = makeWhere(whereStr,"d.ChoCode",KEY_MUKE_CODEStr,"")
		case else
	   		whereStr = makeWhere(whereStr,"y.KEY_MUKE_CODE",KEY_MUKE_CODEStr,"")
		end select
   		whereStr = makeWhere(whereStr,"INS_NOW",Request.QueryString("INS_NOW"),"")
   		whereStr = makeWhere(whereStr,"UPD_NOW",Request.QueryString("UPD_NOW"),"")

		sqlStr = "select "
		if lngMax > 0 then
			sqlStr = sqlStr & " top " & lngMax
		end if
		select case ptypeStr
		case "pTable","pTableCnt"
			sqlStr = sqlStr & " " & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("事業場")
			sqlStr = sqlStr & "," & GetSelect("直送")
			dim	strSyukaSaki
			strSyukaSaki = ""
			select case ptypeStr
			case "pTable"
				strSyukaSaki = "出荷先"
				sqlStr = sqlStr & "," & GetSelect(strSyukaSaki)
			case "pTableCnt"
				strSyukaSaki = "出荷先 直送集計"
				sqlStr = sqlStr & "," & GetSelect(strSyukaSaki)
				sqlStr = sqlStr & "," & GetSelect("出荷先数<br>")
			end select
			sqlStr = sqlStr & "," & GetSelect("合計<br>")
			sqlStr = sqlStr & "," & GetSelect("<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("月切")
			sqlStr = sqlStr & "," & GetSelect("月切<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("緊急")
			sqlStr = sqlStr & "," & GetSelect("緊急<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("補充")
			sqlStr = sqlStr & "," & GetSelect("補充<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("その他")
			sqlStr = sqlStr & "," & GetSelect("その他<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("個数")
			sqlStr = sqlStr & "," & GetSelect("才数")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & ",""事業場"""
			sqlStr = sqlStr & ",""直送"""
			sqlStr = sqlStr & ",""" & strSyukaSaki & """"
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & ",""事業場"""
			sqlStr = sqlStr & ",""直送"""
			sqlStr = sqlStr & ",""" & strSyukaSaki & """"
		case "pTable1500"	' 集計表(注文区分/15:00前後)
			sqlStr = sqlStr & " " & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("事業場")
			sqlStr = sqlStr & "," & GetSelect("直送")
			sqlStr = sqlStr & "," & GetSelect("出荷先")
			sqlStr = sqlStr & "," & GetSelect("15時以降")
			sqlStr = sqlStr & "," & GetSelect("合計<br>")
			sqlStr = sqlStr & "," & GetSelect("<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("月切")
			sqlStr = sqlStr & "," & GetSelect("月切<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("緊急")
			sqlStr = sqlStr & "," & GetSelect("緊急<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("補充")
			sqlStr = sqlStr & "," & GetSelect("補充<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("その他")
			sqlStr = sqlStr & "," & GetSelect("その他<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("個数")
			sqlStr = sqlStr & "," & GetSelect("才数")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & ",""事業場"""
			sqlStr = sqlStr & ",""直送"""
			sqlStr = sqlStr & ",""出荷先"""
			sqlStr = sqlStr & ",""15時以降"""
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & ",""15時以降"""
			sqlStr = sqlStr & ",""事業場"""
			sqlStr = sqlStr & ",""直送"""
			sqlStr = sqlStr & ",""出荷先"""
		case "pTable3"	' 集計表(出荷日別)
			sqlStr = sqlStr & " " & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("合計<br>")
			sqlStr = sqlStr & "," & GetSelect("<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("月切")
			sqlStr = sqlStr & "," & GetSelect("月切<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("緊急")
			sqlStr = sqlStr & "," & GetSelect("緊急<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("補充")
			sqlStr = sqlStr & "," & GetSelect("補充<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("その他")
			sqlStr = sqlStr & "," & GetSelect("その他<br>検品済")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("個数")
			sqlStr = sqlStr & "," & GetSelect("才数")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""出荷日"""
		case "pTable2"	' 集計表(時間別) 
			sqlStr = sqlStr & " " & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("取引<br>区分")
			sqlStr = sqlStr & "," & GetSelect("注文<br>区分")
			sqlStr = sqlStr & "," & GetSelect("直送")
			sqlStr = sqlStr & "," & GetSelect("出荷先")
			sqlStr = sqlStr & "," & GetSelect("件数")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("出庫<br>済")
			sqlStr = sqlStr & "," & GetSelect("出庫<br>残")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("検品<br>済")
			sqlStr = sqlStr & "," & GetSelect("検品<br>残")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("送信<br>残")
			sqlStr = sqlStr & "," & GetSelect("実績<br>残")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("前日")
			sqlStr = sqlStr & "," & GetSelect("09時")
			sqlStr = sqlStr & "," & GetSelect("10時")
			sqlStr = sqlStr & "," & GetSelect("12時")
			sqlStr = sqlStr & "," & GetSelect("13時")
			sqlStr = sqlStr & "," & GetSelect("14時")
			sqlStr = sqlStr & "," & GetSelect("15時")
			sqlStr = sqlStr & "," & GetSelect("16時")
			sqlStr = sqlStr & "," & GetSelect("17時")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("個数")
			sqlStr = sqlStr & "," & GetSelect("才数")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("HMTAH015")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & ",""取引<br>区分"""
			sqlStr = sqlStr & ",""注文<br>区分"""
			sqlStr = sqlStr & ",""直送"""
			sqlStr = sqlStr & ",""出荷先"""
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & ",""取引<br>区分"""
			sqlStr = sqlStr & ",""直送"""
			sqlStr = sqlStr & ",""注文<br>区分"""
			sqlStr = sqlStr & ",""出荷先"""
		case "pTableDay"	' 集計表(時間別件数個数)
			sqlStr = "select KEY_SYUKA_YMD"
			sqlStr = sqlStr & ",CHOKU_KBN + if(CHOKU_KBN='1',' 直送','') as ""直送"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) <= '18',1,0)) as ""前日"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) >  '18',1,0)) as ""前日<br>夜間"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('07','08','09','10','11'),1,0)) as ""AM"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('12','13','14','15'),1,0)) as ""15:00<br>まで"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('16','17','18'),1,0)) as ""15:00<br>以降"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as qty"
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) <= '18',convert(SURYO,SQL_DECIMAL),0)) as ""前日"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) >  '18',convert(SURYO,SQL_DECIMAL),0)) as ""前日<br>夜間"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('07','08','09','10','11'),convert(SURYO,SQL_DECIMAL),0)) as ""AM"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('12','13','14','15'),convert(SURYO,SQL_DECIMAL),0)) as ""15:00<br>まで"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('16','17','18'),convert(SURYO,SQL_DECIMAL),0)) as ""15:00<br>以降"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by KEY_SYUKA_YMD,CHOKU_KBN"
			sqlStr = sqlStr & " order by KEY_SYUKA_YMD,CHOKU_KBN"
		case "pTable4"	' 集計表(品番/出荷先)
			sqlStr = sqlStr & " KEY_HIN_NO as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",KEY_MUKE_CODE + ' ' + MUKE_NAME  as ""出荷先"""
			sqlStr = sqlStr & ",KEPIN_KAIJYO  as ""欠品解除"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & ",sum(convert(i.SAI_SU,SQL_DECIMAL)*convert(SURYO,SQL_DECIMAL)) as ""才数"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""品番"",""品名"",""出荷先"",""欠品解除"""
			sqlStr = sqlStr & " order by ""件数"" desc,""品番"""
		case "pTablePnZaiko"	' 集計表(品番/在庫)
			sqlStr = sqlStr & " " & GetSelect("事")
			sqlStr = sqlStr & "," & GetSelect("品番")
			sqlStr = sqlStr & "," & GetSelect("品名")
			sqlStr = sqlStr & "," & GetSelect("標準棚番")
'			sqlStr = sqlStr & ",i.GENSANKOKU"
			sqlStr = sqlStr & "," & GetSelect("出荷件数")
			sqlStr = sqlStr & "," & GetSelect("出荷数")
			sqlStr = sqlStr & "," & GetSelect("商品化(済)")
			sqlStr = sqlStr & "," & GetSelect("商品化(未)")
			sqlStr = sqlStr & "," & GetSelect("要商品化数")
			sqlStr = sqlStr & "," & GetSelect("才数")
			sqlStr = sqlStr & "," & GetSelect("国内供給区分")
			sqlStr = sqlStr & "," & GetSelect("海外供給区分")
			sqlStr = sqlStr & "," & GetSelect("ユニット区分")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("Zaiko")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事"",""品番"",""品名"",""標準棚番"",""商品化(済)"",""商品化(未)"",""国内供給区分"",""海外供給区分"",""ユニット区分"""
			sqlStr = sqlStr & " order by ""事"",""標準棚番"",""品番"""
		case "pTablePnSumzai"	' 集計表(品番/在庫集計)
			sqlStr = sqlStr & " y.JGYOBU as ""事業部"""
			sqlStr = sqlStr & ",y.KEY_HIN_NO as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",count(distinct y.KEY_ID_NO) as ""出荷件数"""
			sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_DECIMAL)) as ""出荷数"""
			sqlStr = sqlStr & ",z.z_qty as ""POS在庫数"""
			sqlStr = sqlStr & ",convert(sz.bu_zai_qty,SQL_DECIMAL) as ""BU在庫数"""
			sqlStr = sqlStr & ",convert(sz.ppsc_zai_qty,SQL_DECIMAL) as ""PPSC在庫数"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(convert(YUKO_Z_QTY,SQL_DECIMAL)) as z_qty from zaiko group by JGYOBU,NAIGAI,HIN_GAI) as z on ("
			sqlStr = sqlStr & "     z.JGYOBU = y.JGYOBU"
			sqlStr = sqlStr & " and z.NAIGAI = y.NAIGAI"
			sqlStr = sqlStr & " and z.HIN_GAI = y.KEY_HIN_NO)"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & " left outer join sumzai as sz on (y.jgyobu = sz.jgyobu and y.naigai = sz.naigai and y.key_hin_no = sz.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事業部"",""品番"",""品名"",""POS在庫数"",""BU在庫数"",""PPSC在庫数"""
			sqlStr = sqlStr & " order by ""事業部"",""品番"""
		case "pTableShizai"	' 集計表(資材)
			sqlStr = sqlStr & " k.KO_HIN_GAI as ""資材品番"""
			sqlStr = sqlStr & ",si.HIN_NAME as ""資材品名"""
			sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_DECIMAL)*convert(k.KO_QTY,SQL_DECIMAL)) as ""資材<br>数量"""
			sqlStr = sqlStr & ",convert(si.G_ST_SHITAN,SQL_DECIMAL) as ""仕入単価"""
			sqlStr = sqlStr & ",round(sum(convert(y.SURYO,SQL_DECIMAL) * convert(k.KO_QTY,SQL_DECIMAL) * convert(si.G_ST_SHITAN,SQL_DECIMAL)),0) as ""仕入金額"""
			sqlStr = sqlStr & ",convert(si.G_ST_URITAN,SQL_DECIMAL) as ""販売単価"""
			sqlStr = sqlStr & ",round(sum(convert(y.SURYO,SQL_DECIMAL) * convert(k.KO_QTY,SQL_DECIMAL) * convert(si.G_ST_URITAN,SQL_DECIMAL)),0) as ""販売金額"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & 	" inner join item as i"
			sqlStr = sqlStr &				" on (y.JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and y.NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and y.KEY_HIN_NO =i.HIN_GAI)"
			sqlStr = sqlStr &	" inner join P_COMPO_K as k"
			sqlStr = sqlStr &				" on (i.SHIMUKE_CODE=k.SHIMUKE_CODE"
			sqlStr = sqlStr &				" and i.JGYOBU=k.JGYOBU"
			sqlStr = sqlStr &				" and i.NAIGAI=k.NAIGAI"
			sqlStr = sqlStr &				" and i.HIN_GAI=k.HIN_GAI"
			sqlStr = sqlStr &				" and k.DATA_KBN<>'2'"
			sqlStr = sqlStr &				" and k.KO_JGYOBU='S')"
			sqlStr = sqlStr &	" inner join ITEM as si"
			sqlStr = sqlStr &				" on (k.KO_JGYOBU	=si.JGYOBU"
			sqlStr = sqlStr &				" and k.KO_NAIGAI	=si.NAIGAI"
			sqlStr = sqlStr &				" and k.KO_HIN_GAI	=si.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""資材品番"",""資材品名"",""仕入単価"",""販売単価"""
			sqlStr = sqlStr & " order by ""資材品番"""
		case "DelToY"		' 削除済→予定もどし
			dim dlmStr
			dim sqlDlmStr

			sqlStr = "select "
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			set rsList = db.Execute(sqlStr)

			insStr = "insert into y_syuka"
			sqlStr = "select "
			dlmStr = "("
			sqlDlmStr = " "
			For i=0 To rsList.Fields.Count-1
				insStr = insStr & dlmStr & rsList.Fields(i).name
				sqlStr = sqlStr & sqlDlmStr & rsList.Fields(i).name
				dlmStr = ","
				sqlDlmStr = ","
			next
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & andStr & " KEY_ID_NO not in (select KEY_ID_NO from y_syuka)"
			insStr = insStr & ")" & sqlStr
			set rsList = db.Execute(insStr)

		case "pTableSeikyu"		' 集計表(商品化請求内訳)
			sqlStr = sqlStr & " y.JGYOBU as ""事業部"""
			sqlStr = sqlStr & ",y.KEY_HIN_NO as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",i.S_SEIKYU_F as ""請求区分"""
			sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_SET_DATE)='',-999999.99,convert(i.S_KOUSU_BAIKA ,SQL_DECIMAL)) as ""工料＠"""
			sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_SET_DATE)='',-999999.99,convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""資材仕入＠"""
			sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_SET_DATE)='',-999999.99,convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""資材販売＠"""
			sqlStr = sqlStr & ",count(*) as ""出荷件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""出荷数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)) as ""工料金額"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""資材仕入金額"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""資材販売金額"""
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事業部"",""品番"",""品名"",""請求区分"",""工料＠"",""資材仕入＠"",""資材販売＠"""
			sqlStr = sqlStr & " order by ""事業部"",""品番"""
			db.CommandTimeout=900
		case "pTableMonth"		' 集計表(月別)
			sqlStr = sqlStr & " left(KEY_SYUKA_YMD,6)  as ""出荷年月"""
			sqlStr = sqlStr & ",count(*) as ""出荷件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""出荷数"""
			sqlStr = sqlStr & ",round(sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)),0) as ""工料"""
			sqlStr = sqlStr & ",round(sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)),0) as ""資材"""
			sqlStr = sqlStr & ",count(distinct KEY_HIN_NO) as ""アイテム"""
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷年月"""
			sqlStr = sqlStr & " order by ""出荷年月"""
			db.CommandTimeout=900

		case "pTableSeikyuKaigai"	' 集計表(商品化請求内訳/国内外単価対応)
			sqlStr = sqlStr & " y.JGYOBU as ""事業部"""
			sqlStr = sqlStr & ",y.KEY_HIN_NO as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",if(rtrim(ifnull(i.S_KOUSU_SET_DATE,'')) ='',-999999.99,convert(i.S_KOUSU_BAIKA ,SQL_DECIMAL)) as ""工料＠"""
			sqlStr = sqlStr & ",if(rtrim(ifnull(i.S_KOUSU_SET_DATE,''))='',         0,convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""資材仕入＠"""
			sqlStr = sqlStr & ",if(rtrim(ifnull(i.S_KOUSU_SET_DATE,''))='',-999999.99,convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""資材販売＠"""
			sqlStr = sqlStr & ",count(*) as ""出荷件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""出荷数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(ifnull(i.S_KOUSU_BAIKA,'0'),SQL_DECIMAL)) as ""工料金額"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(ifnull(i.S_SHIZAI_GENKA,'0'),SQL_DECIMAL)) as ""資材仕入金額"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(ifnull(i.S_SHIZAI_BAIKA,'0'),SQL_DECIMAL)) as ""資材販売金額"""
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.HAN_KBN = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事業部"",""品番"",""品名"",""工料＠"",""資材仕入＠"",""資材販売＠"""
			sqlStr = sqlStr & " order by ""事業部"",""品番"""
			db.CommandTimeout=900

		case "pTablePnMonth","pTableSMonth"	' 集計表(品番／月別 出荷数量) 集計表(出荷先／月別 出荷数量)
			dim fromStr
			dim groupByStr
			dim	orderByStr
			dim	sumStr
			dim	sqlStr2

			db.CommandTimeout=300
			sqlStr2 = "select distinct"
			sqlStr2 = sqlStr2 & " left(KEY_SYUKA_YMD,6)  as syukaYM"
			sqlStr2 = sqlStr2 & " From " & tblStr & " as y"
			sqlStr2 = sqlStr2 & whereStr
			sqlStr2 = sqlStr2 & " order by syukaYM"
			set rsList = db.Execute(sqlStr2)

			select case ptypeStr
				case "pTablePnMonth"
					sqlStr = sqlStr & " KEY_HIN_NO as ""品番"""
					sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
'2009.10.06					sqlStr = sqlStr & ",k.KO_HIN_GAI as ""資材品番"""
'2009.10.06					sqlStr = sqlStr & ",k.KO_HIN_NAME as ""資材品名"""
					sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_BAIKA) ='',-999999.99,convert(i.S_KOUSU_BAIKA ,SQL_DECIMAL)) as ""工料＠"""
					sqlStr = sqlStr & ",if(rtrim(i.S_SHIZAI_GENKA)='',         0,convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""資材仕入＠"""
					sqlStr = sqlStr & ",if(rtrim(i.S_SHIZAI_BAIKA)='',-999999.99,convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""資材販売＠"""
					sqlStr = sqlStr & ",ifnull(z.sumi_qty,0) as ""商品化(済)"""
					sqlStr = sqlStr & ",ifnull(z.mi_qty,0) as ""商品化(未)"""
					sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""出荷数"""
				'	sqlStr = sqlStr & ",sum(distinct if(z.GOODS_ON =  '0',convert(z.YUKO_Z_QTY,SQL_DECIMAL),0)) as ""商品化(済)"""
				'	sqlStr = sqlStr & ",sum(distinct if(z.GOODS_ON <> '0',convert(z.YUKO_Z_QTY,SQL_DECIMAL),0)) as ""商品化(未)"""
					sumStr = "convert(SURYO,SQL_DECIMAL)"
					fromStr = " from " & tblStr & " as y"
					fromStr = fromStr & " left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(if(GOODS_ON =  '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) as sumi_qty,sum(if(GOODS_ON <> '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) as mi_qty from zaiko group by JGYOBU,NAIGAI,HIN_GAI) as z on ("
					fromStr = fromStr & "     z.JGYOBU = y.JGYOBU"
					fromStr = fromStr & " and z.NAIGAI = y.NAIGAI"
					fromStr = fromStr & " and z.HIN_GAI = y.KEY_HIN_NO)"
					fromStr = fromStr & " inner join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
'2009.10.06					fromStr = fromStr &	" inner join (select a.SHIMUKE_CODE,a.JGYOBU,a.NAIGAI,a.HIN_GAI,a.KO_HIN_GAI,b.HIN_NAME as KO_HIN_NAME"
'2009.10.06					fromStr = fromStr &	            " from P_COMPO_K as a"
'2009.10.06					fromStr = fromStr &	            " inner join ITEM as b on (a.KO_JGYOBU=b.JGYOBU and a.KO_NAIGAI=b.NAIGAI and a.KO_HIN_GAI=b.HIN_GAI)"
'2009.10.06					fromStr = fromStr &				" where a.DATA_KBN='1'"
'2009.10.06					fromStr = fromStr &				" and a.SEQNO='010'"
'2009.10.06					fromStr = fromStr &				" ) as k"
'2009.10.06'					fromStr = fromStr &				" on (if(ifnull(i.SHIMUKE_CODE,'') = '' ,'01',i.SHIMUKE_CODE) =k.SHIMUKE_CODE"
'2009.10.06					fromStr = fromStr &				" on ('" & strShimuke & "'=k.SHIMUKE_CODE"
'2009.10.06					fromStr = fromStr &				" and y.JGYOBU=k.JGYOBU"
'2009.10.06					fromStr = fromStr &				" and y.NAIGAI=k.NAIGAI"
'2009.10.06					fromStr = fromStr &				" and y.KEY_HIN_NO=k.HIN_GAI"
'2009.10.06					fromStr = fromStr &				" )"
					groupByStr = " group by ""品番"",""品名"",""工料＠"",""資材仕入＠"",""資材販売＠"",""商品化(済)"",""商品化(未)"""
					orderByStr = " order by ""出荷数"" desc ,""品番"""
				case "pTableSMonth"
					sqlStr = sqlStr & " y.LK_MUKE_CODE + ' ' + y.MUKE_NAME ""出荷先"""	'" KEY_MUKE_CODE + ' ' + Mts.MUKE_NAME as ""出荷先"""
					sqlStr = sqlStr & ",count(*) as ""出荷件数"""
					sumStr = "1"
					fromStr = " from " & tblStr & " as y left outer join Mts on KEY_MUKE_CODE = Mts.MUKE_CODE and Mts.SS_CODE = ''"
					groupByStr = " group by ""出荷先"""
					orderByStr = " order by ""出荷先"""
			end select
			sqlStr = sqlStr & ",' ' as ""span1"""
			Do While Not rsList.EOF
				sqlStr = sqlStr & ",sum(if(left(KEY_SYUKA_YMD,6) ='" & rsList.Fields("syukaYM") & "'," & sumStr & ",0)) as """ & rsList.Fields("syukaYM") & """"
				rsList.Movenext
			loop
			sqlStr = sqlStr & fromStr
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & groupByStr
			sqlStr = sqlStr & orderByStr
			db.CommandTimeout=900
		case "pTablePnJumpUp"	' 集計表(品番 出荷急増)
			dim	workDay
			sqlStr2 = "select distinct"
			sqlStr2 = sqlStr2 & " count(distinct KEY_SYUKA_YMD) as wDay"
			sqlStr2 = sqlStr2 & " From del_syuka"
			sqlStr2 = sqlStr2 & " where KEY_SYUKA_YMD between '" & dtStr & "' and '" & dtToStr & "'"
			set rsList = db.Execute(sqlStr2)
			workDay = 0
			workDay = rsList.Fields("wDay")

			sqlStr = sqlStr & " y.KEY_HIN_NO as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",count(*) as ""出荷件数(当日)"""
			sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_DECIMAL)) as ""出荷数(当日)"""
			sqlStr = sqlStr & ",if(ifnull(d.qty,0) = 0,sum(convert(y.SURYO,SQL_DECIMAL))*10000,sum(convert(y.SURYO,SQL_DECIMAL))*100/(d.qty/" & workDay & ")) as ""増減比"""
			sqlStr = sqlStr & ",ifnull(d.qty/" & workDay & ",0) as ""出荷数(平均)"""
			sqlStr = sqlStr & ",ifnull(d.cnt,0) as ""出荷件数(過去" & workDay & "日)"""
			sqlStr = sqlStr & ",ifnull(d.qty,0) as ""出荷数(過去" & workDay & "日)"""
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & " left outer join ("
			sqlStr = sqlStr &   " select "
			sqlStr = sqlStr &   " JGYOBU,"
			sqlStr = sqlStr &   " NAIGAI,"
			sqlStr = sqlStr &   " KEY_HIN_NO,"
			sqlStr = sqlStr &   " count(*) as cnt,"
			sqlStr = sqlStr &   " sum(convert(SURYO,SQL_DECIMAL)) as qty"
			sqlStr = sqlStr &   " from del_syuka"
			sqlStr = sqlStr &   " where KEY_SYUKA_YMD between '" & dtStr & "' and '" & dtToStr & "'"
			sqlStr = sqlStr &   "   and KEY_HIN_NO in (select distinct KEY_HIN_NO from " & tblStr & " " & replace(whereStr,"y.","") & ")"
			sqlStr = sqlStr &   " group by JGYOBU,NAIGAI,KEY_HIN_NO"
			sqlStr = sqlStr & " ) as d on (d.JGYOBU = y.JGYOBU"
			sqlStr = sqlStr & " and d.NAIGAI = y.NAIGAI"
			sqlStr = sqlStr & " and d.KEY_HIN_NO = y.KEY_HIN_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""品番"",""品名"",""出荷件数(過去" & workDay & "日)"",""出荷数(過去" & workDay & "日)"",d.qty"
			sqlStr = sqlStr & " order by ""増減比"" desc"
			db.CommandTimeout=900
		case "pTable5"	' 集計表(日別／向先別)
			sqlStr = sqlStr & " KEY_MUKE_CODE as ""出荷先"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '01',1,0)) as ""01"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '02',1,0)) as ""02"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '03',1,0)) as ""03"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '04',1,0)) as ""04"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '05',1,0)) as ""05"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '06',1,0)) as ""06"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '07',1,0)) as ""07"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '08',1,0)) as ""08"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '09',1,0)) as ""09"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '10',1,0)) as ""10"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '11',1,0)) as ""11"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '12',1,0)) as ""12"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '13',1,0)) as ""13"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '14',1,0)) as ""14"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '15',1,0)) as ""15"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '16',1,0)) as ""16"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '17',1,0)) as ""17"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '18',1,0)) as ""18"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '19',1,0)) as ""19"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '20',1,0)) as ""20"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '21',1,0)) as ""21"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '22',1,0)) as ""22"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '23',1,0)) as ""23"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '24',1,0)) as ""24"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '25',1,0)) as ""25"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '26',1,0)) as ""26"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '27',1,0)) as ""27"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '28',1,0)) as ""28"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '29',1,0)) as ""29"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '30',1,0)) as ""30"""
			sqlStr = sqlStr & ",sum(if(right(KEY_SYUKA_YMD,2) = '31',1,0)) as ""31"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by KEY_MUKE_CODE"
			sqlStr = sqlStr & " order by KEY_MUKE_CODE"
		case "pTableDaySaki"	' 集計表(出荷日別／向先別)
			db.CommandTimeout=900
			sqlStr = sqlStr & " KEY_SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',1,0)) as ""A1<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',convert(SURYO,SQL_DECIMAL),0)) as ""A1<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',1,0)) as ""A2<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',convert(SURYO,SQL_DECIMAL),0)) as ""A2<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',1,0)) as ""A3<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',convert(SURYO,SQL_DECIMAL),0)) as ""A3<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',1,0)) as ""A4<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',convert(SURYO,SQL_DECIMAL),0)) as ""A4<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',1,0)) as ""A5<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',convert(SURYO,SQL_DECIMAL),0)) as ""A5<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',1,0)) as ""A6<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',convert(SURYO,SQL_DECIMAL),0)) as ""A6<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',1,0)) as ""A7<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',convert(SURYO,SQL_DECIMAL),0)) as ""A7<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',1,0)) as ""A8<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',convert(SURYO,SQL_DECIMAL),0)) as ""A8<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),1,0)) as ""その他<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),convert(SURYO,SQL_DECIMAL),0)) as ""その他<br>数量"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"""
			sqlStr = sqlStr & " order by ""出荷日"""
		case "pTableDayTime"	' 集計表(出荷日／受信時刻別)
			db.CommandTimeout=900
			sqlStr = sqlStr & " KEY_SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",left(INS_NOW,8) as ""受信日"""
			sqlStr = sqlStr & ",substring(INS_NOW,9,4) as ""受信時刻"""
			sqlStr = sqlStr & ",if(KEY_SYUKA_YMD = left(INS_NOW,8),substring(INS_NOW,9,2) + '時','前日') as ""当日時刻"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',1,0)) as ""A1<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',convert(SURYO,SQL_DECIMAL),0)) as ""A1<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',1,0)) as ""A2<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',convert(SURYO,SQL_DECIMAL),0)) as ""A2<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',1,0)) as ""A3<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',convert(SURYO,SQL_DECIMAL),0)) as ""A3<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',1,0)) as ""A4<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',convert(SURYO,SQL_DECIMAL),0)) as ""A4<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',1,0)) as ""A5<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',convert(SURYO,SQL_DECIMAL),0)) as ""A5<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',1,0)) as ""A6<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',convert(SURYO,SQL_DECIMAL),0)) as ""A6<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',1,0)) as ""A7<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',convert(SURYO,SQL_DECIMAL),0)) as ""A7<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',1,0)) as ""A8<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',convert(SURYO,SQL_DECIMAL),0)) as ""A8<br>数量"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),1,0)) as ""その他<br>件数"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),convert(SURYO,SQL_DECIMAL),0)) as ""その他<br>数量"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""受信日"",""受信時刻"",""当日時刻"""
			sqlStr = sqlStr & " order by ""出荷日"",""受信日"",""受信時刻"""
		case "pTableSoko"	' 集計表(倉庫別)
'			sqlStr = sqlStr & " left(TANABAN1,2) as ""倉庫"""
			sqlStr = sqlStr & " left(i.ST_SOKO,2) + ' ' + sk.soko_name as ""倉庫"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',1,0)) as ""出庫<br>済"" "
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',0,1)) as ""出庫<br>残"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',1,0)) as ""検品<br>済"" "
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',0,1)) as ""検品<br>残"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as qty"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & " left outer join soko as sk on (i.st_soko = sk.soko_no)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""倉庫"""
			sqlStr = sqlStr & " order by ""倉庫"""
		case "pTableOrderNo"	' 集計表(オーダーNo)
			sqlStr = sqlStr & " KEY_SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",ODER_NO as ""オーダーNo"""
			sqlStr = sqlStr & ",if(LK_MUKE_CODE <> '',LK_MUKE_CODE,rtrim(KEY_MUKE_CODE)) + ' ' + MUKE_NAME as ""出荷先"""
			sqlStr = sqlStr & ",CYU_KBN_NAME as ""注文区分名"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',1,0)) as ""出庫<br>済"" "
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',0,1)) as ""出庫<br>残"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',1,0)) as ""検品<br>済"" "
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',0,1)) as ""検品<br>残"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as qty"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""オーダーNo"",""出荷先"",""注文区分名"""
			sqlStr = sqlStr & " order by ""出荷日"",""オーダーNo"""
		case "pList"	' 一覧表
			sqlStr = sqlStr & " " & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("事業場")
			sqlStr = sqlStr & "," & GetSelect("取引<br>区分")
			sqlStr = sqlStr & "," & GetSelect("注文<br>区分")
			sqlStr = sqlStr & "," & GetSelect("欠品解除")
			sqlStr = sqlStr & "," & GetSelect("出荷先")
			sqlStr = sqlStr & "," & GetSelect("事業部")
			sqlStr = sqlStr & "," & GetSelect("在庫<br>収支")
			sqlStr = sqlStr & "," & GetSelect("ID")
			sqlStr = sqlStr & "," & GetSelect("伝票No")
			sqlStr = sqlStr & "," & GetSelect("オーダーNo")
			sqlStr = sqlStr & "," & GetSelect("アイテムNo")
			sqlStr = sqlStr & "," & GetSelect("品番")
			sqlStr = sqlStr & "," & GetSelect("品名")
			sqlStr = sqlStr & "," & GetSelect("数量")
			sqlStr = sqlStr & "," & GetSelect("(出庫済)")
			sqlStr = sqlStr & "," & GetSelect(".才数")
			sqlStr = sqlStr & "," & GetSelect("備考1")
			sqlStr = sqlStr & "," & GetSelect("備考2")
			sqlStr = sqlStr & "," & GetSelect("完了区分")
			sqlStr = sqlStr & "," & GetSelect("出庫日時")
			sqlStr = sqlStr & "," & GetSelect("検品日時")
			sqlStr = sqlStr & "," & GetSelect("検品担当者")
			sqlStr = sqlStr & "," & GetSelect("事前チェック")
			sqlStr = sqlStr & "," & GetSelect("移管状況")
			sqlStr = sqlStr & "," & GetSelect("検品メッセージ")
			sqlStr = sqlStr & "," & GetSelect("棚番")
			sqlStr = sqlStr & "," & GetSelect("登録日時")
			sqlStr = sqlStr & "," & GetSelect("更新日時")
'			sqlStr = sqlStr & "," & GetSelect("データ区分")
'			sqlStr = sqlStr & "," & GetSelect("販売区分")
'			sqlStr = sqlStr & "," & GetSelect("直送先")
'			sqlStr = sqlStr & "," & GetSelect("出庫表印刷")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""出荷日"""
			sqlStr = sqlStr & ",""注文<br>区分"""
			sqlStr = sqlStr & ",""出荷先"""
			sqlStr = sqlStr & ",""ID"""
		case "pTest"	' 出庫テスト
			sqlStr = sqlStr & " " & GetSelect("z棚番_BC")
			sqlStr = sqlStr & "," & GetSelect("事")
			sqlStr = sqlStr & "," & GetSelect("品番_BC")
			sqlStr = sqlStr & "," & GetSelect("数量")
			sqlStr = sqlStr & "," & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("注文区分")
			sqlStr = sqlStr & "," & GetSelect("出荷先")
			sqlStr = sqlStr & "," & GetSelect("出荷先_BC")
			sqlStr = sqlStr & "," & GetSelect("(出庫済)")
			sqlStr = sqlStr & "," & GetSelect("完了区分")
			sqlStr = sqlStr & "," & GetSelect("ID_BC")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Zaiko.Tana")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""ID_BC"""
			sqlStr = sqlStr & ",""z棚番_BC"""
		case "pPickingTest"	' 検品テスト
			sqlStr = sqlStr & " " & GetSelect("棚番_BC")
			sqlStr = sqlStr & "," & GetSelect("事")
			sqlStr = sqlStr & "," & GetSelect("品番_BC")
			sqlStr = sqlStr & "," & GetSelect("数量")
			sqlStr = sqlStr & "," & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("出荷先LK")
			sqlStr = sqlStr & "," & GetSelect("注文区分")
			sqlStr = sqlStr & "," & GetSelect("(出庫済)")
			sqlStr = sqlStr & "," & GetSelect("完了区分")
			sqlStr = sqlStr & "," & GetSelect("ID_BC")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""ID_BC"""
			sqlStr = sqlStr & ",""棚番_BC"""
		case "pPicking"	' 出庫表
			sqlStr = sqlStr & " " & GetSelect("標準棚番")
			sqlStr = sqlStr & "," & GetSelect("事")
			sqlStr = sqlStr & "," & GetSelect("品番")
			sqlStr = sqlStr & "," & GetSelect("数量")
			sqlStr = sqlStr & "," & GetSelect("出荷日")
			sqlStr = sqlStr & "," & GetSelect("出荷先")
			sqlStr = sqlStr & "," & GetSelect("(出庫済)")
			sqlStr = sqlStr & "," & GetSelect("完了区分")
			sqlStr = sqlStr & "," & GetSelect("出庫日時")
			sqlStr = sqlStr & "," & GetSelect("ID_BC")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by 1,2,3"
		case "pListIri","pTableIri"	' 一覧表(外装入数),集計表(外装入数)
			sqlStr = sqlStr & " KEY_MUKE_CODE + ' ' + MUKE_NAME ""出荷先"""
			sqlStr = sqlStr & ",KEY_CYU_KBN + if(KEY_CYU_KBN = '1',' 月切',if(KEY_CYU_KBN = '2',' 緊急',if(KEY_CYU_KBN = '3',' 補充',if(KEY_CYU_KBN = 'E',' 貿易','')))) ""注文<br>区分"""
            if pTypeStr = "pListIri" then  ' 一覧表(外装入数)
    			sqlStr = sqlStr & ",y.JGYOBU ""事"""
    			sqlStr = sqlStr & ",KEPIN_KAIJYO ""欠品<br>解除"""
            end if
			sqlStr = sqlStr & ",KEY_SYUKA_YMD ""出荷日"""
            if pTypeStr = "pListIri" then  ' 一覧表(外装入数)
			    sqlStr = sqlStr & ",KEY_ID_NO ""ID"""
			    sqlStr = sqlStr & ",KEY_HIN_NO ""品番"""
			    sqlStr = sqlStr & ",y.HIN_NAME ""品名"""
			    sqlStr = sqlStr & ",convert(SURYO,SQL_DECIMAL) ""数量"""
			    sqlStr = sqlStr & ",mod5(convert(SURYO,SQL_DECIMAL),g_qty_1,g_qty_2,g_qty_3,g_qty_4,g_qty_5) ""端数"""
			    sqlStr = sqlStr & ",g_qty_1 ""入数1"""
			    sqlStr = sqlStr & ",g_qty_2 ""入数2"""
			    sqlStr = sqlStr & ",g_qty_3 ""入数3"""
			    sqlStr = sqlStr & ",g_qty_4 ""入数4"""
			    sqlStr = sqlStr & ",g_qty_5 ""入数5"""
			    sqlStr = sqlStr & ",convert(i.SAI_SU,SQL_DECIMAL)*convert(SURYO,SQL_DECIMAL) as ""才数"""
			    sqlStr = sqlStr & ",BIKOU1 as ""備考1"""
			    sqlStr = sqlStr & ",BIKOU2 as ""備考2"""
            else                        ' 集計表(外装入数)
			    sqlStr = sqlStr & ",count(*) ""件数"""
			    sqlStr = sqlStr & ",sum(if(mod5(convert(SURYO,SQL_DECIMAL),g_qty_1,g_qty_2,g_qty_3,g_qty_4,g_qty_5)=0,1,0)) ""端数<br>０"""
			    sqlStr = sqlStr & ",sum(if(mod5(convert(SURYO,SQL_DECIMAL),g_qty_1,g_qty_2,g_qty_3,g_qty_4,g_qty_5)=0,0,1)) ""端数<br>あり"""
            end if
			sqlStr = sqlStr & " From " & tblStr & " y"
            if pTypeStr = "pListIri" then  ' 一覧表(外装入数)
    			sqlStr = sqlStr & " left outer join item i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
            end if
			sqlStr = sqlStr & " inner join ("
			sqlStr = sqlStr & " 	select"
'			sqlStr = sqlStr & " 	//top 100"
			sqlStr = sqlStr & " 	 k.SHIMUKE_CODE                     						SHIMUKE"
			sqlStr = sqlStr & " 	,k.JGYOBU													JGYOBU"
			sqlStr = sqlStr & " 	,k.NAIGAI													NAIGAI"
			sqlStr = sqlStr & " 	,k.HIN_GAI													HIN_GAI"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '010',convert(k.KO_QTY,SQL_decimal),0))	g_qty_1"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '020',convert(k.KO_QTY,SQL_decimal),0))	g_qty_2"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '030',convert(k.KO_QTY,SQL_decimal),0))	g_qty_3"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '040',convert(k.KO_QTY,SQL_decimal),0))	g_qty_4"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '050',convert(k.KO_QTY,SQL_decimal),0))	g_qty_5"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '010',rtrim(k.KO_HIN_GAI),''))					g_hin_1"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '020',rtrim(k.KO_HIN_GAI),''))					g_hin_2"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '030',rtrim(k.KO_HIN_GAI),''))					g_hin_3"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '040',rtrim(k.KO_HIN_GAI),''))					g_hin_4"
			sqlStr = sqlStr & " 	,max(if(k.SEQNO = '050',rtrim(k.KO_HIN_GAI),''))					g_hin_5"
			sqlStr = sqlStr & " 	from P_COMPO_K k"
			sqlStr = sqlStr & " 		where k.DATA_KBN = '2'"
		    select case centerStr
		    case "小野PC"
    			sqlStr = sqlStr & " 		  and k.SHIMUKE_CODE in ('01','02','03')"
		    case else
    			sqlStr = sqlStr & " 		  and k.SHIMUKE_CODE in ('01')"
		    end select
			sqlStr = sqlStr & " 		group by"
			sqlStr = sqlStr & " 		 SHIMUKE"
			sqlStr = sqlStr & " 		,k.JGYOBU"
			sqlStr = sqlStr & " 		,k.NAIGAI"
			sqlStr = sqlStr & " 		,k.HIN_GAI"
			sqlStr = sqlStr & " 	) g on (g.JGYOBU = y.JGYOBU and g.NAIGAI = y.NAIGAI and g.HIN_GAI = y.HIN_NO)"
			sqlStr = sqlStr & whereStr
            if pTypeStr = "pListIri" then  ' 一覧表(外装入数)
			    sqlStr = sqlStr & " order by ""出荷日"""
			    sqlStr = sqlStr &          ",""注文<br>区分"""
			    sqlStr = sqlStr &          ",""出荷先"""
			    sqlStr = sqlStr &          ",""事"""
			    sqlStr = sqlStr &          ",KEY_ID_NO"
            else  ' 集計表(外装入数)
    			sqlStr = sqlStr & " group by"
			    sqlStr = sqlStr & " ""出荷日"""
			    sqlStr = sqlStr & ",""注文<br>区分"""
'			    sqlStr = sqlStr & ",""欠品<br>解除"""
			    sqlStr = sqlStr & ",""出荷先"""
'			    sqlStr = sqlStr & ",""事"""
    			sqlStr = sqlStr & " order by"
			    sqlStr = sqlStr & " ""出荷先"""
			    sqlStr = sqlStr & ",""注文<br>区分"""
'			    sqlStr = sqlStr & ",""事"""
'			    sqlStr = sqlStr & ",""欠品<br>解除"""
			    sqlStr = sqlStr & ",""出荷日"""
            end if
		case "pListAll"	' 一覧表(全項目)
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
'			sqlStr = sqlStr & " order by KEY_SYUKA_YMD"
		case "uSyuko"	' 出庫OK
'			const	vbYesNo	=	4
'			const	vbYes	=	6
			strHMS = right("0" & Hour(now()),2) & right("0" & Minute(now()),2)  & right("0" & Second(now()),2) 
			strYMD = year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2) 
			sqlStr = "update " & tblStr & " as y"
			sqlStr = sqlStr & " set "
			sqlStr = sqlStr & "     KAN_YMD = '" & strYMD & "'"
			sqlStr = sqlStr & "    ,KAN_KBN = '9'"
			sqlStr = sqlStr & "    ,JITU_SURYO = SURYO"
			sqlStr = sqlStr & whereStr
'			if Msg("指定した条件全て、検品済にします。" & vbcrlf & sqlStr,vbYesNo) = vbYes then
				set rsList = db.Execute(sqlStr)
'			end if
			sqlStr = "select"
			sqlStr = sqlStr & " KEY_ID_NO"
			sqlStr = sqlStr & ",KAN_YMD"
			sqlStr = sqlStr & ",KAN_KBN"
			sqlStr = sqlStr & ",JITU_SURYO"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by KEY_ID_NO"
		case "uSyukoCancel"	' 出庫OKキャンセル
			sqlStr = "update " & tblStr & " as y"
			sqlStr = sqlStr & " set "
			sqlStr = sqlStr & "     KAN_KBN = '0'"
			sqlStr = sqlStr & "	   ,JITU_SURYO = '0'"
			sqlStr = sqlStr & whereStr
'			if Msg("指定した条件全て、検品済にします。" & vbcrlf & sqlStr,vbYesNo) = vbYes then
				set rsList = db.Execute(sqlStr)
'			end if
			sqlStr = "select"
			sqlStr = sqlStr & " KEY_ID_NO"
			sqlStr = sqlStr & ",KAN_YMD"
			sqlStr = sqlStr & ",KAN_KBN"
			sqlStr = sqlStr & ",JITU_SURYO"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by KEY_ID_NO"
		case "uPrintCancel"	' 出庫表OKキャンセル
			sqlStr = "update " & tblStr & " as y"
			sqlStr = sqlStr & " set "
			sqlStr = sqlStr & "     PRINT_YMD = ''"
			sqlStr = sqlStr & whereStr
			set rsList = db.Execute(sqlStr)
			sqlStr = "select"
			sqlStr = sqlStr & " KEY_ID_NO"
			sqlStr = sqlStr & ",KAN_YMD"
			sqlStr = sqlStr & ",KAN_KBN"
			sqlStr = sqlStr & ",JITU_SURYO"
			sqlStr = sqlStr & ",PRINT_YMD"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by KEY_ID_NO"

		case "uKenpin"	' 検品OK
'			const	vbYesNo	=	4
'			const	vbYes	=	6
			strHMS = right("0" & Hour(now()),2) & right("0" & Minute(now()),2)  & right("0" & Second(now()),2) 
			strYMD = year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2) 
			sqlStr = "update " & tblStr & " as y"
			sqlStr = sqlStr & " set KENPIN_TANTO_CODE = 'SQL'"
			sqlStr = sqlStr & "    ,KENPIN_YMD = '" & strYMD & "'"
			sqlStr = sqlStr & "    ,KENPIN_HMS = '" & strHMS & "'"
			sqlStr = sqlStr & "    ,JITU_SURYO = SURYO"
			sqlStr = sqlStr & "    ,LK_SEQ_NO = ''"
			sqlStr = sqlStr & whereStr
'			if Msg("指定した条件全て、検品済にします。" & vbcrlf & sqlStr,vbYesNo) = vbYes then
				set rsList = db.Execute(sqlStr)
'			end if

			sqlStr = "select"
			sqlStr = sqlStr & " KEY_ID_NO"
			sqlStr = sqlStr & ",KENPIN_TANTO_CODE"
			sqlStr = sqlStr & ",KENPIN_YMD"
			sqlStr = sqlStr & ",KENPIN_HMS"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by KEY_ID_NO"
		case "uKenpinCancel"	' 検品キャンセル
'			const	vbYesNo	=	4
'			const	vbYes	=	6
			strHMS = right("0" & Hour(now()),2) & right("0" & Minute(now()),2)  & right("0" & Second(now()),2) 
			strYMD = year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2) 
			sqlStr = "update " & tblStr & " as y"
			sqlStr = sqlStr & " set KENPIN_TANTO_CODE = ''"
			sqlStr = sqlStr & "    ,KENPIN_YMD = ''"
			sqlStr = sqlStr & "    ,KENPIN_HMS = ''"
			sqlStr = sqlStr & "    ,JITU_SURYO = '000000'"
			sqlStr = sqlStr & "    ,LK_SEQ_NO = ''"
			sqlStr = sqlStr & "    ,UPD_NOW = left(replace(replace(replace(convert(CURRENT_TIMESTAMP(),SQL_CHAR),'-',''),':',''),' ',''),14)"
			sqlStr = sqlStr & whereStr
'			if Msg("指定した条件全て、検品済にします。" & vbcrlf & sqlStr,vbYesNo) = vbYes then
				set rsList = db.Execute(sqlStr)
'			end if

			sqlStr = "select"
			sqlStr = sqlStr & " KEY_ID_NO"
			sqlStr = sqlStr & ",KENPIN_TANTO_CODE"
			sqlStr = sqlStr & ",KENPIN_YMD"
			sqlStr = sqlStr & ",KENPIN_HMS"
			sqlStr = sqlStr & ",JITU_SURYO"
			sqlStr = sqlStr & ",LK_SEQ_NO"
			sqlStr = sqlStr & ",UPD_NOW"
			sqlStr = sqlStr & " From " & tblStr & " y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by KEY_ID_NO"
		case "uJizenCancel"	' 事前チェックキャンセル
			sqlStr = "update " & tblStr
			sqlStr = sqlStr & " set LK_SEQ_NO = ''"
			sqlStr = sqlStr & whereStr
			set rsList = db.Execute(sqlStr)

			sqlStr = "select"
			sqlStr = sqlStr & " KEY_ID_NO"
			sqlStr = sqlStr & ",KENPIN_TANTO_CODE"
			sqlStr = sqlStr & ",KENPIN_YMD"
			sqlStr = sqlStr & ",KENPIN_HMS"
			sqlStr = sqlStr & ",LK_SEQ_NO"
			sqlStr = sqlStr & ",'事前チェックキャンセル'"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by KEY_ID_NO"
		case "dData"	' 削除
			sqlStr = "delete from " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			set rsList = db.Execute(sqlStr)

			sqlStr = "select @@rowcount as ""削除件数"""
'			sqlStr = sqlStr & " From " & tblStr
		end select
		db.CommandTimeout=900
		set rsList = db.Execute(sqlStr)
	%>
		<%
			strTh = "<TR BGCOLOR=""#ccffee"">"
			strTh = strTh & "<TH>No.</TH>"

			dim totalArray(255)
			For i=0 To rsList.Fields.Count-1
			    rem FILLER は表示しない
				if left(rsList.Fields(i).Name,6) <> "_filler" then
					select case rsList.Fields(i).name
					case "KEY_SYUKA_YMD"	fname = "伝票日付"
					case "KEY_MUKE_CODE"	fname = "出荷先"
					case "saki"		fname = "出荷先"
					case "allcnt"	fname = "合計<br>　"
					case "kan"		fname = "<br>(済)"
					case "per"		fname = "<br>(%)"
					case "c1"		fname = "月切<br>　"
					case "c1kan"	fname = "<br>(済)"
					case "c2"		fname = "スポット<br>　"
					case "c2kan"	fname = "<br>(済)"
					case "c3"		fname = "補充<br>　"
					case "c3kan"	fname = "<br>(済)"
					case "c9"		fname = "その他<br>　"
					case "c9kan"	fname = "<br>(済)"
					case "choku"	fname = "直送"
					case "qty"		fname = "個数"

					case "useId"	fname = "使用端末ID"
					case "usePrg"	fname = "使用中プログラム"
					case "jKbn"		fname = "事業部<br>区分"
					case "KAN_KBN"	fname = "完了<br>区分"
					case "KEY_CYU_KBN"	fname = "注文<br>区分"
					case "KEY_CYU_KBN"	fname = "注文<br>区分<br>(ﾎｽﾄ)"
					case "mCode"	fname = "出荷先"
					case "KEY_MUKE_CODE"	fname = "出荷先<br>読替"
					case "KEY_SYUKA_YMD"	fname = "伝票日付"
					case "denNo"	fname = "伝票№"
					case "ssNo"		fname = "ＳＳ追番"
					case "kNaiGai"	fname = "国内外"
					case "pn"		fname = "品番"
					case "dataType"	fname = "データ<br>種別"
					case "yoteiQty"	fname = "予定<br>数量"
					case "kakuQty"	fname = "確定<br>数量"
					case "textNo"	fname = "テキスト№"
					case "chokuKbn"	fname = "直送<br>区分"
					case "ioKbn"	fname = "入出庫<br>区分"
					case "rbKbn"	fname = "赤黒<br>区分"
					case "denType"	fname = "伝票<br>種別"
					case "pnNai"	fname = "品番(内部)"
					case "pName"	fname = "品名"
					case "mYosan"	fname = "予算単位<br>(元)"
					case "sYosan"	fname = "予算単位<br>(先)"
					case "hSoko"	fname = "倉庫<br>(ﾎｽﾄ)"
					case "hTana"	fname = "棚番<br>(ﾎｽﾄ)"
					case "sCode"	fname = "出荷先"
					case "sName"	fname = "出荷先名"
					case "kanDate"	fname = "完了日付"
					case "kenDate"	fname = "検品日付"
					case "TOK_KBN"	fname = "特売<br>区分"
					case "span1"	fname = ""
					case "span2"	fname = ""
					case else		fname = rsList.Fields(i).name
					end select
					strTh = strTh & "<TH title='" & rsList.Fields(i).name & "'>" & fname & "</TH>"
				end if
			Next
			strTh = strTh & "</TR>"
		%>
		<%=strTh%>
		<%
			cnt = 0
			Do While Not rsList.EOF
				cnt = cnt + 1
		%>
				<TR VALIGN='TOP'>
				<TD nowrap id="Integer"><%=cnt%></TD>
		<%
				For i=0 To rsList.Fields.Count-1
			        rem FILLER は表示しない
					if left(rsList.Fields(i).Name,6) <> "_filler" then
						' 値
						fValue = rtrim(rsList.Fields(i))
						if rsList.Fields(i).Name = "per" then
							tdTag = "<TD nowrap id=""Integer"">"
						elseif right(rsList.Fields(i).Name,1) = "＠" then
							tdTag = "<TD nowrap id=""Integer"">"
							if fValue < 0 then
								fValue = "未設定"
							else
								fValue = formatnumber(fValue,2,,,-1)
							end if
						elseif right(rsList.Fields(i).Name,1) = "比" then
							tdTag = "<TD nowrap id=""Integer"">"
							fValue = formatnumber(fValue,0,,,0) & "%"
						elseif right(rsList.Fields(i).Name,4) = "(平均)" then
							tdTag = "<TD nowrap id=""Integer"">"
							fValue = formatnumber(fValue,2,,,-1)
						elseif right(rsList.Fields(i).Name,2) = "金額" or right(rsList.Fields(i).Name,2) = "才数" then
							tdTag = "<TD nowrap id=""Integer"">"
							if isnull(fValue) = False then
								if i = 0 then
									totalArray(i) = cdbl(fValue)
								else
										totalArray(i) = totalArray(i) + cdbl(fValue)
								end if
							end if
							if fValue = 0 then
								fValue = ""
							elseif isnull(fValue) = True then
								fValue = ""
							else
								fValue = formatnumber(fValue,2,,,-1)
							end if
						elseif rsList.Fields(i).Name = "棚番_BC" then
							tdTag = "<TD class=""_BC"">"
							dim	src
							src = "https://www.fabrice.co.jp/cgi-bin/code128.cgi?code=/" & fValue
							fValue = fValue & "<p><a href=""" & src & """><img src=""" & src & """></href>"
						elseif right(rsList.Fields(i).Name,3) = "_BC" then
'							tdTag = "<TD align=""center"" style=""border-style: none;"">"
							tdTag = "<TD class=""_BC"">"
							src = "https://www.fabrice.co.jp/cgi-bin/code128.cgi?code=" & fValue
							src = "http://barcodes4.me/barcode/c39/" & fValue & ".png"
							fValue = fValue & "<p><a href=""" & src & """><img src=""" & src & """></href>"
						else
							' 位置定義（型）
							select case rsList.Fields(i).type
							Case 2		' 数値(Integer)
								tdTag = "<TD nowrap id=""Integer"">"
								if fValue = "-32768" then
									fValue = ""
								end if
							Case 2 , 3 , 5 ,131	' 数値(Integer)
								if i = 0 then
									totalArray(i) = clng(fValue)
								elseif fValue <> "" then
									totalArray(i) = totalArray(i) + clng(fValue)
								end if
								if fValue = 0 then
									fValue = ""
								elseif fValue <> "" then
									fValue = formatnumber(fValue,0,,,-1)
								end if
								tdTag = "<TD nowrap id=""Integer"">"
							Case 133		' 日付(Date)	
								tdTag = "<TD nowrap id=""Date"">"
								if len(fValue) > 0 then
									fValue = year(rsList.Fields(i)) & "/"
									if month(rsList.Fields(i)) < 10 then
										fValue = fValue & "0" & month(rsList.Fields(i))
									else
										fValue = fValue & month(rsList.Fields(i))
									end if
									fValue = fValue & "/"
									if day(rsList.Fields(i)) < 10 then
										fValue = fValue & "0" & day(rsList.Fields(i))
									else
										fValue = fValue & day(rsList.Fields(i))
									end if
								end if
							Case 129		' 文字列(Charactor)
								tdTag = "<TD nowrap id=""Charactor"">"
							Case else		' その他
								tdTag = "<TD nowrap>"
							end select
						end if
						Response.Write tdTag & fValue & "</TD>"
					end if
				Next
		    	Response.Write "</TR>"
				rsList.Movenext
			Loop
		%>
		<!-- 合計 -->
		<TR VALIGN='TOP'>
			<TD nowrap align="center">計</TD>
			<%For i=0 To rsList.Fields.Count-1
				select case rsList.Fields(i).type
				Case 2 , 3 , 5 ,131	' 数値(Integer)
					if right(rsList.Fields(i).Name,2) = "才数" then
			%>
						<TD nowrap id="Integer"><%=formatnumber(totalArray(i),2,,,-1)%></TD>
			<%
					else
			%>
						<TD nowrap id="Integer"><%=formatnumber(totalArray(i),0,,,-1)%></TD>
			<%		end if %>
			<%	Case else		' その他	%>
					<TD></TD>
			<%	end select	%>
			<%Next%>
    	</TR>
		<%=strTh%>
	</TABLE></div>
	<hr>
	<a href="javascript:showhide('sql')" title="SQLを 表示/非表示">SQL</a>
	<div id="sql" style="display:none;">
		<%=sqlStr%>
		<%=insStr%>
	</div>
	<%
		rsList.Close
		db.Close
		set rsList = nothing
		set db = nothing
	%>
	<SCRIPT LANGUAGE=javascript>
	<!--
		sqlForm.disabled = false;
		btnClip.disabled = false;
		$('#btnClip').text('結果をコピー');
//		cpTblBtn.value = "結果をコピー";
//		autoBtn.disabled = false;
	//-->
	</SCRIPT>
<% end if %>
</BODY>
</HTML>
