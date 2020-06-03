<%
Option Explicit
Response.Buffer = false
Response.Expires = -1
' formatnumber()
const TristateTrue			= -1	'ゼロを表示します
const TristateFalse			= 0		'ゼロを表示しません
const TristateUseDefault	= -2	'「地域のプロパティ」の設定値を使用します

Function GetVersion()
	dim	strVersion
	' 2007.06.18 y_sagyo_log.asp
	' 2007.06.29 検索条件:向け先 の対応
	strVersion = "2007.06.29 検索条件:向け先 の対応"
	strVersion = "2007.09.07 検索条件:ID_NO の対応"
	strVersion = "2007.09.12 検索条件:処理日 範囲指定 対応"
	strVersion = "2007.09.12 検索条件:事業部 対応"
	strVersion = "2007.09.12 出力形式：品番×月別 対応"
	strVersion = "2008.02.08 出力形式：倉庫別 対応"
	strVersion = "2008.02.09 出力形式：倉庫別 対応 TimeOut対応"
	strVersion = "2008.08.19 作業時間対応"
	strVersion = "2008.09.26 伝票枚数(伝票IDの数)対応"
	strVersion = "2008.10.08 事業部=S の検索対応"
	strVersion = "2008.10.09 タイムアウト時間を10分(600s)に変更"
	strVersion = "2009.12.24 倉庫別 標準棚番からではなく、移動元で集計するように変更"
	strVersion = "2009.02.22 出力形式：品番・棚別(E4:在庫精査 件数確認用) 追加"
	strVersion = "2009.07.13 検索条件：移動元／移動先 追加"
	strVersion = "2009.07.21 出力形式：メニュー・処理日別 追加"
	strVersion = "2009.08.26 出力形式：一覧表 PRG_ID 追加"
	strVersion = "2009.10.07 出力形式：出退社時刻 追加"
	strVersion = "2009.10.08 検索条件：事業部 -S 資材除く の対応"
	strVersion = "2009.10.08 検索条件：要因 複数指定(カンマ , で区切る) 対応"
	strVersion = "2009.10.15 出力形式：一覧表(対内,品名) 追加"
	strVersion = "2009.11.09 出力形式：向け先別 追加／検索条件 向け先 先頭ハイフン(-)でNOT検索対応"
	strVersion = "2009.11.25 出力形式：品番・入出庫回数 対応"
	strVersion = "2010.03.03 検索条件：向け先 の文字数制限(8)を解除"
	strVersion = "2010.06.09 出力形式：メニュー別：移動件数(在庫移動した件数) を追加"
	strVersion = "2010.08.11 伝票枚数(伝票IDの数) 同一IDを1件とカウントするように変更"
	strVersion = "2010.08.20 ポップアップメニュー対応"
	strVersion = "2010.09.13 一覧表 指示書No. ラベルCheck 現品票Check 追加/makeWhere対応"
	strVersion = "2012.04.12 出力形式：品番・入出庫回数 エラー「ﾇﾙが不正です。」の対応"
	strVersion = "2012.05.17 出力形式：品番(エアコン移管用)の対応"
	strVersion = "2013.07.23 床暖対応：GetDbName()に変更"
	GetVersion = "2016.08.10 検索条件(追加)端末ID、変数をGetRequest()で取得に変更"
	GetVersion = "2017.09.21 出力形式：品番(金額),事(金額)"
	GetVersion = "2017.09.22 出力形式：端末ID・要因別"
	GetVersion = "2018.12.25 出力形式：一覧表 項目追加[Memo][外装Check]"
	GetVersion = "2019.01.10 出力形式：一覧表 項目追加[JAN]"
	GetVersion = "2020.06.03 検索条件：処理時刻 対応"
End Function
Function GetToday()
	GetToday = right("0000" & year(now),4) & right("00" & month(now),2) & right("00" & day(now),2)
End Function
%>
<!--#include file="makeWhere.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="result.css" TITLE="CSS">
<TITLE><%=GetCenterName()%> 作業ログ</TITLE>
<!-- jdMenu head用 include 開始 -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<!-- jdMenu head用 include 終了 -->
<SCRIPT LANGUAGE="JavaScript"><!--
navi = navigator.userAgent;

function DoCopy(arg){
	var doc = document.body.createTextRange();
	doc.moveToElementText(document.all(arg));
	doc.execCommand("copy");
	window.alert("クリップボードへコピーしました。\n貼り付けできます。" );
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
--></SCRIPT>
</HEAD>
<BODY>
<!-- jdMenu body用 include 開始 -->
<!--#include file="jdmenu-sdc.asp" -->
<!-- jdMenu body用 include 終了 -->

  <FORM name="sqlForm">
  <div id="sqlDiv">
	<table id="sqlTbl">
		<caption style="text-align:left;">作業ログ検索</caption>
		<tr>
			<th>事業部</th>
			<th>処理日</th>
			<th>処理時刻</th>
			<th>担当者</th>
			<th>メニューNo</th>
			<th>要因</th>
			<th>端末ID</th>
			<th>品番</th>
			<th>向け先</th>
			<th>ID-No</th>
            <th>指示書No</th>
			<th>移動元</th>
			<th>移動先</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="JGYOBU" id="JGYOBU" VALUE="<%=GetRequest("JGYOBU","")%>" size="2" style="text-align:center;">
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="dt" id="dt" VALUE="<%=GetRequest("dt",GetToday())%>" size="10">
				～
				<INPUT TYPE="text" NAME="dtTo" id="dtTo" VALUE="<%=GetRequest("dtTo","")%>" size="10">
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="tmFr" id="tmFr" VALUE="<%=GetRequest("tmFr", "")%>" size="5">
				～
				<INPUT TYPE="text" NAME="tmTo" id="tmTo" VALUE="<%=GetRequest("tmTo", "")%>" size="5">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="Tanto" id="Tanto" VALUE="<%=GetRequest("Tanto","")%>" size="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="MENU_NO" id="MENU_NO" VALUE="<%=GetRequest("MENU_NO","")%>" size="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="Yoin" id="Yoin" VALUE="<%=GetRequest("Yoin","")%>" size="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="WEL_ID" id="WEL_ID" VALUE="<%=GetRequest("WEL_ID","")%>" size="6">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="pn" id="pn" VALUE="<%=GetRequest("pn","")%>" size="20">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="MukeCode" id="MukeCode" VALUE="<%=GetRequest("MukeCode","")%>" size="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="ID_NO" id="ID_NO" VALUE="<%=GetRequest("ID_NO","")%>" size="15">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="SHIJI_NO" id="SHIJI_NO" VALUE="<%=GetRequest("SHIJI_NO","")%>" size="12">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="TanaFROM" id="TanaFROM" VALUE="<%=GetRequest("TanaFROM","")%>" size="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="TanaTO" id="TanaTO" VALUE="<%=GetRequest("TanaTO","")%>" size="10">
			</td>
		</tr>
		<tr>
			<td colspan="13" nowrap>
				<table>
				<tr>
				<td><b>出力形式：</b></td>
				<td>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable" id="pTable">
					<label for="pTable">要因</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMenu" id="pTableMenu">
					<label for="pTableMenu">メニュー</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTantoMenu" id="pTableTantoMenu">
					<label for="pTableTantoMenu">担当者・メニュー</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTantoYoin" id="pTableTantoYoin">
					<label for="pTableTantoYoin">担当者・要因</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTanto" id="pTableTanto">
					<label for="pTableTanto">担当者</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMenuYoin" id="pTableMenuYoin">
					<label for="pTableMenuYoin">メニュー・要因</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMenuDt" id="pTableMenuDt">
					<label for="pTableMenuDt">メニュー・処理日</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableInOut" id="pTableInOut">
					<label for="pTableInOut">品番・入出庫回数</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSoko" id="pTableSoko">
					<label for="pTableSoko">倉庫</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMukeCode" id="pTableMukeCode">
					<label for="pTableMukeCode">向け先</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnMonth" id="pTablePnMonth">
					<label for="pTablePnMonth">品番×月</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnTana" id="pTablePnTana">
					<label for="pTablePnTana">品番・棚</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTana" id="pTableTana">
					<label for="pTableTana">棚</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableWelID" id="pTableWelID">
					<label for="pTableWelID">スキャナ最新処理</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableWork" id="pTableWork">
					<label for="pTableWork">出退社時刻</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
					<label for="pList">一覧表</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListItem" id="pListItem">
					<label for="pListItem">一覧表(対内,品名)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
					<label for="pListAll">一覧表(全項目)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePn" id="pTablePn">
					<label for="pTablePn">品番(エアコン移管用)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnAmt" id="pTablePnAmt">
					<label for="pTablePnAmt">品番(金額)</label><!--Amt(Amount:金額)-->
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableJAmt" id="pTableJAmt">
					<label for="pTableJAmt">事(金額)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableId" id="pTableId">
					<label for="pTableId">端末ID・要因</label>
				</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr bordercolor=White>
			<td colspan="13" nowrap>
				<INPUT TYPE="submit" value="検索" id=submit1 name=submit1>
				<INPUT TYPE="reset" value="リセット" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
				最大件数：<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=GetRequest("max","1000")%>" size="8" maxlength="6">
				<%=GetVersion()%>
			</td>
		</tr>
	</table>
	</div>
  </FORM>
<SCRIPT LANGUAGE='JavaScript'>
	ptypeChange('<%=GetRequest("ptype","pTable")%>');
</SCRIPT>
<%	if len(GetRequest("submit1","")) > 0 then %>
	<SCRIPT LANGUAGE=javascript><!--
		sqlForm.disabled = true;
	//--></SCRIPT>
	<%
		dim db
		Set db = Server.CreateObject("ADODB.Connection")
		db.open GetDbName()
		dim	sqlStr
		sqlStr = ""
		dim	whereStr
		whereStr = ""
		dim	andStr
		andStr = " where"

		whereStr = makeWhere(whereStr,"p.JITU_DT",GetRequest("dt",GetToday()),GetRequest("dtTo",""))
		whereStr = makeWhere(whereStr,"p.JITU_TM",GetRequest("tmFr", ""), GetRequest("tmTo", ""))
		whereStr = makeWhere(whereStr,"p.JGYOBU",GetRequest("JGYOBU",""),"")
		whereStr = makeWhere(whereStr,"p.TANTO_CODE",GetRequest("Tanto",""),"")
		whereStr = makeWhere(whereStr,"p.MENU_NO",GetRequest("MENU_NO",""),"")
		whereStr = makeWhere(whereStr,"p.RIRK_ID",GetRequest("Yoin",""),"")
		whereStr = makeWhere(whereStr,"p.MUKE_CODE",GetRequest("MukeCode",""),"")
		whereStr = makeWhere(whereStr,"p.ID_NO",GetRequest("ID_NO",""),"")
		whereStr = makeWhere(whereStr,"p.SHIJI_NO",GetRequest("SHIJI_NO",""),"")
		whereStr = makeWhere(whereStr,"p.HIN_GAI",GetRequest("pn",""),"")
		whereStr = makeWhere(whereStr,"p.FROM_SOKO + p.FROM_RETU + p.FROM_REN + p.FROM_DAN"	,GetRequest("TanaFROM",""),"")
		whereStr = makeWhere(whereStr,"p.TO_SOKO   + p.TO_RETU   + p.TO_REN   + p.TO_DAN"	,GetRequest("TanaTO",""),"")
		whereStr = makeWhere(whereStr,"p.WEL_ID",GetRequest("WEL_ID",""),"")

		sqlStr = "select "
		dim	lngMax
		lngMax = CLng(GetRequest("max","1000"))
		if lngMax > 0 then
			sqlStr = sqlStr & " top " & lngMax
		end if
		select case GetRequest("ptype","pTable")
		case "pTable"	' 要因別 集計表
			sqlStr = sqlStr & " p.RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""要因"""
			sqlStr = sqlStr & ",sum(if(p.RIRK_ID = 'ST' or p.RIRK_ID = 'EN',0,1)) ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) + convert(p.MI_JITU_QTY,SQL_decimal)) ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(p.WORK_TM,SQL_decimal))/60 ""作業時間(分)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(p.ID_NO)= '',null(),rtrim(p.ID_NO))) ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""要因"""
			sqlStr = sqlStr & " order by ""要因"""
		case "pTableTanto"	    ' 担当者別 集計表
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""担当者"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""作業時間(分)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""担当者"""
			sqlStr = sqlStr & " order by ""担当者"""
		case "pTableWelID"	    ' スキャナ最新処理
			sqlStr = sqlStr & " p.WEL_ID ""端末ID"""
			sqlStr = sqlStr & ",p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') ""担当者"""
			sqlStr = sqlStr & ",JITU_DT + ' ' + JITU_TM ""処理日時"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",p.RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""要因"""
			sqlStr = sqlStr & ",ID_NO ""ID-No."""
			sqlStr = sqlStr & ",HIN_GAI ""品番"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC) ""数量"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) ""商済"""
			sqlStr = sqlStr & ",convert(MI_JITU_QTY,SQL_NUMERIC) ""未商"""
			sqlStr = sqlStr & ",MUKE_CODE ""向け先"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN ""移動元"""
			sqlStr = sqlStr & ",TO_SOKO   + TO_RETU   + TO_REN   + TO_DAN   ""移動先"""
			sqlStr = sqlStr & ",convert(WORK_TM,SQL_NUMERIC) ""作業時間(秒)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & "  inner join (select WEL_ID,max(JITU_DT+JITU_TM+WEL_ID+RIRK_ID) as maxtm from P_SAGYO_LOG "
			sqlStr = sqlStr & makeWhere("","JITU_DT",GetRequest("dt",GetToday()),GetRequest("dtTo",""))
'			if len(dtStr) + len(dtToStr) > 0 then
'				if len(dtStr) > 0 and len(dtToStr) > 0 then
'					sqlStr = sqlStr & " (JITU_DT between '" & dtStr & "' and '" & dtToStr & "')"
'				else
'					sqlStr = sqlStr & " JITU_DT like '" & dtStr & "%'"
'				end if
'			end if
			sqlStr = sqlStr & " group by WEL_ID) mx on (p.JITU_DT + p.JITU_TM + p.WEL_ID + p.RIRK_ID = maxtm)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""端末ID"""
		case "pTableTantoYoin"	' 担当者・要因別 集計表
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') ""担当者"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""要因"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 ""作業時間(分)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""担当者"",""MENU"",""要因"""
			sqlStr = sqlStr & " order by ""担当者"",""MENU"",""要因"""
		case "pTableId"		' 端末ID・要因別
			sqlStr = sqlStr & " p.WEL_ID ""端末ID"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""要因"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 ""作業時間(分)"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
'			sqlStr = sqlStr & "  left outer join TANTO t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""端末ID"",""MENU"",""要因"""
			sqlStr = sqlStr & " order by ""端末ID"",""MENU"",""要因"""
		case "pTableMenuYoin"	' メニュー・要因別 集計表
			sqlStr = sqlStr & " p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""要因"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""作業時間(分)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""MENU"",""要因"""
			sqlStr = sqlStr & " order by ""MENU"",""要因"""
		case "pTableMenuDt"	' メニュー・処理日別 集計表
			sqlStr = sqlStr & " p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",left(p.JITU_DT,8) as ""処理日"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""作業時間(分)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""MENU"",""処理日"""
			sqlStr = sqlStr & " order by ""MENU"",""処理日"""
		case "pTableMenu"	    ' メニュー別 集計表
			sqlStr = sqlStr & " p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""作業件数"""
			sqlStr = sqlStr & ",sum(if(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)=0,0,1)) as ""移動件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""作業時間(分)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""MENU"""
			sqlStr = sqlStr & " order by ""MENU"""
		case "pTableTantoMenu"	' 担当者・メニュー別 集計表
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""担当者"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""作業数量"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""作業時間(分)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""伝票枚数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""担当者"",""MENU"""
			sqlStr = sqlStr & " order by ""担当者"",""MENU"""
		case "pTableSoko"	' 倉庫別 集計表
			sqlStr = sqlStr & " RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""要因"""
'			sqlStr = sqlStr & ",i.ST_Soko,'')  as ""倉庫"""
			sqlStr = sqlStr & ",FROM_SOKO as ""倉庫(元)"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""作業数量"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""伝票枚数"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""作業時間(分)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
'			sqlStr = sqlStr & "  left outer join ITEM as i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""要因"",""倉庫(元)"""
			sqlStr = sqlStr & " order by ""要因"",""倉庫(元)"""
		case "pTableMukeCode"	' 向け先別 集計表
			sqlStr = sqlStr & " MUKE_CODE as ""向け先"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""要因"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""作業件数"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""作業数量"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""伝票枚数"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""伝票枚数"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""作業時間(分)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""向け先"",""要因"""
			sqlStr = sqlStr & " order by ""向け先"",""要因"""
		case "pTableInOut"		' 品番・入出庫回数
			sqlStr = "select"
			sqlStr = sqlStr & " p.JGYOBU ""事"""
			sqlStr = sqlStr & ",p.HIN_GAI ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME ""品名"""
			sqlStr = sqlStr & ",z.qty ""在庫数"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '7',1,0)) ""入庫<br>回数"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '7',convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC),0)) ""入庫<br>数"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '4',1,0)) ""出庫<br>回数"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '4',convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC),0)) ""出庫<br>数"""
			sqlStr = sqlStr & ",convert(i.G_ST_SHITAN,sql_numeric) ""仕入＠"""
			sqlStr = sqlStr & ",convert(i.G_ST_SHITAN,sql_numeric) * z.qty ""在庫金額"""
'			sqlStr = sqlStr & ",(year(now())*100+month(now()))-convert(left(max(p.JITU_DT),6),sql_numeric) as ""不移動<br>月数"""
			sqlStr = sqlStr & ",datediff(month,convert(left(max(p.JITU_DT),4)+'-'+SUBSTRING(max(p.JITU_DT),5,2)+'-'+right(max(p.JITU_DT),2),sql_date),now()) ""不移動<br>月数"""
			sqlStr = sqlStr & ",max(p.JITU_DT) ""最終移動日"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join ITEM i on (i.JGYOBU = p.JGYOBU and i.NAIGAI = '1' and i.HIN_GAI = p.HIN_GAI)"
			sqlStr = sqlStr & "  left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(convert(YUKO_Z_QTY,sql_numeric)) qty from zaiko group by JGYOBU,NAIGAI,HIN_GAI) z on (z.JGYOBU = i.JGYOBU and z.NAIGAI = i.NAIGAI and z.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事"",""品番"",""品名"",""在庫数"",""仕入＠"""
			sqlStr = sqlStr & " order by ""出庫<br>回数"" desc, ""事"",""品番"""

		case "pTablePnMonth"	' 品番×月別 集計表
			dim	sumStr
			dim	sqlStr2
			db.CommandTimeout		= 180	' 180

			sqlStr2 = "select distinct"
			sqlStr2 = sqlStr2 & " left(JITU_DT,6)  as ym"
			sqlStr2 = sqlStr2 & " From P_SAGYO_LOG as p"
			sqlStr2 = sqlStr2 & whereStr
			sqlStr2 = sqlStr2 & " order by ym"
			set rsList = db.Execute(sqlStr2)

			sqlStr = sqlStr & " HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + y.YOIN_DNAME as ""要因"""
			sumStr = "convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)"
			Do While Not rsList.EOF
				sqlStr = sqlStr & ",sum(if(left(JITU_DT,6) ='" & rsList.Fields("ym") & "'," & sumStr & ",0)) as """ & rsList.Fields("ym") & """"
				rsList.Movenext
			loop
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""品番"",""要因"""
			sqlStr = sqlStr & " order by ""品番"",""要因"""

		case "pTablePn"		' 品番
			sqlStr = sqlStr & " p.HIN_GAI ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME ""品名"""
'			sqlStr = sqlStr & ",i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN ""標準棚番"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '7',convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC),0)) ""移動数"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '1',convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC),0)) ""訂正＋"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '2',convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC),0)) ""訂正▲"""
			sqlStr = sqlStr & ",z.ac as ""AC在庫数"""
			sqlStr = sqlStr & ",z.az as ""AZ在庫数"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join ITEM i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & "  left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(if(Soko_No='AC',convert(YUKO_Z_QTY,sql_numeric),0)) as AC,sum(if(Soko_No='AZ',convert(YUKO_Z_QTY,sql_numeric),0)) as AZ from zaiko group by JGYOBU,NAIGAI,HIN_GAI) as z on (z.JGYOBU = i.JGYOBU and z.NAIGAI = i.NAIGAI and z.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""品番"",""品名"",""AC在庫数"",""AZ在庫数"""
			sqlStr = sqlStr & " having ""移動数""<>0"
			sqlStr = sqlStr & " order by ""品番"""
		case "pTablePnAmt"		' 品番(金額)
			sqlStr = sqlStr & " p.JGYOBU ""事"""
			sqlStr = sqlStr & ",p.HIN_GAI ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME ""品名"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY,SQL_decimal)) ""未商"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal)) ""商済"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)) ""商済<br>工料"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) ""商済<br>箱代"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & " left outer join ITEM i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事"",""品番"",""品名"""
			sqlStr = sqlStr & " order by ""事"",""品番"""
		case "pTableJAmt"		'事(金額)
			sqlStr = sqlStr & " p.JGYOBU ""事"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY,SQL_decimal)) ""未商"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal)) ""商済"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)) ""商済<br>工料"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) ""商済<br>箱代"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & " left outer join ITEM i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事"""
			sqlStr = sqlStr & " order by ""事"""
		case "pTablePnTana"	' 品番・棚別 集計表
			sqlStr = sqlStr & " p.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""移動元"""
			sqlStr = sqlStr & ",count(*) as ""回数"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC)) as ""数量"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC)) as ""商済"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY  ,SQL_NUMERIC)) as ""未商"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC)) as ""作業時間(秒)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""品番"",""移動元"""
			sqlStr = sqlStr & " order by ""品番"",""移動元"""
		case "pTableTana"	' 棚別 集計表
			sqlStr = sqlStr & " FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""移動元"""
			sqlStr = sqlStr & ",count(*) as ""回数"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC)) as ""数量"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC)) as ""商済"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY  ,SQL_NUMERIC)) as ""未商"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC)) as ""作業時間(秒)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""移動元"""
			sqlStr = sqlStr & " order by ""移動元"""
		case "pTableWork"	' 出退社時刻
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""担当者"""
			sqlStr = sqlStr & ",JITU_DT as ""処理日"""
			sqlStr = sqlStr & ",min(left(JITU_TM,2)+':'+SUBSTRING(JITU_TM,3,2)) as ""出社"""
			sqlStr = sqlStr & ",if(min(left(JITU_TM,4))<>max(left(JITU_TM,4)),max(left(JITU_TM,2)+':'+SUBSTRING(JITU_TM,3,2)),'') as ""退社"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""要因"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""担当者"",""処理日"",""MENU"",""要因"""
			sqlStr = sqlStr & " order by ""担当者"",""処理日"""
		case "pList"	' 一覧表
			sqlStr = sqlStr & " JITU_DT + ' ' + JITU_TM ""処理日時"""
			sqlStr = sqlStr & ",p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') ""担当者"""
			sqlStr = sqlStr & ",WEL_ID ""WEL_ID"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""要因"""
			sqlStr = sqlStr & ",ID_NO as ""ID-No."""
			sqlStr = sqlStr & ",HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC) as ""数量"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) as ""商済"""
			sqlStr = sqlStr & ",convert(MI_JITU_QTY,SQL_NUMERIC) as ""未商"""
			sqlStr = sqlStr & ",MUKE_CODE as ""向け先"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""移動元"""
			sqlStr = sqlStr & ",TO_SOKO   + TO_RETU   + TO_REN   + TO_DAN   as ""移動先"""
			sqlStr = sqlStr & ",p.Memo ""Memo"""
			sqlStr = sqlStr & ",SHIJI_No ""指示書No."""
			sqlStr = sqlStr & ",convert(p.HIN_CHECK_LABEL_CNT,SQL_decimal) ""ラベルCheck"""
			sqlStr = sqlStr & ",convert(p.HIN_CHECK_GENPIN_CNT,SQL_decimal) ""現品票Check"""
			sqlStr = sqlStr & ",convert(p.HIN_CHECK_GAISOU_CNT,SQL_decimal) ""外装Check"""
			sqlStr = sqlStr & ",JAN_CODE ""JAN"""
			sqlStr = sqlStr & ",convert(WORK_TM,SQL_NUMERIC) ""作業時間(秒)"""
			sqlStr = sqlStr & ",PRG_ID"
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""処理日時"""
		case "pListItem"	' 一覧表(対内品番,品名)
			sqlStr = sqlStr & " JITU_DT + ' ' + JITU_TM as ""処理日時"""
			sqlStr = sqlStr & ",p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""担当者"""
			sqlStr = sqlStr & ",WEL_ID as ""WEL_ID"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""要因"""
			sqlStr = sqlStr & ",ID_NO as ""ID-No."""
			sqlStr = sqlStr & ",p.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAI as ""対内品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC) as ""数量"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) as ""商済"""
			sqlStr = sqlStr & ",convert(MI_JITU_QTY,SQL_NUMERIC) as ""未商"""
			sqlStr = sqlStr & ",MUKE_CODE as ""向け先"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""移動元"""
			sqlStr = sqlStr & ",TO_SOKO   + TO_RETU   + TO_REN   + TO_DAN   as ""移動先"""
			sqlStr = sqlStr & ",convert(WORK_TM,SQL_NUMERIC) as ""作業時間(秒)"""
			sqlStr = sqlStr & ",PRG_ID"
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join ITEM as i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""処理日時"""
		case "pListAll"	' 一覧表(全項目)
			sqlStr = sqlStr & " * From P_SAGYO_LOG as p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by JITU_DT,JITU_TM"
		end select
		db.CommandTimeout		= 600	' 180
		Server.ScriptTimeout	= 600	' 180
	%>
	<div>
		<INPUT TYPE="button" onClick="DoCopy('resultDiv')"
			value="検索中...ScriptTimeout=<%=Server.ScriptTimeout%>/CommandTimeout=<%=db.CommandTimeout%>"
			id="cpTblBtn" disabled>
	</div>
	<%
		dim rsList
		set rsList = db.Execute(sqlStr)
	%>

	<SCRIPT LANGUAGE=javascript><!--
		cpTblBtn.value = "テーブル出力中...";
	//--></SCRIPT>

	<div id='resultDiv'>
	<table id="resultTbl">
		<caption align='left'></caption>
		<TR BGCOLOR=#ccffee>
			<TH>No</TH>
		<%
			dim	i
			For i=0 To rsList.Fields.Count-1
				dim	fName
				select case rsList.Fields(i).name
				case else		fName = rsList.Fields(i).name
				end select
				Response.Write "<TH title='" & rsList.Fields(i).name & " " & rsList.Fields(i).type & "'>" & fName & "</TH>"
			Next
		%>
		</TR>
		<%
			dim	cnt
			cnt = 0
			Do While Not rsList.EOF
				cnt = cnt + 1
'				lngMax = clng(maxStr)
		%>
				<TR VALIGN='TOP'>
				<TD nowrap id="Integer"><%=cnt%></TD>
		<%
				For i=0 To rsList.Fields.Count-1
					' 値
					dim	fValue
					fValue = ""
					if isnull(rsList.Fields(i)) = False then
						if isempty(rsList.Fields(i)) = False then
							dim	v
							v = rsList.Fields(i)
							fValue = rtrim(v)
						end if
					end if
                    fName = rsList.Fields(i).name
					' 位置定義（型）
					dim	tdTag
					if right(fName,1) = "＠" then
						tdTag = "<TD nowrap id=""Integer"">"
						if isnull(fValue) = true then
							fValue = ""
						elseif fValue = "" then
							fValue = ""
						else
							fValue = formatnumber(fValue,2,,,-1)
						end if
					elseif right(fName,2) = "金額" then
						tdTag = "<TD nowrap id=""Integer"">"
						if isnull(fValue) = true then
							fValue = ""
						elseif fValue = "" then
							fValue = ""
						else
							fValue = formatnumber(fValue,0,,,-1)
						end if
					else
						select case rsList.Fields(i).type
						Case 2		' 数値(Integer)
							tdTag = "<TD nowrap id=""Integer"">"
							if fValue = "-32768" then
								fValue = ""
							end if
						Case 2 , 3 , 5	,131' 数値(Integer)
							if fName = "作業時間(分)" then
								fValue = formatnumber(round(fValue,1),1,true,false,TristateTrue)
							else
								if fValue <> "" then
								    if fValue = 0 then
									    fValue = ""
									end if
							    end if
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
							tdTag = "<TD>"
						end select
					end if
					Response.Write tdTag & fValue & "</TD>"
				Next
		    	Response.Write "</TR>"
				rsList.Movenext
			Loop
		%>
	</TABLE></div>
	<p>
	<div id="sql">
		<%=sqlStr%><br>
	</div>
	<%
		rsList.Close
		db.Close
		set rsList = nothing
		set db = nothing
	%>
	<SCRIPT LANGUAGE=javascript><!--
		sqlForm.disabled = false;
		cpTblBtn.disabled = false;
		cpTblBtn.value = "結果をコピー";
	//--></SCRIPT>
<% end if %>
</BODY>
</HTML>
