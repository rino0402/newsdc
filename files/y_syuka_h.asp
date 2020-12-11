<% Option Explicit	%>
<% Response.Buffer = false %>
<% Response.Expires = -1 %>
<% 
	Server.ScriptTimeout = 900

	dim	versionStr

	versionStr = "2007.10.16 検索条件：品番で検索できない不具合修正"
	versionStr = "2007.10.29 久留米運輸の問合せNoにリンクを追加"
	versionStr = "2007.11.02 項目名がおかしくなったのを修正"
	versionStr = "2010.02.24 出力形式：集計表(便別送状枚数) 復旧/一覧表 才数・重量 追加"
	versionStr = "2010.02.25 検索条件：才数 追加"
	versionStr = "2010.03.03 出力形式：集計表(個口)  正しい「口数」を集計できるように変更"
	versionStr = "2010.04.07 出力形式：一覧表 Tel/郵便番号 追加"
	versionStr = "2010.04.07 出力形式：集計表(個口) をクリックすると、集計表(便別) が選択される不具合修正"
	versionStr = "2010.05.11 出力形式：集計表(運送会社別) 送状枚数：ID_NO 7桁で集計"
	versionStr = "2010.08.20 ポップアップメニュー対応"
	versionStr = "2010.09.22 出力形式：一覧表：集約送り先コード 追加"
	versionStr = "2011.05.06 出力形式：一覧表(邸別照合) ：追加(邸別注文データと照合)"
	versionStr = "2011.06.06 出力形式：集計表(件管/品管No) ：追加"
	versionStr = "2012.10.11 検索条件：便 追加"
	versionStr = "2012.10.29 検索条件：送り先名 追加"
	versionStr = "2013.07.18 出力形式：一覧表(邸別照合)：項目追加 件管No. , 品管No."
	versionStr = "2014.07.30 検索条件：運送会社 追加"
	versionStr = "2016.05.31 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～福島県)"
	versionStr = "2016.06.02 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～栃木県)"
	versionStr = "2016.06.03 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～栃木県) 解除:奥尻郡奥尻町 常総市菅生町"
	versionStr = "2016.06.06 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～埼玉県)"
	versionStr = GetVersion()
	versionStr = "2020.07.10 クリップボード出力対応(Chrome)"
	versionStr = "2020.07.11 検索できない不具合修正"
	versionStr = "2020.07.28 IE11で配達不可が先頭に表示されるように修正"
%>
<HTML>
<HEAD>
<%
	dim	fname
	dim	tblStr
	dim	ptypeStr
	dim	delStr
	dim	dtStr
	dim	dtToStr
	dim	MUKE_CODEStr
	dim	CYU_KBNStr
	dim	DATA_KBNStr
	dim	HAN_KBNStr
	dim	JGYOBUStr
	dim	TOK_KBNStr
	dim	KEY_ID_NOStr
	dim	DEN_NOStr
	dim	pnStr
	dim	KAN_KBNStr
	dim	SYUKO_SYUSIStr
	dim	SAI_SUStr
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
	dim	OKURI_NOStr
	dim	CANCEL_FStr
	dim	dbName
	dim	lngMax
	dim	maxStr
	dim	COL_OKURISAKI_CDStr
	dim	strOKURISAKI_CD
	dim	strINS_BIN
	dim	strMUKE_NAME

	dbName	= "newsdc"
	submitStr			= Request.QueryString("submit1")
	autoStr = 0
	
	tblStr = ucase(Request.QueryString("tbl"))
	if len(tblStr) = 0 then
		tblStr = "Y_SYUKA_H"
	end if
	if tblStr = "Y_SYUKA_H" then
		delStr = ""
	else
		delStr = "削除済"
	end if
	ptypeStr = Request.QueryString("ptype")
	if len(ptypeStr) = 0 then
		ptypeStr = "pTable"
	end if
	dtStr = Request.QueryString("dt")
	dtToStr = Request.QueryString("dtTo")
	MUKE_CODEStr	= ucase(Request.QueryString("MUKE_CODE"))
	CYU_KBNStr		= ucase(Request.QueryString("CYU_KBN"))
	DATA_KBNStr		= ucase(Request.QueryString("DATA_KBN"))
	HAN_KBNStr		= ucase(Request.QueryString("HAN_KBN"))
	JGYOBUStr		= ucase(Request.QueryString("JGYOBU"))
	TOK_KBNStr		= ucase(Request.QueryString("TOK_KBN"))
	KEY_ID_NOStr	= ucase(Request.QueryString("KEY_ID_NO"))
	DEN_NOStr		= ucase(Request.QueryString("DEN_NO"))
	pnStr			= ucase(Request.QueryString("pn"))
	KAN_KBNStr		= ucase(Request.QueryString("KAN_KBN"))
	SYUKO_SYUSIStr	= ucase(Request.QueryString("SYUKO_SYUSI"))
	SAI_SUStr		= ucase(Request.QueryString("SAI_SU"))
	submitStr		= Request.QueryString("submit1")
	OKURI_NOStr		= ucase(Request.QueryString("OKURI_NO"))
	CANCEL_FStr		= ucase(Request.QueryString("CANCEL_F"))
	COL_OKURISAKI_CDStr		= ucase(Request.QueryString("COL_OKURISAKI_CD"))
	strOKURISAKI_CD		= ucase(Request.QueryString("OKURISAKI_CD"))
	strINS_BIN		= ucase(Request.QueryString("INS_BIN"))
	strMUKE_NAME		= Request.QueryString("MUKE_NAME")

	maxStr			= ucase(Request.QueryString("max"))
	if maxStr = "" then
		maxStr = 2000
	end if
	lngMax = clng(maxStr)

%>
<!--#include file="info.txt" -->
<!--#include file="makeWhere.asp" -->
<!--#include file="GetHaitatsu.vbs" -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="result.css" TITLE="CSS">
<TITLE><%=centerStr%> 出荷予定</TITLE>
<!-- jdMenu head用 include 開始 -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="https://cdn.jsdelivr.net/clipboard.js/1.5.3/clipboard.min.js"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<script src="jquery.tablesorter.js" type="text/javascript"></script>
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

	function uKenpinClick() {
		if ( window.confirm("検品OKにします") == false ) {
			ptypeChange("pTable");
		}
	}
	function DeleteClick() {
		if ( window.confirm("データを削除します\n＊元に戻せませんが＊よろしいですか？") == false ) {
			ptypeChange("pTable");
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
		dtStr = year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2)
	end if
%>
  <FORM name="sqlForm">
  <div id="sqlDiv">
	<table id="sqlTbl">
		<caption style="text-align:left;"><%=centerStr%> 出荷予定検索 <%=delStr%></caption>
		<tr>
			<th>伝票日付</th>
			<th>向け先</th>
			<th>ID-No</th>
			<th>伝票No</th>
			<th>品番</th>
			<th>才数</th>
			<th>問合せNo</th>
			<th>便</th>
			<th>キャンセル</th>
			<th>送り先</th>
			<!--th>集約送り先コード</th-->
			<!--th>送り先名</th-->
			<th>運送会社</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="dt" id="dt" VALUE="<%=dtStr%>" size="10" maxlength="8"><br>
				～<br>
				<INPUT TYPE="text" NAME="dtTo" id="dtTo" VALUE="<%=dtToStr%>" size="10" maxlength="8">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="MUKE_CODE" id="MUKE_CODE" VALUE="<%=MUKE_CODEStr%>" size="14" maxlength="8"><br>
				<div class="note">
				積水全て<br>
				積水注文あり<br>
				積水注文なし
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEY_ID_NO" id="KEY_ID_NO" VALUE="<%=KEY_ID_NOStr%>" size="15" maxlength="12">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="DEN_NO" id="DEN_NO" VALUE="<%=DEN_NOStr%>" size="10" maxlength="8">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="pn" id="pn" VALUE="<%=pnStr%>" size="20">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="SAI_SU" id="SAI_SU" VALUE="<%=SAI_SUStr%>" size="5">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="OKURI_NO" id="OKURI_NO" VALUE="<%=OKURI_NOStr%>" size="15">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="INS_BIN" id="INS_BIN" VALUE="<%=strINS_BIN%>" size="4"><br>
				<div class="note">
				01:1便<br>
				02:2便
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="CANCEL_F" id="CANCEL_F" VALUE="<%=CANCEL_FStr%>" size="2"><br>
				<div class="note">
				0:キャンセル除く<br>
				1:キャンセルのみ
				</div>
			</td>
			<td align="center">
				<div><INPUT TYPE="text" NAME="OKURISAKI_CD" 		VALUE="<%=GetRequest("OKURISAKI_CD","")%>"		size="15" placeholder = "送り先コード"></div>
				<div><INPUT TYPE="text" NAME="COL_OKURISAKI_CD" 	VALUE="<%=GetRequest("COL_OKURISAKI_CD","")%>"	size="15" placeholder = "集約送り先コード"></div>
				<div><INPUT TYPE="text" NAME="MUKE_NAME" 		VALUE="<%=GetRequest("MUKE_NAME","")%>" 			size="15" placeholder = "送り先名"></div>
				<div><INPUT TYPE="text" NAME="JYUSHO"			VALUE="<%=GetRequest("JYUSHO","")%>"				size="15" placeholder = "送り先住所"></div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="UNSOU_KAISHA" id="UNSOU_KAISHA" VALUE="<%=GetRequest("UNSOU_KAISHA","")%>" size="12">
			</td>
		</tr>
		<tr>
			<td align="right">
				<b>出力形式：</b>
			</td>
			<td colspan="10">
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable" id="pTable">
					<label for="pTable">集計表(便別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKoguchi" id="pTableKoguchi">
					<label for="pTableKoguchi">集計表(個口)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableOkuri" id="pTableOkuri">
					<label for="pTableOkuri">集計表(便別送状枚数)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKaisya" id="pTableKaisya">
					<label for="pTableKaisya">集計表(運送会社別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableColOkurisaki" id="pTableColOkurisaki">
					<label for="pTableColOkurisaki">集計表(集約送り先別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableOkurisaki" id="pTableOkurisaki">
					<label for="pTableOkurisaki">集計表(送り先別)</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnMonth" id="pTablePnMonth">
					<label for="pTablePnMonth">集計表(品番/月別)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKenHin" id="pTableKenHin">
					<label for="pTableKenHin">集計表(件管/品管No)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
					<label for="pList">一覧表</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListTei" id="pListTei">
					<label for="pListTei">一覧表(邸別照合)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
					<label for="pListAll">一覧表(全項目)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="dData" id="dData" onclick="DeleteClick();" disabled>
					<label for="dData">削除</label>
			</td>
		</tr>
		<tr>
			<td align="right">
				<b>対象：</b>
			</td>
			<td colspan="10">
				<INPUT TYPE="radio" NAME="tbl" VALUE="Y_SYUKA_H" id="Y_SYUKA_H">
					<label for="Y_SYUKA_H"><b>出荷予定(Y_SYUKA_H)</b></label>
				<INPUT TYPE="radio" NAME="tbl" VALUE="DEL_SYUKA_H" id="DEL_SYUKA_H">
					<label for="DEL_SYUKA_H"><b>削除済(DEL_SYUKA_H)</b></label>
			</td>
		</tr>
	</table>
	<tr>
		<td>
		<INPUT TYPE="submit" value="検索" id=submit1 name=submit1>
		<INPUT TYPE="reset" value="リセット" id=reset1 name=reset1 onClick="location.href='y_syuka_h.asp?tbl=<%=tblStr%>';">
				最大件数：<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=maxStr%>" size="8">
		<%=versionStr%>
		</td>
	</tr>
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
	<!--table id="resultTbl" class="tablesorter"-->
	<table id="resultTbl" class="tablesorter">
	<%
		Set db = Server.CreateObject("ADODB.Connection")
		db.open GetDbName()
		sqlStr = ""
		whereStr = ""
		andStr = " where"

		whereStr = makeWhere(whereStr,"s.SYUKA_YMD"						,dtStr				,dtToStr	)
		whereStr = makeWhere(whereStr,"s.ID_NO"							,KEY_ID_NOStr		,""			)
		whereStr = makeWhere(whereStr,"s.OKURI_NO"						,OKURI_NOStr		,""			)
		whereStr = makeWhere(whereStr,"s.DEN_NO"						,DEN_NOStr			,""			)
		whereStr = makeWhere(whereStr,"s.MUKE_CODE"						,MUKE_CODEStr		,""			)
		whereStr = makeWhere(whereStr,"s.HIN_NO"						,pnStr				,""			)
		whereStr = makeWhere(whereStr,"convert(s.SAI_SU,SQL_DECIMAL)"	,SAI_SUStr			,""			)
		whereStr = makeWhere(whereStr,"s.CANCEL_F"						,CANCEL_FStr		,""			)
		whereStr = makeWhere(whereStr,"s.OKURISAKI_CD"					,GetRequest("OKURISAKI_CD","")	,""			)
		whereStr = makeWhere(whereStr,"s.COL_OKURISAKI_CD"				,COL_OKURISAKI_CDStr,""			)
		whereStr = makeWhere(whereStr,"s.INS_BIN"						,strINS_BIN			,""			)
		whereStr = makeWhere(whereStr,"s.MUKE_NAME"						,GetRequest("MUKE_NAME","")		,""			)
		whereStr = makeWhere(whereStr,"s.JYUSHO"						,GetRequest("JYUSHO","")		,""			)
		whereStr = makeWhere(whereStr,"s.UNSOU_KAISHA"					,GetRequest("UNSOU_KAISHA",""),"")

		sqlStr = "select distinct "
		if lngMax > 0 then
			sqlStr = sqlStr & " top " & lngMax
		end if
		select case ptypeStr
		case "pTable"	' 集計表
			sqlStr = sqlStr & " SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",CANCEL_F as ""キャンセル"""
			sqlStr = sqlStr & ",left(s.ID_NO,7) ""ID"""
			sqlStr = sqlStr & ",s.MUKE_CODE + ' ' + Mts.MUKE_NAME as ""向け先"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",count(distinct left(s.ID_NO,7)) as ""送状"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""<br>検品済"""
'			sqlStr = sqlStr & ",(convert(floor(sum(if(KENPIN_NOW <> '',1,0)) / count(*) * 100),SQL_CHAR) + '%') as ""<br>％"""
			sqlStr = sqlStr & ",' ' as "" """
			sqlStr = sqlStr & ",sum(if(INS_BIN = '01',1,0)) as ""1便<br>件数"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '01' and KENPIN_NOW <> '',1,0)) as ""1便<br>検品済"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '02',1,0)) as ""2便"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '02' and KENPIN_NOW <> '',1,0)) as ""<br>検品済"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '03',1,0)) as ""3便"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '03' and KENPIN_NOW <> '',1,0)) as ""<br>検品済"""
			sqlStr = sqlStr & ",sum(if(INS_BIN not in ('01','02','03'),1,0)) as ""その他"""
			sqlStr = sqlStr & ",sum(if(INS_BIN not in ('01','02','03') and KENPIN_NOW <> '',1,0)) as ""<br>検品済"""
			sqlStr = sqlStr & ",' ' as "" """
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & " left outer join Mts on s.MUKE_CODE = Mts.MUKE_CODE and Mts.SS_CODE = ''"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""キャンセル"",""ID"",""向け先"""
			sqlStr = sqlStr & " order by ""出荷日"",""キャンセル"",""向け先"",""ID"""
		case "pTableOkuri"	' 集計表
			sqlStr = sqlStr & " SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",INS_BIN as ""便"""
			sqlStr = sqlStr & ",CANCEL_F as ""キャンセル"""
			sqlStr = sqlStr & ",count(distinct left(s.ID_NO,7)) as ""送状"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""検品済"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""便"",""キャンセル"""
			sqlStr = sqlStr & " order by ""出荷日"",""便"",""キャンセル"""
		case "pTableKaisya"	' 集計表(運送会社別)
			sqlStr = sqlStr & " SYUKA_YMD ""出荷日"""
			sqlStr = sqlStr & ",UNSOU_KAISHA ""運送会社"""
			sqlStr = sqlStr & ",count(distinct LEFT(ID_NO,7)) ""送状"""
			sqlStr = sqlStr & ",count(*) ""件数"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) ""検品済"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) ""数量"""
'			sqlStr = sqlStr & ",sum(convert(KUTI_SU,SQL_DECIMAL)) ""口数"""
'			sqlStr = sqlStr & ",sum(convert(SAI_SU,SQL_DECIMAL)) ""才数"""
'			sqlStr = sqlStr & ",sum(convert(JURYO,SQL_DECIMAL)) ""重量"""
			sqlStr = sqlStr & " From " & tblStr & " s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""運送会社"""
			sqlStr = sqlStr & " order by ""出荷日"",""運送会社"""
		case "pTableColOkurisaki"	' 集計表(集約送り先別)
			sqlStr = sqlStr & " SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",COL_OKURISAKI_CD as ""集約送り先コード"""
			sqlStr = sqlStr & ",if(COL_OKURISAKI_CD <> '',' ' + OKURISAKI,'') as ""送り先名"""
			sqlStr = sqlStr & ",convert(INS_BIN,SQL_DECIMAL) as ""便"""
			sqlStr = sqlStr & ",count(distinct if(OKURI_NO = '',null(),OKURI_NO)) as ""送状"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""検品済"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & ",max(convert(KUTI_SU,SQL_DECIMAL)) as ""口数"""
			sqlStr = sqlStr & ",max(convert(SAI_SU,SQL_DECIMAL)) as ""才数"""
			sqlStr = sqlStr & ",max(convert(JURYO,SQL_DECIMAL)) as ""重量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""集約送り先コード"",""送り先名"",""便"""
			sqlStr = sqlStr & " order by ""出荷日"",""集約送り先コード"",""便"""
		case "pTableOkurisaki"	' 集計表(送り先別)
			sqlStr = sqlStr & " SYUKA_YMD ""出荷日"""
			sqlStr = sqlStr & ",COL_OKURISAKI_CD ""集約送り先コード"""
 			sqlStr = sqlStr & ",OKURISAKI_CD ""送り先コード"""
 			sqlStr = sqlStr & ",OKURISAKI ""送り先"""
 			sqlStr = sqlStr & ",JYUSHO ""送り先住所"""
 			sqlStr = sqlStr & ",'' ""配達不可"""
			sqlStr = sqlStr & ",UNSOU_KAISHA ""運送会社"""
			sqlStr = sqlStr & ",convert(INS_BIN,SQL_DECIMAL) as ""便"""
			sqlStr = sqlStr & ",count(distinct if(OKURI_NO = '',null(),OKURI_NO)) as ""送状"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""検品済"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & ",max(convert(KUTI_SU,SQL_DECIMAL)) as ""口数"""
			sqlStr = sqlStr & ",max(convert(SAI_SU,SQL_DECIMAL)) as ""才数"""
			sqlStr = sqlStr & ",max(convert(JURYO,SQL_DECIMAL)) as ""重量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""集約送り先コード"",""送り先コード"",""送り先"",""送り先住所"",""配達不可"",""運送会社"",""便"""
			sqlStr = sqlStr & " order by ""出荷日"",""集約送り先コード"",""送り先コード"",""便"""
		case "pTableKoguchi"	' 集計表
			sqlStr = sqlStr & " s.SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",UNSOU_KAISHA as ""運送会社"""
			sqlStr = sqlStr & ",OKURI_NO as ""問合せNo"""
			sqlStr = sqlStr & ",left(ID_NO,7) as ""ID-No"""
			sqlStr = sqlStr & ",convert(KUTI_SU,SQL_DECIMAL) as ""口数"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & ",convert(SAI_SU,SQL_DECIMAL) as ""才数"""
			sqlStr = sqlStr & ",convert(JURYO,SQL_DECIMAL) as ""重量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""運送会社"",""問合せNo"",""ID-No"",""口数"",""才数"",""重量"""
			sqlStr = sqlStr & " order by ""出荷日"",""運送会社"",""問合せNo"""
		case "pTablePnMonth"	' 集計表(品番／月別 出荷数量)
			sqlStr = sqlStr & " HIN_NO as ""品番"""
			sqlStr = sqlStr & ",left(s.SYUKA_YMD,6) as ""出荷年月"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""品番"",""出荷年月"""
			sqlStr = sqlStr & " order by ""品番"",""出荷年月"""
		case "pTableKenHin"	' 集計表(件管/品管No)
			sqlStr = sqlStr & " SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",INS_BIN as ""便"""
			sqlStr = sqlStr & ",SEK_KEN_NO as ""件管No."""
			sqlStr = sqlStr & ",SEK_HIn_NO as ""品管No."""
			sqlStr = sqlStr & ",CANCEL_F as ""キャンセル"""
			sqlStr = sqlStr & ",count(distinct left(s.ID_NO,7)) as ""送状"""
			sqlStr = sqlStr & ",count(*) as ""件数"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""検品済"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""数量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""出荷日"",""便"",""件管No."",""品管No."",""キャンセル"""
			sqlStr = sqlStr & " order by ""出荷日"",""便"",""件管No."",""品管No."",""キャンセル"""
		case "pList"	' 一覧表
			sqlStr = sqlStr & " SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",convert(INS_BIN,SQL_DECIMAL) as ""便"""
			sqlStr = sqlStr & ",CANCEL_F as ""キャンセル"""
			sqlStr = sqlStr & ",URIDEN as ""売伝"""
 			sqlStr = sqlStr & ",MUKE_CODE + ' ' + MUKE_NAME  as ""得意先"""
 			sqlStr = sqlStr & ",COL_OKURISAKI_CD as ""集約送り先コード"""
 			sqlStr = sqlStr & ",OKURISAKI_CD ""送り先コード"""
 			sqlStr = sqlStr & ",OKURISAKI as ""送り先"""
 			sqlStr = sqlStr & ",JYUSHO as ""送り先住所"""
			sqlStr = sqlStr & ",ID_NO"
			sqlStr = sqlStr & ",DEN_NO as ""伝票No"""
			sqlStr = sqlStr & ",ODER_NO as ""オーダーNo"""
			sqlStr = sqlStr & ",HIN_NO as ""品番"""
			sqlStr = sqlStr & ",convert(SURYO,SQL_DECIMAL) as ""数量"""
			sqlStr = sqlStr & ",BIKOU as ""備考"""
			sqlStr = sqlStr & ",TEL_No as ""Tel"""
			sqlStr = sqlStr & ",YUBIN_No as ""郵便番号"""
			sqlStr = sqlStr & ",left(KENPIN_NOW,8) + '-' + left(right(KENPIN_NOW,6),4) as ""検品日時"""
			sqlStr = sqlStr & ",KENPIN_TANTO_CODE as ""検品担当者"""
			sqlStr = sqlStr & ",UNSOU_KAISHA as ""運送会社"""
			sqlStr = sqlStr & ",OKURI_NO as ""問合せNo"""
			sqlStr = sqlStr & ",convert(SEQ_NO,SQL_DECIMAL) as ""小口No"""
			sqlStr = sqlStr & ",convert(KUTI_SU,SQL_DECIMAL) as ""口数"""
			sqlStr = sqlStr & ",convert(SAI_SU,SQL_DECIMAL) as ""才数"""
			sqlStr = sqlStr & ",convert(JURYO,SQL_DECIMAL) as ""重量"""
			sqlStr = sqlStr & ",INS_TANTO ""登録ID"""
			sqlStr = sqlStr & ",left(INS_DATETIME,8) + '-' +  right(INS_DATETIME,6) ""登録日時"""
			sqlStr = sqlStr & ",UPD_TANTO ""更新ID"""
			sqlStr = sqlStr & ",left(UPD_DATETIME,8) + '-' +  right(UPD_DATETIME,6) ""更新日時"""
'			sqlStr = sqlStr & ",UPD_DATETIME"
'			sqlStr = sqlStr & ",left(UPD_DATETIME,8)"
'			sqlStr = sqlStr & ",left(ltrim(UPD_DATETIME),8)"
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""出荷日"",""便"",""運送会社"",ID_NO,""問合せNo"",""キャンセル"",""小口No"",""送り先"""
		case "pListTei"	' 一覧表(邸別照合)
			sqlStr = sqlStr & " s.SYUKA_YMD as ""出荷日"""
			sqlStr = sqlStr & ",convert(s.INS_BIN,SQL_DECIMAL) as ""便"""
			sqlStr = sqlStr & ",s.CANCEL_F as ""キャンセル"""
			sqlStr = sqlStr & ",s.URIDEN as ""売伝"""
 			sqlStr = sqlStr & ",s.COL_OKURISAKI_CD as ""集約送り先コード"""
 			sqlStr = sqlStr & ",s.OKURISAKI as ""送り先"""
 			sqlStr = sqlStr & ",s.MUKE_CODE + ' ' + MUKE_NAME  as ""得意先"""
			sqlStr = sqlStr & ",s.ID_NO"
			sqlStr = sqlStr & ",s.DEN_NO as ""伝票No"""
			sqlStr = sqlStr & ",s.ODER_NO as ""オーダーNo"""
			sqlStr = sqlStr & ",s.SEK_KEN_NO as ""件管No."""
			sqlStr = sqlStr & ",s.SEK_HIn_NO as ""品管No."""
			sqlStr = sqlStr & ",t.CHU_CD ""(邸別)注文№<br>■指図№(上)"""
			sqlStr = sqlStr & ",t.THINB_CD ""得意先品番<br>■品番(上)"""
			sqlStr = sqlStr & ",t.HINB_CD ""品番<br>■品番(下)"""
			sqlStr = sqlStr & ",s.HIN_NO as ""品番"""
			sqlStr = sqlStr & ",convert(ifnull(t.JUC_SUU,0),sql_numeric) ""(邸別)受注数量"""
			sqlStr = sqlStr & ",convert(s.SURYO,SQL_DECIMAL) as ""数量"""
			sqlStr = sqlStr & ",t.SND_YMD + '-' + t.SND_HMS ""(邸別)データ作成日時"""
			sqlStr = sqlStr & ",t.SYU_JUN ""(邸別)出荷順番<br>■指図№(下・左)"""
			sqlStr = sqlStr & ",t.TEI_NM ""(邸別)邸名<br>■指図№(下・右)"""
			sqlStr = sqlStr & ",t.TEI_LABELID ""邸別ラベルID"""
			sqlStr = sqlStr & ",t.KONPO_ID ""集合梱包ID"""
			sqlStr = sqlStr & ",s.BIKOU as ""備考"""
			sqlStr = sqlStr & ",s.TEL_No as ""Tel"""
			sqlStr = sqlStr & ",s.YUBIN_No as ""郵便番号"""
			sqlStr = sqlStr & ",left(s.KENPIN_NOW,8) + '-' + left(right(s.KENPIN_NOW,6),4) as ""検品日時"""
			sqlStr = sqlStr & ",s.KENPIN_TANTO_CODE as ""検品担当者"""
			sqlStr = sqlStr & ",s.UNSOU_KAISHA as ""運送会社"""
			sqlStr = sqlStr & ",s.OKURI_NO as ""問合せNo"""
			sqlStr = sqlStr & ",convert(s.SEQ_NO,SQL_DECIMAL) as ""小口No"""
			sqlStr = sqlStr & ",convert(s.KUTI_SU,SQL_DECIMAL) as ""口数"""
			sqlStr = sqlStr & ",convert(s.SAI_SU,SQL_DECIMAL) as ""才数"""
			sqlStr = sqlStr & ",convert(s.JURYO,SQL_DECIMAL) as ""重量"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
'			sqlStr = sqlStr & " left outer join y_syuka_tei t on (t.TOK_CD = s.MUKE_CODE and t.CHU_CD = s.ODER_NO and t.HINB_CD = s.HIN_NO)"
			sqlStr = sqlStr & " left outer join y_syuka_tei t on (t.KEN_NO = s.SEK_KEN_NO and t.HIN_NO = s.SEK_HIN_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""出荷日"",""便"",""運送会社"",ID_NO,""問合せNo"",""キャンセル"",""小口No"",""送り先"""
		case "pListAll"	' 一覧表(全項目)
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
'			sqlStr = sqlStr & " order by SYUKA_YMD"
		case "dData"	' 削除
			sqlStr = "delete from " & tblStr
			sqlStr = sqlStr & whereStr
			set rsList = db.Execute(sqlStr)

			sqlStr = "select @@rowcount"
			sqlStr = sqlStr & " From " & tblStr
		end select
'		db.CommandTimeout=900
		set rsList = db.Execute(sqlStr)
	%>
		<caption style="text-align:left;"><%=now%> 現在</caption>
		
		<%
			strTh = "<TR BGCOLOR=""#ccffee"">"
			strTh = strTh & "<TH>No.</TH>"

			dim totalArray(255)
			For i=0 To rsList.Fields.Count-1
				strTh = strTh & "<TH title='" & rsList.Fields(i).name & "'>" & rsList.Fields(i).name & "</TH>"
			Next
			strTh = strTh & "</TR>"
		%>
		<thead>
		<%=strTh%>
		</thead>
		<tbody>
		<%
			cnt = 0
			Do While Not rsList.EOF
				cnt = cnt + 1
		%>
				<TR VALIGN='TOP'>
				<TD nowrap id="Integer"><%=cnt%></TD>
		<%
				For i=0 To rsList.Fields.Count-1
					' 値
					fValue = rtrim(rsList.Fields(i))
					if rsList.Fields(i).Name = "問合せNo" then
						tdTag = "<TD nowrap id=""Charactor"">"
						if rtrim(rsList.Fields("運送会社")) = "久留米運輸" then
							fValue = "<a href=""http://www4.kisc.co.jp/kurume-trans/kamotsu.asp?w_no=" & fValue & """>" & fValue & "</a>"
						end if
					elseif rsList.Fields(i).Name = "才数" or rsList.Fields(i).Name = "重量" then
						if i = 0 then
							totalArray(i) = clng(fValue)
						else
							totalArray(i) = totalArray(i) + cdbl(fValue)
						end if
						if fValue = 0 then
							fValue = ""
						else
							fValue = formatnumber(fValue,2,,,-1)
						end if
						tdTag = "<TD nowrap id=""Integer"">"
					elseif rsList.Fields(i).Name = "送り先住所" then
						tdTag = "<TD id=""Charactor"">"
					elseif rsList.Fields(i).Name = "配達不可" then
						tdTag = "<TD id=""Charactor"">"
						fValue = GetHaitatsu(rsList.Fields("送り先住所") & " " & rsList.Fields("送り先"))
						if inStr(fValue,"世田谷") > 0 then
							if cLng(rsList.Fields("才数")) < 10 then
								fValue = ""
							end if
						end if
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
							else
								totalArray(i) = totalArray(i) + clng(fValue)
							end if
							if fValue = 0 then
								fValue = ""
							else
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
				Next
		    	Response.Write "</TR>"
				rsList.Movenext
			Loop
		%>
		</tbody>
		<!-- 合計 -->
		<TR VALIGN='TOP'>
			<TD nowrap align="center">計</TD>
			<%For i=0 To rsList.Fields.Count-1
				select case rsList.Fields(i).type
				Case 2 , 3 , 5 ,131	' 数値(Integer)
			%>
					<TD nowrap id="Integer"><%=formatnumber(totalArray(i),0,,,-1)%></TD>
			<%	Case else		' その他	%>
					<TD></TD>
			<%	end select	%>
			<%Next%>
    	</TR>
		<!--%=strTh%-->
	</TABLE></div>
	<hr>
	<div id="sql">
		<%=sqlStr%>
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
//		cpTblBtn.disabled = false;
//		cpTblBtn.value = "結果をコピー";
//		autoBtn.disabled = false;
		if ('<%=ptypeStr%>' == 'pTableOkurisaki') {
		    $("#resultTbl").tablesorter({ 
		        sortList: [[6,1],[0,0]] 
		    }); 
		}
	//-->
	</SCRIPT>
<% end if %>
</BODY>
</HTML>
<SCRIPT LANGUAGE=javascript>
/*
$(document).ready(function() 
    {
		if ('<%=ptypeStr%>' == 'pTableOkurisaki') {
		    $("#resultTbl").tablesorter({ 
		        sortList: [[6,1],[0,0]] 
		    }); 
		} else {
		    $("#resultTbl").tablesorter({ 
		    }); 
		}
    }
); 
*/
</SCRIPT>
