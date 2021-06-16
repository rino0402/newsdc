<%@Language="VBScript"%><%Option Explicit%>
<% Response.Buffer = false %>
<% Response.Expires = -1 %>
<%
Function myNumber(myData,para1)
'	myNumber = FormatNumber(myData,0,,,para1)
	myNumber = myData
End Function

Function roundUp(v)
	dim	retV
	retV = v + 0.5
	retV = round(retV)
	roundUp = retV
End Function

	dim	versionStr
'	dim	dbName

	versionStr = "2010.03.04 バージョン表示対応"
	versionStr = "2010.03.04 最大行数対応"
	versionStr = "2010.03.04 一覧表(外装) 対応"
	versionStr = "2010.08.18 ポップアップメニュー対応/集計表(更新日)対応"
	versionStr = "2019.10.31 "

'	dbName	= "newsdc"
%>

<%
	dim objFS
	dim objF
	
	dim SHIMUKE_CODEStr
	dim DATA_KBNStr
	dim CLASS_CODEStr
	dim F_CLASS_CODEStr
	dim N_CLASS_CODEStr
	dim pnStr
	dim BIKOUStr
	dim	KO_HIN_GAIStr

	dim submitStr
	dim ptypeStr
	dim cmpStr
	dim db
	dim rsList
	dim rsRow
	dim	sqlStr
	dim	whereStr
	dim andStr
	dim cnt,i
	dim fValue,fname,tdTag
	dim fType
	dim centerStr
	dim	maxStr
	dim	lngMax

	BIKOUStr = ucase(Request.QueryString("BIKOU"))
	SHIMUKE_CODEStr = ucase(Request.QueryString("SHIMUKE_CODE"))
	DATA_KBNStr = ucase(Request.QueryString("DATA_KBN"))
	pnStr = ucase(Request.QueryString("pn"))
	CLASS_CODEStr = ucase(Request.QueryString("CLASS_CODE"))
	F_CLASS_CODEStr = ucase(Request.QueryString("F_CLASS_CODE"))
	N_CLASS_CODEStr = ucase(Request.QueryString("N_CLASS_CODE"))
	KO_HIN_GAIStr	= ucase(Request.QueryString("KO_HIN_GAI"))
	submitStr = Request.QueryString("submit1")

	maxStr			= ucase(Request.QueryString("max"))
	if len(maxStr) = 0 then
		maxStr	= 100
	end if
	lngMax = clng(maxStr)

	ptypeStr = Request.QueryString("ptype")
	if len(ptypeStr) = 0 then
'		ptypeStr = "pListClass"
		ptypeStr = "pTableKbn"
'		DATA_KBNStr = "0"
	end if
%>
<!--#include file="info.txt" -->
<!--#include file="makeWhere.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="result.css" TITLE="CSS">
<TITLE><%=centerStr%> 構成マスター</TITLE>
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

	function DoCopyxx(arg){
		var	clpText	= "";
		var loop = resultTbl.rows.length;
		for(var i = 0 ; i < loop ; i++)	{
			window.status = i + "/" + (loop - 1);
			var	colCount = resultTbl.rows[i].cells.length;
			for (var j = 0; j < colCount; j++) {
				var strText = "" + resultTbl.rows[i].cells[j].innerText;
				if ( i == 0 ) {
					strText = strText.replace("\x0d\x0a","");
//					strText = strText.replace("\n","");
//					strText = strText.replace("\f","");
				}
				if ( j > 0 ) {
					clpText = clpText + "\t";
				}
				clpText = clpText + strText;
			}
			clpText = clpText + "\x0d\x0a";
		}
		window.clipboardData.setData('text',clpText);
		DoneMes();
	}

	function ptypeChange(typ) {
//		sqlForm.ptype[typ].checked = "true";
		for(var	i = 0;i < document.sqlForm.elements.length;i++) {
//			window.alert(document.sqlForm.elements[i].id);
			if ( document.sqlForm.elements[i].id == typ ) {
				document.sqlForm.elements[i].checked = "true"
				break;
			}
		}
	}
	function setValue(f,s) {
		f.value = s;
	}
	function pListClassClick() {
		sqlForm.JGYOBU.value = "S";
	}
	// (すべての変数に格納する値は0オリジンとする) 
	function myFormatNumber(x) { // 引数の例としては 95839285734.3245
	    var s = "" + x; // 確実に文字列型に変換する。例では "95839285734.3245"
	    var p = s.indexOf("."); // 小数点の位置を0オリジンで求める。例では 11
	    if (p < 0) { // 小数点が見つからなかった時
	        p = s.length; // 仮想的な小数点の位置とする
	    }
	    var r = s.substring(p, s.length); // 小数点の桁と小数点より右側の文字列。例では ".3245"
	    for (var i = 0; i < p; i++) { // (10 ^ i) の位について
	        var c = s.substring(p - 1 - i, p - 1 - i + 1); // (10 ^ i) の位のひとつの桁の数字。例では "4", "3", "7", "5", "8", "2", "9", "3", "8", "5", "9" の順になる。
	        if (c < "0" || c > "9") { // 数字以外のもの(符合など)が見つかった
	            r = s.substring(0, p - i) + r; // 残りを全部付加する
	            break;
	        }
	        if (i > 0 && i % 3 == 0) { // 3 桁ごと、ただし初回は除く
	            r = "," + r; // カンマを付加する
	        }
	        r = c + r; // 数字を一桁追加する。
	    }
	    return r; // 例では "95,839,285,734.3245"
	}
//--></SCRIPT>
</HEAD>
<BODY>
<!-- jdMenu body用 include 開始 -->
<!--#include file="jdmenu-sdc.asp" -->
<!-- jdMenu body用 include 終了 -->
  <FORM name="sqlForm">
  <div id="sqlDiv">
	<table id="sqlTbl">
		<caption style="text-align:left;">構成マスター検索</caption>
		<tr>
			<th>仕向先</th>
			<th>データ区分</th>
			<th>親品番</th>
			<th>子品番</th>
			<th>商品化クラス</th>
			<th>付加クラス</th>
			<th>内職クラス</th>
			<th>備考</th>
			<th>担当者</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="SHIMUKE_CODE" id="SHIMUKE_CODE" VALUE="<%=SHIMUKE_CODEStr%>"
						size="2" maxlength="2" style="text-align:center;">
				<div align="left">
				<%
					Set db = Server.CreateObject("ADODB.Connection")
					db.open GetRequest("dbName","newsdc")
					sqlStr = "select C_Code,C_NAME from p_code where DATA_KBN ='04' order by C_Code"
					set rsList = db.Execute(sqlStr)
					Do While Not rsList.EOF
				%>
						<%=rsList.Fields("C_Code")%>:<%=rsList.Fields("C_NAME")%><br>
				<%
						rsList.Movenext
					loop
					set db = nothing
				%>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="DATA_KBN" id="DATA_KBN" VALUE="<%=DATA_KBNStr%>"
						size="2" maxlength="2" style="text-align:center;">
				<div align="left">
				0:クラス<br>
				1:個装資材<br>
				2:外装資材<br>
				3:構成部品
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="pn" id="pn" VALUE="<%=pnStr%>" size="22" maxlength="20"><br>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KO_HIN_GAI" id="KO_HIN_GAI" VALUE="<%=KO_HIN_GAIStr%>" size="22" maxlength="20"><br>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="CLASS_CODE" id="CLASS_CODEStr" VALUE="<%=CLASS_CODEStr%>" size="12" maxlength="10">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="F_CLASS_CODE" id="F_CLASS_CODEStr" VALUE="<%=F_CLASS_CODEStr%>" size="7" maxlength="5">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="N_CLASS_CODE" id="N_CLASS_CODEStr" VALUE="<%=N_CLASS_CODEStr%>" size="7" maxlength="5">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="BIKOU" id="BIKOUStr" VALUE="<%=BIKOUStr%>" size="20">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="UPD_TANTO" id="UPD_TANTO" VALUE="<%=GetRequest("UPD_TANTO","")%>" size="14" style="text-align:left;">
			</td>
		</tr>
		<tr>
			<td colspan="9">
				<table border="0" cellspacing="0" bordercolor="White">
					<tr>
					<td valign="top">
					<b>出力形式：</b>
					</td>
					<td>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKbn" id="pTableKbn">
						<label for="pTableKbn">集計表(仕向先/区分別 登録件数)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableUpd" id="pTableUpd">
						<label for="pTableUpd">集計表(更新日)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKosei" id="pTableKosei">
						<label for="pTableKosei">集計表(同梱)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableChild" id="pTableChild">
						<label for="pTableChild">集計表(構成子)</label>
					<br>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
						<label for="pList">一覧表</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListClass" id="pListClass" onclick="sqlForm.DATA_KBN.value='0';">
						<label for="pListClass">一覧表(クラス)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListCompo" id="pListCompo">
						<label for="pListCompo">一覧表(構成)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListGaiso" id="pListGaiso" onclick="sqlForm.DATA_KBN.value='2';">
						<label for="pListGaiso">一覧表(外装)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
						<label for="pListAll">一覧表(全項目)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListKAll" id="pListKAll">
						<label for="pListKAll">一覧表(全項目:子)</label>
					</td>
					</tr>
					<tr>
						<td align="right">
							<b>対象：</b>
						</td>
						<td>
							<span>
							<INPUT class="input" TYPE="text" NAME="dbName" id="dbName" VALUE="<%=GetRequest("dbName","newsdc")%>" size="10">
							</span>
						</td>
					</tr>
				</table>                                                                            
			</td>                                                                            
		</tr>
		<tr bordercolor=White>                                                                            
			<td colspan="9">
			<INPUT TYPE="submit" value="検索" id=submit1 name=submit1>
			<INPUT TYPE="reset" value="リセット" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
			最大件数：<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=maxStr%>" size="8" maxlength="6">
			<%=versionStr%>
			</td>
		</tr>                                                                            
	</table>                                                                            
	</div>                                                                            
  </FORM>                                                                            
<%	Response.Write "<SCRIPT LANGUAGE='JavaScript'>" & "ptypeChange('" & ptypeStr & "');</SCRIPT>"	%>                                                                            
<%	if len(submitStr) > 0 then %>                                                                            
	<SCRIPT LANGUAGE="javascript"><!--
		sqlForm.disabled = true;
	//--></SCRIPT>
	<div>                                                                            
	<INPUT TYPE='button' onClick="DoCopy('resultDiv')" value='検索中...' id='cpTblBtn' disabled>                                                                            
	</div>                                                                            
                                                                            
	<%
		dim strTName
'		Response.Flush()                                                                            
		Set db = Server.CreateObject("ADODB.Connection")                                                                            
		db.open GetRequest("dbName","newsdc")                                                                            
		sqlStr = ""                                                                            
		whereStr = ""                                                                            
		andStr = " where "
		if ptypeStr = "pTableKosei" then
			strTName = " "
		else
			strTName = " P_COMPO."
		end if

		if len(SHIMUKE_CODEStr) > 0 then
			if left(SHIMUKE_CODEStr,1) = "-" then
				whereStr = whereStr & andStr & "p.SHIMUKE_CODE <> '" & SHIMUKE_CODEStr & "'"
			else
				whereStr = whereStr & andStr & "p.SHIMUKE_CODE = '" & SHIMUKE_CODEStr & "'"
			end if
			andStr = " and "
		end if

		if len(DATA_KBNStr) > 0 then
			if left(DATA_KBNStr,1) = "-" then
				whereStr = whereStr & andStr & "p.DATA_KBN <> '" & DATA_KBNStr & "'"
			else
				whereStr = whereStr & andStr & "p.DATA_KBN = '" & DATA_KBNStr & "'"
			end if
			andStr = " and "
		end if

		if len(pnStr) > 0 then
			if instr(1,pnStr,"%") > 0 then
				cmpStr = "like"
			else
				cmpStr = "="
			end if
			whereStr = whereStr & andStr & "p.hin_gai " & cmpStr & " '" & pnStr & "'"
			andStr = " and "                                                                            
		end if

		if len(KO_HIN_GAIStr) > 0 then
			if instr(1,KO_HIN_GAIStr,"%") > 0 then
				cmpStr = "like"
			else
				cmpStr = "="
			end if
			whereStr = whereStr & andStr & "p.KO_HIN_GAI " & cmpStr & " '" & KO_HIN_GAIStr & "'"
			andStr = " and "                                                                            
		end if

		if len(CLASS_CODEStr) > 0 then
			if instr(1,CLASS_CODEStr,"%") > 0 then
				cmpStr = "like"
			else
				cmpStr = "="
			end if
			whereStr = whereStr & andStr & "p.CLASS_CODE " & cmpStr & " '" & CLASS_CODEStr & "'"
			andStr = " and "
		end if

		if len(F_CLASS_CODEStr) > 0 then
			if instr(1,F_CLASS_CODEStr,"%") > 0 then
				cmpStr = "like"
			else
				cmpStr = "="
			end if
			whereStr = whereStr & andStr & "p.F_CLASS_CODE " & cmpStr & " '" & F_CLASS_CODEStr & "'"
			andStr = " and "
		end if
		
		if len(N_CLASS_CODEStr) > 0 then
			if instr(1,N_CLASS_CODEStr,"%") > 0 then
				cmpStr = "like "
			else
				cmpStr = "="
			end if
			whereStr = whereStr & andStr & "p.N_CLASS_CODE " & cmpStr & " '" & N_CLASS_CODEStr & "'"
			andStr = " and "
		end if

		if len(BIKOUStr) > 0 then
			if instr(1,BIKOUStr,"%") > 0 then
				cmpStr = "like "
			else
				cmpStr = "="
			end if
			whereStr = whereStr & andStr & "p.BIKOU " & cmpStr & " '" & BIKOUStr & "'"
			andStr = " and "
		end if
		whereStr = makeWhere(whereStr,"p.UPD_TANTO",GetRequest("UPD_TANTO",""),"")

		sqlStr = "select "
		if lngMax > 0 then
			sqlStr = sqlStr & " top " & lngMax
		end if

		select case ptypeStr                                       
		case "pTableKbn"	' 集計表(区分別)                                                                            
			db.CommandTimeout = 360
			sqlStr = sqlStr & " p.SHIMUKE_CODE + rtrim(' ' + ifnull(c.C_NAME,'')) ""仕向先"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '0',1,0)) ""クラス<br>件数"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '1',1,0)) ""資材<br>件数"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '2',1,0)) ""外装<br>件数"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '3',1,0)) ""構成部品<br>件数"""
			sqlStr = sqlStr & " From P_COMPO p"
			sqlStr = sqlStr & " left outer join p_code c on (p.SHIMUKE_CODE=c.C_Code and c.DATA_KBN='04')"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""仕向先"""
			sqlStr = sqlStr & " order by ""仕向先"""
		case "pTableUpd"	' 集計表(更新日)
			sqlStr = sqlStr & " left(UPD_DATETIME,8) ""更新日"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '0',1,0)) as ""クラス<br>件数"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '1',1,0)) as ""資材<br>件数"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '2',1,0)) as ""外装<br>件数"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '3',1,0)) as ""構成部品<br>件数"""
			sqlStr = sqlStr & " From P_COMPO p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""更新日"""
			sqlStr = sqlStr & " order by ""更新日"" desc"
		case "pTableKosei"	' 集計表(同梱)                                                                      			db.CommandTimeout = 360
			sqlStr = sqlStr & " HIN_GAI"
			sqlStr = sqlStr & ",count(*)"
			sqlStr = sqlStr & " From P_COMPO_K p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by HIN_GAI"
			sqlStr = sqlStr & " order by HIN_GAI desc"
		case "pTableChild"	' 集計表(構成子)
			db.CommandTimeout = 360
			sqlStr = sqlStr & " p.KO_JGYOBU ""事業部(子)"""
			sqlStr = sqlStr & ",p.KO_NAIGAI ""内外(子)"""
			sqlStr = sqlStr & ",p.KO_HIN_GAI ""品番(子)"""
			sqlStr = sqlStr & ",i.HIN_NAME ""品名(子)"""
			sqlStr = sqlStr & ",count(*) ""件数"""
			sqlStr = sqlStr & ",sum(convert(p.KO_QTY,SQL_decimal)) ""員数計"""
			sqlStr = sqlStr & " From P_COMPO_K p"
			sqlStr = sqlStr & "    left outer join ITEM i"
			sqlStr = sqlStr &				" on (p.KO_JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and p.KO_NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and p.KO_HIN_GAI	=i.HIN_GAI)"
			sqlStr = sqlStr & replace(whereStr,"P_COMPO.","k.",1,-1)
			sqlStr = sqlStr & andStr & " p.DATA_KBN <> '0'"
'			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""事業部(子)"",""内外(子)"",""品番(子)"",""品名(子)"""
			sqlStr = sqlStr & " order by ""事業部(子)"",""内外(子)"",""品番(子)"""
		case "pList"	' 一覧表
			sqlStr = sqlStr & " p.SHIMUKE_CODE ""仕向先"""
			sqlStr = sqlStr & ",p.HIN_GAI ""品番"""
			sqlStr = sqlStr & ",p.DATA_KBN"
			sqlStr = sqlStr & ",p.SEQNO"
			sqlStr = sqlStr & ",p.KO_SYUBETSU"
			sqlStr = sqlStr & ",p.KO_JGYOBU"
			sqlStr = sqlStr & ",p.KO_NAIGAI"
			sqlStr = sqlStr & ",p.KO_HIN_GAI"
			sqlStr = sqlStr & ",convert(p.KO_QTY,sql_decimal) KO_QTY"
			sqlStr = sqlStr & ",p.BIKOU ""備考"""
			sqlStr = sqlStr & ",p.UPD_TANTO"
			sqlStr = sqlStr & ",p.UPD_DATETIME"
			sqlStr = sqlStr & " From ("
			sqlStr = sqlStr & " select"
			sqlStr = sqlStr & " SHIMUKE_CODE"
			sqlStr = sqlStr & ",JGYOBU"
			sqlStr = sqlStr & ",NAIGAI"
			sqlStr = sqlStr & ",HIN_GAI"
			sqlStr = sqlStr & ",DATA_KBN"
			sqlStr = sqlStr & ",SEQNO"
			sqlStr = sqlStr & ",'' KO_SYUBETSU"
			sqlStr = sqlStr & ",'' KO_JGYOBU"
			sqlStr = sqlStr & ",'' KO_NAIGAI"
			sqlStr = sqlStr & ",'' KO_HIN_GAI"
			sqlStr = sqlStr & ",Null KO_QTY"
			sqlStr = sqlStr & ",BIKOU"
			sqlStr = sqlStr & ",UPD_TANTO"
			sqlStr = sqlStr & ",UPD_DATETIME"
			sqlStr = sqlStr & " from p_compo"
			sqlStr = sqlStr & " where DATA_KBN='0'"
			sqlStr = sqlStr & " union select"
			sqlStr = sqlStr & " SHIMUKE_CODE"
			sqlStr = sqlStr & ",JGYOBU"
			sqlStr = sqlStr & ",NAIGAI"
			sqlStr = sqlStr & ",HIN_GAI"
			sqlStr = sqlStr & ",DATA_KBN"
			sqlStr = sqlStr & ",SEQNO"
			sqlStr = sqlStr & ",KO_SYUBETSU"
			sqlStr = sqlStr & ",KO_JGYOBU"
			sqlStr = sqlStr & ",KO_NAIGAI"
			sqlStr = sqlStr & ",KO_HIN_GAI"
			sqlStr = sqlStr & ",KO_QTY"
			sqlStr = sqlStr & ",KO_BIKOU BIKOU"
			sqlStr = sqlStr & ",UPD_TANTO"
			sqlStr = sqlStr & ",UPD_DATETIME"
			sqlStr = sqlStr & " from p_compo_k"
			sqlStr = sqlStr & " where DATA_KBN<>'0'"
			sqlStr = sqlStr & " ) p"
			sqlStr = sqlStr & " " & whereStr
		case "pListClass"	' 一覧表(商品化クラス)                                                      
			Server.ScriptTimeout = 900
			db.CommandTimeout = 360
			sqlStr = sqlStr & " p.SHIMUKE_CODE ""仕向先"""
'			sqlStr = sqlStr & ",c.DATA_KBN as ""データ<br>区分"""
			sqlStr = sqlStr & ",p.HIN_GAI ""品番"""
			sqlStr = sqlStr & ",rtrim(p.CLASS_CODE) + ' ' + rtrim(CLASS_NAME) ""商品化クラス"""
			sqlStr = sqlStr & ",p.F_CLASS_CODE ""付加クラス"""
			sqlStr = sqlStr & ",p.N_CLASS_CODE ""内職クラス"""
			sqlStr = sqlStr & ",convert(TANKA,SQL_NUMERIC) ""単価"""
			sqlStr = sqlStr & ",' ' span1"
			sqlStr = sqlStr & ",convert(KOUSU,SQL_NUMERIC) ""工数"""
			sqlStr = sqlStr & ",convert(KOURYOU,SQL_NUMERIC) ""工料"""
			sqlStr = sqlStr & ",convert(ETC,SQL_NUMERIC) ""その他"""
			sqlStr = sqlStr & ",' ' span2"
			sqlStr = sqlStr & ",count(k.KO_HIN_GAI) ""資材件数"""
			sqlStr = sqlStr & ",sum(convert(k.KO_QTY,SQL_NUMERIC)) ""資材員数"""
			sqlStr = sqlStr & ",sum(convert(G_ST_URITAN,SQL_NUMERIC)*convert(k.KO_QTY,SQL_NUMERIC)) ""販売単価"""
			sqlStr = sqlStr & ",sum(convert(G_ST_SHITAN,SQL_NUMERIC)*convert(k.KO_QTY,SQL_NUMERIC)) ""仕入単価"""
			sqlStr = sqlStr & " From P_COMPO p"
			sqlStr = sqlStr &	" left outer join P_CLASS s"
			sqlStr = sqlStr &				" on (p.SHIMUKE_CODE=s.SHIMUKE_CODE"
			sqlStr = sqlStr &				" and p.CLASS_CODE=s.CLASS_CODE)"
			sqlStr = sqlStr &	" left outer join P_COMPO_K k"
			sqlStr = sqlStr &				" on (p.SHIMUKE_CODE=k.SHIMUKE_CODE"
			sqlStr = sqlStr &				" and p.JGYOBU=k.JGYOBU"
			sqlStr = sqlStr &				" and p.NAIGAI=k.NAIGAI"
			sqlStr = sqlStr &				" and p.HIN_GAI=k.HIN_GAI"
			sqlStr = sqlStr &				" and k.DATA_KBN='1')"
			sqlStr = sqlStr & "    left outer join ITEM i"
			sqlStr = sqlStr &				" on (k.KO_JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and k.KO_NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and k.KO_HIN_GAI	=i.HIN_GAI)"
'			sqlStr = sqlStr & replace(whereStr,"P_COMPO.","p.",1,-1)
			sqlStr = sqlStr & " group by ""仕向先"",""品番"",""商品化クラス"",""付加クラス"",""内職クラス"",""単価"",span1,""工数"",""工料"",""その他"",span2"
			sqlStr = sqlStr & " order by ""商品化クラス"",""品番"""
		case "pListCompo"	' 一覧表(商品化構成)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " k.SHIMUKE_CODE as ""仕向先"""
			sqlStr = sqlStr & ",k.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME ""品名"""
'			sqlStr = sqlStr & ",if(k.DATA_KBN = '0',k.CLASS_CODE,'') as ""商品化クラス"""
'			sqlStr = sqlStr & ",k.F_CLASS_CODE as ""付加クラス"""
'			sqlStr = sqlStr & ",k.N_CLASS_CODE as ""内職クラス"""
			sqlStr = sqlStr & ",k.DATA_KBN as ""種別"""
			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',k.KO_SYUBETSU,'') as ""構成種別"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',k.KO_HIN_GAI,'') as ""構成品番"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',s.HIN_NAME,'') as ""構成品名"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',convert(k.KO_QTY,SQL_NUMERIC),'') as ""構成員数"""
			sqlStr = sqlStr & " From P_COMPO_K as k"
			sqlStr = sqlStr &   " left outer join ITEM as i"
			sqlStr = sqlStr &				" on (k.JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and k.NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and k.HIN_GAI	=i.HIN_GAI)"
			sqlStr = sqlStr &   " left outer join ITEM as s"
			sqlStr = sqlStr &				" on (k.KO_JGYOBU	=s.JGYOBU"
			sqlStr = sqlStr &				" and k.KO_NAIGAI	=s.NAIGAI"
			sqlStr = sqlStr &				" and k.KO_HIN_GAI	=s.HIN_GAI)"
			sqlStr = sqlStr & replace(whereStr,"P_COMPO.","k.",1,-1)
			sqlStr = sqlStr & " order by ""仕向先"",""品番"",""種別"",""No"""
		case "pListCompoXX"	' 一覧表(商品化構成)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " c.SHIMUKE_CODE as ""仕向先"""
			sqlStr = sqlStr & ",c.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",rtrim(c.CLASS_CODE) + ' ' + rtrim(CLASS_NAME) as ""商品化クラス"""
			sqlStr = sqlStr & ",c.F_CLASS_CODE as ""付加クラス"""
			sqlStr = sqlStr & ",c.N_CLASS_CODE as ""内職クラス"""
			sqlStr = sqlStr & ",convert(TANKA,SQL_NUMERIC) as ""商品化単価"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",convert(KOURYOU,SQL_NUMERIC) as ""工料"""
			sqlStr = sqlStr & ",convert(TANKA,SQL_NUMERIC) - convert(KOURYOU,SQL_NUMERIC) as ""箱代"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",convert(KOUSU,SQL_NUMERIC) as ""工数"""
			sqlStr = sqlStr & ",convert(ETC,SQL_NUMERIC) as ""その他"""
			sqlStr = sqlStr & ",' ' as span2"
			sqlStr = sqlStr & ",k.DATA_KBN as ""種別"""
			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",KO_HIN_GAI as ""資材品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""資材品名"""
			sqlStr = sqlStr & ",convert(k.KO_QTY,SQL_NUMERIC) as ""資材員数"""
			sqlStr = sqlStr & ",convert(i.G_ST_URITAN,SQL_NUMERIC) as ""販売単価"""
			sqlStr = sqlStr & ",convert(i.G_ST_SHITAN,SQL_NUMERIC) as ""仕入単価"""
			sqlStr = sqlStr & " From P_COMPO_K as k"
			sqlStr = sqlStr &	" left outer join P_COMPO as c"
			sqlStr = sqlStr &				" on (c.SHIMUKE_CODE=k.SHIMUKE_CODE"
			sqlStr = sqlStr &				" and c.JGYOBU=k.JGYOBU"
			sqlStr = sqlStr &				" and c.NAIGAI=k.NAIGAI"
			sqlStr = sqlStr &				" and c.HIN_GAI=k.HIN_GAI"
			sqlStr = sqlStr &				" and c.DATA_KBN='0')"
			sqlStr = sqlStr &	" left outer join P_CLASS as s"
			sqlStr = sqlStr &				" on (c.SHIMUKE_CODE=s.SHIMUKE_CODE"
			sqlStr = sqlStr &				" and c.CLASS_CODE=s.CLASS_CODE)"
			sqlStr = sqlStr &   " left outer join ITEM as i"
			sqlStr = sqlStr &				" on (k.KO_JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and k.KO_NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and k.KO_HIN_GAI	=i.HIN_GAI)"
			sqlStr = sqlStr & replace(whereStr,"P_COMPO.","c.",1,-1)
			sqlStr = sqlStr & " order by ""商品化クラス"",""品番"",""種別"",""No"""
		case "pListCompoOld"	' 一覧表(商品化構成)
			sqlStr = sqlStr & " c.SHIMUKE_CODE as ""仕向先"""
			sqlStr = sqlStr & ",c.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",right(c.CLASS_CODE,length(c.CLASS_CODE)-2) as ""商品化クラス"""
			sqlStr = sqlStr & ",c.F_CLASS_CODE as ""付加クラス"""
			sqlStr = sqlStr & ",c.N_CLASS_CODE as ""内職クラス"""
'			sqlStr = sqlStr & ",k.DATA_KBN as ""種別"""
'			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",k.KO_HIN_GAI as ""資材"""
			sqlStr = sqlStr & ",convert(k.KO_QTY,SQL_NUMERIC) as ""員数"""
			sqlStr = sqlStr & ",convert(G_ST_URITAN,SQL_NUMERIC) as ""販売単価"""
			sqlStr = sqlStr & ",convert(G_ST_SHITAN,SQL_NUMERIC) as ""仕入単価"""
			sqlStr = sqlStr & " From P_COMPO as c"
			sqlStr = sqlStr & "    left outer join P_COMPO_K as k"
			sqlStr = sqlStr &				" on (c.SHIMUKE_CODE=k.SHIMUKE_CODE"
			sqlStr = sqlStr &				" and c.JGYOBU=k.JGYOBU"
			sqlStr = sqlStr &				" and c.NAIGAI=k.NAIGAI"
			sqlStr = sqlStr &				" and c.HIN_GAI=k.HIN_GAI"
			sqlStr = sqlStr &				" and k.DATA_KBN='1')"
			sqlStr = sqlStr & "    left outer join ITEM as i"
			sqlStr = sqlStr &				" on (k.KO_JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and k.KO_NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and k.KO_HIN_GAI	=i.HIN_GAI)"
			sqlStr = sqlStr & replace(whereStr,"P_COMPO.","c.",1,-1)
			sqlStr = sqlStr & " order by c.CLASS_CODE,c.HIN_GAI,k.DATA_KBN,k.SEQNO"
		case "pListGaiso"	' 一覧表(外装)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " k.SHIMUKE_CODE as ""仕向先"""
			sqlStr = sqlStr & ",k.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",k.DATA_KBN as ""種別"""
			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",k.KO_HIN_GAI as ""外装品番"""
			sqlStr = sqlStr & ",ki.HIN_NAME as ""外装品名"""
			sqlStr = sqlStr & ",convert(k.KO_QTY,SQL_NUMERIC) as ""入数"""
			sqlStr = sqlStr & ",ki.D_SIZE_W as ""W"""
			sqlStr = sqlStr & ",ki.D_SIZE_D as ""D"""
			sqlStr = sqlStr & ",ki.D_SIZE_H as ""H"""
			sqlStr = sqlStr & " From P_COMPO_K as k"
			sqlStr = sqlStr & "  inner join ITEM as i"
			sqlStr = sqlStr &				" on (k.JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and k.NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and k.HIN_GAI	=i.HIN_GAI)"
			sqlStr = sqlStr &  " left outer join ITEM as ki"
			sqlStr = sqlStr &				" on (k.KO_JGYOBU	=ki.JGYOBU"
			sqlStr = sqlStr &				" and k.KO_NAIGAI	=ki.NAIGAI"
			sqlStr = sqlStr &				" and k.KO_HIN_GAI	=ki.HIN_GAI)"
			sqlStr = sqlStr & replace(whereStr,"P_COMPO.","k.",1,-1)
			sqlStr = sqlStr & " order by ""仕向先"",""品番"",""種別"",""入数"" desc,""No"" desc"
		case "pListAll"	' 一覧表(全項目)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From P_COMPO p"
			sqlStr = sqlStr & whereStr                                                                            
			sqlStr = sqlStr & " order by 1,2,3,4,5,6"                                                                            
		case "pListKAll"	' 一覧表(全項目:子)
			server.scripttimeout = 900
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From P_COMPO_K p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by 1,2,3,4,5,6"                                                                            
		end select                                                                            
		                                                                            
		set rsList = db.Execute(sqlStr)                                                                            
	%>                                                                            
	<SCRIPT LANGUAGE="javascript"><!--
			cpTblBtn.value = "テーブル作成中..." + "<%=db.CommandTimeout%>";
	//--></SCRIPT>
	<div id='resultDiv'>                                                                            
	<table id="resultTbl">                                                                            
		<caption align='left'></caption>
		<TR BGCOLOR=#ccffee>
		<TH>No</TH>
		<%
'			Response.Flush()
			For i=0 To rsList.Fields.Count-1
				select case rsList.Fields(i).name
				case "span1"	fname = ""
				case "span2"	fname = ""
				case else		fname = rsList.Fields(i).name
				end select
				Response.Write "<TH title='" & rsList.Fields(i).name & "'>" & fname & "</TH>"
			Next
		%>
		</TR>
		<%
		const TristateTrue			= -1	'ゼロを表示します
		const TristateFalse			= 0		'ゼロを表示しません
		const TristateUseDefault	= -2	'「地域のプロパティ」の設定値を使用します
		dim		tankaS		' 商品化単価
		dim		tankaF		' 付加クラス単価
		dim		tankaG		' 外装クラス単価
		dim		strClass	' クラスコード
		dim		strGaiso	' 外装品番
		dim		strClassSQL	' 商品化クラス算出SQL
		dim		rsClass		' 商品化クラスレコードセット
        dim		strCnt
		dim		strPnPrev
		dim		strShimukePrev
		dim		strPnCurr
		dim		strShimukeCurr
		dim		bOutput

			cnt = 0
			lngMax = clng(maxStr)
			Do While Not rsList.EOF
				bOutput = true
				if ptypeStr = "pListGaiso" then
					strShimukeCurr = rsList.Fields("仕向先")
					strPnCurr = rsList.Fields("品番")
					if strShimukeCurr = strShimukePrev then
						if strPnCurr = strPnPrev then
							bOutput = false
						end if
					end if
					strShimukePrev	= strShimukeCurr
					strPnPrev		= strPnCurr
				end if
				if bOutput = true then
					cnt = cnt + 1
'					if lngMax > 0 then
'						if cnt > lngMax then
'							exit do
'						end if
'					end if
					strCnt = cnt

					tankaS = 0
					tankaF = 0
					tankaG = 0
			%>
					<TR VALIGN='TOP'>
					<TD nowrap id="Integer"><%=strCnt%></TD>                                                                            
			<%
					For i=0 To rsList.Fields.Count-1
							' 値
							fValue	= rtrim(rsList.Fields(i) & "")
							fType	= rsList.Fields(i).type
							select case rsList.Fields(i).name
							end select                                                                            
							' 位置定義（型）                                                                            
							select case fType                                                                            
							Case 2		' 数値(Integer)                                                                            
								tdTag = "<TD nowrap id=""Integer"">"                                                                            
	'							if fValue = "-32768" then                                                                            
	'								fValue = ""                                                                            
	'							end if                                                                            
							Case 2 , 3 , 5 , 131	' 数値(Integer)                                                                            
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
								select case rsList.Fields(i).Name
								case else
									tdTag = "<TD nowrap align=""left"" id=""Charactor"">"
								end select
							Case else		' その他
								tdTag = "<TD nowrap>"
							end select
							Response.Write tdTag & fValue & "</TD>"
					Next
			    	Response.Write "</TR>"
				end if
				rsList.Movenext
			Loop
		%>
	</TABLE></div>
	<hr>
	<div id="sql">
		<%=sqlStr%><br>
	</div>
<%
		rsList.Close                                                                            
		db.Close                                                                            
		set rsList = nothing                                                                            
		set db = nothing                                       
%>
	<SCRIPT LANGUAGE="javascript"><!--
		sqlForm.disabled = false;
		cpTblBtn.disabled = false;
		cpTblBtn.value = "結果をコピー";
	//--></SCRIPT>                                                                            
<% end if %>
</BODY>
</HTML>
