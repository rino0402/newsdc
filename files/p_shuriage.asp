<%@Language="VBScript"%>
<%Option Explicit%>
<% Response.Buffer = false %>
<% Response.Expires = -1 %>
<%
function GetTimeOut()
	GetTimeOut = Server.ScriptTimeout
end function

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
	dim	strVersion
	' 2007.09.27 p_shuriage.asp 売上実績データ検索
	strVersion = "2007.09.27 p_shuriage.asp 売上実績データ検索"
	strVersion = "2007.10.24 数値をカンマ区切り＆[単価]は少数2桁表示に変更"
	strVersion = "2007.11.05 一覧表:「収支」「販売区分」を追加"
	strVersion = "2007.11.05 一覧表:「得意先」順に変更"
	strVersion = "2007.11.05 一覧表:「得意先」「収支」「販売区分」に名称を追加"
	strVersion = "2010.08.20 ポップアップメニュー対応"
	strVersion = "2020.03.31 クリップボードコピー対応(IE以外)"

	dim objFS
	dim objF

	dim SHIMUKE_CODEStr
	dim UKEHARAI_CODEStr
	dim TORI_KBNStr
	dim SHIJI_FStr
	dim pnStr
	dim HAKKO_DTSStr
	dim HAKKO_DTEStr
	dim UKEIRE_DTSStr
	dim UKEIRE_DTEStr
	dim KAN_DTSStr
	dim KAN_DTEStr
	dim submitStr
	dim ptypeStr
	dim cmpStr
	dim db
	dim rsList
	dim rsRow
	dim dbName
	dim	sqlStr
	dim	whereStr
	dim andStr
	dim cnt,i
	dim fValue,fname,tdTag
	dim fType
	dim centerStr
	dim SHIJI_NOStr
	dim	KEIJYO_YMSStr
	dim	KEIJYO_YMEStr
	dim	maxStr
	dim	lngMax

	dbName = "newsdc"
	SHIJI_NOStr		= ucase(Request.QueryString("SHIJI_NO"))
	SHIMUKE_CODEStr	= ucase(Request.QueryString("SHIMUKE_CODE"))
	SHIJI_FStr		= ucase(Request.QueryString("SHIJI_F"))
	TORI_KBNStr		= ucase(Request.QueryString("TORI_KBN"))
	
	UKEHARAI_CODEStr = ucase(Request.QueryString("UKEHARAI_CODE"))

	HAKKO_DTSStr = Request.QueryString("HAKKO_DTS")
	HAKKO_DTEStr = Request.QueryString("HAKKO_DTE")

	UKEIRE_DTSStr = Request.QueryString("UKEIRE_DTS")
	UKEIRE_DTEStr = Request.QueryString("UKEIRE_DTE")

	KAN_DTSStr = Request.QueryString("KAN_DTS")
	KAN_DTEStr = Request.QueryString("KAN_DTE")

	KEIJYO_YMSStr	= ucase(Request.QueryString("KEIJYO_YMS"))
	KEIJYO_YMEStr	= ucase(Request.QueryString("KEIJYO_YME"))

	pnStr = ucase(Request.QueryString("pn"))
	submitStr = Request.QueryString("submit1")
	ptypeStr = Request.QueryString("ptype")
	if len(ptypeStr) = 0 then
		ptypeStr = "pList"
		KEIJYO_YMSStr = year(now()) & right("0" & month(now()),2)
	end if
	maxStr			= rtrim(ucase(Request.QueryString("max")))
	if len(maxStr) = 0 then
		maxStr	= 100
	end if
	lngMax = clng(maxStr)
%>
<!--#include file="info.txt" -->
<!--#include file="makeWhere.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="result.css" TITLE="CSS">
<TITLE><%=centerStr%> 売上実績</TITLE>
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
//		DoneMes();
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
		<caption style="text-align:left;">売上実績データ検索</caption>
		<tr>
			<th>注文No</th>
			<th></th>
			<th></th>
			<th></th>
			<th>品番</th>
			<th>受入日</th>
			<th>計上年月</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="SHIJI_NO" id="SHIJI_NO" VALUE="<%=SHIJI_NOStr%>" size="6" maxlength="5">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="SHIMUKE_CODE" id="SHIMUKE_CODE" VALUE="<%=SHIMUKE_CODEStr%>"
						size="2" maxlength="2" style="text-align:center;">
				<div align="left">
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="TORI_KBN" id="TORI_KBN" VALUE="<%=TORI_KBNStr%>"
						size="4" maxlength="3" style="text-align:center;">
			</td>

			<td align="center">
				<INPUT TYPE="text" NAME="UKEHARAI_CODE" id="UKEHARAI_CODE" VALUE="<%=UKEHARAI_CODEStr%>"
						size="6" maxlength="5" style="text-align:center;">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="pn" id="pn" VALUE="<%=pnStr%>" size="22" maxlength="20"><br>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="UKEIRE_DTS" id="UKEIRE_DTS" VALUE="<%=UKEIRE_DTSStr%>" size="10" maxlength="8">
				～
				<INPUT TYPE="text" NAME="UKEIRE_DTE" id="UKEIRE_DTE" VALUE="<%=UKEIRE_DTEStr%>" size="10" maxlength="8">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEIJYO_YMS" id="KEIJYO_YMS" VALUE="<%=KEIJYO_YMSStr%>" size="10" maxlength="8">
				～
				<INPUT TYPE="text" NAME="KEIJYO_YME" id="KEIJYO_YME" VALUE="<%=KEIJYO_YMEStr%>" size="10" maxlength="8">
			</td>

		</tr>
		<tr>
			<td colspan="7">
				<table border="0" cellspacing="0" bordercolor="White">
					<tr>
					<td valign="top">
						<b>出力形式：</b>
					</td>
					<td>
						<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
							<label for="pList">一覧表</label>
						<INPUT TYPE="radio" NAME="ptype" VALUE="pTable" id="pTable">
							<label for="pTable">集計表(品番・月別)</label>
						<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
							<label for="pListAll">一覧表(全項目)</label>
					</td>                                                                            
					</tr>                                                                            
				</table>                                                                            
			</td>                                                                            
		</tr>                                                                            
		<tr bordercolor=White>                                                                            
			<td colspan="7">                                                                            
			<INPUT TYPE="submit" value="検索" id=submit1 name=submit1>                                                                            
			<INPUT TYPE="reset" value="リセット" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
				最大件数：<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=maxStr%>" size="8" maxlength="6">
				<%=strVersion%>
			</td>
		</tr>                                                                            
	</table>                                                                            
	</div>                                                                            
  </FORM>
<SCRIPT LANGUAGE='JavaScript'>
	ptypeChange('<%=ptypeStr%>');
</SCRIPT>
                                                                  
<%	if len(submitStr) > 0 then %>                                                                            
	<SCRIPT LANGUAGE="javascript"><!--
		sqlForm.disabled = true;
	//--></SCRIPT>
	<%
		dim	P_SSHIJI_O
		Set db = Server.CreateObject("ADODB.Connection")
		db.CommandTimeout = 360

		db.open dbName
		sqlStr = ""
		whereStr = ""
		andStr = " where"
		P_SSHIJI_O = false
		if len(SHIJI_NOStr) > 0 then
			if left(SHIJI_NOStr,1) = "-" then
				whereStr = whereStr & andStr & " u.ORDER_NO <> '" & SHIJI_NOStr & "'"
			else
				whereStr = whereStr & andStr & " u.ORDER_NO = '" & SHIJI_NOStr & "'"
			end if
			andStr = " and"
		end if

		if len(SHIMUKE_CODEStr) > 0 then
			if left(SHIMUKE_CODEStr,1) = "-" then
				whereStr = whereStr & andStr & " P_SUKEIRE.SHIMUKE_CODE <> '" & SHIMUKE_CODEStr & "'"
			else
				whereStr = whereStr & andStr & " P_SUKEIRE.SHIMUKE_CODE = '" & SHIMUKE_CODEStr & "'"
			end if
			andStr = " and"
		end if

		if len(SHIJI_FStr) > 0 then
			if left(SHIJI_FStr,1) = "-" then
				whereStr = whereStr & andStr & " SHIJI_F <> '" & SHIJI_FStr & "'"
			else
				whereStr = whereStr & andStr & " SHIJI_F = '" & SHIJI_FStr & "'"
			end if
			andStr = " and"
			P_SSHIJI_O = true
		end if

		if len(TORI_KBNStr) > 0 then
			if left(TORI_KBNStr,1) = "-" then
				whereStr = whereStr & andStr & " P_UKEHARAI.TORI_KBN <> '" & TORI_KBNStr & "'"
			else
				whereStr = whereStr & andStr & " P_UKEHARAI.TORI_KBN = '" & TORI_KBNStr & "'"
			end if
			andStr = " and"
		end if


		if len(UKEHARAI_CODEStr) > 0 then
			if left(UKEHARAI_CODEStr,1) = "-" then
				whereStr = whereStr & andStr & " TORI_CODE <> '" & UKEHARAI_CODEStr & "'"
			else
				if InStr(1,UKEHARAI_CODEStr,"%") > 0 then
					whereStr = whereStr & andStr & " TORI_CODE like '" & UKEHARAI_CODEStr & "'"
				else
					whereStr = whereStr & andStr & " TORI_CODE = '" & UKEHARAI_CODEStr & "'"
				end if
			end if
			andStr = " and"
		end if

		if len(HAKKO_DTSStr) > 0 then
			whereStr = whereStr & andStr & " HAKKO_DT >= '" & HAKKO_DTSStr & "'"
			andStr = " and"
		end if
		if len(HAKKO_DTEStr) > 0 then
			whereStr = whereStr & andStr & " HAKKO_DT <= '" & HAKKO_DTEStr & "'"
			andStr = " and"
		end if

		dim	strKanDtFname
		if ptypeStr = "pList" then
			strKanDtFname = "UKEIRE_DT"
		else
			strKanDtFname = "UKEIRE_DT"
'			strKanDtFname = "KAN_DT"
		end if

		if len(UKEIRE_DTSStr) > 0 then
			whereStr = whereStr & andStr & " UKEIRE_DT >= '" & UKEIRE_DTSStr & "'"
			andStr = " and"
		end if
		if len(UKEIRE_DTEStr) > 0 then
			whereStr = whereStr & andStr & " UKEIRE_DT <= '" & UKEIRE_DTEStr & "'"
			andStr = " and"
		end if
		if len(KEIJYO_YMSStr) > 0 then
			whereStr = whereStr & andStr & " KEIJYO_YM >= '" & KEIJYO_YMSStr & "'"
			andStr = " and"
		end if
		if len(KEIJYO_YMEStr) > 0 then
			whereStr = whereStr & andStr & " KEIJYO_YM <= '" & KEIJYO_YMEStr & "'"
			andStr = " and"
		end if

		if len(pnStr) > 0 then
			if instr(1,pnStr,"%") > 0 then
				cmpStr = "like"
			else
				cmpStr = "="
			end if
			whereStr = whereStr & andStr & " u.hin_gai " & cmpStr & " '" & pnStr & "'"
			andStr = " and"
		end if

		sqlStr = "select "
		if lngMax > 0 then
			sqlStr = sqlStr & " top " & lngMax
		end if
		select case ptypeStr
		case "pList"	' 一覧表
			sqlStr = sqlStr & " u.KEIJYO_YM as ""計上年月"""
			sqlStr = sqlStr & ",u.TOKUI_CODE + ' ' + t.UKEHARAI_NAME as ""得意先"""
			sqlStr = sqlStr & ",u.G_SYUSHI     + ' ' + p03.C_RNAME as ""収支"""
			sqlStr = sqlStr & ",u.G_HANBAI_KBN + ' ' + p02.C_RNAME as ""販売区分"""
			sqlStr = sqlStr & ",u.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			sqlStr = sqlStr & ",convert(u.TANKA,SQL_NUMERIC) as ""単価"""
			sqlStr = sqlStr & ",convert(u.URIAGE_QTY,SQL_NUMERIC) as ""売上数"""
			sqlStr = sqlStr & ",convert(u.KINGAKU,SQL_NUMERIC) as ""金額"""
			sqlStr = sqlStr & ",u.URIAGE_DT as ""売上日"""
			sqlStr = sqlStr & ",u.URIAGE_NO as ""売上No"""
			sqlStr = sqlStr & ",u.SEIKU_F as ""請求F"""
			sqlStr = sqlStr & " From P_SHURIAGE as u"
			sqlStr = sqlStr & " LEFT OUTER JOIN ITEM as i on i.JGYOBU='S' and i.NAIGAI='1' and u.HIN_GAI = i.HIN_GAI"
			sqlStr = sqlStr & " LEFT OUTER JOIN P_UKEHARAI as t on u.TOKUI_CODE = t.UKEHARAI_CODE"
			sqlStr = sqlStr & " LEFT OUTER JOIN P_CODE as p03 on u.G_SYUSHI     = p03.C_Code and p03.DATA_KBN = '03'"
			sqlStr = sqlStr & " LEFT OUTER JOIN P_CODE as p02 on u.G_HANBAI_KBN = p02.C_Code and p02.DATA_KBN = '02'"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""計上年月"",""得意先"",""収支"",""販売区分"",""品番"",""売上日"""
		case "pTable"	' 集計表(品番・月別)
			dim	sqlStr2
			sqlStr2 = "select distinct"
			sqlStr2 = sqlStr2 & " u.KEIJYO_YM as ym"
			sqlStr2 = sqlStr2 & " From P_SHURIAGE as u"
			sqlStr2 = sqlStr2 & whereStr
			sqlStr2 = sqlStr2 & " order by ym"
			set rsList = db.Execute(sqlStr2)

			sqlStr = sqlStr & " u.TOKUI_CODE + ' ' + t.UKEHARAI_NAME as ""得意先"""
			sqlStr = sqlStr & ",u.G_SYUSHI     + ' ' + p03.C_RNAME as ""収支"""
			sqlStr = sqlStr & ",u.G_HANBAI_KBN + ' ' + p02.C_RNAME as ""販売区分"""
			sqlStr = sqlStr & ",u.HIN_GAI as ""品番"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""品名"""
			Do While Not rsList.EOF
				sqlStr = sqlStr & ",sum(if(left(u.KEIJYO_YM,6) ='" & trim(rsList.Fields("ym")) & "',convert(u.KINGAKU,SQL_NUMERIC),0)) as """ & trim(rsList.Fields("ym")) & """"
				rsList.Movenext
			loop
			sqlStr = sqlStr & " From P_SHURIAGE as u"
			sqlStr = sqlStr & " LEFT OUTER JOIN ITEM as i on i.JGYOBU='S' and i.NAIGAI='1' and u.HIN_GAI = i.HIN_GAI"
			sqlStr = sqlStr & " LEFT OUTER JOIN P_UKEHARAI as t on u.TOKUI_CODE = t.UKEHARAI_CODE"
			sqlStr = sqlStr & " LEFT OUTER JOIN P_CODE as p03 on u.G_SYUSHI     = p03.C_Code and p03.DATA_KBN = '03'"
			sqlStr = sqlStr & " LEFT OUTER JOIN P_CODE as p02 on u.G_HANBAI_KBN = p02.C_Code and p02.DATA_KBN = '02'"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""得意先"",""収支"",""販売区分"",""品番"",""品名"""
			sqlStr = sqlStr & " order by ""得意先"",""収支"",""販売区分"",""品番"""
		case "pListAll"	' 一覧表(全項目)
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From P_SHURIAGE as u"
			sqlStr = sqlStr & whereStr
		end select
%>
	<div>
		<!--INPUT TYPE="button" onClick="DoCopy('resultDiv')"
			value="検索中...ScriptTimeout=<%=Server.ScriptTimeout%>"
			id="cpTblBtn" disabled-->
		<button id="btnClip" class="btn" data-clipboard-target="#resultTbl" disabled onClick="DoCopy('resultDiv');">
			検索中...ScriptTimeout=<%=Server.ScriptTimeout%>
		</button>
	</div>
	<!--%=sqlStr%-->
<%
		set rsList = db.Execute(sqlStr)
%>
		<SCRIPT LANGUAGE="javascript"><!--
//				cpTblBtn.value = "テーブル作成中...";
				$('#btnClip').text('テーブル作成中...')
		//--></SCRIPT>
	<div id='resultDiv'>
	<table id="resultTbl">
		<thead>
			<TR>
			<TH>No</TH>
		<%
			dim strThTag
%>
<%
				For i=0 To rsList.Fields.Count-1
					strThTag = ""
					fname = rsList.Fields(i).name
					select case fname
					case "span1"	fname = ""
					case else		
					end select
%>
					<TH<%=strThTag%>><%=fname%></TH>
<%
				Next
%>
			</TR>
		</thead>
		<tbody>
		<%
		const TristateTrue			= -1	'ゼロを表示します
		const TristateFalse			= 0		'ゼロを表示しません
		const TristateUseDefault	= -2	'「地域のプロパティ」の設定値を使用します
        dim		strCnt
		dim		tdTitle
		dim		ttlShijiQty
		dim		ttlUkeQty
		dim		ttlNaiMny
		dim		totalArray(255)
		dim		j
		
			cnt			= 0
			for i = 0 to ubound(totalArray)
				totalArray(i) = 0
			next
			Do While Not rsList.EOF
				cnt = cnt + 1
				strCnt = cnt
		%>
				<TR VALIGN='TOP'>
				<TD nowrap id="Integer"><%=strCnt%></TD>                                                                            
		<%
				For i=0 To rsList.Fields.Count-1
						' 値
						tdTitle = ""
						fValue	= rtrim(rsList.Fields(i))
						fType	= rsList.Fields(i).type
						' 位置定義（型）
						select case fType
						Case 2 , 3 , 5 , 131	' 数値(Integer)
							tdTag = "<TD nowrap id=""Integer"">"
							if isnumeric(fValue) then
								totalArray(i) = totalArray(i) + cdbl(fValue)
							end if
							if fValue = 0 then
								fValue = ""
							else
								if rsList.Fields(i).Name = "単価" then
									fValue = formatnumber(fValue,2,,,-1)
								else
									fValue = formatnumber(fValue,0,,,-1)
								end if
							end if
						Case 133				' 日付(Date)
							tdTag = "<TD nowrap id=""Date"">"
							if len(fValue) > 0 then
								fValue = year(rsList.Fields(i)) & "/"
								fValue = fValue & rtrim("0" & month(rsList.Fields(i)),2)
								fValue = fValue & "/"                                                                            
								fValue = fValue & rtrim("0" & day(rsList.Fields(i)),2)
							end if
						Case 129				' 文字列(Charactor)
							tdTag = "<TD nowrap align=""left"" id=""Charactor""" & tdTitle & ">"
						Case else				' その他
							tdTag = "<TD nowrap" & tdTitle & ">"
						end select
%>
						<%=tdTag%><%=fValue%></TD>
<%
				Next
%>
		    	</TR>
<%
				rsList.Movenext
			Loop
%>                                                                            
		</tbody>
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
		//	テーブルの先頭に合計行を追加

		function InnerText(c,t) {
			if (navigator.userAgent.indexOf("Firefox") > -1) {
				c.textContent　= t;
			} else {
				c.innerText = t;
			}
		}
		function GetInnerText(c) {
			if (navigator.userAgent.indexOf("Firefox") > -1) {
				return c.textContent;
			} else {
				return c.innerText;
			}
		}
	
	
//		cpTblBtn.disabled = false;
//		cpTblBtn.value = "合計集計中...";         
//		cpTblBtn.disabled = true;
		var	colCount = resultTbl.rows[0].cells.length;
//				window.alert(colCount);
		var row = resultTbl.insertRow(1);
		
		for (var j = 0; j < colCount; j++) {                                 
			var cell	= row.insertCell(j);
			cell.noWrap = true;
			if (j == 0) {
				var cnt = 0;
				var loop = resultTbl.rows.length;             
				for(var i = 1 ; i < loop ; i++)	{    
					var	strText = "" + GetInnerText(resultTbl.rows[i].cells[j]);    
					if(strText == "")	{    
					} else { 
						strText = strText.replace(",","");   
						cnt = parseInt(strText);     
					}             
				}          
				cell.align = "right";
				InnerText(cell,myFormatNumber(cnt));
			}
			if (j == 1) {
				InnerText(cell,'合計');
			}
			var	strColTitle = "" + GetInnerText(resultTbl.rows[0].cells[j]);
			if (strColTitle == '指示数' ||
				strColTitle == '完了数' ||
				strColTitle == '受入数' ||
				strColTitle == '件数' ||
				strColTitle == '金額' ||
				strColTitle == '外装金額' ||
				strColTitle == '外装数量' ||
				strColTitle == '商品化金額') {
				var sum = 0;
				var loop = resultTbl.rows.length;
				for(var i = 1 ; i < loop ; i++)	{  
					var	strText = "" + GetInnerText(resultTbl.rows[i].cells[j]);
					if(strText == "")	{
					} else { 
						strText = strText.replace(",","");
						sum += parseInt(strText);
					}
				}
				cell.align = "right";
				InnerText(cell,myFormatNumber(sum));
			}
			if (strColTitle.substring(0,4) == '作業時間' ||
				strColTitle.substring(0,2) == '人数' ||
				strColTitle.substring(0,1) == '分') {
				var sum = 0;
				var loop = resultTbl.rows.length;
				for(var i = 1 ; i < loop ; i++)	{
					var	strText = "" + GetInnerText(resultTbl.rows[i].cells[j]);
					if(strText == "")	{
					} else if(strText == "0")	{
						InnerText(resultTbl.rows[i].cells[j],"");
					} else {
						strText = strText.replace(",","");
						sum += parseFloat(strText);
					}
				}      
				if (strColTitle.substring(0,4) == '作業時間') {
					cell.align = "right";
					InnerText(cell,myFormatNumber(sum));
				}
			}
		}
	//--></SCRIPT>
	<SCRIPT LANGUAGE="javascript"><!--
		sqlForm.disabled = false;
		btnClip.disabled = false;
		$('#btnClip').text('結果をコピー')
//		cpTblBtn.disabled = false;
//		cpTblBtn.value = "結果をコピー";
	//--></SCRIPT>                                                                            
<% end if %>
</BODY>
</HTML>
