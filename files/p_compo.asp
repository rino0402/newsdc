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

	versionStr = "2010.03.04 �o�[�W�����\���Ή�"
	versionStr = "2010.03.04 �ő�s���Ή�"
	versionStr = "2010.03.04 �ꗗ�\(�O��) �Ή�"
	versionStr = "2010.08.18 �|�b�v�A�b�v���j���[�Ή�/�W�v�\(�X�V��)�Ή�"
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
<TITLE><%=centerStr%> �\���}�X�^�[</TITLE>
<!-- jdMenu head�p include �J�n -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<!-- jdMenu head�p include �I�� -->
<SCRIPT LANGUAGE="JavaScript"><!--
	navi = navigator.userAgent;

	function DoCopy(arg){
		var doc = document.body.createTextRange();
		doc.moveToElementText(document.all(arg));
		doc.execCommand("copy");
		window.alert("�N���b�v�{�[�h�փR�s�[���܂����B\n�\��t���ł��܂��B" );
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
	// (���ׂĂ̕ϐ��Ɋi�[����l��0�I���W���Ƃ���) 
	function myFormatNumber(x) { // �����̗�Ƃ��Ă� 95839285734.3245
	    var s = "" + x; // �m���ɕ�����^�ɕϊ�����B��ł� "95839285734.3245"
	    var p = s.indexOf("."); // �����_�̈ʒu��0�I���W���ŋ��߂�B��ł� 11
	    if (p < 0) { // �����_��������Ȃ�������
	        p = s.length; // ���z�I�ȏ����_�̈ʒu�Ƃ���
	    }
	    var r = s.substring(p, s.length); // �����_�̌��Ə����_���E���̕�����B��ł� ".3245"
	    for (var i = 0; i < p; i++) { // (10 ^ i) �̈ʂɂ���
	        var c = s.substring(p - 1 - i, p - 1 - i + 1); // (10 ^ i) �̈ʂ̂ЂƂ̌��̐����B��ł� "4", "3", "7", "5", "8", "2", "9", "3", "8", "5", "9" �̏��ɂȂ�B
	        if (c < "0" || c > "9") { // �����ȊO�̂���(�����Ȃ�)����������
	            r = s.substring(0, p - i) + r; // �c���S���t������
	            break;
	        }
	        if (i > 0 && i % 3 == 0) { // 3 �����ƁA����������͏���
	            r = "," + r; // �J���}��t������
	        }
	        r = c + r; // �������ꌅ�ǉ�����B
	    }
	    return r; // ��ł� "95,839,285,734.3245"
	}
//--></SCRIPT>
</HEAD>
<BODY>
<!-- jdMenu body�p include �J�n -->
<!--#include file="jdmenu-sdc.asp" -->
<!-- jdMenu body�p include �I�� -->
  <FORM name="sqlForm">
  <div id="sqlDiv">
	<table id="sqlTbl">
		<caption style="text-align:left;">�\���}�X�^�[����</caption>
		<tr>
			<th>�d����</th>
			<th>�f�[�^�敪</th>
			<th>�e�i��</th>
			<th>�q�i��</th>
			<th>���i���N���X</th>
			<th>�t���N���X</th>
			<th>���E�N���X</th>
			<th>���l</th>
			<th>�S����</th>
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
				0:�N���X<br>
				1:������<br>
				2:�O������<br>
				3:�\�����i
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
					<b>�o�͌`���F</b>
					</td>
					<td>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKbn" id="pTableKbn">
						<label for="pTableKbn">�W�v�\(�d����/�敪�� �o�^����)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableUpd" id="pTableUpd">
						<label for="pTableUpd">�W�v�\(�X�V��)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKosei" id="pTableKosei">
						<label for="pTableKosei">�W�v�\(����)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pTableChild" id="pTableChild">
						<label for="pTableChild">�W�v�\(�\���q)</label>
					<br>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
						<label for="pList">�ꗗ�\</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListClass" id="pListClass" onclick="sqlForm.DATA_KBN.value='0';">
						<label for="pListClass">�ꗗ�\(�N���X)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListCompo" id="pListCompo">
						<label for="pListCompo">�ꗗ�\(�\��)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListGaiso" id="pListGaiso" onclick="sqlForm.DATA_KBN.value='2';">
						<label for="pListGaiso">�ꗗ�\(�O��)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
						<label for="pListAll">�ꗗ�\(�S����)</label>
					<INPUT TYPE="radio" NAME="ptype" VALUE="pListKAll" id="pListKAll">
						<label for="pListKAll">�ꗗ�\(�S����:�q)</label>
					</td>
					</tr>
					<tr>
						<td align="right">
							<b>�ΏہF</b>
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
			<INPUT TYPE="submit" value="����" id=submit1 name=submit1>
			<INPUT TYPE="reset" value="���Z�b�g" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
			�ő匏���F<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=maxStr%>" size="8" maxlength="6">
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
	<INPUT TYPE='button' onClick="DoCopy('resultDiv')" value='������...' id='cpTblBtn' disabled>                                                                            
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
		case "pTableKbn"	' �W�v�\(�敪��)                                                                            
			db.CommandTimeout = 360
			sqlStr = sqlStr & " p.SHIMUKE_CODE + rtrim(' ' + ifnull(c.C_NAME,'')) ""�d����"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '0',1,0)) ""�N���X<br>����"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '1',1,0)) ""����<br>����"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '2',1,0)) ""�O��<br>����"""
			sqlStr = sqlStr & ",sum(if(p.DATA_KBN = '3',1,0)) ""�\�����i<br>����"""
			sqlStr = sqlStr & " From P_COMPO p"
			sqlStr = sqlStr & " left outer join p_code c on (p.SHIMUKE_CODE=c.C_Code and c.DATA_KBN='04')"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�d����"""
			sqlStr = sqlStr & " order by ""�d����"""
		case "pTableUpd"	' �W�v�\(�X�V��)
			sqlStr = sqlStr & " left(UPD_DATETIME,8) ""�X�V��"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '0',1,0)) as ""�N���X<br>����"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '1',1,0)) as ""����<br>����"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '2',1,0)) as ""�O��<br>����"""
			sqlStr = sqlStr & ",sum(if(P_COMPO.DATA_KBN = '3',1,0)) as ""�\�����i<br>����"""
			sqlStr = sqlStr & " From P_COMPO p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�X�V��"""
			sqlStr = sqlStr & " order by ""�X�V��"" desc"
		case "pTableKosei"	' �W�v�\(����)                                                                      			db.CommandTimeout = 360
			sqlStr = sqlStr & " HIN_GAI"
			sqlStr = sqlStr & ",count(*)"
			sqlStr = sqlStr & " From P_COMPO_K p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by HIN_GAI"
			sqlStr = sqlStr & " order by HIN_GAI desc"
		case "pTableChild"	' �W�v�\(�\���q)
			db.CommandTimeout = 360
			sqlStr = sqlStr & " p.KO_JGYOBU ""���ƕ�(�q)"""
			sqlStr = sqlStr & ",p.KO_NAIGAI ""���O(�q)"""
			sqlStr = sqlStr & ",p.KO_HIN_GAI ""�i��(�q)"""
			sqlStr = sqlStr & ",i.HIN_NAME ""�i��(�q)"""
			sqlStr = sqlStr & ",count(*) ""����"""
			sqlStr = sqlStr & ",sum(convert(p.KO_QTY,SQL_decimal)) ""�����v"""
			sqlStr = sqlStr & " From P_COMPO_K p"
			sqlStr = sqlStr & "    left outer join ITEM i"
			sqlStr = sqlStr &				" on (p.KO_JGYOBU	=i.JGYOBU"
			sqlStr = sqlStr &				" and p.KO_NAIGAI	=i.NAIGAI"
			sqlStr = sqlStr &				" and p.KO_HIN_GAI	=i.HIN_GAI)"
			sqlStr = sqlStr & replace(whereStr,"P_COMPO.","k.",1,-1)
			sqlStr = sqlStr & andStr & " p.DATA_KBN <> '0'"
'			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""���ƕ�(�q)"",""���O(�q)"",""�i��(�q)"",""�i��(�q)"""
			sqlStr = sqlStr & " order by ""���ƕ�(�q)"",""���O(�q)"",""�i��(�q)"""
		case "pList"	' �ꗗ�\
			sqlStr = sqlStr & " p.SHIMUKE_CODE ""�d����"""
			sqlStr = sqlStr & ",p.HIN_GAI ""�i��"""
			sqlStr = sqlStr & ",p.DATA_KBN"
			sqlStr = sqlStr & ",p.SEQNO"
			sqlStr = sqlStr & ",p.KO_SYUBETSU"
			sqlStr = sqlStr & ",p.KO_JGYOBU"
			sqlStr = sqlStr & ",p.KO_NAIGAI"
			sqlStr = sqlStr & ",p.KO_HIN_GAI"
			sqlStr = sqlStr & ",convert(p.KO_QTY,sql_decimal) KO_QTY"
			sqlStr = sqlStr & ",p.BIKOU ""���l"""
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
		case "pListClass"	' �ꗗ�\(���i���N���X)                                                      
			Server.ScriptTimeout = 900
			db.CommandTimeout = 360
			sqlStr = sqlStr & " p.SHIMUKE_CODE ""�d����"""
'			sqlStr = sqlStr & ",c.DATA_KBN as ""�f�[�^<br>�敪"""
			sqlStr = sqlStr & ",p.HIN_GAI ""�i��"""
			sqlStr = sqlStr & ",rtrim(p.CLASS_CODE) + ' ' + rtrim(CLASS_NAME) ""���i���N���X"""
			sqlStr = sqlStr & ",p.F_CLASS_CODE ""�t���N���X"""
			sqlStr = sqlStr & ",p.N_CLASS_CODE ""���E�N���X"""
			sqlStr = sqlStr & ",convert(TANKA,SQL_NUMERIC) ""�P��"""
			sqlStr = sqlStr & ",' ' span1"
			sqlStr = sqlStr & ",convert(KOUSU,SQL_NUMERIC) ""�H��"""
			sqlStr = sqlStr & ",convert(KOURYOU,SQL_NUMERIC) ""�H��"""
			sqlStr = sqlStr & ",convert(ETC,SQL_NUMERIC) ""���̑�"""
			sqlStr = sqlStr & ",' ' span2"
			sqlStr = sqlStr & ",count(k.KO_HIN_GAI) ""���ތ���"""
			sqlStr = sqlStr & ",sum(convert(k.KO_QTY,SQL_NUMERIC)) ""���ވ���"""
			sqlStr = sqlStr & ",sum(convert(G_ST_URITAN,SQL_NUMERIC)*convert(k.KO_QTY,SQL_NUMERIC)) ""�̔��P��"""
			sqlStr = sqlStr & ",sum(convert(G_ST_SHITAN,SQL_NUMERIC)*convert(k.KO_QTY,SQL_NUMERIC)) ""�d���P��"""
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
			sqlStr = sqlStr & " group by ""�d����"",""�i��"",""���i���N���X"",""�t���N���X"",""���E�N���X"",""�P��"",span1,""�H��"",""�H��"",""���̑�"",span2"
			sqlStr = sqlStr & " order by ""���i���N���X"",""�i��"""
		case "pListCompo"	' �ꗗ�\(���i���\��)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " k.SHIMUKE_CODE as ""�d����"""
			sqlStr = sqlStr & ",k.HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME ""�i��"""
'			sqlStr = sqlStr & ",if(k.DATA_KBN = '0',k.CLASS_CODE,'') as ""���i���N���X"""
'			sqlStr = sqlStr & ",k.F_CLASS_CODE as ""�t���N���X"""
'			sqlStr = sqlStr & ",k.N_CLASS_CODE as ""���E�N���X"""
			sqlStr = sqlStr & ",k.DATA_KBN as ""���"""
			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',k.KO_SYUBETSU,'') as ""�\�����"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',k.KO_HIN_GAI,'') as ""�\���i��"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',s.HIN_NAME,'') as ""�\���i��"""
			sqlStr = sqlStr & ",if(k.DATA_KBN <> '0',convert(k.KO_QTY,SQL_NUMERIC),'') as ""�\������"""
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
			sqlStr = sqlStr & " order by ""�d����"",""�i��"",""���"",""No"""
		case "pListCompoXX"	' �ꗗ�\(���i���\��)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " c.SHIMUKE_CODE as ""�d����"""
			sqlStr = sqlStr & ",c.HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",rtrim(c.CLASS_CODE) + ' ' + rtrim(CLASS_NAME) as ""���i���N���X"""
			sqlStr = sqlStr & ",c.F_CLASS_CODE as ""�t���N���X"""
			sqlStr = sqlStr & ",c.N_CLASS_CODE as ""���E�N���X"""
			sqlStr = sqlStr & ",convert(TANKA,SQL_NUMERIC) as ""���i���P��"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",convert(KOURYOU,SQL_NUMERIC) as ""�H��"""
			sqlStr = sqlStr & ",convert(TANKA,SQL_NUMERIC) - convert(KOURYOU,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",convert(KOUSU,SQL_NUMERIC) as ""�H��"""
			sqlStr = sqlStr & ",convert(ETC,SQL_NUMERIC) as ""���̑�"""
			sqlStr = sqlStr & ",' ' as span2"
			sqlStr = sqlStr & ",k.DATA_KBN as ""���"""
			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",KO_HIN_GAI as ""���ޕi��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""���ޕi��"""
			sqlStr = sqlStr & ",convert(k.KO_QTY,SQL_NUMERIC) as ""���ވ���"""
			sqlStr = sqlStr & ",convert(i.G_ST_URITAN,SQL_NUMERIC) as ""�̔��P��"""
			sqlStr = sqlStr & ",convert(i.G_ST_SHITAN,SQL_NUMERIC) as ""�d���P��"""
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
			sqlStr = sqlStr & " order by ""���i���N���X"",""�i��"",""���"",""No"""
		case "pListCompoOld"	' �ꗗ�\(���i���\��)
			sqlStr = sqlStr & " c.SHIMUKE_CODE as ""�d����"""
			sqlStr = sqlStr & ",c.HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",right(c.CLASS_CODE,length(c.CLASS_CODE)-2) as ""���i���N���X"""
			sqlStr = sqlStr & ",c.F_CLASS_CODE as ""�t���N���X"""
			sqlStr = sqlStr & ",c.N_CLASS_CODE as ""���E�N���X"""
'			sqlStr = sqlStr & ",k.DATA_KBN as ""���"""
'			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",k.KO_HIN_GAI as ""����"""
			sqlStr = sqlStr & ",convert(k.KO_QTY,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",convert(G_ST_URITAN,SQL_NUMERIC) as ""�̔��P��"""
			sqlStr = sqlStr & ",convert(G_ST_SHITAN,SQL_NUMERIC) as ""�d���P��"""
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
		case "pListGaiso"	' �ꗗ�\(�O��)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " k.SHIMUKE_CODE as ""�d����"""
			sqlStr = sqlStr & ",k.HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
			sqlStr = sqlStr & ",k.DATA_KBN as ""���"""
			sqlStr = sqlStr & ",k.SEQNO as ""No"""
			sqlStr = sqlStr & ",k.KO_HIN_GAI as ""�O���i��"""
			sqlStr = sqlStr & ",ki.HIN_NAME as ""�O���i��"""
			sqlStr = sqlStr & ",convert(k.KO_QTY,SQL_NUMERIC) as ""����"""
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
			sqlStr = sqlStr & " order by ""�d����"",""�i��"",""���"",""����"" desc,""No"" desc"
		case "pListAll"	' �ꗗ�\(�S����)
			db.CommandTimeout = 900
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From P_COMPO p"
			sqlStr = sqlStr & whereStr                                                                            
			sqlStr = sqlStr & " order by 1,2,3,4,5,6"                                                                            
		case "pListKAll"	' �ꗗ�\(�S����:�q)
			server.scripttimeout = 900
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From P_COMPO_K p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by 1,2,3,4,5,6"                                                                            
		end select                                                                            
		                                                                            
		set rsList = db.Execute(sqlStr)                                                                            
	%>                                                                            
	<SCRIPT LANGUAGE="javascript"><!--
			cpTblBtn.value = "�e�[�u���쐬��..." + "<%=db.CommandTimeout%>";
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
		const TristateTrue			= -1	'�[����\�����܂�
		const TristateFalse			= 0		'�[����\�����܂���
		const TristateUseDefault	= -2	'�u�n��̃v���p�e�B�v�̐ݒ�l���g�p���܂�
		dim		tankaS		' ���i���P��
		dim		tankaF		' �t���N���X�P��
		dim		tankaG		' �O���N���X�P��
		dim		strClass	' �N���X�R�[�h
		dim		strGaiso	' �O���i��
		dim		strClassSQL	' ���i���N���X�Z�oSQL
		dim		rsClass		' ���i���N���X���R�[�h�Z�b�g
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
					strShimukeCurr = rsList.Fields("�d����")
					strPnCurr = rsList.Fields("�i��")
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
							' �l
							fValue	= rtrim(rsList.Fields(i) & "")
							fType	= rsList.Fields(i).type
							select case rsList.Fields(i).name
							end select                                                                            
							' �ʒu��`�i�^�j                                                                            
							select case fType                                                                            
							Case 2		' ���l(Integer)                                                                            
								tdTag = "<TD nowrap id=""Integer"">"                                                                            
	'							if fValue = "-32768" then                                                                            
	'								fValue = ""                                                                            
	'							end if                                                                            
							Case 2 , 3 , 5 , 131	' ���l(Integer)                                                                            
								tdTag = "<TD nowrap id=""Integer"">"                                                                            
							Case 133		' ���t(Date)	                                                                            
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
							Case 129		' ������(Charactor)
								select case rsList.Fields(i).Name
								case else
									tdTag = "<TD nowrap align=""left"" id=""Charactor"">"
								end select
							Case else		' ���̑�
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
		cpTblBtn.value = "���ʂ��R�s�[";
	//--></SCRIPT>                                                                            
<% end if %>
</BODY>
</HTML>
