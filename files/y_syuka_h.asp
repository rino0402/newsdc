<% Option Explicit	%>
<% Response.Buffer = false %>
<% Response.Expires = -1 %>
<% 
	Server.ScriptTimeout = 900

	dim	versionStr

	versionStr = "2007.10.16 ���������F�i�ԂŌ����ł��Ȃ��s��C��"
	versionStr = "2007.10.29 �v���ĉ^�A�̖⍇��No�Ƀ����N��ǉ�"
	versionStr = "2007.11.02 ���ږ������������Ȃ����̂��C��"
	versionStr = "2010.02.24 �o�͌`���F�W�v�\(�֕ʑ��󖇐�) ����/�ꗗ�\ �ː��E�d�� �ǉ�"
	versionStr = "2010.02.25 ���������F�ː� �ǉ�"
	versionStr = "2010.03.03 �o�͌`���F�W�v�\(��)  �������u�����v���W�v�ł���悤�ɕύX"
	versionStr = "2010.04.07 �o�͌`���F�ꗗ�\ Tel/�X�֔ԍ� �ǉ�"
	versionStr = "2010.04.07 �o�͌`���F�W�v�\(��) ���N���b�N����ƁA�W�v�\(�֕�) ���I�������s��C��"
	versionStr = "2010.05.11 �o�͌`���F�W�v�\(�^����Е�) ���󖇐��FID_NO 7���ŏW�v"
	versionStr = "2010.08.20 �|�b�v�A�b�v���j���[�Ή�"
	versionStr = "2010.09.22 �o�͌`���F�ꗗ�\�F�W�񑗂��R�[�h �ǉ�"
	versionStr = "2011.05.06 �o�͌`���F�ꗗ�\(�@�ʏƍ�) �F�ǉ�(�@�ʒ����f�[�^�Əƍ�)"
	versionStr = "2011.06.06 �o�͌`���F�W�v�\(����/�i��No) �F�ǉ�"
	versionStr = "2012.10.11 ���������F�� �ǉ�"
	versionStr = "2012.10.29 ���������F����於 �ǉ�"
	versionStr = "2013.07.18 �o�͌`���F�ꗗ�\(�@�ʏƍ�)�F���ڒǉ� ����No. , �i��No."
	versionStr = "2014.07.30 ���������F�^����� �ǉ�"
	versionStr = "2016.05.31 �o�͌`���F�W�v�\(������)�F�z�B�s��(�Ή���...�� �k�C���`������)"
	versionStr = "2016.06.02 �o�͌`���F�W�v�\(������)�F�z�B�s��(�Ή���...�� �k�C���`�Ȗ،�)"
	versionStr = "2016.06.03 �o�͌`���F�W�v�\(������)�F�z�B�s��(�Ή���...�� �k�C���`�Ȗ،�) ����:���K�S���K�� �푍�s������"
	versionStr = "2016.06.06 �o�͌`���F�W�v�\(������)�F�z�B�s��(�Ή���...�� �k�C���`��ʌ�)"
	versionStr = GetVersion()
	versionStr = "2020.07.10 �N���b�v�{�[�h�o�͑Ή�(Chrome)"
	versionStr = "2020.07.11 �����ł��Ȃ��s��C��"
	versionStr = "2020.07.28 IE11�Ŕz�B�s���擪�ɕ\�������悤�ɏC��"
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
		delStr = "�폜��"
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
<TITLE><%=centerStr%> �o�ח\��</TITLE>
<!-- jdMenu head�p include �J�n -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="https://cdn.jsdelivr.net/clipboard.js/1.5.3/clipboard.min.js"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<script src="jquery.tablesorter.js" type="text/javascript"></script>
<script src="./clipboard.js" type="text/javascript"></script>
<!-- jdMenu head�p include �I�� -->
<SCRIPT LANGUAGE="JavaScript"><!--
navi = navigator.userAgent;

function DoCopy(arg){
	var doc = document.body.createTextRange();
	doc.moveToElementText(document.all(arg));
	doc.execCommand("copy");
	window.alert("�N���b�v�{�[�h�փR�s�[���܂����B\n�\��t���ł��܂��B" );
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
		window.alert("�����X�V" + autoBtn.value + "��Off");
//		DispMsg("�����X�V" + autoBtn.value + "��Off");
		autoBtn.value = "Off";
		autoValue.value = "On";
	} else {
		window.alert("�����X�V" + autoBtn.value + "��On");
//		DispMsg("�����X�V" + autoBtn.value + "��On");
		autoBtn.value = "On";
		autoValue.value = "Off";
	}
}

	function uKenpinClick() {
		if ( window.confirm("���iOK�ɂ��܂�") == false ) {
			ptypeChange("pTable");
		}
	}
	function DeleteClick() {
		if ( window.confirm("�f�[�^���폜���܂�\n�����ɖ߂��܂��񂪁���낵���ł����H") == false ) {
			ptypeChange("pTable");
		}
	}
--></SCRIPT>
</HEAD>
<BODY>
<!-- jdMenu body�p include �J�n -->
<!--#include file="jdmenu-sdc.asp" -->
<!-- jdMenu body�p include �I�� -->
<%
	if len(submitStr) + len(dtStr) = 0 then
		dtStr = year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2)
	end if
%>
  <FORM name="sqlForm">
  <div id="sqlDiv">
	<table id="sqlTbl">
		<caption style="text-align:left;"><%=centerStr%> �o�ח\�茟�� <%=delStr%></caption>
		<tr>
			<th>�`�[���t</th>
			<th>������</th>
			<th>ID-No</th>
			<th>�`�[No</th>
			<th>�i��</th>
			<th>�ː�</th>
			<th>�⍇��No</th>
			<th>��</th>
			<th>�L�����Z��</th>
			<th>�����</th>
			<!--th>�W�񑗂��R�[�h</th-->
			<!--th>����於</th-->
			<th>�^�����</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="dt" id="dt" VALUE="<%=dtStr%>" size="10" maxlength="8"><br>
				�`<br>
				<INPUT TYPE="text" NAME="dtTo" id="dtTo" VALUE="<%=dtToStr%>" size="10" maxlength="8">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="MUKE_CODE" id="MUKE_CODE" VALUE="<%=MUKE_CODEStr%>" size="14" maxlength="8"><br>
				<div class="note">
				�ϐ��S��<br>
				�ϐ���������<br>
				�ϐ������Ȃ�
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
				01:1��<br>
				02:2��
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="CANCEL_F" id="CANCEL_F" VALUE="<%=CANCEL_FStr%>" size="2"><br>
				<div class="note">
				0:�L�����Z������<br>
				1:�L�����Z���̂�
				</div>
			</td>
			<td align="center">
				<div><INPUT TYPE="text" NAME="OKURISAKI_CD" 		VALUE="<%=GetRequest("OKURISAKI_CD","")%>"		size="15" placeholder = "�����R�[�h"></div>
				<div><INPUT TYPE="text" NAME="COL_OKURISAKI_CD" 	VALUE="<%=GetRequest("COL_OKURISAKI_CD","")%>"	size="15" placeholder = "�W�񑗂��R�[�h"></div>
				<div><INPUT TYPE="text" NAME="MUKE_NAME" 		VALUE="<%=GetRequest("MUKE_NAME","")%>" 			size="15" placeholder = "����於"></div>
				<div><INPUT TYPE="text" NAME="JYUSHO"			VALUE="<%=GetRequest("JYUSHO","")%>"				size="15" placeholder = "�����Z��"></div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="UNSOU_KAISHA" id="UNSOU_KAISHA" VALUE="<%=GetRequest("UNSOU_KAISHA","")%>" size="12">
			</td>
		</tr>
		<tr>
			<td align="right">
				<b>�o�͌`���F</b>
			</td>
			<td colspan="10">
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable" id="pTable">
					<label for="pTable">�W�v�\(�֕�)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKoguchi" id="pTableKoguchi">
					<label for="pTableKoguchi">�W�v�\(��)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableOkuri" id="pTableOkuri">
					<label for="pTableOkuri">�W�v�\(�֕ʑ��󖇐�)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKaisya" id="pTableKaisya">
					<label for="pTableKaisya">�W�v�\(�^����Е�)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableColOkurisaki" id="pTableColOkurisaki">
					<label for="pTableColOkurisaki">�W�v�\(�W�񑗂���)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableOkurisaki" id="pTableOkurisaki">
					<label for="pTableOkurisaki">�W�v�\(������)</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnMonth" id="pTablePnMonth">
					<label for="pTablePnMonth">�W�v�\(�i��/����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableKenHin" id="pTableKenHin">
					<label for="pTableKenHin">�W�v�\(����/�i��No)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
					<label for="pList">�ꗗ�\</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListTei" id="pListTei">
					<label for="pListTei">�ꗗ�\(�@�ʏƍ�)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
					<label for="pListAll">�ꗗ�\(�S����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="dData" id="dData" onclick="DeleteClick();" disabled>
					<label for="dData">�폜</label>
			</td>
		</tr>
		<tr>
			<td align="right">
				<b>�ΏہF</b>
			</td>
			<td colspan="10">
				<INPUT TYPE="radio" NAME="tbl" VALUE="Y_SYUKA_H" id="Y_SYUKA_H">
					<label for="Y_SYUKA_H"><b>�o�ח\��(Y_SYUKA_H)</b></label>
				<INPUT TYPE="radio" NAME="tbl" VALUE="DEL_SYUKA_H" id="DEL_SYUKA_H">
					<label for="DEL_SYUKA_H"><b>�폜��(DEL_SYUKA_H)</b></label>
			</td>
		</tr>
	</table>
	<tr>
		<td>
		<INPUT TYPE="submit" value="����" id=submit1 name=submit1>
		<INPUT TYPE="reset" value="���Z�b�g" id=reset1 name=reset1 onClick="location.href='y_syuka_h.asp?tbl=<%=tblStr%>';">
				�ő匏���F<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=maxStr%>" size="8">
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
			value="������...ScriptTimeout=<%=Server.ScriptTimeout%>"
			id="cpTblBtn" disabled-->
		<button id="btnClip" class="btn" data-clipboard-target="#resultTbl" disabled onClick="DoCopy('resultDiv');">
			������...ScriptTimeout=<%=Server.ScriptTimeout%>
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
		case "pTable"	' �W�v�\
			sqlStr = sqlStr & " SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",CANCEL_F as ""�L�����Z��"""
			sqlStr = sqlStr & ",left(s.ID_NO,7) ""ID"""
			sqlStr = sqlStr & ",s.MUKE_CODE + ' ' + Mts.MUKE_NAME as ""������"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",count(distinct left(s.ID_NO,7)) as ""����"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""<br>���i��"""
'			sqlStr = sqlStr & ",(convert(floor(sum(if(KENPIN_NOW <> '',1,0)) / count(*) * 100),SQL_CHAR) + '%') as ""<br>��"""
			sqlStr = sqlStr & ",' ' as "" """
			sqlStr = sqlStr & ",sum(if(INS_BIN = '01',1,0)) as ""1��<br>����"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '01' and KENPIN_NOW <> '',1,0)) as ""1��<br>���i��"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '02',1,0)) as ""2��"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '02' and KENPIN_NOW <> '',1,0)) as ""<br>���i��"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '03',1,0)) as ""3��"""
			sqlStr = sqlStr & ",sum(if(INS_BIN = '03' and KENPIN_NOW <> '',1,0)) as ""<br>���i��"""
			sqlStr = sqlStr & ",sum(if(INS_BIN not in ('01','02','03'),1,0)) as ""���̑�"""
			sqlStr = sqlStr & ",sum(if(INS_BIN not in ('01','02','03') and KENPIN_NOW <> '',1,0)) as ""<br>���i��"""
			sqlStr = sqlStr & ",' ' as "" """
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & " left outer join Mts on s.MUKE_CODE = Mts.MUKE_CODE and Mts.SS_CODE = ''"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""�L�����Z��"",""ID"",""������"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""�L�����Z��"",""������"",""ID"""
		case "pTableOkuri"	' �W�v�\
			sqlStr = sqlStr & " SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",INS_BIN as ""��"""
			sqlStr = sqlStr & ",CANCEL_F as ""�L�����Z��"""
			sqlStr = sqlStr & ",count(distinct left(s.ID_NO,7)) as ""����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""���i��"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""��"",""�L�����Z��"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""��"",""�L�����Z��"""
		case "pTableKaisya"	' �W�v�\(�^����Е�)
			sqlStr = sqlStr & " SYUKA_YMD ""�o�ד�"""
			sqlStr = sqlStr & ",UNSOU_KAISHA ""�^�����"""
			sqlStr = sqlStr & ",count(distinct LEFT(ID_NO,7)) ""����"""
			sqlStr = sqlStr & ",count(*) ""����"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) ""���i��"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) ""����"""
'			sqlStr = sqlStr & ",sum(convert(KUTI_SU,SQL_DECIMAL)) ""����"""
'			sqlStr = sqlStr & ",sum(convert(SAI_SU,SQL_DECIMAL)) ""�ː�"""
'			sqlStr = sqlStr & ",sum(convert(JURYO,SQL_DECIMAL)) ""�d��"""
			sqlStr = sqlStr & " From " & tblStr & " s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""�^�����"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""�^�����"""
		case "pTableColOkurisaki"	' �W�v�\(�W�񑗂���)
			sqlStr = sqlStr & " SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",COL_OKURISAKI_CD as ""�W�񑗂��R�[�h"""
			sqlStr = sqlStr & ",if(COL_OKURISAKI_CD <> '',' ' + OKURISAKI,'') as ""����於"""
			sqlStr = sqlStr & ",convert(INS_BIN,SQL_DECIMAL) as ""��"""
			sqlStr = sqlStr & ",count(distinct if(OKURI_NO = '',null(),OKURI_NO)) as ""����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""���i��"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & ",max(convert(KUTI_SU,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & ",max(convert(SAI_SU,SQL_DECIMAL)) as ""�ː�"""
			sqlStr = sqlStr & ",max(convert(JURYO,SQL_DECIMAL)) as ""�d��"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""�W�񑗂��R�[�h"",""����於"",""��"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""�W�񑗂��R�[�h"",""��"""
		case "pTableOkurisaki"	' �W�v�\(������)
			sqlStr = sqlStr & " SYUKA_YMD ""�o�ד�"""
			sqlStr = sqlStr & ",COL_OKURISAKI_CD ""�W�񑗂��R�[�h"""
 			sqlStr = sqlStr & ",OKURISAKI_CD ""�����R�[�h"""
 			sqlStr = sqlStr & ",OKURISAKI ""�����"""
 			sqlStr = sqlStr & ",JYUSHO ""�����Z��"""
 			sqlStr = sqlStr & ",'' ""�z�B�s��"""
			sqlStr = sqlStr & ",UNSOU_KAISHA ""�^�����"""
			sqlStr = sqlStr & ",convert(INS_BIN,SQL_DECIMAL) as ""��"""
			sqlStr = sqlStr & ",count(distinct if(OKURI_NO = '',null(),OKURI_NO)) as ""����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""���i��"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & ",max(convert(KUTI_SU,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & ",max(convert(SAI_SU,SQL_DECIMAL)) as ""�ː�"""
			sqlStr = sqlStr & ",max(convert(JURYO,SQL_DECIMAL)) as ""�d��"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""�W�񑗂��R�[�h"",""�����R�[�h"",""�����"",""�����Z��"",""�z�B�s��"",""�^�����"",""��"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""�W�񑗂��R�[�h"",""�����R�[�h"",""��"""
		case "pTableKoguchi"	' �W�v�\
			sqlStr = sqlStr & " s.SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",UNSOU_KAISHA as ""�^�����"""
			sqlStr = sqlStr & ",OKURI_NO as ""�⍇��No"""
			sqlStr = sqlStr & ",left(ID_NO,7) as ""ID-No"""
			sqlStr = sqlStr & ",convert(KUTI_SU,SQL_DECIMAL) as ""����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & ",convert(SAI_SU,SQL_DECIMAL) as ""�ː�"""
			sqlStr = sqlStr & ",convert(JURYO,SQL_DECIMAL) as ""�d��"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""�^�����"",""�⍇��No"",""ID-No"",""����"",""�ː�"",""�d��"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""�^�����"",""�⍇��No"""
		case "pTablePnMonth"	' �W�v�\(�i�ԁ^���� �o�א���)
			sqlStr = sqlStr & " HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",left(s.SYUKA_YMD,6) as ""�o�הN��"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�i��"",""�o�הN��"""
			sqlStr = sqlStr & " order by ""�i��"",""�o�הN��"""
		case "pTableKenHin"	' �W�v�\(����/�i��No)
			sqlStr = sqlStr & " SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",INS_BIN as ""��"""
			sqlStr = sqlStr & ",SEK_KEN_NO as ""����No."""
			sqlStr = sqlStr & ",SEK_HIn_NO as ""�i��No."""
			sqlStr = sqlStr & ",CANCEL_F as ""�L�����Z��"""
			sqlStr = sqlStr & ",count(distinct left(s.ID_NO,7)) as ""����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(if(KENPIN_NOW <> '',1,0)) as ""���i��"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""��"",""����No."",""�i��No."",""�L�����Z��"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""��"",""����No."",""�i��No."",""�L�����Z��"""
		case "pList"	' �ꗗ�\
			sqlStr = sqlStr & " SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",convert(INS_BIN,SQL_DECIMAL) as ""��"""
			sqlStr = sqlStr & ",CANCEL_F as ""�L�����Z��"""
			sqlStr = sqlStr & ",URIDEN as ""���`"""
 			sqlStr = sqlStr & ",MUKE_CODE + ' ' + MUKE_NAME  as ""���Ӑ�"""
 			sqlStr = sqlStr & ",COL_OKURISAKI_CD as ""�W�񑗂��R�[�h"""
 			sqlStr = sqlStr & ",OKURISAKI_CD ""�����R�[�h"""
 			sqlStr = sqlStr & ",OKURISAKI as ""�����"""
 			sqlStr = sqlStr & ",JYUSHO as ""�����Z��"""
			sqlStr = sqlStr & ",ID_NO"
			sqlStr = sqlStr & ",DEN_NO as ""�`�[No"""
			sqlStr = sqlStr & ",ODER_NO as ""�I�[�_�[No"""
			sqlStr = sqlStr & ",HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",convert(SURYO,SQL_DECIMAL) as ""����"""
			sqlStr = sqlStr & ",BIKOU as ""���l"""
			sqlStr = sqlStr & ",TEL_No as ""Tel"""
			sqlStr = sqlStr & ",YUBIN_No as ""�X�֔ԍ�"""
			sqlStr = sqlStr & ",left(KENPIN_NOW,8) + '-' + left(right(KENPIN_NOW,6),4) as ""���i����"""
			sqlStr = sqlStr & ",KENPIN_TANTO_CODE as ""���i�S����"""
			sqlStr = sqlStr & ",UNSOU_KAISHA as ""�^�����"""
			sqlStr = sqlStr & ",OKURI_NO as ""�⍇��No"""
			sqlStr = sqlStr & ",convert(SEQ_NO,SQL_DECIMAL) as ""����No"""
			sqlStr = sqlStr & ",convert(KUTI_SU,SQL_DECIMAL) as ""����"""
			sqlStr = sqlStr & ",convert(SAI_SU,SQL_DECIMAL) as ""�ː�"""
			sqlStr = sqlStr & ",convert(JURYO,SQL_DECIMAL) as ""�d��"""
			sqlStr = sqlStr & ",INS_TANTO ""�o�^ID"""
			sqlStr = sqlStr & ",left(INS_DATETIME,8) + '-' +  right(INS_DATETIME,6) ""�o�^����"""
			sqlStr = sqlStr & ",UPD_TANTO ""�X�VID"""
			sqlStr = sqlStr & ",left(UPD_DATETIME,8) + '-' +  right(UPD_DATETIME,6) ""�X�V����"""
'			sqlStr = sqlStr & ",UPD_DATETIME"
'			sqlStr = sqlStr & ",left(UPD_DATETIME,8)"
'			sqlStr = sqlStr & ",left(ltrim(UPD_DATETIME),8)"
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""�o�ד�"",""��"",""�^�����"",ID_NO,""�⍇��No"",""�L�����Z��"",""����No"",""�����"""
		case "pListTei"	' �ꗗ�\(�@�ʏƍ�)
			sqlStr = sqlStr & " s.SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",convert(s.INS_BIN,SQL_DECIMAL) as ""��"""
			sqlStr = sqlStr & ",s.CANCEL_F as ""�L�����Z��"""
			sqlStr = sqlStr & ",s.URIDEN as ""���`"""
 			sqlStr = sqlStr & ",s.COL_OKURISAKI_CD as ""�W�񑗂��R�[�h"""
 			sqlStr = sqlStr & ",s.OKURISAKI as ""�����"""
 			sqlStr = sqlStr & ",s.MUKE_CODE + ' ' + MUKE_NAME  as ""���Ӑ�"""
			sqlStr = sqlStr & ",s.ID_NO"
			sqlStr = sqlStr & ",s.DEN_NO as ""�`�[No"""
			sqlStr = sqlStr & ",s.ODER_NO as ""�I�[�_�[No"""
			sqlStr = sqlStr & ",s.SEK_KEN_NO as ""����No."""
			sqlStr = sqlStr & ",s.SEK_HIn_NO as ""�i��No."""
			sqlStr = sqlStr & ",t.CHU_CD ""(�@��)������<br>���w�}��(��)"""
			sqlStr = sqlStr & ",t.THINB_CD ""���Ӑ�i��<br>���i��(��)"""
			sqlStr = sqlStr & ",t.HINB_CD ""�i��<br>���i��(��)"""
			sqlStr = sqlStr & ",s.HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",convert(ifnull(t.JUC_SUU,0),sql_numeric) ""(�@��)�󒍐���"""
			sqlStr = sqlStr & ",convert(s.SURYO,SQL_DECIMAL) as ""����"""
			sqlStr = sqlStr & ",t.SND_YMD + '-' + t.SND_HMS ""(�@��)�f�[�^�쐬����"""
			sqlStr = sqlStr & ",t.SYU_JUN ""(�@��)�o�׏���<br>���w�}��(���E��)"""
			sqlStr = sqlStr & ",t.TEI_NM ""(�@��)�@��<br>���w�}��(���E�E)"""
			sqlStr = sqlStr & ",t.TEI_LABELID ""�@�ʃ��x��ID"""
			sqlStr = sqlStr & ",t.KONPO_ID ""�W������ID"""
			sqlStr = sqlStr & ",s.BIKOU as ""���l"""
			sqlStr = sqlStr & ",s.TEL_No as ""Tel"""
			sqlStr = sqlStr & ",s.YUBIN_No as ""�X�֔ԍ�"""
			sqlStr = sqlStr & ",left(s.KENPIN_NOW,8) + '-' + left(right(s.KENPIN_NOW,6),4) as ""���i����"""
			sqlStr = sqlStr & ",s.KENPIN_TANTO_CODE as ""���i�S����"""
			sqlStr = sqlStr & ",s.UNSOU_KAISHA as ""�^�����"""
			sqlStr = sqlStr & ",s.OKURI_NO as ""�⍇��No"""
			sqlStr = sqlStr & ",convert(s.SEQ_NO,SQL_DECIMAL) as ""����No"""
			sqlStr = sqlStr & ",convert(s.KUTI_SU,SQL_DECIMAL) as ""����"""
			sqlStr = sqlStr & ",convert(s.SAI_SU,SQL_DECIMAL) as ""�ː�"""
			sqlStr = sqlStr & ",convert(s.JURYO,SQL_DECIMAL) as ""�d��"""
			sqlStr = sqlStr & " From " & tblStr & " as s"
'			sqlStr = sqlStr & " left outer join y_syuka_tei t on (t.TOK_CD = s.MUKE_CODE and t.CHU_CD = s.ODER_NO and t.HINB_CD = s.HIN_NO)"
			sqlStr = sqlStr & " left outer join y_syuka_tei t on (t.KEN_NO = s.SEK_KEN_NO and t.HIN_NO = s.SEK_HIN_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""�o�ד�"",""��"",""�^�����"",ID_NO,""�⍇��No"",""�L�����Z��"",""����No"",""�����"""
		case "pListAll"	' �ꗗ�\(�S����)
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From " & tblStr & " as s"
			sqlStr = sqlStr & whereStr
'			sqlStr = sqlStr & " order by SYUKA_YMD"
		case "dData"	' �폜
			sqlStr = "delete from " & tblStr
			sqlStr = sqlStr & whereStr
			set rsList = db.Execute(sqlStr)

			sqlStr = "select @@rowcount"
			sqlStr = sqlStr & " From " & tblStr
		end select
'		db.CommandTimeout=900
		set rsList = db.Execute(sqlStr)
	%>
		<caption style="text-align:left;"><%=now%> ����</caption>
		
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
					' �l
					fValue = rtrim(rsList.Fields(i))
					if rsList.Fields(i).Name = "�⍇��No" then
						tdTag = "<TD nowrap id=""Charactor"">"
						if rtrim(rsList.Fields("�^�����")) = "�v���ĉ^�A" then
							fValue = "<a href=""http://www4.kisc.co.jp/kurume-trans/kamotsu.asp?w_no=" & fValue & """>" & fValue & "</a>"
						end if
					elseif rsList.Fields(i).Name = "�ː�" or rsList.Fields(i).Name = "�d��" then
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
					elseif rsList.Fields(i).Name = "�����Z��" then
						tdTag = "<TD id=""Charactor"">"
					elseif rsList.Fields(i).Name = "�z�B�s��" then
						tdTag = "<TD id=""Charactor"">"
						fValue = GetHaitatsu(rsList.Fields("�����Z��") & " " & rsList.Fields("�����"))
						if inStr(fValue,"���c�J") > 0 then
							if cLng(rsList.Fields("�ː�")) < 10 then
								fValue = ""
							end if
						end if
					else
						' �ʒu��`�i�^�j
						select case rsList.Fields(i).type
						Case 2		' ���l(Integer)
							tdTag = "<TD nowrap id=""Integer"">"
							if fValue = "-32768" then
								fValue = ""
							end if
						Case 2 , 3 , 5 ,131	' ���l(Integer)
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
							tdTag = "<TD nowrap id=""Charactor"">"
						Case else		' ���̑�
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
		<!-- ���v -->
		<TR VALIGN='TOP'>
			<TD nowrap align="center">�v</TD>
			<%For i=0 To rsList.Fields.Count-1
				select case rsList.Fields(i).type
				Case 2 , 3 , 5 ,131	' ���l(Integer)
			%>
					<TD nowrap id="Integer"><%=formatnumber(totalArray(i),0,,,-1)%></TD>
			<%	Case else		' ���̑�	%>
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
		$('#btnClip').text('���ʂ��R�s�[');
//		cpTblBtn.disabled = false;
//		cpTblBtn.value = "���ʂ��R�s�[";
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
