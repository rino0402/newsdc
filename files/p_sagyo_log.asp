<%
Option Explicit
Response.Buffer = false
Response.Expires = -1
' formatnumber()
const TristateTrue			= -1	'�[����\�����܂�
const TristateFalse			= 0		'�[����\�����܂���
const TristateUseDefault	= -2	'�u�n��̃v���p�e�B�v�̐ݒ�l���g�p���܂�

Function GetVersion()
	dim	strVersion
	' 2007.06.18 y_sagyo_log.asp
	' 2007.06.29 ��������:������ �̑Ή�
	strVersion = "2007.06.29 ��������:������ �̑Ή�"
	strVersion = "2007.09.07 ��������:ID_NO �̑Ή�"
	strVersion = "2007.09.12 ��������:������ �͈͎w�� �Ή�"
	strVersion = "2007.09.12 ��������:���ƕ� �Ή�"
	strVersion = "2007.09.12 �o�͌`���F�i�ԁ~���� �Ή�"
	strVersion = "2008.02.08 �o�͌`���F�q�ɕ� �Ή�"
	strVersion = "2008.02.09 �o�͌`���F�q�ɕ� �Ή� TimeOut�Ή�"
	strVersion = "2008.08.19 ��Ǝ��ԑΉ�"
	strVersion = "2008.09.26 �`�[����(�`�[ID�̐�)�Ή�"
	strVersion = "2008.10.08 ���ƕ�=S �̌����Ή�"
	strVersion = "2008.10.09 �^�C���A�E�g���Ԃ�10��(600s)�ɕύX"
	strVersion = "2009.12.24 �q�ɕ� �W���I�Ԃ���ł͂Ȃ��A�ړ����ŏW�v����悤�ɕύX"
	strVersion = "2009.02.22 �o�͌`���F�i�ԁE�I��(E4:�݌ɐ��� �����m�F�p) �ǉ�"
	strVersion = "2009.07.13 ���������F�ړ����^�ړ��� �ǉ�"
	strVersion = "2009.07.21 �o�͌`���F���j���[�E�������� �ǉ�"
	strVersion = "2009.08.26 �o�͌`���F�ꗗ�\ PRG_ID �ǉ�"
	strVersion = "2009.10.07 �o�͌`���F�o�ގЎ��� �ǉ�"
	strVersion = "2009.10.08 ���������F���ƕ� -S ���ޏ��� �̑Ή�"
	strVersion = "2009.10.08 ���������F�v�� �����w��(�J���} , �ŋ�؂�) �Ή�"
	strVersion = "2009.10.15 �o�͌`���F�ꗗ�\(�Γ�,�i��) �ǉ�"
	strVersion = "2009.11.09 �o�͌`���F������� �ǉ��^�������� ������ �擪�n�C�t��(-)��NOT�����Ή�"
	strVersion = "2009.11.25 �o�͌`���F�i�ԁE���o�ɉ� �Ή�"
	strVersion = "2010.03.03 ���������F������ �̕���������(8)������"
	strVersion = "2010.06.09 �o�͌`���F���j���[�ʁF�ړ�����(�݌Ɉړ���������) ��ǉ�"
	strVersion = "2010.08.11 �`�[����(�`�[ID�̐�) ����ID��1���ƃJ�E���g����悤�ɕύX"
	strVersion = "2010.08.20 �|�b�v�A�b�v���j���[�Ή�"
	strVersion = "2010.09.13 �ꗗ�\ �w����No. ���x��Check ���i�[Check �ǉ�/makeWhere�Ή�"
	strVersion = "2012.04.12 �o�͌`���F�i�ԁE���o�ɉ� �G���[�u�ق��s���ł��B�v�̑Ή�"
	strVersion = "2012.05.17 �o�͌`���F�i��(�G�A�R���ڊǗp)�̑Ή�"
	strVersion = "2013.07.23 ���g�Ή��FGetDbName()�ɕύX"
	GetVersion = "2016.08.10 ��������(�ǉ�)�[��ID�A�ϐ���GetRequest()�Ŏ擾�ɕύX"
	GetVersion = "2017.09.21 �o�͌`���F�i��(���z),��(���z)"
	GetVersion = "2017.09.22 �o�͌`���F�[��ID�E�v����"
	GetVersion = "2018.12.25 �o�͌`���F�ꗗ�\ ���ڒǉ�[Memo][�O��Check]"
	GetVersion = "2019.01.10 �o�͌`���F�ꗗ�\ ���ڒǉ�[JAN]"
	GetVersion = "2020.06.03 ���������F�������� �Ή�"
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
<TITLE><%=GetCenterName()%> ��ƃ��O</TITLE>
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
<!-- jdMenu body�p include �J�n -->
<!--#include file="jdmenu-sdc.asp" -->
<!-- jdMenu body�p include �I�� -->

  <FORM name="sqlForm">
  <div id="sqlDiv">
	<table id="sqlTbl">
		<caption style="text-align:left;">��ƃ��O����</caption>
		<tr>
			<th>���ƕ�</th>
			<th>������</th>
			<th>��������</th>
			<th>�S����</th>
			<th>���j���[No</th>
			<th>�v��</th>
			<th>�[��ID</th>
			<th>�i��</th>
			<th>������</th>
			<th>ID-No</th>
            <th>�w����No</th>
			<th>�ړ���</th>
			<th>�ړ���</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="JGYOBU" id="JGYOBU" VALUE="<%=GetRequest("JGYOBU","")%>" size="2" style="text-align:center;">
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="dt" id="dt" VALUE="<%=GetRequest("dt",GetToday())%>" size="10">
				�`
				<INPUT TYPE="text" NAME="dtTo" id="dtTo" VALUE="<%=GetRequest("dtTo","")%>" size="10">
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="tmFr" id="tmFr" VALUE="<%=GetRequest("tmFr", "")%>" size="5">
				�`
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
				<td><b>�o�͌`���F</b></td>
				<td>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable" id="pTable">
					<label for="pTable">�v��</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMenu" id="pTableMenu">
					<label for="pTableMenu">���j���[</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTantoMenu" id="pTableTantoMenu">
					<label for="pTableTantoMenu">�S���ҁE���j���[</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTantoYoin" id="pTableTantoYoin">
					<label for="pTableTantoYoin">�S���ҁE�v��</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTanto" id="pTableTanto">
					<label for="pTableTanto">�S����</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMenuYoin" id="pTableMenuYoin">
					<label for="pTableMenuYoin">���j���[�E�v��</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMenuDt" id="pTableMenuDt">
					<label for="pTableMenuDt">���j���[�E������</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableInOut" id="pTableInOut">
					<label for="pTableInOut">�i�ԁE���o�ɉ�</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSoko" id="pTableSoko">
					<label for="pTableSoko">�q��</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMukeCode" id="pTableMukeCode">
					<label for="pTableMukeCode">������</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnMonth" id="pTablePnMonth">
					<label for="pTablePnMonth">�i�ԁ~��</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnTana" id="pTablePnTana">
					<label for="pTablePnTana">�i�ԁE�I</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableTana" id="pTableTana">
					<label for="pTableTana">�I</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableWelID" id="pTableWelID">
					<label for="pTableWelID">�X�L���i�ŐV����</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableWork" id="pTableWork">
					<label for="pTableWork">�o�ގЎ���</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
					<label for="pList">�ꗗ�\</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListItem" id="pListItem">
					<label for="pListItem">�ꗗ�\(�Γ�,�i��)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
					<label for="pListAll">�ꗗ�\(�S����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePn" id="pTablePn">
					<label for="pTablePn">�i��(�G�A�R���ڊǗp)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnAmt" id="pTablePnAmt">
					<label for="pTablePnAmt">�i��(���z)</label><!--Amt(Amount:���z)-->
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableJAmt" id="pTableJAmt">
					<label for="pTableJAmt">��(���z)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableId" id="pTableId">
					<label for="pTableId">�[��ID�E�v��</label>
				</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr bordercolor=White>
			<td colspan="13" nowrap>
				<INPUT TYPE="submit" value="����" id=submit1 name=submit1>
				<INPUT TYPE="reset" value="���Z�b�g" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
				�ő匏���F<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=GetRequest("max","1000")%>" size="8" maxlength="6">
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
		case "pTable"	' �v���� �W�v�\
			sqlStr = sqlStr & " p.RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""�v��"""
			sqlStr = sqlStr & ",sum(if(p.RIRK_ID = 'ST' or p.RIRK_ID = 'EN',0,1)) ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) + convert(p.MI_JITU_QTY,SQL_decimal)) ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(p.WORK_TM,SQL_decimal))/60 ""��Ǝ���(��)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(p.ID_NO)= '',null(),rtrim(p.ID_NO))) ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�v��"""
			sqlStr = sqlStr & " order by ""�v��"""
		case "pTableTanto"	    ' �S���ҕ� �W�v�\
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""�S����"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""��Ǝ���(��)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�S����"""
			sqlStr = sqlStr & " order by ""�S����"""
		case "pTableWelID"	    ' �X�L���i�ŐV����
			sqlStr = sqlStr & " p.WEL_ID ""�[��ID"""
			sqlStr = sqlStr & ",p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') ""�S����"""
			sqlStr = sqlStr & ",JITU_DT + ' ' + JITU_TM ""��������"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",p.RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""�v��"""
			sqlStr = sqlStr & ",ID_NO ""ID-No."""
			sqlStr = sqlStr & ",HIN_GAI ""�i��"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC) ""����"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) ""����"""
			sqlStr = sqlStr & ",convert(MI_JITU_QTY,SQL_NUMERIC) ""����"""
			sqlStr = sqlStr & ",MUKE_CODE ""������"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN ""�ړ���"""
			sqlStr = sqlStr & ",TO_SOKO   + TO_RETU   + TO_REN   + TO_DAN   ""�ړ���"""
			sqlStr = sqlStr & ",convert(WORK_TM,SQL_NUMERIC) ""��Ǝ���(�b)"""
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
			sqlStr = sqlStr & " order by ""�[��ID"""
		case "pTableTantoYoin"	' �S���ҁE�v���� �W�v�\
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') ""�S����"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""�v��"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 ""��Ǝ���(��)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�S����"",""MENU"",""�v��"""
			sqlStr = sqlStr & " order by ""�S����"",""MENU"",""�v��"""
		case "pTableId"		' �[��ID�E�v����
			sqlStr = sqlStr & " p.WEL_ID ""�[��ID"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""�v��"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 ""��Ǝ���(��)"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
'			sqlStr = sqlStr & "  left outer join TANTO t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�[��ID"",""MENU"",""�v��"""
			sqlStr = sqlStr & " order by ""�[��ID"",""MENU"",""�v��"""
		case "pTableMenuYoin"	' ���j���[�E�v���� �W�v�\
			sqlStr = sqlStr & " p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""�v��"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""��Ǝ���(��)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""MENU"",""�v��"""
			sqlStr = sqlStr & " order by ""MENU"",""�v��"""
		case "pTableMenuDt"	' ���j���[�E�������� �W�v�\
			sqlStr = sqlStr & " p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",left(p.JITU_DT,8) as ""������"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""��Ǝ���(��)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""MENU"",""������"""
			sqlStr = sqlStr & " order by ""MENU"",""������"""
		case "pTableMenu"	    ' ���j���[�� �W�v�\
			sqlStr = sqlStr & " p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""��ƌ���"""
			sqlStr = sqlStr & ",sum(if(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)=0,0,1)) as ""�ړ�����"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""��Ǝ���(��)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""MENU"""
			sqlStr = sqlStr & " order by ""MENU"""
		case "pTableTantoMenu"	' �S���ҁE���j���[�� �W�v�\
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""�S����"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""��Ɛ���"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""��Ǝ���(��)"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""�`�[����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�S����"",""MENU"""
			sqlStr = sqlStr & " order by ""�S����"",""MENU"""
		case "pTableSoko"	' �q�ɕ� �W�v�\
			sqlStr = sqlStr & " RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""�v��"""
'			sqlStr = sqlStr & ",i.ST_Soko,'')  as ""�q��"""
			sqlStr = sqlStr & ",FROM_SOKO as ""�q��(��)"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""��Ɛ���"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""�`�[����"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""��Ǝ���(��)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
'			sqlStr = sqlStr & "  left outer join ITEM as i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�v��"",""�q��(��)"""
			sqlStr = sqlStr & " order by ""�v��"",""�q��(��)"""
		case "pTableMukeCode"	' ������� �W�v�\
			sqlStr = sqlStr & " MUKE_CODE as ""������"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""�v��"""
			sqlStr = sqlStr & ",sum(if(RIRK_ID = 'ST' or RIRK_ID = 'EN',0,1)) as ""��ƌ���"""
			sqlStr = sqlStr & ",sum(convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)) as ""��Ɛ���"""
'			sqlStr = sqlStr & ",sum(if(rtrim(ID_NO) = '',0,1)) as ""�`�[����"""
			sqlStr = sqlStr & ",count(distinct if(rtrim(ID_NO)= '',null(),rtrim(ID_NO))) as ""�`�[����"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC))/60 as ""��Ǝ���(��)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""������"",""�v��"""
			sqlStr = sqlStr & " order by ""������"",""�v��"""
		case "pTableInOut"		' �i�ԁE���o�ɉ�
			sqlStr = "select"
			sqlStr = sqlStr & " p.JGYOBU ""��"""
			sqlStr = sqlStr & ",p.HIN_GAI ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME ""�i��"""
			sqlStr = sqlStr & ",z.qty ""�݌ɐ�"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '7',1,0)) ""����<br>��"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '7',convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC),0)) ""����<br>��"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '4',1,0)) ""�o��<br>��"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '4',convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC),0)) ""�o��<br>��"""
			sqlStr = sqlStr & ",convert(i.G_ST_SHITAN,sql_numeric) ""�d����"""
			sqlStr = sqlStr & ",convert(i.G_ST_SHITAN,sql_numeric) * z.qty ""�݌ɋ��z"""
'			sqlStr = sqlStr & ",(year(now())*100+month(now()))-convert(left(max(p.JITU_DT),6),sql_numeric) as ""�s�ړ�<br>����"""
			sqlStr = sqlStr & ",datediff(month,convert(left(max(p.JITU_DT),4)+'-'+SUBSTRING(max(p.JITU_DT),5,2)+'-'+right(max(p.JITU_DT),2),sql_date),now()) ""�s�ړ�<br>����"""
			sqlStr = sqlStr & ",max(p.JITU_DT) ""�ŏI�ړ���"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join ITEM i on (i.JGYOBU = p.JGYOBU and i.NAIGAI = '1' and i.HIN_GAI = p.HIN_GAI)"
			sqlStr = sqlStr & "  left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(convert(YUKO_Z_QTY,sql_numeric)) qty from zaiko group by JGYOBU,NAIGAI,HIN_GAI) z on (z.JGYOBU = i.JGYOBU and z.NAIGAI = i.NAIGAI and z.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""��"",""�i��"",""�i��"",""�݌ɐ�"",""�d����"""
			sqlStr = sqlStr & " order by ""�o��<br>��"" desc, ""��"",""�i��"""

		case "pTablePnMonth"	' �i�ԁ~���� �W�v�\
			dim	sumStr
			dim	sqlStr2
			db.CommandTimeout		= 180	' 180

			sqlStr2 = "select distinct"
			sqlStr2 = sqlStr2 & " left(JITU_DT,6)  as ym"
			sqlStr2 = sqlStr2 & " From P_SAGYO_LOG as p"
			sqlStr2 = sqlStr2 & whereStr
			sqlStr2 = sqlStr2 & " order by ym"
			set rsList = db.Execute(sqlStr2)

			sqlStr = sqlStr & " HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + y.YOIN_DNAME as ""�v��"""
			sumStr = "convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC)"
			Do While Not rsList.EOF
				sqlStr = sqlStr & ",sum(if(left(JITU_DT,6) ='" & rsList.Fields("ym") & "'," & sumStr & ",0)) as """ & rsList.Fields("ym") & """"
				rsList.Movenext
			loop
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�i��"",""�v��"""
			sqlStr = sqlStr & " order by ""�i��"",""�v��"""

		case "pTablePn"		' �i��
			sqlStr = sqlStr & " p.HIN_GAI ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME ""�i��"""
'			sqlStr = sqlStr & ",i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN ""�W���I��"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '7',convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC),0)) ""�ړ���"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '1',convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC),0)) ""�����{"""
			sqlStr = sqlStr & ",sum(if(left(RIRK_ID,1) = '2',convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC),0)) ""������"""
			sqlStr = sqlStr & ",z.ac as ""AC�݌ɐ�"""
			sqlStr = sqlStr & ",z.az as ""AZ�݌ɐ�"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join ITEM i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & "  left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(if(Soko_No='AC',convert(YUKO_Z_QTY,sql_numeric),0)) as AC,sum(if(Soko_No='AZ',convert(YUKO_Z_QTY,sql_numeric),0)) as AZ from zaiko group by JGYOBU,NAIGAI,HIN_GAI) as z on (z.JGYOBU = i.JGYOBU and z.NAIGAI = i.NAIGAI and z.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�i��"",""�i��"",""AC�݌ɐ�"",""AZ�݌ɐ�"""
			sqlStr = sqlStr & " having ""�ړ���""<>0"
			sqlStr = sqlStr & " order by ""�i��"""
		case "pTablePnAmt"		' �i��(���z)
			sqlStr = sqlStr & " p.JGYOBU ""��"""
			sqlStr = sqlStr & ",p.HIN_GAI ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME ""�i��"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY,SQL_decimal)) ""����"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal)) ""����"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)) ""����<br>�H��"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) ""����<br>����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & " left outer join ITEM i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""��"",""�i��"",""�i��"""
			sqlStr = sqlStr & " order by ""��"",""�i��"""
		case "pTableJAmt"		'��(���z)
			sqlStr = sqlStr & " p.JGYOBU ""��"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY,SQL_decimal)) ""����"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal)) ""����"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)) ""����<br>�H��"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_decimal) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) ""����<br>����"""
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & " left outer join ITEM i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""��"""
			sqlStr = sqlStr & " order by ""��"""
		case "pTablePnTana"	' �i�ԁE�I�� �W�v�\
			sqlStr = sqlStr & " p.HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""�ړ���"""
			sqlStr = sqlStr & ",count(*) as ""��"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC)) as ""����"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC)) as ""����"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY  ,SQL_NUMERIC)) as ""����"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC)) as ""��Ǝ���(�b)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�i��"",""�ړ���"""
			sqlStr = sqlStr & " order by ""�i��"",""�ړ���"""
		case "pTableTana"	' �I�� �W�v�\
			sqlStr = sqlStr & " FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""�ړ���"""
			sqlStr = sqlStr & ",count(*) as ""��"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC) + convert(p.MI_JITU_QTY,SQL_NUMERIC)) as ""����"""
			sqlStr = sqlStr & ",sum(convert(p.SUMI_JITU_QTY,SQL_NUMERIC)) as ""����"""
			sqlStr = sqlStr & ",sum(convert(p.MI_JITU_QTY  ,SQL_NUMERIC)) as ""����"""
			sqlStr = sqlStr & ",sum(convert(WORK_TM,SQL_NUMERIC)) as ""��Ǝ���(�b)"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�ړ���"""
			sqlStr = sqlStr & " order by ""�ړ���"""
		case "pTableWork"	' �o�ގЎ���
			sqlStr = sqlStr & " p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""�S����"""
			sqlStr = sqlStr & ",JITU_DT as ""������"""
			sqlStr = sqlStr & ",min(left(JITU_TM,2)+':'+SUBSTRING(JITU_TM,3,2)) as ""�o��"""
			sqlStr = sqlStr & ",if(min(left(JITU_TM,4))<>max(left(JITU_TM,4)),max(left(JITU_TM,2)+':'+SUBSTRING(JITU_TM,3,2)),'') as ""�ގ�"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""�v��"""
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�S����"",""������"",""MENU"",""�v��"""
			sqlStr = sqlStr & " order by ""�S����"",""������"""
		case "pList"	' �ꗗ�\
			sqlStr = sqlStr & " JITU_DT + ' ' + JITU_TM ""��������"""
			sqlStr = sqlStr & ",p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') ""�S����"""
			sqlStr = sqlStr & ",WEL_ID ""WEL_ID"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') ""�v��"""
			sqlStr = sqlStr & ",ID_NO as ""ID-No."""
			sqlStr = sqlStr & ",HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",convert(MI_JITU_QTY,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",MUKE_CODE as ""������"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""�ړ���"""
			sqlStr = sqlStr & ",TO_SOKO   + TO_RETU   + TO_REN   + TO_DAN   as ""�ړ���"""
			sqlStr = sqlStr & ",p.Memo ""Memo"""
			sqlStr = sqlStr & ",SHIJI_No ""�w����No."""
			sqlStr = sqlStr & ",convert(p.HIN_CHECK_LABEL_CNT,SQL_decimal) ""���x��Check"""
			sqlStr = sqlStr & ",convert(p.HIN_CHECK_GENPIN_CNT,SQL_decimal) ""���i�[Check"""
			sqlStr = sqlStr & ",convert(p.HIN_CHECK_GAISOU_CNT,SQL_decimal) ""�O��Check"""
			sqlStr = sqlStr & ",JAN_CODE ""JAN"""
			sqlStr = sqlStr & ",convert(WORK_TM,SQL_NUMERIC) ""��Ǝ���(�b)"""
			sqlStr = sqlStr & ",PRG_ID"
			sqlStr = sqlStr & " From P_SAGYO_LOG p"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""��������"""
		case "pListItem"	' �ꗗ�\(�Γ��i��,�i��)
			sqlStr = sqlStr & " JITU_DT + ' ' + JITU_TM as ""��������"""
			sqlStr = sqlStr & ",p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') as ""�S����"""
			sqlStr = sqlStr & ",WEL_ID as ""WEL_ID"""
			sqlStr = sqlStr & ",p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') as ""MENU"""
			sqlStr = sqlStr & ",RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') as ""�v��"""
			sqlStr = sqlStr & ",ID_NO as ""ID-No."""
			sqlStr = sqlStr & ",p.HIN_GAI as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAI as ""�Γ��i��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",convert(MI_JITU_QTY,SQL_NUMERIC) as ""����"""
			sqlStr = sqlStr & ",MUKE_CODE as ""������"""
			sqlStr = sqlStr & ",FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN as ""�ړ���"""
			sqlStr = sqlStr & ",TO_SOKO   + TO_RETU   + TO_REN   + TO_DAN   as ""�ړ���"""
			sqlStr = sqlStr & ",convert(WORK_TM,SQL_NUMERIC) as ""��Ǝ���(�b)"""
			sqlStr = sqlStr & ",PRG_ID"
			sqlStr = sqlStr & " From P_SAGYO_LOG as p"
			sqlStr = sqlStr & "  left outer join ITEM as i on (p.JGYOBU = i.JGYOBU and p.NAIGAI = i.NAIGAI and p.HIN_GAI = i.HIN_GAI)"
			sqlStr = sqlStr & "  left outer join YOIN as y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))"
			sqlStr = sqlStr & "  left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)"
			sqlStr = sqlStr & "  left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by ""��������"""
		case "pListAll"	' �ꗗ�\(�S����)
			sqlStr = sqlStr & " * From P_SAGYO_LOG as p"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by JITU_DT,JITU_TM"
		end select
		db.CommandTimeout		= 600	' 180
		Server.ScriptTimeout	= 600	' 180
	%>
	<div>
		<INPUT TYPE="button" onClick="DoCopy('resultDiv')"
			value="������...ScriptTimeout=<%=Server.ScriptTimeout%>/CommandTimeout=<%=db.CommandTimeout%>"
			id="cpTblBtn" disabled>
	</div>
	<%
		dim rsList
		set rsList = db.Execute(sqlStr)
	%>

	<SCRIPT LANGUAGE=javascript><!--
		cpTblBtn.value = "�e�[�u���o�͒�...";
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
					' �l
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
					' �ʒu��`�i�^�j
					dim	tdTag
					if right(fName,1) = "��" then
						tdTag = "<TD nowrap id=""Integer"">"
						if isnull(fValue) = true then
							fValue = ""
						elseif fValue = "" then
							fValue = ""
						else
							fValue = formatnumber(fValue,2,,,-1)
						end if
					elseif right(fName,2) = "���z" then
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
						Case 2		' ���l(Integer)
							tdTag = "<TD nowrap id=""Integer"">"
							if fValue = "-32768" then
								fValue = ""
							end if
						Case 2 , 3 , 5	,131' ���l(Integer)
							if fName = "��Ǝ���(��)" then
								fValue = formatnumber(round(fValue,1),1,true,false,TristateTrue)
							else
								if fValue <> "" then
								    if fValue = 0 then
									    fValue = ""
									end if
							    end if
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
		cpTblBtn.value = "���ʂ��R�s�[";
	//--></SCRIPT>
<% end if %>
</BODY>
</HTML>
