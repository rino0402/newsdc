<% Option Explicit	%>
<% Response.Buffer = false %>
<% Response.Expires = -1 %>
<%
Function GetVersion()
' 2008.06.23 ���������F�����敪 ��ǉ�/�W�v�\(���i����������)"
' 2008.06.19 �o��OK���� ��ǉ� (�����敪=9 �� 0)"
' 2008.04.12 ���������F����敪 �ǉ�"
' 2008.02.07 �W�v�\(�q��) �q��2���ŏW�v����悤�ɕύX"
' 2008.01.11 G���Ə�(JGYOBA)�Ή�(�ꗗ�\�A��������)"
' 2007.12.26 �폜�ρ��\��̖߂��Ή�
' 2007.10.04 ���������F���i�敪�ǉ�(�����f�|�o�ז��בΉ�)"
' 2007.08.08 �W�v�\(�i��/�݌�)�E�������P�[�V�����̏o�א��ʁA���������������s��C��"
' 2007.08.09 �W�v�\(�i��/�݌�)�E�v���i�����̌v�Z����(���ʁ|���i����)�ɕύX"
' 2007.08.06 �W�v�\(�i��/�݌�) �쐬
' ver1.17 2007.07.07 ���ԕʌ������W�v�\�̑Ή�
' ver1.16 2007.05.31 ���iOK�����̑Ή�
' ver1.15 2006.08.07 ���ԕʏW�v�\�F�����敪�ǉ�/7��(7�`9��)10��(10�`11��)�ǉ�
' ver1.13 2006.07.08 �o�ɍς̌������@(-0)�\��"
' ver1.12 2006.07.08 �o�ɍς̌������@(0)�\��
' ver1.11 2006.07.03 Response.Buffer = false �ɕύX
' ver1.10 2006.06.28 �W�v�\(���ԕ�)�̎���敪�ǉ�,�����敪�̖��̕\��/���i�ς̕\�L�ǉ�
' ver1.09 2006.06.24 �W�v�\(���ԕ�) �� 12���`18���͓`�[���t�Ɠ����ꍇ�̂݃J�E���g����悤�ɕύX
' ver1.08 2006.06.24 12���𕪂��ďW�v
' ver1.07 2006.06.20 �`�[No/ID-NO �����Ή�
' ver1.06 2006.06.15 �W�v�\(�����敪��/�o�ד�) �̑Ή�
' ver1.05 2006.06.15 group by , ourder by ���C��
' ver1.04 2006.06.13 ���ږ����e�[�u���̍ŏI�s�ɂ��\��
' ver1.03 2006.06.13 ������}�X�^�[�Q�ƑΉ�
	GetVersion = "2008.07.02 �W�v�\(�i��/�݌�) �G���[�ɂȂ���̑Ώ�"
	GetVersion = "2008.07.08 ���iOK �G���[�ɂȂ���̑Ώ�( as y �ǋL)"
	GetVersion = "2008.07.17 �P���ݒ�(0:�P�����o�^/1:�P���o�^��)�̑Ή�"
	GetVersion = "<font color='red'>2008.07.18 �P���ݒ�(0:�P�����o�^/1:�P���o�^��)�̌����s��C��</font>"
	GetVersion = "<font color='red'>2008.07.19 �P���ݒ�(0:�P�����o�^/1:�P���o�^��)�̌����s��C��(���i���������� �ȊO�̏o�͌`���Ή�)</font>"
	GetVersion = "2008.10.28 �W�v�\(����)�̑Ή�"
	GetVersion = "2008.11.04 dbName �ϐ���"
	GetVersion = "2008.12.12 dbName �ϐ���"
	GetVersion = "2008.12.24 �������� ���i���� �ǉ�"
	GetVersion = "2009.02.06 �o�͌`���̃f�[�^����֘A���\��"
	GetVersion = "2009.02.24 �ꗗ�\�F�I�[�_�[No �ǉ�"
	GetVersion = "2009.02.24 �o�͌`���F�W�v�\(�I�[�_�[No) �ǉ�"
	GetVersion = "2009.04.21 �W�v�\(�q��)�F�q�ɖ���\��"
	GetVersion = "2009.05.08 �ꗗ�\�F�ڊǏ� �폜/���i���b�Z�[�W �ǉ�"
	GetVersion = "2009.06.17 �y�d�v�z�������� �P���ݒ�(0:�P�����o�^/1:�P���o�^��)�̒P���ݒ���Ŕ��f����悤�ɕύX"
	GetVersion = "2009.08.12 ������A�����������\���ɂ���悤���P(�������ʂ�傫���\�������)"
	GetVersion = "2009.10.07 �ꗗ�\�F�o�ɕ\��� �ǉ�"
	GetVersion = "2009.10.22 ���sSQL���\���ɕύX �����N�ŕ\��/��\���ؑ�"
	GetVersion = "2009.10.26 �W�v�\(�i��/�݌ɏW�v) �Ή��E�E�E�p���f�[�^��BU/PPSC�݌ɐ����ƍ�"
	GetVersion = "2009.11.05 �������� ���ƕ� �̕����w��Ή� ��F4,D"
	GetVersion = "2009.11.12 �W�v�\(�i�� �o�׋}��) �쐬"
	GetVersion = "2009.11.12 �W�v�\(�i�� �o�׋}��) �폜��(DEL_SYUKA) ���w�肵�ē��삷��悤�ɕύX"
	GetVersion = "2010.01.15 �W�v�\(����) �y�A�C�e���z�ǉ�"
	GetVersion = "2010.03.23 �W�v�\(�����敪:1500�O��) �Ή�(�ޗǃZ���^�[�q�ɔ����̈�)"
	GetVersion = "2010.04.07 �ꗗ�\�F�A�C�e��No �ǉ�"
	GetVersion = "2010.06.09 �W�v�\(�����)/�W�v�\(�����/15:00�O��)�F�ː� �ǉ�"
	GetVersion = "2010.06.10 �ː� �̕\���`��������2���ɕύX"
	GetVersion = "2010.06.10 �W�v�\(�i��/������) �� �i�� ��ǉ�"
	GetVersion = "2010.06.25 �W�v�\(���ԕ�) �� �ː� ��ǉ�"
	GetVersion = "2010.08.20 �|�b�v�A�b�v���j���[�Ή�"
	GetVersion = "2010.09.01 �W�v�\(�i��/�݌�)  �W�v�\(�i��/�݌ɏW�v) �݌ɐ��̕s��C��"
	GetVersion = "2011.04.10 �ꗗ�\ �X�V����,�������� �ǉ�"
	GetVersion = "2011.07.01 �����悪�}�X�^�[���o�^�̏ꍇ�ł��A�o�א�R�[�h��\������悤�ɏC��"
	GetVersion = "2011.08.30 �W�v�\(�i��/�݌�) �����敪(�����E�C�O)�^���j�b�g�敪��ǉ�"
	GetVersion = "2012.04.06 �������� �i��(�ΊO) �̑O����v����������"
	GetVersion = "<font color=red>2012.10.17 �W�v�\(�i��/�݌�) �̒I�Ԃ�W���I�ԂɕύX</font>"
	GetVersion = "2013.02.26 �W�v�\(���i����������) �� ���ƕ� ��ǉ�"
	GetVersion = "2015.06.29 ���̕ύX:�����恨�o�א�A�o�א�Œ��������������悤�ɕύX�A�ꗗ�\�̏o�א�ɊC�O�������ǉ�"
	GetVersion = "2015.06.29 �ǉ��F�ꗗ�\(�O������),�W�v�\(�O������)"
	GetVersion = "2016.10.03 ��������(���l)�ǉ�"
	GetVersion = "2016.10.13 �o�א於�F������}�X�^�[���Q�Ƃ��Ȃ��悤�ɕύX"
	GetVersion = "2016.10.14 �o�א�F�Y�@�����̏o�א於��\���^�����F�Y�@�������u1 �Y�@�v�\����"
	GetVersion = "2016.10.28 �ː� �̐��xUp�F�ː��e�[�u��(ItemSize)�Q��"
	GetVersion = "2016.10.31 �W�v�\(�i��/�݌�)�F�\�[�g���ύX(��/�I��/�i��)"
	GetVersion = "2017.01.27 �W�v�\(�����/��������)�F�Y�@�ƒ����̍��v���W�v"
	GetVersion = "2017.04.11 �Y�@ ������̌����Ή�"
	GetVersion = "2017.09.24 �o�ɕ\"
	GetVersion = "2019.06.05 ���������F�X�V����"
	GetVersion = "2019.11.01 ���iOK�����FJITU_SURYO��0���Z�b�g����悤�ɏC��"
	GetVersion = "2020.03.23 �N���b�v�{�[�h�R�s�[�Ή�(IE�ȊO)"
	GetVersion = "2020.07.21 ���ƕ��̌����s��C��"
	GetVersion = "2020.07.21 �N���b�v�{�[�h�R�s�[�͈͕ύX(�������������O)"
	GetVersion = "2020.08.18 �o�Ƀe�X�g�p"
End Function

'----------------------------------------------------------
'select from �e�[�u��
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
'select �t�B�[���h
'----------------------------------------------------------
Function GetSelect(byVal strName)
	GetSelect = ""
	select case strName
	case "�o�ד�"
		GetSelect = GetSelect & "KEY_SYUKA_YMD"
	case "���Ə�"
		GetSelect = GetSelect & "JGYOBA"
	case "���ƕ�","��"
		GetSelect = GetSelect & "y.JGYOBU"
	case "�݌�<br>���x"
		GetSelect = GetSelect & "y.SYUKO_SYUSI"
	case "ID","ID_BC"
		GetSelect = GetSelect & "y.KEY_ID_NO"
	case "�`�[No"
		GetSelect = GetSelect & "y.DEN_NO"
	case "�I�[�_�[No"
		GetSelect = GetSelect & "y.ODER_NO"
	case "�A�C�e��No"
		GetSelect = GetSelect & "y.ITEM_NO"
	case "�i��","�i��_BC"
		GetSelect = GetSelect & "y.KEY_HIN_NO"
	case "�i��"
		GetSelect = GetSelect & "y.HIN_NAME"
	case "����"
		GetSelect = GetSelect & "convert(y.SURYO,SQL_DECIMAL)"
	case "(�o�ɍ�)"
		GetSelect = GetSelect & "convert(y.JITU_SURYO,SQL_INTEGER)"
	case "����"
		GetSelect = GetSelect & "CHOKU_KBN + if(CHOKU_KBN='1',if(y.LK_MUKE_CODE = '00027768',' �Y�@',' ����'),'')"
	case "�o�א�"
		GetSelect = GetSelect & "if(ifnull(d.ChoCode,'')=''"
		GetSelect = GetSelect & " ,LK_MUKE_CODE + ' ' + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME)"
		GetSelect = GetSelect & " ,d.ChoCode + ' ' + d.ChoName)"
	case "�o�א�LK"
		GetSelect = GetSelect & " y.KEY_MUKE_CODE + ' ' + y.MUKE_NAME + '<br>' + y.LK_MUKE_CODE"
	case "�o�א�_BC"
		GetSelect = GetSelect & " y.LK_MUKE_CODE"
	case "�o�א搔<br>"
		GetSelect = GetSelect & "count(distinct if(ifnull(d.ChoCode,'')='',LK_MUKE_CODE + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME),d.ChoCode + d.ChoName))"
	case "�o�א� �����W�v"
'		GetSelect = GetSelect & "if(CHOKU_KBN<>'1' or ifnull(d.ChoCode,'')<>'',LK_MUKE_CODE + ' ' + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME),'���̑�')"
		GetSelect = GetSelect & "if(CHOKU_KBN = '1',if(LK_MUKE_CODE = '00027768','P�Y�@','���̑�'),LK_MUKE_CODE + ' ' + if(convert(m.DISPLAY_RANKING,sql_decimal)>0,m.MUKE_NAME,y.MUKE_NAME))"
	case "15���ȍ~","�f�[�^<br>��M����"
		GetSelect = GetSelect & "if(left(INS_NOW,10) < KEY_SYUKA_YMD + '15','','15���ȍ~')"
	case "���<br>�敪"
		GetSelect = GetSelect & "TORI_KBN"
		GetSelect = GetSelect & "+case TORI_KBN"
		GetSelect = GetSelect & " when '25' then ' ����'"
		GetSelect = GetSelect & " when '29' then ' �U�֏o��'"
		GetSelect = GetSelect & " when '19' then ' �U�֓���'"
		GetSelect = GetSelect & " end"
	case "����<br>�敪","�����敪"
		GetSelect = GetSelect & "KEY_CYU_KBN + if(KEY_CYU_KBN = '1',' ����',if(KEY_CYU_KBN = '2',' �ً}',if(KEY_CYU_KBN = '3',' ��[',if(KEY_CYU_KBN = 'E',' �f��',''))))"
	case "���i����"
		GetSelect = GetSelect & "KEPIN_KAIJYO"
	case "�o�׌���"
		GetSelect = GetSelect & "count(distinct y.KEY_ID_NO)"
	case "�o�א�"
		GetSelect = GetSelect & "sum(convert(y.SURYO,SQL_DECIMAL))"
	case "���i��(��)"
		GetSelect = GetSelect & "z.sumi_qty"
	case "���i��(��)"
		GetSelect = GetSelect & "z.mi_qty"
	case "�v���i����"
		GetSelect = GetSelect & "if(sum(convert(y.SURYO,SQL_DECIMAL)) >= z.sumi_qty,sum(convert(y.SURYO,SQL_DECIMAL)) - z.sumi_qty,0)"
	case "���������敪"
		GetSelect = GetSelect & "i.NAI_BUHIN + ' ' + case i.NAI_BUHIN when '1' then '�Ώ�' when '2' then '�Ő؈ē���'  when '3' then '�Ő�' when '3' then '�P�i���j�b�g' else ''  end"
	case "�C�O�����敪"
		GetSelect = GetSelect & "i.GAI_BUHIN + ' ' + case i.GAI_BUHIN when '1' then '�Ώ�' when '2' then '�Ő؈ē���'  when '3' then '�Ő�' when '3' then '�P�i���j�b�g' else ''  end"
	case "���j�b�g�敪"
		GetSelect = GetSelect & "i.UNIT_BUHIN + ' ' + case i.UNIT_BUHIN when '0' then '�P�i' when '1' then '���j�b�g�e'  when '2' then '���j�b�g�q' when '3' then '�P�i���j�b�g' else '' end"
	case "����","���v<br>"
		GetSelect = GetSelect & "count(*)"
	case "����"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '1',1,0))"
	case "����<br>���i��"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '1' and KENPIN_YMD <> '',1,0))"
		strName = "<br>���i��"
	case "�ً}"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '2',1,0))"
	case "�ً}<br>���i��"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '2' and KENPIN_YMD <> '',1,0))"
		strName = "<br>���i��"
	case "��["
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '3',1,0))"
	case "��[<br>���i��"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN = '3' and KENPIN_YMD <> '',1,0))"
		strName = "<br>���i��"
	case "���̑�"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN not in ('1','2','3'),1,0))"
	case "���̑�<br>���i��"
		GetSelect = GetSelect & "sum(if(KEY_CYU_KBN not in ('1','2','3') and KENPIN_YMD <> '',1,0))"
		strName = "<br>���i��"
	case "span1"
		GetSelect = GetSelect & "' '"
	case "�o��<br>��"
		GetSelect = GetSelect & "sum(if(KAN_YMD <> '',1,0))"
	case "�o��<br>�c"
		GetSelect = GetSelect & "sum(if(KAN_YMD <> '',0,1))"
	case "���i<br>��","<br>���i��"
		GetSelect = GetSelect & "sum(if(KENPIN_YMD <> '',1,0))"
	case "���i<br>�c"
		GetSelect = GetSelect & "sum(if(KENPIN_YMD <> '',0,1))"
	case "���M<br>�c"
		GetSelect = GetSelect & "sum(if(g.IDno is null or KEY_CYU_KBN = 'E' or RTrim(LK_SEQ_NO)<>'',0,1))"
	case "����<br>�c"
		GetSelect = GetSelect & "sum(if(g.IDno is null,0,1))"
	case "�O��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8),1,0))"
	case "09��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) <= '09',1,0))"
	case "10��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('10','11'),1,0))"
	case "12��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('12'),1,0))"
	case "13��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('13'),1,0))"
	case "14��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('14'),1,0))"
	case "15��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('15'),1,0))"
	case "16��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('16'),1,0))"
	case "17��"
		GetSelect = GetSelect & "sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) >= '17',1,0))"
	case "��"
		GetSelect = GetSelect & "sum(convert(SURYO,SQL_DECIMAL))"
	case "�ː�"
		GetSelect = GetSelect & "sum(iSize.Size * convert(y.SURYO,SQL_DECIMAL))"
	case ".�ː�"
		GetSelect = GetSelect & "iSize.Size * convert(y.SURYO,SQL_DECIMAL)"
	case "xx�ː�"
		GetSelect = GetSelect & "sum(convert(i.SAI_SU,SQL_DECIMAL)*convert(SURYO,SQL_DECIMAL))"
	case ".xx�ː�"
		GetSelect = GetSelect & "convert(i.SAI_SU,SQL_DECIMAL)*convert(y.SURYO,SQL_DECIMAL)"
		strName = "�ː�"
	case "���l1"
		GetSelect = GetSelect & "y.BIKOU1"
	case "���l2"
		GetSelect = GetSelect & "y.BIKOU2"
	case "�����敪"
		GetSelect = GetSelect & "y.KAN_KBN"
	case "�o�ɓ�"
		GetSelect = GetSelect & "y.KAN_YMD"
	case "�o�Ɏ���"
		GetSelect = GetSelect & "y.KAN_HMS"
	case "�o�ɓ���"
		GetSelect = GetSelect & "y.KAN_YMD + '-' + left(y.KAN_HMS,4)"
	case "���i����"
		GetSelect = GetSelect & "y.KENPIN_YMD + '-' + left(y.KENPIN_HMS,4)"
	case "���i�S����"
		GetSelect = GetSelect & "y.KENPIN_TANTO_CODE"
	case "���O�`�F�b�N"
		GetSelect = GetSelect & "y.LK_SEQ_NO"
	case "�ڊǏ�"
		GetSelect = GetSelect & "i.BIKOU_TANA"
	case "���i���b�Z�[�W"
		GetSelect = GetSelect & "i.INSP_MESSAGE"
	case "�W���I��"
		GetSelect = GetSelect & "rtrim(i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN)"
	case "�I��"
		GetSelect = GetSelect & "y.HTANABAN"
	case "�W���I��"
		GetSelect = GetSelect & "rtrim(i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN)"
	case "�I��_BC"
		GetSelect = GetSelect & "if(rtrim(i.ST_SOKO)<>'',rtrim(i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN),y.HTANABAN)"
	case "z�I��_BC"
		GetSelect = GetSelect & "'.' + z.Tana"
	case "�o�^����"
		GetSelect = GetSelect & "left(y.INS_NOW,8) + '-' + substring(y.INS_NOW,9,4)"
	case "�X�V����"
		GetSelect = GetSelect & "left(y.UPD_NOW,8) + '-' + substring(y.UPD_NOW,9,4)"
	case "�f�[�^�敪"
		GetSelect = GetSelect & "y.DATA_KBN"
	case "�̔��敪"
		GetSelect = GetSelect & "y.HAN_KBN"
	case "������"
		GetSelect = GetSelect & "y.LK_MUKE_CODE"
	case "�o�ɕ\���"
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
		delStr = "�폜��"
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
<TITLE><%=centerStr%> �o�ח\��</TITLE>
<!-- jdMenu head�p include �J�n -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="https://cdn.jsdelivr.net/clipboard.js/1.5.3/clipboard.min.js"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
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

	function uSyukoClick() {
		if ( window.confirm("�o��OK�ɂ��܂�") == false ) {
			ptypeChange("pTable");
		}
	}
	function uSyukoClickCancel() {
		if ( window.confirm("�o��OK�������ɂ��܂�") == false ) {
			ptypeChange("pTable");
		}
	}
	function uKenpinClick() {
		if ( window.confirm("���iOK�ɂ��܂�") == false ) {
			ptypeChange("pTable");
		}
	}
	function uKenpinCancelClick() {
		if ( window.confirm("���iOK���������܂�") == false ) {
			ptypeChange("pTable");
		}
	}
	function DeleteClick() {
		if ( window.confirm("�f�[�^���폜���܂�\n�����ɖ߂��܂��񂪁���낵���ł����H") == false ) {
			ptypeChange("pTable");
		}
	}
	function DelToYClick() {
		if ( window.confirm("�폜�σf�[�^��\��ɖ߂��܂�\n�����ɖ߂��܂��񂪁���낵���ł����H") == false ) {
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
<!-- jdMenu body�p include �J�n -->
<!--#include file="jdmenu-sdc.asp" -->
<!-- jdMenu body�p include �I�� -->
<%
	if len(submitStr) + len(dtStr) = 0 then
		dtStr = "today"
		select case centerStr
		case "����PC","����PC"
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
	<div><%=centerStr%> �o�ח\�茟�� <%=delStr%></div>
	<table id="sqlTbl">
		<tr>
			<th>�`�[���t</th>
			<th title="���Ə�R�[�h">���Ə�</th>
			<th>����敪</th>
			<th>DATA�敪</th>
			<th>�̔��敪</th>
			<th>�����敪</th>
			<th>���ƕ�</th>
			<th>����</th>
			<th>�o�א�</th>
			<th>ID-No</th>
			<th>�`�[No</th>
			<th>�i��(�ΊO)</th>
			<th>�����敪</th>
			<th>���i�敪</th>
			<th>�����Ώ�</th><!-- 2009.06.17 ���̕ύX -->
			<th>�P���ݒ�</th>
			<th>���i����</th>
			<th>�݌Ɏ��x</th>
			<!--th>SS<br>(�g���b�NNo)</th-->
			<th>���i���b�Z�[�W</th>
			<th>���l</th>
			<th>�o�^����</th>
			<th>�X�V����</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT TYPE="text" NAME="dt" id="dt" VALUE="<%=dtStr%>" size="10" maxlength="8"><br>
				�`<br>
				<INPUT TYPE="text" NAME="dtTo" id="dtTo" VALUE="<%=dtToStr%>" size="10" maxlength="8">
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="JGYOBA" id="JGYOBA" VALUE="<%=JGYOBAStr%>" size="10" maxlength="8" style="text-align:left;">
				<div style="text-align:left;">
					<font size="-2">00036003�FAP��CS</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="TORI_KBN" id="TORI_KBN" VALUE="<%=TORI_KBNStr%>" size="5" maxlength="2" style="text-align:left;">
				<div style="text-align:left;">
					<font size="-2">25:����</font><br>
					<font size="-2">29:�U�֏o��</font>
					<font size="-2">19:�U�֓���</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="DATA_KBN" id="DATA_KBN" VALUE="<%=DATA_KBNStr%>" size="4" maxlength="" style="text-align:center;">
				<div style="text-align:left;">
				<font size="-2">1�F����<br>3�F�U��<br>7�F�ȖڐU��</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="HAN_KBN" id="HAN_KBN" VALUE="<%=HAN_KBNStr%>" size="4" maxlength="3" style="text-align:center;">
				<div style="text-align:left;">
				<font size="-2">1�F����<br>2�F�A�o</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEY_CYU_KBN" id="KEY_CYU_KBN" VALUE="<%=KEY_CYU_KBNStr%>" size="8" maxl3ength="8" style="text-align:center;">
				<div style="text-align:left;">
				<font size="-2">1�F����<br>2�F�ً}<br>3�F��[<br>E�F�f��<br>1,2,3�F�f�Տ���</font>
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="JGYOBU" id="JGYOBU" VALUE="<%=JGYOBUStr%>" size="4" style="text-align:center;">
				<!--div style="text-align:left;">
				1:����ذBU<br>
				4:CABU<br>
				D:IHBU<br>
				7:�ذŰBU<br>
				A:����BU
				</div-->
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="TOK_KBN" id="TOK_KBN" VALUE="<%=TOK_KBNStr%>" size="1" maxlength="1" style="text-align:center;"><br>
				1�F����
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
				%�F�����܂�����<br>
				�@�� AZC81%
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KAN_KBN" id="KAN_KBN" VALUE="<%=KAN_KBNStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0�F���o��<br>
				9�F�o�ɍ�<br>
				=�F
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="KEN_KBN" id="KEN_KBN" VALUE="<%=KEN_KBNStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0�F�����i<br>
				9�F���i��
				</div>
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="SEI_KBN" id="SEI_KBN" VALUE="<%=SEI_KBNStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0�F�����ΏۊO<br>
				1�F�����Ώ�
				</div>
			</td>
			<td align="center">
				<INPUT TYPE="text" NAME="sTanka" id="sTanka" VALUE="<%=sTankaStr%>" size="1" maxlength="1"><br>
				<div style="text-align:left;"><font size="-2">
					0:�P�����o�^<br>
					1:�P���o�^��<br>
					2:�P��0�ȏ�<br>
				</font></div>
			</td>
			<td align="center" nowrap>
				<INPUT TYPE="text" NAME="KEPIN_KAIJYO" id="KEPIN_KAIJYO" VALUE="<%=KEPIN_KAIJYOStr%>" size="2" maxlength="2" style="text-align:center;"><br>
				<div style="text-align:left;"><font size="-2">
				0�F�ʏ����<br>
				1�F���i����
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
				-�F���i���b�Z�[�W����
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
				<b>�o�͌`���F</b>
			</td>
			<td colspan="21" nowrap>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable" id="pTable">
					<label for="pTable">�W�v�\(�����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableCnt" id="pTableCnt">
					<label for="pTableCnt">�W�v�\(�����/�����W�v)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable1500" id="pTable1500">
					<label for="pTable1500">�W�v�\(�����/15:00�O��)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable2" id="pTable2">
					<label for="pTable2">�W�v�\(���ԕ�)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable3" id="pTable3">
					<label for="pTable3">�W�v�\(�o�ד���)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnMonth" id="pTablePnMonth">
					<label for="pTablePnMonth">�W�v�\(�i��/����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSMonth" id="pTableSMonth">
					<label for="pTableSMonth">�W�v�\(�o�א�/����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableMonth" id="pTableMonth">
					<label for="pTableMonth">�W�v�\(����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableDay" id="pTableDay">
					<label for="pTableDay">�W�v�\(���ԕʌ�����)</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSoko" id="pTableSoko">
					<label for="pTableSoko">�W�v�\(�q��)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable4" id="pTable4">
					<label for="pTable4">�W�v�\(�i��/�o�א�)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnZaiko" id="pTablePnZaiko">
					<label for="pTablePnZaiko">�W�v�\(�i��/�݌�)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnSumzai" id="pTablePnSumzai">
					<label for="pTablePnSumzai">�W�v�\(�i��/�݌ɏW�v)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTable5" id="pTable5">
					<label for="pTable5">�W�v�\(���ʁ^�o�א��)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableDaySaki" id="pTableDaySaki">
					<label for="pTableDaySaki">�W�v�\(�o�ד��ʁ^�o�א��)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableDayTime" id="pTableDayTime">
					<label for="pTableDayTime">�W�v�\(�o�ד��E��M���E������)</label>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pList" id="pList">
					<label for="pList">�ꗗ�\</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableIri" id="pTableIri">
					<label for="pTableIri">�W�v�\(�O������)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListIri" id="pListIri">
					<label for="pListIri">�ꗗ�\(�O������)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pPicking" id="pPicking">
					<label for="pPicking">�o�ɕ\</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pListAll" id="pListAll">
					<label for="pListAll">�ꗗ�\(�S����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSeikyu" id="pTableSeikyu">
					<label for="pTableSeikyu">�W�v�\(���i����������)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableSeikyuKaigai" id="pTableSeikyuKaigai">
					<label for="pTableSeikyuKaigai">�W�v�\(���i����������/�����O�P��)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableShizai" id="pTableShizai">
					<label for="pTableShizai">�W�v�\(����)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTableOrderNo" id="pTableOrderNo">
					<label for="pTableOrderNo">�W�v�\(�I�[�_�[No)</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTablePnJumpUp" id="pTablePnJumpUp">
					<label for="pTablePnJumpUp">�W�v�\(�i�� �o�ב�)</label>
<% if adminStr = "admin" then %>
				<br>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pTest" id="pTest">
					<label for="pTest">�o�Ƀe�X�g</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="pPickingTest" id="pPickingTest">
					<label for="pPickingTest">���i�e�X�g</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uSyuko" id="uSyuko" onclick="uSyukoClick();">
					<label for="uSyuko">�o��OK</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uSyukoCancel" id="uSyukoCancel" onclick="uSyukoClickCancel();">
					<label for="uSyukoCancel">�o��OK����</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uKenpin" id="uKenpin" onclick="uKenpinClick();">
					<label for="uKenpin">���iOK</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uKenpinCancel" id="uKenpinCancel" onclick="uKenpinCancelClick();">
					<label for="uKenpinCancel">���iOK����</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uJizenCancel" id="uJizenCancel" onclick="uClick('���O�`�F�b�N�ς��������܂�');">
					<label for="uJizenCancel">���O�`�F�b�N����</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="uPrintCancel" id="uPrintCancel">
					<label for="uPrintCancel">�o�ɕ\OK����</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="dData" id="dData" onclick="DeleteClick();">
					<label for="dData">�폜</label>
				<INPUT TYPE="radio" NAME="ptype" VALUE="DelToY" id="DelToY" onclick="DelToYClick();">
					<label for="DelToY">�폜�ρ��\��߂�</label>
<% end if	%>
			</td>
		</tr>
		<tr>
			<td align="right">
				<b>�ΏہF</b>
			</td>
			<td colspan="21">
				<INPUT TYPE="radio" NAME="tbl" VALUE="Y_Syuka" id="Y_Syuka">
					<label for="Y_Syuka"><b>�o�ח\��(Y_Syuka)</b></label>
				<INPUT TYPE="radio" NAME="tbl" VALUE="DEL_SYUKA" id="DEL_SYUKA">
					<label for="DEL_SYUKA"><b>�폜��(DEL_SYUKA)</b></label>
			</td>
		</tr>
		<tr bordercolor="White">
			<td colspan="22">
			<INPUT TYPE="submit" value="����" id=submit1 name=submit1>
			<INPUT TYPE="reset" value="���Z�b�g" id=reset1 name=reset1 onClick="location.href='y_syuka.asp?tbl=<%=tblStr%>';">
				�ő匏���F<INPUT TYPE="text" NAME="max" id="max" VALUE="<%=maxStr%>" size="8">
			<%=GetVersion()%>
	<%		if len(submitStr) > 0 and ptypeStr = "pTable" and tblStr = "Y_Syuka" then	%>
				<!--span>�@�@�@�@�@�@�@�@�@�@�@�@�@�����X�V 
				<INPUT TYPE="text" NAME="auto" id="auto" VALUE="<%=autoStr%>" size="2" style="text-align : right;">
				��</span-->
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
			value="������...ScriptTimeout=<%=Server.ScriptTimeout%>"
			id="cpTblBtn" disabled-->
		<button id="btnClip" class="btn" data-clipboard-target="#resultTbl" disabled onClick="DoCopy('resultDiv');">
			������...ScriptTimeout=<%=Server.ScriptTimeout%>
		</button>
	</div>
	<div id='resultDiv'>
	<div><%=now%> ����</div>
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

		if len(JGYOBUStr) > 0 and False then	'2020.07.21 False�R�����g
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
			case "0"	' �����ΏۊO
				whereStr = whereStr & andStr & " ((HAN_KBN = '1' and LK_SEQ_NO = '') or (HAN_KBN = '2' and KAN_KBN <> '9'))"
			case "1"	' �����Ώ�
				whereStr = whereStr & andStr & " ((HAN_KBN = '1' and LK_SEQ_NO <> '') or (HAN_KBN = '2' and KAN_KBN = '9'))"
			end select
			andStr = " and"
		end if
		if len(KEPIN_KAIJYOStr) > 0 then
			select case KEPIN_KAIJYOStr
			case "0"	' �ʏ����
				whereStr = whereStr & andStr & " KEPIN_KAIJYO <> '1'"
			case "1"	' ���i����
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
			sqlStr = sqlStr & " " & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("���Ə�")
			sqlStr = sqlStr & "," & GetSelect("����")
			dim	strSyukaSaki
			strSyukaSaki = ""
			select case ptypeStr
			case "pTable"
				strSyukaSaki = "�o�א�"
				sqlStr = sqlStr & "," & GetSelect(strSyukaSaki)
			case "pTableCnt"
				strSyukaSaki = "�o�א� �����W�v"
				sqlStr = sqlStr & "," & GetSelect(strSyukaSaki)
				sqlStr = sqlStr & "," & GetSelect("�o�א搔<br>")
			end select
			sqlStr = sqlStr & "," & GetSelect("���v<br>")
			sqlStr = sqlStr & "," & GetSelect("<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("����<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("�ً}")
			sqlStr = sqlStr & "," & GetSelect("�ً}<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("��[")
			sqlStr = sqlStr & "," & GetSelect("��[<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("���̑�")
			sqlStr = sqlStr & "," & GetSelect("���̑�<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�ː�")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & ",""���Ə�"""
			sqlStr = sqlStr & ",""����"""
			sqlStr = sqlStr & ",""" & strSyukaSaki & """"
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & ",""���Ə�"""
			sqlStr = sqlStr & ",""����"""
			sqlStr = sqlStr & ",""" & strSyukaSaki & """"
		case "pTable1500"	' �W�v�\(�����敪/15:00�O��)
			sqlStr = sqlStr & " " & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("���Ə�")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("�o�א�")
			sqlStr = sqlStr & "," & GetSelect("15���ȍ~")
			sqlStr = sqlStr & "," & GetSelect("���v<br>")
			sqlStr = sqlStr & "," & GetSelect("<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("����<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("�ً}")
			sqlStr = sqlStr & "," & GetSelect("�ً}<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("��[")
			sqlStr = sqlStr & "," & GetSelect("��[<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("���̑�")
			sqlStr = sqlStr & "," & GetSelect("���̑�<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�ː�")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & ",""���Ə�"""
			sqlStr = sqlStr & ",""����"""
			sqlStr = sqlStr & ",""�o�א�"""
			sqlStr = sqlStr & ",""15���ȍ~"""
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & ",""15���ȍ~"""
			sqlStr = sqlStr & ",""���Ə�"""
			sqlStr = sqlStr & ",""����"""
			sqlStr = sqlStr & ",""�o�א�"""
		case "pTable3"	' �W�v�\(�o�ד���)
			sqlStr = sqlStr & " " & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("���v<br>")
			sqlStr = sqlStr & "," & GetSelect("<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("����<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("�ً}")
			sqlStr = sqlStr & "," & GetSelect("�ً}<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("��[")
			sqlStr = sqlStr & "," & GetSelect("��[<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("���̑�")
			sqlStr = sqlStr & "," & GetSelect("���̑�<br>���i��")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�ː�")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""�o�ד�"""
		case "pTable2"	' �W�v�\(���ԕ�) 
			sqlStr = sqlStr & " " & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("���<br>�敪")
			sqlStr = sqlStr & "," & GetSelect("����<br>�敪")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("�o�א�")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("�o��<br>��")
			sqlStr = sqlStr & "," & GetSelect("�o��<br>�c")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("���i<br>��")
			sqlStr = sqlStr & "," & GetSelect("���i<br>�c")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("���M<br>�c")
			sqlStr = sqlStr & "," & GetSelect("����<br>�c")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("�O��")
			sqlStr = sqlStr & "," & GetSelect("09��")
			sqlStr = sqlStr & "," & GetSelect("10��")
			sqlStr = sqlStr & "," & GetSelect("12��")
			sqlStr = sqlStr & "," & GetSelect("13��")
			sqlStr = sqlStr & "," & GetSelect("14��")
			sqlStr = sqlStr & "," & GetSelect("15��")
			sqlStr = sqlStr & "," & GetSelect("16��")
			sqlStr = sqlStr & "," & GetSelect("17��")
			sqlStr = sqlStr & "," & GetSelect("span1")
			sqlStr = sqlStr & "," & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�ː�")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("HMTAH015")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & ",""���<br>�敪"""
			sqlStr = sqlStr & ",""����<br>�敪"""
			sqlStr = sqlStr & ",""����"""
			sqlStr = sqlStr & ",""�o�א�"""
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & ",""���<br>�敪"""
			sqlStr = sqlStr & ",""����"""
			sqlStr = sqlStr & ",""����<br>�敪"""
			sqlStr = sqlStr & ",""�o�א�"""
		case "pTableDay"	' �W�v�\(���ԕʌ�����)
			sqlStr = "select KEY_SYUKA_YMD"
			sqlStr = sqlStr & ",CHOKU_KBN + if(CHOKU_KBN='1',' ����','') as ""����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) <= '18',1,0)) as ""�O��"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) >  '18',1,0)) as ""�O��<br>���"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('07','08','09','10','11'),1,0)) as ""AM"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('12','13','14','15'),1,0)) as ""15:00<br>�܂�"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('16','17','18'),1,0)) as ""15:00<br>�ȍ~"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as qty"
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) <= '18',convert(SURYO,SQL_DECIMAL),0)) as ""�O��"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD <> LEFT(INS_NOW,8) and substring(INS_NOW,9,2) >  '18',convert(SURYO,SQL_DECIMAL),0)) as ""�O��<br>���"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('07','08','09','10','11'),convert(SURYO,SQL_DECIMAL),0)) as ""AM"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('12','13','14','15'),convert(SURYO,SQL_DECIMAL),0)) as ""15:00<br>�܂�"""
			sqlStr = sqlStr & ",sum(if(KEY_SYUKA_YMD = LEFT(INS_NOW,8) and substring(INS_NOW,9,2) in ('16','17','18'),convert(SURYO,SQL_DECIMAL),0)) as ""15:00<br>�ȍ~"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by KEY_SYUKA_YMD,CHOKU_KBN"
			sqlStr = sqlStr & " order by KEY_SYUKA_YMD,CHOKU_KBN"
		case "pTable4"	' �W�v�\(�i��/�o�א�)
			sqlStr = sqlStr & " KEY_HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
			sqlStr = sqlStr & ",KEY_MUKE_CODE + ' ' + MUKE_NAME  as ""�o�א�"""
			sqlStr = sqlStr & ",KEPIN_KAIJYO  as ""���i����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & ",sum(convert(i.SAI_SU,SQL_DECIMAL)*convert(SURYO,SQL_DECIMAL)) as ""�ː�"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�i��"",""�i��"",""�o�א�"",""���i����"""
			sqlStr = sqlStr & " order by ""����"" desc,""�i��"""
		case "pTablePnZaiko"	' �W�v�\(�i��/�݌�)
			sqlStr = sqlStr & " " & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�i��")
			sqlStr = sqlStr & "," & GetSelect("�i��")
			sqlStr = sqlStr & "," & GetSelect("�W���I��")
'			sqlStr = sqlStr & ",i.GENSANKOKU"
			sqlStr = sqlStr & "," & GetSelect("�o�׌���")
			sqlStr = sqlStr & "," & GetSelect("�o�א�")
			sqlStr = sqlStr & "," & GetSelect("���i��(��)")
			sqlStr = sqlStr & "," & GetSelect("���i��(��)")
			sqlStr = sqlStr & "," & GetSelect("�v���i����")
			sqlStr = sqlStr & "," & GetSelect("�ː�")
			sqlStr = sqlStr & "," & GetSelect("���������敪")
			sqlStr = sqlStr & "," & GetSelect("�C�O�����敪")
			sqlStr = sqlStr & "," & GetSelect("���j�b�g�敪")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("Zaiko")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""��"",""�i��"",""�i��"",""�W���I��"",""���i��(��)"",""���i��(��)"",""���������敪"",""�C�O�����敪"",""���j�b�g�敪"""
			sqlStr = sqlStr & " order by ""��"",""�W���I��"",""�i��"""
		case "pTablePnSumzai"	' �W�v�\(�i��/�݌ɏW�v)
			sqlStr = sqlStr & " y.JGYOBU as ""���ƕ�"""
			sqlStr = sqlStr & ",y.KEY_HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
			sqlStr = sqlStr & ",count(distinct y.KEY_ID_NO) as ""�o�׌���"""
			sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_DECIMAL)) as ""�o�א�"""
			sqlStr = sqlStr & ",z.z_qty as ""POS�݌ɐ�"""
			sqlStr = sqlStr & ",convert(sz.bu_zai_qty,SQL_DECIMAL) as ""BU�݌ɐ�"""
			sqlStr = sqlStr & ",convert(sz.ppsc_zai_qty,SQL_DECIMAL) as ""PPSC�݌ɐ�"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join (select JGYOBU,NAIGAI,HIN_GAI,sum(convert(YUKO_Z_QTY,SQL_DECIMAL)) as z_qty from zaiko group by JGYOBU,NAIGAI,HIN_GAI) as z on ("
			sqlStr = sqlStr & "     z.JGYOBU = y.JGYOBU"
			sqlStr = sqlStr & " and z.NAIGAI = y.NAIGAI"
			sqlStr = sqlStr & " and z.HIN_GAI = y.KEY_HIN_NO)"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & " left outer join sumzai as sz on (y.jgyobu = sz.jgyobu and y.naigai = sz.naigai and y.key_hin_no = sz.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""���ƕ�"",""�i��"",""�i��"",""POS�݌ɐ�"",""BU�݌ɐ�"",""PPSC�݌ɐ�"""
			sqlStr = sqlStr & " order by ""���ƕ�"",""�i��"""
		case "pTableShizai"	' �W�v�\(����)
			sqlStr = sqlStr & " k.KO_HIN_GAI as ""���ޕi��"""
			sqlStr = sqlStr & ",si.HIN_NAME as ""���ޕi��"""
			sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_DECIMAL)*convert(k.KO_QTY,SQL_DECIMAL)) as ""����<br>����"""
			sqlStr = sqlStr & ",convert(si.G_ST_SHITAN,SQL_DECIMAL) as ""�d���P��"""
			sqlStr = sqlStr & ",round(sum(convert(y.SURYO,SQL_DECIMAL) * convert(k.KO_QTY,SQL_DECIMAL) * convert(si.G_ST_SHITAN,SQL_DECIMAL)),0) as ""�d�����z"""
			sqlStr = sqlStr & ",convert(si.G_ST_URITAN,SQL_DECIMAL) as ""�̔��P��"""
			sqlStr = sqlStr & ",round(sum(convert(y.SURYO,SQL_DECIMAL) * convert(k.KO_QTY,SQL_DECIMAL) * convert(si.G_ST_URITAN,SQL_DECIMAL)),0) as ""�̔����z"""
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
			sqlStr = sqlStr & " group by ""���ޕi��"",""���ޕi��"",""�d���P��"",""�̔��P��"""
			sqlStr = sqlStr & " order by ""���ޕi��"""
		case "DelToY"		' �폜�ρ��\����ǂ�
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

		case "pTableSeikyu"		' �W�v�\(���i����������)
			sqlStr = sqlStr & " y.JGYOBU as ""���ƕ�"""
			sqlStr = sqlStr & ",y.KEY_HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
			sqlStr = sqlStr & ",i.S_SEIKYU_F as ""�����敪"""
			sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_SET_DATE)='',-999999.99,convert(i.S_KOUSU_BAIKA ,SQL_DECIMAL)) as ""�H����"""
			sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_SET_DATE)='',-999999.99,convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""���ގd����"""
			sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_SET_DATE)='',-999999.99,convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""���ޔ̔���"""
			sqlStr = sqlStr & ",count(*) as ""�o�׌���"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""�o�א�"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)) as ""�H�����z"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""���ގd�����z"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""���ޔ̔����z"""
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""���ƕ�"",""�i��"",""�i��"",""�����敪"",""�H����"",""���ގd����"",""���ޔ̔���"""
			sqlStr = sqlStr & " order by ""���ƕ�"",""�i��"""
			db.CommandTimeout=900
		case "pTableMonth"		' �W�v�\(����)
			sqlStr = sqlStr & " left(KEY_SYUKA_YMD,6)  as ""�o�הN��"""
			sqlStr = sqlStr & ",count(*) as ""�o�׌���"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""�o�א�"""
			sqlStr = sqlStr & ",round(sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_KOUSU_BAIKA,SQL_DECIMAL)),0) as ""�H��"""
			sqlStr = sqlStr & ",round(sum(convert(SURYO,SQL_DECIMAL) * convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)),0) as ""����"""
			sqlStr = sqlStr & ",count(distinct KEY_HIN_NO) as ""�A�C�e��"""
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�הN��"""
			sqlStr = sqlStr & " order by ""�o�הN��"""
			db.CommandTimeout=900

		case "pTableSeikyuKaigai"	' �W�v�\(���i����������/�����O�P���Ή�)
			sqlStr = sqlStr & " y.JGYOBU as ""���ƕ�"""
			sqlStr = sqlStr & ",y.KEY_HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
			sqlStr = sqlStr & ",if(rtrim(ifnull(i.S_KOUSU_SET_DATE,'')) ='',-999999.99,convert(i.S_KOUSU_BAIKA ,SQL_DECIMAL)) as ""�H����"""
			sqlStr = sqlStr & ",if(rtrim(ifnull(i.S_KOUSU_SET_DATE,''))='',         0,convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""���ގd����"""
			sqlStr = sqlStr & ",if(rtrim(ifnull(i.S_KOUSU_SET_DATE,''))='',-999999.99,convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""���ޔ̔���"""
			sqlStr = sqlStr & ",count(*) as ""�o�׌���"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""�o�א�"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(ifnull(i.S_KOUSU_BAIKA,'0'),SQL_DECIMAL)) as ""�H�����z"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(ifnull(i.S_SHIZAI_GENKA,'0'),SQL_DECIMAL)) as ""���ގd�����z"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL) * convert(ifnull(i.S_SHIZAI_BAIKA,'0'),SQL_DECIMAL)) as ""���ޔ̔����z"""
			sqlStr = sqlStr & " from " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.HAN_KBN = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""���ƕ�"",""�i��"",""�i��"",""�H����"",""���ގd����"",""���ޔ̔���"""
			sqlStr = sqlStr & " order by ""���ƕ�"",""�i��"""
			db.CommandTimeout=900

		case "pTablePnMonth","pTableSMonth"	' �W�v�\(�i�ԁ^���� �o�א���) �W�v�\(�o�א�^���� �o�א���)
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
					sqlStr = sqlStr & " KEY_HIN_NO as ""�i��"""
					sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
'2009.10.06					sqlStr = sqlStr & ",k.KO_HIN_GAI as ""���ޕi��"""
'2009.10.06					sqlStr = sqlStr & ",k.KO_HIN_NAME as ""���ޕi��"""
					sqlStr = sqlStr & ",if(rtrim(i.S_KOUSU_BAIKA) ='',-999999.99,convert(i.S_KOUSU_BAIKA ,SQL_DECIMAL)) as ""�H����"""
					sqlStr = sqlStr & ",if(rtrim(i.S_SHIZAI_GENKA)='',         0,convert(i.S_SHIZAI_GENKA,SQL_DECIMAL)) as ""���ގd����"""
					sqlStr = sqlStr & ",if(rtrim(i.S_SHIZAI_BAIKA)='',-999999.99,convert(i.S_SHIZAI_BAIKA,SQL_DECIMAL)) as ""���ޔ̔���"""
					sqlStr = sqlStr & ",ifnull(z.sumi_qty,0) as ""���i��(��)"""
					sqlStr = sqlStr & ",ifnull(z.mi_qty,0) as ""���i��(��)"""
					sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""�o�א�"""
				'	sqlStr = sqlStr & ",sum(distinct if(z.GOODS_ON =  '0',convert(z.YUKO_Z_QTY,SQL_DECIMAL),0)) as ""���i��(��)"""
				'	sqlStr = sqlStr & ",sum(distinct if(z.GOODS_ON <> '0',convert(z.YUKO_Z_QTY,SQL_DECIMAL),0)) as ""���i��(��)"""
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
					groupByStr = " group by ""�i��"",""�i��"",""�H����"",""���ގd����"",""���ޔ̔���"",""���i��(��)"",""���i��(��)"""
					orderByStr = " order by ""�o�א�"" desc ,""�i��"""
				case "pTableSMonth"
					sqlStr = sqlStr & " y.LK_MUKE_CODE + ' ' + y.MUKE_NAME ""�o�א�"""	'" KEY_MUKE_CODE + ' ' + Mts.MUKE_NAME as ""�o�א�"""
					sqlStr = sqlStr & ",count(*) as ""�o�׌���"""
					sumStr = "1"
					fromStr = " from " & tblStr & " as y left outer join Mts on KEY_MUKE_CODE = Mts.MUKE_CODE and Mts.SS_CODE = ''"
					groupByStr = " group by ""�o�א�"""
					orderByStr = " order by ""�o�א�"""
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
		case "pTablePnJumpUp"	' �W�v�\(�i�� �o�׋}��)
			dim	workDay
			sqlStr2 = "select distinct"
			sqlStr2 = sqlStr2 & " count(distinct KEY_SYUKA_YMD) as wDay"
			sqlStr2 = sqlStr2 & " From del_syuka"
			sqlStr2 = sqlStr2 & " where KEY_SYUKA_YMD between '" & dtStr & "' and '" & dtToStr & "'"
			set rsList = db.Execute(sqlStr2)
			workDay = 0
			workDay = rsList.Fields("wDay")

			sqlStr = sqlStr & " y.KEY_HIN_NO as ""�i��"""
			sqlStr = sqlStr & ",i.HIN_NAME as ""�i��"""
			sqlStr = sqlStr & ",count(*) as ""�o�׌���(����)"""
			sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_DECIMAL)) as ""�o�א�(����)"""
			sqlStr = sqlStr & ",if(ifnull(d.qty,0) = 0,sum(convert(y.SURYO,SQL_DECIMAL))*10000,sum(convert(y.SURYO,SQL_DECIMAL))*100/(d.qty/" & workDay & ")) as ""������"""
			sqlStr = sqlStr & ",ifnull(d.qty/" & workDay & ",0) as ""�o�א�(����)"""
			sqlStr = sqlStr & ",ifnull(d.cnt,0) as ""�o�׌���(�ߋ�" & workDay & "��)"""
			sqlStr = sqlStr & ",ifnull(d.qty,0) as ""�o�א�(�ߋ�" & workDay & "��)"""
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
			sqlStr = sqlStr & " group by ""�i��"",""�i��"",""�o�׌���(�ߋ�" & workDay & "��)"",""�o�א�(�ߋ�" & workDay & "��)"",d.qty"
			sqlStr = sqlStr & " order by ""������"" desc"
			db.CommandTimeout=900
		case "pTable5"	' �W�v�\(���ʁ^�����)
			sqlStr = sqlStr & " KEY_MUKE_CODE as ""�o�א�"""
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
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by KEY_MUKE_CODE"
			sqlStr = sqlStr & " order by KEY_MUKE_CODE"
		case "pTableDaySaki"	' �W�v�\(�o�ד��ʁ^�����)
			db.CommandTimeout=900
			sqlStr = sqlStr & " KEY_SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',1,0)) as ""A1<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',convert(SURYO,SQL_DECIMAL),0)) as ""A1<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',1,0)) as ""A2<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',convert(SURYO,SQL_DECIMAL),0)) as ""A2<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',1,0)) as ""A3<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',convert(SURYO,SQL_DECIMAL),0)) as ""A3<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',1,0)) as ""A4<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',convert(SURYO,SQL_DECIMAL),0)) as ""A4<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',1,0)) as ""A5<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',convert(SURYO,SQL_DECIMAL),0)) as ""A5<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',1,0)) as ""A6<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',convert(SURYO,SQL_DECIMAL),0)) as ""A6<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',1,0)) as ""A7<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',convert(SURYO,SQL_DECIMAL),0)) as ""A7<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',1,0)) as ""A8<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',convert(SURYO,SQL_DECIMAL),0)) as ""A8<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),1,0)) as ""���̑�<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),convert(SURYO,SQL_DECIMAL),0)) as ""���̑�<br>����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"""
			sqlStr = sqlStr & " order by ""�o�ד�"""
		case "pTableDayTime"	' �W�v�\(�o�ד��^��M������)
			db.CommandTimeout=900
			sqlStr = sqlStr & " KEY_SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",left(INS_NOW,8) as ""��M��"""
			sqlStr = sqlStr & ",substring(INS_NOW,9,4) as ""��M����"""
			sqlStr = sqlStr & ",if(KEY_SYUKA_YMD = left(INS_NOW,8),substring(INS_NOW,9,2) + '��','�O��') as ""��������"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',1,0)) as ""A1<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A1',convert(SURYO,SQL_DECIMAL),0)) as ""A1<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',1,0)) as ""A2<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A2',convert(SURYO,SQL_DECIMAL),0)) as ""A2<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',1,0)) as ""A3<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A3',convert(SURYO,SQL_DECIMAL),0)) as ""A3<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',1,0)) as ""A4<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A4',convert(SURYO,SQL_DECIMAL),0)) as ""A4<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',1,0)) as ""A5<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A5',convert(SURYO,SQL_DECIMAL),0)) as ""A5<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',1,0)) as ""A6<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A6',convert(SURYO,SQL_DECIMAL),0)) as ""A6<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',1,0)) as ""A7<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A7',convert(SURYO,SQL_DECIMAL),0)) as ""A7<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',1,0)) as ""A8<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE = 'A8',convert(SURYO,SQL_DECIMAL),0)) as ""A8<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),1,0)) as ""���̑�<br>����"""
			sqlStr = sqlStr & ",sum(if(KEY_MUKE_CODE not in ('A1','A2','A3','A4','A5','A6','A7','A8'),convert(SURYO,SQL_DECIMAL),0)) as ""���̑�<br>����"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as ""����"""
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""��M��"",""��M����"",""��������"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""��M��"",""��M����"""
		case "pTableSoko"	' �W�v�\(�q�ɕ�)
'			sqlStr = sqlStr & " left(TANABAN1,2) as ""�q��"""
			sqlStr = sqlStr & " left(i.ST_SOKO,2) + ' ' + sk.soko_name as ""�q��"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',1,0)) as ""�o��<br>��"" "
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',0,1)) as ""�o��<br>�c"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',1,0)) as ""���i<br>��"" "
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',0,1)) as ""���i<br>�c"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as qty"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			sqlStr = sqlStr & " left outer join soko as sk on (i.st_soko = sk.soko_no)"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�q��"""
			sqlStr = sqlStr & " order by ""�q��"""
		case "pTableOrderNo"	' �W�v�\(�I�[�_�[No)
			sqlStr = sqlStr & " KEY_SYUKA_YMD as ""�o�ד�"""
			sqlStr = sqlStr & ",ODER_NO as ""�I�[�_�[No"""
			sqlStr = sqlStr & ",if(LK_MUKE_CODE <> '',LK_MUKE_CODE,rtrim(KEY_MUKE_CODE)) + ' ' + MUKE_NAME as ""�o�א�"""
			sqlStr = sqlStr & ",CYU_KBN_NAME as ""�����敪��"""
			sqlStr = sqlStr & ",count(*) as ""����"""
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',1,0)) as ""�o��<br>��"" "
			sqlStr = sqlStr & ",sum(if(KAN_YMD <> '',0,1)) as ""�o��<br>�c"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',1,0)) as ""���i<br>��"" "
			sqlStr = sqlStr & ",sum(if(KENPIN_YMD <> '',0,1)) as ""���i<br>�c"" "
			sqlStr = sqlStr & ",' ' as span1"
			sqlStr = sqlStr & ",sum(convert(SURYO,SQL_DECIMAL)) as qty"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " group by ""�o�ד�"",""�I�[�_�[No"",""�o�א�"",""�����敪��"""
			sqlStr = sqlStr & " order by ""�o�ד�"",""�I�[�_�[No"""
		case "pList"	' �ꗗ�\
			sqlStr = sqlStr & " " & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("���Ə�")
			sqlStr = sqlStr & "," & GetSelect("���<br>�敪")
			sqlStr = sqlStr & "," & GetSelect("����<br>�敪")
			sqlStr = sqlStr & "," & GetSelect("���i����")
			sqlStr = sqlStr & "," & GetSelect("�o�א�")
			sqlStr = sqlStr & "," & GetSelect("���ƕ�")
			sqlStr = sqlStr & "," & GetSelect("�݌�<br>���x")
			sqlStr = sqlStr & "," & GetSelect("ID")
			sqlStr = sqlStr & "," & GetSelect("�`�[No")
			sqlStr = sqlStr & "," & GetSelect("�I�[�_�[No")
			sqlStr = sqlStr & "," & GetSelect("�A�C�e��No")
			sqlStr = sqlStr & "," & GetSelect("�i��")
			sqlStr = sqlStr & "," & GetSelect("�i��")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("(�o�ɍ�)")
			sqlStr = sqlStr & "," & GetSelect(".�ː�")
			sqlStr = sqlStr & "," & GetSelect("���l1")
			sqlStr = sqlStr & "," & GetSelect("���l2")
			sqlStr = sqlStr & "," & GetSelect("�����敪")
			sqlStr = sqlStr & "," & GetSelect("�o�ɓ���")
			sqlStr = sqlStr & "," & GetSelect("���i����")
			sqlStr = sqlStr & "," & GetSelect("���i�S����")
			sqlStr = sqlStr & "," & GetSelect("���O�`�F�b�N")
			sqlStr = sqlStr & "," & GetSelect("�ڊǏ�")
			sqlStr = sqlStr & "," & GetSelect("���i���b�Z�[�W")
			sqlStr = sqlStr & "," & GetSelect("�I��")
			sqlStr = sqlStr & "," & GetSelect("�o�^����")
			sqlStr = sqlStr & "," & GetSelect("�X�V����")
'			sqlStr = sqlStr & "," & GetSelect("�f�[�^�敪")
'			sqlStr = sqlStr & "," & GetSelect("�̔��敪")
'			sqlStr = sqlStr & "," & GetSelect("������")
'			sqlStr = sqlStr & "," & GetSelect("�o�ɕ\���")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""�o�ד�"""
			sqlStr = sqlStr & ",""����<br>�敪"""
			sqlStr = sqlStr & ",""�o�א�"""
			sqlStr = sqlStr & ",""ID"""
		case "pTest"	' �o�Ƀe�X�g
			sqlStr = sqlStr & " " & GetSelect("z�I��_BC")
			sqlStr = sqlStr & "," & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�i��_BC")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("�����敪")
			sqlStr = sqlStr & "," & GetSelect("�o�א�")
			sqlStr = sqlStr & "," & GetSelect("�o�א�_BC")
			sqlStr = sqlStr & "," & GetSelect("(�o�ɍ�)")
			sqlStr = sqlStr & "," & GetSelect("�����敪")
			sqlStr = sqlStr & "," & GetSelect("ID_BC")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Zaiko.Tana")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""ID_BC"""
			sqlStr = sqlStr & ",""z�I��_BC"""
		case "pPickingTest"	' ���i�e�X�g
			sqlStr = sqlStr & " " & GetSelect("�I��_BC")
			sqlStr = sqlStr & "," & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�i��_BC")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("�o�א�LK")
			sqlStr = sqlStr & "," & GetSelect("�����敪")
			sqlStr = sqlStr & "," & GetSelect("(�o�ɍ�)")
			sqlStr = sqlStr & "," & GetSelect("�����敪")
			sqlStr = sqlStr & "," & GetSelect("ID_BC")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by"
			sqlStr = sqlStr & " ""ID_BC"""
			sqlStr = sqlStr & ",""�I��_BC"""
		case "pPicking"	' �o�ɕ\
			sqlStr = sqlStr & " " & GetSelect("�W���I��")
			sqlStr = sqlStr & "," & GetSelect("��")
			sqlStr = sqlStr & "," & GetSelect("�i��")
			sqlStr = sqlStr & "," & GetSelect("����")
			sqlStr = sqlStr & "," & GetSelect("�o�ד�")
			sqlStr = sqlStr & "," & GetSelect("�o�א�")
			sqlStr = sqlStr & "," & GetSelect("(�o�ɍ�)")
			sqlStr = sqlStr & "," & GetSelect("�����敪")
			sqlStr = sqlStr & "," & GetSelect("�o�ɓ���")
			sqlStr = sqlStr & "," & GetSelect("ID_BC")
			sqlStr = sqlStr & GetFrom(tblStr)
			sqlStr = sqlStr & GetFrom("Item")
			sqlStr = sqlStr & GetFrom("ItemSize")
			sqlStr = sqlStr & GetFrom("HtDrctId")
			sqlStr = sqlStr & GetFrom("Mts")
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by 1,2,3"
		case "pListIri","pTableIri"	' �ꗗ�\(�O������),�W�v�\(�O������)
			sqlStr = sqlStr & " KEY_MUKE_CODE + ' ' + MUKE_NAME ""�o�א�"""
			sqlStr = sqlStr & ",KEY_CYU_KBN + if(KEY_CYU_KBN = '1',' ����',if(KEY_CYU_KBN = '2',' �ً}',if(KEY_CYU_KBN = '3',' ��[',if(KEY_CYU_KBN = 'E',' �f��','')))) ""����<br>�敪"""
            if pTypeStr = "pListIri" then  ' �ꗗ�\(�O������)
    			sqlStr = sqlStr & ",y.JGYOBU ""��"""
    			sqlStr = sqlStr & ",KEPIN_KAIJYO ""���i<br>����"""
            end if
			sqlStr = sqlStr & ",KEY_SYUKA_YMD ""�o�ד�"""
            if pTypeStr = "pListIri" then  ' �ꗗ�\(�O������)
			    sqlStr = sqlStr & ",KEY_ID_NO ""ID"""
			    sqlStr = sqlStr & ",KEY_HIN_NO ""�i��"""
			    sqlStr = sqlStr & ",y.HIN_NAME ""�i��"""
			    sqlStr = sqlStr & ",convert(SURYO,SQL_DECIMAL) ""����"""
			    sqlStr = sqlStr & ",mod5(convert(SURYO,SQL_DECIMAL),g_qty_1,g_qty_2,g_qty_3,g_qty_4,g_qty_5) ""�[��"""
			    sqlStr = sqlStr & ",g_qty_1 ""����1"""
			    sqlStr = sqlStr & ",g_qty_2 ""����2"""
			    sqlStr = sqlStr & ",g_qty_3 ""����3"""
			    sqlStr = sqlStr & ",g_qty_4 ""����4"""
			    sqlStr = sqlStr & ",g_qty_5 ""����5"""
			    sqlStr = sqlStr & ",convert(i.SAI_SU,SQL_DECIMAL)*convert(SURYO,SQL_DECIMAL) as ""�ː�"""
			    sqlStr = sqlStr & ",BIKOU1 as ""���l1"""
			    sqlStr = sqlStr & ",BIKOU2 as ""���l2"""
            else                        ' �W�v�\(�O������)
			    sqlStr = sqlStr & ",count(*) ""����"""
			    sqlStr = sqlStr & ",sum(if(mod5(convert(SURYO,SQL_DECIMAL),g_qty_1,g_qty_2,g_qty_3,g_qty_4,g_qty_5)=0,1,0)) ""�[��<br>�O"""
			    sqlStr = sqlStr & ",sum(if(mod5(convert(SURYO,SQL_DECIMAL),g_qty_1,g_qty_2,g_qty_3,g_qty_4,g_qty_5)=0,0,1)) ""�[��<br>����"""
            end if
			sqlStr = sqlStr & " From " & tblStr & " y"
            if pTypeStr = "pListIri" then  ' �ꗗ�\(�O������)
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
		    case "����PC"
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
            if pTypeStr = "pListIri" then  ' �ꗗ�\(�O������)
			    sqlStr = sqlStr & " order by ""�o�ד�"""
			    sqlStr = sqlStr &          ",""����<br>�敪"""
			    sqlStr = sqlStr &          ",""�o�א�"""
			    sqlStr = sqlStr &          ",""��"""
			    sqlStr = sqlStr &          ",KEY_ID_NO"
            else  ' �W�v�\(�O������)
    			sqlStr = sqlStr & " group by"
			    sqlStr = sqlStr & " ""�o�ד�"""
			    sqlStr = sqlStr & ",""����<br>�敪"""
'			    sqlStr = sqlStr & ",""���i<br>����"""
			    sqlStr = sqlStr & ",""�o�א�"""
'			    sqlStr = sqlStr & ",""��"""
    			sqlStr = sqlStr & " order by"
			    sqlStr = sqlStr & " ""�o�א�"""
			    sqlStr = sqlStr & ",""����<br>�敪"""
'			    sqlStr = sqlStr & ",""��"""
'			    sqlStr = sqlStr & ",""���i<br>����"""
			    sqlStr = sqlStr & ",""�o�ד�"""
            end if
		case "pListAll"	' �ꗗ�\(�S����)
			sqlStr = sqlStr & " *"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			if instr(whereStr,"i.") > 0 then
				sqlStr = sqlStr & " left outer join item as i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
			end if
			sqlStr = sqlStr & whereStr
'			sqlStr = sqlStr & " order by KEY_SYUKA_YMD"
		case "uSyuko"	' �o��OK
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
'			if Msg("�w�肵�������S�āA���i�ςɂ��܂��B" & vbcrlf & sqlStr,vbYesNo) = vbYes then
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
		case "uSyukoCancel"	' �o��OK�L�����Z��
			sqlStr = "update " & tblStr & " as y"
			sqlStr = sqlStr & " set "
			sqlStr = sqlStr & "     KAN_KBN = '0'"
			sqlStr = sqlStr & "	   ,JITU_SURYO = '0'"
			sqlStr = sqlStr & whereStr
'			if Msg("�w�肵�������S�āA���i�ςɂ��܂��B" & vbcrlf & sqlStr,vbYesNo) = vbYes then
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
		case "uPrintCancel"	' �o�ɕ\OK�L�����Z��
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

		case "uKenpin"	' ���iOK
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
'			if Msg("�w�肵�������S�āA���i�ςɂ��܂��B" & vbcrlf & sqlStr,vbYesNo) = vbYes then
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
		case "uKenpinCancel"	' ���i�L�����Z��
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
'			if Msg("�w�肵�������S�āA���i�ςɂ��܂��B" & vbcrlf & sqlStr,vbYesNo) = vbYes then
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
		case "uJizenCancel"	' ���O�`�F�b�N�L�����Z��
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
			sqlStr = sqlStr & ",'���O�`�F�b�N�L�����Z��'"
			sqlStr = sqlStr & " From " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			sqlStr = sqlStr & " order by KEY_ID_NO"
		case "dData"	' �폜
			sqlStr = "delete from " & tblStr & " as y"
			sqlStr = sqlStr & whereStr
			set rsList = db.Execute(sqlStr)

			sqlStr = "select @@rowcount as ""�폜����"""
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
			    rem FILLER �͕\�����Ȃ�
				if left(rsList.Fields(i).Name,6) <> "_filler" then
					select case rsList.Fields(i).name
					case "KEY_SYUKA_YMD"	fname = "�`�[���t"
					case "KEY_MUKE_CODE"	fname = "�o�א�"
					case "saki"		fname = "�o�א�"
					case "allcnt"	fname = "���v<br>�@"
					case "kan"		fname = "<br>(��)"
					case "per"		fname = "<br>(%)"
					case "c1"		fname = "����<br>�@"
					case "c1kan"	fname = "<br>(��)"
					case "c2"		fname = "�X�|�b�g<br>�@"
					case "c2kan"	fname = "<br>(��)"
					case "c3"		fname = "��[<br>�@"
					case "c3kan"	fname = "<br>(��)"
					case "c9"		fname = "���̑�<br>�@"
					case "c9kan"	fname = "<br>(��)"
					case "choku"	fname = "����"
					case "qty"		fname = "��"

					case "useId"	fname = "�g�p�[��ID"
					case "usePrg"	fname = "�g�p���v���O����"
					case "jKbn"		fname = "���ƕ�<br>�敪"
					case "KAN_KBN"	fname = "����<br>�敪"
					case "KEY_CYU_KBN"	fname = "����<br>�敪"
					case "KEY_CYU_KBN"	fname = "����<br>�敪<br>(ν�)"
					case "mCode"	fname = "�o�א�"
					case "KEY_MUKE_CODE"	fname = "�o�א�<br>�Ǒ�"
					case "KEY_SYUKA_YMD"	fname = "�`�[���t"
					case "denNo"	fname = "�`�[��"
					case "ssNo"		fname = "�r�r�ǔ�"
					case "kNaiGai"	fname = "�����O"
					case "pn"		fname = "�i��"
					case "dataType"	fname = "�f�[�^<br>���"
					case "yoteiQty"	fname = "�\��<br>����"
					case "kakuQty"	fname = "�m��<br>����"
					case "textNo"	fname = "�e�L�X�g��"
					case "chokuKbn"	fname = "����<br>�敪"
					case "ioKbn"	fname = "���o��<br>�敪"
					case "rbKbn"	fname = "�ԍ�<br>�敪"
					case "denType"	fname = "�`�[<br>���"
					case "pnNai"	fname = "�i��(����)"
					case "pName"	fname = "�i��"
					case "mYosan"	fname = "�\�Z�P��<br>(��)"
					case "sYosan"	fname = "�\�Z�P��<br>(��)"
					case "hSoko"	fname = "�q��<br>(ν�)"
					case "hTana"	fname = "�I��<br>(ν�)"
					case "sCode"	fname = "�o�א�"
					case "sName"	fname = "�o�א於"
					case "kanDate"	fname = "�������t"
					case "kenDate"	fname = "���i���t"
					case "TOK_KBN"	fname = "����<br>�敪"
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
			        rem FILLER �͕\�����Ȃ�
					if left(rsList.Fields(i).Name,6) <> "_filler" then
						' �l
						fValue = rtrim(rsList.Fields(i))
						if rsList.Fields(i).Name = "per" then
							tdTag = "<TD nowrap id=""Integer"">"
						elseif right(rsList.Fields(i).Name,1) = "��" then
							tdTag = "<TD nowrap id=""Integer"">"
							if fValue < 0 then
								fValue = "���ݒ�"
							else
								fValue = formatnumber(fValue,2,,,-1)
							end if
						elseif right(rsList.Fields(i).Name,1) = "��" then
							tdTag = "<TD nowrap id=""Integer"">"
							fValue = formatnumber(fValue,0,,,0) & "%"
						elseif right(rsList.Fields(i).Name,4) = "(����)" then
							tdTag = "<TD nowrap id=""Integer"">"
							fValue = formatnumber(fValue,2,,,-1)
						elseif right(rsList.Fields(i).Name,2) = "���z" or right(rsList.Fields(i).Name,2) = "�ː�" then
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
						elseif rsList.Fields(i).Name = "�I��_BC" then
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
								elseif fValue <> "" then
									totalArray(i) = totalArray(i) + clng(fValue)
								end if
								if fValue = 0 then
									fValue = ""
								elseif fValue <> "" then
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
					end if
				Next
		    	Response.Write "</TR>"
				rsList.Movenext
			Loop
		%>
		<!-- ���v -->
		<TR VALIGN='TOP'>
			<TD nowrap align="center">�v</TD>
			<%For i=0 To rsList.Fields.Count-1
				select case rsList.Fields(i).type
				Case 2 , 3 , 5 ,131	' ���l(Integer)
					if right(rsList.Fields(i).Name,2) = "�ː�" then
			%>
						<TD nowrap id="Integer"><%=formatnumber(totalArray(i),2,,,-1)%></TD>
			<%
					else
			%>
						<TD nowrap id="Integer"><%=formatnumber(totalArray(i),0,,,-1)%></TD>
			<%		end if %>
			<%	Case else		' ���̑�	%>
					<TD></TD>
			<%	end select	%>
			<%Next%>
    	</TR>
		<%=strTh%>
	</TABLE></div>
	<hr>
	<a href="javascript:showhide('sql')" title="SQL�� �\��/��\��">SQL</a>
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
		$('#btnClip').text('���ʂ��R�s�[');
//		cpTblBtn.value = "���ʂ��R�s�[";
//		autoBtn.disabled = false;
	//-->
	</SCRIPT>
<% end if %>
</BODY>
</HTML>
