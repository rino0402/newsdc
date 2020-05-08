Option Explicit
'-----------------------------------------------------------------------
'���C��
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
	'�g�p���@
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "hmem500.vbs [option] [filename]"
		Echo "Ex."
		Echo "cscript//nologo hmem500.vbs /db:newsdc1 hmem508szz.dat.20170804-133832.3397"
		Echo "cscript//nologo hmem500.vbs /db:newsdc4 hmem507szz.dat.20170922-185045.19583"
		Echo "cscript//nologo hmem500.vbs /db:newsdc4 hmem507szz.dat.20170922-185045.19583 /list:1 �����׃��X�g"
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
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
    Private Function Init()
		Debug ".Init()"
		dim	strArg
		Init = False
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Echo "�I�v�V�����G���[:" & strArg
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
				Echo "�I�v�V�����G���[:" & strArg
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
	'Run() ���s����
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
	'YNyuka() ���׃f�[�^�o�^
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
	'YNyukaInsert() ���׃f�[�^Insert
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
		AddSql " KAN_KBN			"'  1)	//�����敪
		AddSql ",DT_SYU				"'  1)	//�f�[�^���
		AddSql ",JGYOBU				"'  1)	//���ƕ��敪	key0
		AddSql ",NAIGAI				"'  1)	//�����O
		AddSql ",TEXT_NO			"'  9)	//�e�L�X�g��	key0
		AddSql ",JGYOBA				"'  8)	//���Ə꺰��
		AddSql ",DATA_KBN			"'  1)	//�f�[�^�敪
		AddSql ",TORI_KBN			"'  2)	//����敪
		AddSql ",ID_NO				"' 12)	//ID-NO
		AddSql ",KAIKEI_JGYOBA		"'  8)	//��v�p���Ə꺰��
		AddSql ",SHISAN_JGYOBA		"'  8)	//���Y�Ǘ��p���Ə꺰��
		AddSql ",HIN_NO				"' 20)	//�i�ڔԍ�
		AddSql ",DEN_NO				"' 10)	//�`�[�ԍ�
		AddSql ",SURYO				"'  7)	//�o�א���
		AddSql ",MUKE_CODE			"'  8)	//���Ӑ�R�[�h
		AddSql ",SYUKO_SYUSI		"'  8)	//�݌Ɏ��x
		AddSql ",SHISAN_SYUSI		"'  8)	//���Y�Ǘ��p�݌Ɏ��x����
		AddSql ",HOJYO_SYUSI		"'  8)	//�⏕�݌Ɏ��x����
		AddSql ",SYUKO_YMD			"'  8)	//�o�ɓ��t
		AddSql ",TANKA				"' 10)	//���ےP��
		AddSql ",ODER_NO			"' 12)	//�I�[�_�[�ԍ�
		AddSql ",ITEM_NO			"'  5)	//�A�C�e���ԍ�
		AddSql ",ODER_NO_R			"'  5)	//�����Ǘ��ԍ�����
		AddSql ",KOSO_KEITAI		"' 10)	//���`�Ժ���
		AddSql ",SYUKA_YMD			"'  8)	//�o�ח\���	key0
		AddSql ",TANABAN1			"' 10)	//۹����1
		AddSql ",TANABAN2			"' 10)	//۹����2
		AddSql ",TANABAN3			"' 10)	//۹����3
		AddSql ",MUKE_NAME			"' 24)	//���Ӑ於��
		AddSql ",CYU_KBN			"'  1)	//�����敪
		AddSql ",CYU_KBN_NAME		"' 10)	//�����敪����
		AddSql ",ORIGIN1			"' 10)	//���Y��1
		AddSql ",ORIGIN2			"' 10)	//���Y��2
		AddSql ",BIKOU2				"' 40)	//���l2
		AddSql ",HAN_KBN			"'  1)	//�̔��敪
		AddSql ",CHOKU_KBN			"'  1)	//�����w���敪
		AddSql ",UNIT_ID_NO			"' 12)	//�ƯďC���Ǘ��ԍ�
		AddSql ",ZAIKO_HIKIATE		"'  3)	//�݌Ɉ�������
		AddSql ",GOKON_KANRI_NO		"'  8)	//�����Ǘ��ԍ�
		AddSql ",JYUCHU_ZAN			"'  7)	//�󒍎c����
		AddSql ",KYOKYU_KBN			"'  1)	//�����敪
		AddSql ",SHOHIN_SYUSI		"'  8)	//���i���[�i�݌Ɏ��x����
		AddSql ",S_SHISAN_SYUSI		"'  8)	//���i���[�i���Y�Ǘ����x����
		AddSql ",S_HOJYO_SYUSI		"'  8)	//���i���[�i�⏕���x����
		AddSql ",BIKOU1				"' 40)	//���l1
		AddSql ",CHOHA_KBN			"'  1)	//���[�敪
		AddSql ",JYU_HIN_NO			"' 20)	//��t�i�ڔԍ�
		AddSql ",HIN_NAME			"' 20)	//�i��
		AddSql ",HIN_CHANGE_KBN		"'  1)	//�i�ڔԍ��ύX�敪
		AddSql ",MODULE_EXCHANGE	"'  1)	//Ӽޭ�ٌ����敪
		AddSql ",ZAIKO_SYUSI		"'  8)	//�c�݌ɂ܂Ƃߍ݌Ɏ��x����
		AddSql ",ZAN_SHISAN_SYUSI	"'  8)	//�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
		AddSql ",ZAN_HOJYO_SYUSI	"'  8)	//�c�݌ɂ܂Ƃߕ⏕���x����
		AddSql ",NOUKI_YMD			"'  8)	//�w��[��
		AddSql ",SERVICE_KANRI_NO	"'  9)	//���޽��ЊǗ��ԍ�
		AddSql ",KI_HIN_NO			"'  3)	//�@��i�ں���
		AddSql ",ENVIRONMENT_KBN	"'  1)	//����敔�i�敪
		AddSql ",SS_CODE			"'  8)	//��������溰��
		AddSql ",KEPIN_KAIJYO		"'  1)	//���i�����敪
		AddSql ",KAN_DT				"'  8)	//�������t
		AddSql ",BEF_NYU_QTY		"'  8)	//��s���א�
		AddSql ",YOSAN_FROM			"'  5)	//�\�Z�P�ʁi���j
		AddSql ",YOSAN_TO			"'  5)	//�\�Z�P�ʁi��j
		AddSql ",HTANABAN			"'  8)	//�W���I��
		AddSql ",HIN_NAI			"' 13)	//�i�ԁi�����j
		AddSql ",H_SOKO				"'  2)	//νđq�� 2006.10.17
		AddSql ",NYU_LIST_OUT		"'  1)	//���ɗ\��o���׸� 2007.06.12    ���ݖ��g�p 0:�f�[�^�o�͑Ώ� 9:�o�͍�(�������͏o�͑ΏۊO)
		AddSql ",GENSANKOKU			"' 20)	//���Y����
		AddSql ",GEN_GENSANKOKU		"' 20)	//�����\�����Y����
		AddSql ",SHIIRE_WORK_CENTER	"'  8)	//���ގd����ܰ�����
		AddSql ",KANKYO_KBN			"'  3)	//����ދ敪
		AddSql ",KANKYO_KBN_ST		"'  8)	//����ދ敪�K�p�J�n
		AddSql ",KANKYO_KBN_SURYO	"' 10)	//����ދ敪����
		AddSql ",ID_NO2				"' 12)	//ID_NO
		AddSql ",AITESAKI_CODE		"' 16)	//����溰��
		AddSql ",JYUCHU_YMD			"'  8)	//�󒍔N����
		AddSql ",SHITEI_NOUKI_YMD	"'  8)	//�w��[���N����
		AddSql ",LIST_OUT_END_F		"'  1)	//���Ɋ֘Aؽďo��F 0:�������Y�����i���ɊǗ�ؽĂ܂��͓��Ɂ^�I������ؽĂ��� '9:�������Y�����i���ɊǗ�ؽĂ����Ɂ^�I������ؽĂ�������
		AddSql ",LIST_NYU_KANRI_F	"'  1)	//���ɊǗ�ؽďo��F�u�������Y�����i���ɊǗ�ؽėp�v 0:����Ώ�(�����) 8:����ΏۊO�@9:�����	(0��9)
		AddSql ",LIST_NYU_CHECK_F	"'  1)	//��������ؽďo��F�u���Ɂ^�I������ؽėp�v�@0:����� 9:�����
		AddSql ",NYUKO_TANABAN		"'  8)	//���ɒI��
		AddSql ",MAEGARI_SURYO		"'  8)	//�O�ؑ��E��
		AddSql ",INS_TANTO			"'  5)	//�ǉ��@�S����
		AddSql ",Ins_DateTime		"' 14)	//�ǉ��@����  
		AddSql ",UPD_TANTO			"'  5)	//�X�V�@�S����
		AddSql ",UPD_DATETIME		"' 14)	//�X�V�@����  
		AddSql ",MOTO_PROG_ID		"'  8)	//�������v���O����
		AddSql ",MOTO_TEXT_NO		"'  9)	//���e�L�X�g��
		AddSql ",JITU_SURYO			"'  7)	//���ѐ���
		AddSql ") values ("
		AddSql " '0'"	'  1)	//�����敪
		AddSql ",'0'"	'  1)	//�f�[�^���
		AddSql ",'" & RTrim(objRs.Fields("JGyobu")) & "'"'	  1)	//���ƕ��敪	key0
		AddSql ",'1'"	'  1)	//�����O
		AddSql ",'" & RTrim(objRs.Fields("SyoriMD")) & RTrim(objRs.Fields("Bin")) & RTrim(objRs.Fields("SeqNo")) & "'"	'  9)	//�e�L�X�g��	key0
		AddSql ",''"	'  8)	//���Ə꺰��
		AddSql ",''"	'  1)	//�f�[�^�敪
		AddSql ",''"	'  2)	//����敪
		AddSql ",'" & RTrim(objRs.Fields("ID_NO")) & "'"	' 12)	//ID-NO
		AddSql ",''"	'  8)	//��v�p���Ə꺰��
		AddSql ",''"	'  8)	//���Y�Ǘ��p���Ə꺰��
		AddSql ",'" & RTrim(objRs.Fields("Pn")) & "'" 	' 20)	//�i�ڔԍ�
		AddSql ",'" & RTrim(objRs.Fields("DenNo")) & "'"	' 10)	//�`�[�ԍ�
		AddSql ",'" & RTrim(objRs.Fields("Qty")) & "'"	'  7)	//�o�א���
		AddSql ",''"	'  8)	//���Ӑ�R�[�h
		AddSql ",''"	'  8)	//�݌Ɏ��x
		AddSql ",''"	'  8)	//���Y�Ǘ��p�݌Ɏ��x����
		AddSql ",''"	'  8)	//�⏕�݌Ɏ��x����
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	'  8)	//�o�ɓ��t
		AddSql ",''"	' 10)	//���ےP��
		AddSql ",''"	' 12)	//�I�[�_�[�ԍ�
		AddSql ",''"	'  5)	//�A�C�e���ԍ�
		AddSql ",''"	'  5)	//�����Ǘ��ԍ�����
		AddSql ",''"	' 10)	//���`�Ժ���
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	'  8)	//�o�ח\���	key0
		AddSql ",''"	' 10)	//۹����1
		AddSql ",''"	' 10)	//۹����2
		AddSql ",''"	' 10)	//۹����3
		AddSql ",''"	' 24)	//���Ӑ於��
		AddSql ",''"	'  1)	//�����敪
		AddSql ",''"	' 10)	//�����敪����
		AddSql ",''"	' 10)	//���Y��1
		AddSql ",''"	' 10)	//���Y��2
		AddSql ",''"	' 40)	//���l2
		AddSql ",''"	'  1)	//�̔��敪
		AddSql ",''"	'  1)	//�����w���敪
		AddSql ",''"	' 12)	//�ƯďC���Ǘ��ԍ�
		AddSql ",''"	'  3)	//�݌Ɉ�������
		AddSql ",''"	'  8)	//�����Ǘ��ԍ�
		AddSql ",''"	'  7)	//�󒍎c����
		AddSql ",''"	'  1)	//�����敪
		AddSql ",''"	'  8)	//���i���[�i�݌Ɏ��x����
		AddSql ",''"	'  8)	//���i���[�i���Y�Ǘ����x����
		AddSql ",''"	'  8)	//���i���[�i�⏕���x����
		AddSql ",''"	' 40)	//���l1
		AddSql ",''"	'  1)	//���[�敪
		AddSql ",''"	' 20)	//��t�i�ڔԍ�
		AddSql2 ",'",RTrim(objRs.Fields("PName"))	' 20)	//�i��
		AddSql ",''"	'  1)	//�i�ڔԍ��ύX�敪
		AddSql ",''"	'  1)	//Ӽޭ�ٌ����敪
		AddSql ",''"	'  8)	//�c�݌ɂ܂Ƃߍ݌Ɏ��x����
		AddSql ",''"	'  8)	//�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
		AddSql ",''"	'  8)	//�c�݌ɂ܂Ƃߕ⏕���x����
		AddSql ",'" & RTrim(objRs.Fields("SHITEI_NOUKI_YMD")) & "'"	'  8)	//�w��[��
		AddSql ",''"	'  9)	//���޽��ЊǗ��ԍ�
		AddSql ",''"	'  3)	//�@��i�ں���
		AddSql ",''"	'  1)	//����敔�i�敪
		AddSql ",''"	'  8)	//��������溰��
		AddSql ",''"	'  1)	//���i�����敪
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	'  8)	//�������t
		AddSql ",''"	'  8)	//��s���א�
		AddSql ",'" & RTrim(objRs.Fields("SyukoCd")) & "'"	'  5)	//�\�Z�P�ʁi���j
		AddSql ",'" & RTrim(objRs.Fields("NyukoCd")) & "'"	'  5)	//�\�Z�P�ʁi��j
		AddSql ",'" & RTrim(objRs.Fields("Loc1")) & "'"		'  8)	//�W���I��
		AddSql ",'" & RTrim(objRs.Fields("PnNai")) & "'"		' 13)	//�i�ԁi�����j
		AddSql ",'" & RTrim(objRs.Fields("SyushiCd")) & "'"	'  2)	//νđq�� 2006.10.17
		AddSql ",''"	'  1)	//���ɗ\��o���׸� ���ݖ��g�p 0:�f�[�^�o�͑Ώ� 9:�o�͍�(�������͏o�͑ΏۊO)
		AddSql ",'" & RTrim(objRs.Fields("GENSANKOKU")) & "'"	' 20)	//���Y����
		AddSql ",'" & RTrim(objRs.Fields("GEN_GENSANKOKU")) & "'"	' 20)	//�����\�����Y����
		AddSql ",'" & RTrim(objRs.Fields("SHIIRE_WORK_CENTER")) & "'"	'  8)	//���ގd����ܰ�����
		AddSql ",'" & RTrim(objRs.Fields("KANKYO_KBN")) & "'"	'  3)	//����ދ敪
		AddSql ",'" & RTrim(objRs.Fields("KANKYO_KBN_ST")) & "'"	'  8)	//����ދ敪�K�p�J�n
		AddSql ",'" & RTrim(objRs.Fields("KANKYO_KBN_SURYO")) & "'"	' 10)	//����ދ敪����
		AddSql ",'" & RTrim(objRs.Fields("ID_NO")) & "'"			' 12)	//ID_NO
		AddSql ",'" & RTrim(objRs.Fields("AITESAKI_CODE")) & "'"	' 16)	//����溰��
		AddSql ",'" & RTrim(objRs.Fields("JYUCHU_YMD")) & "'"	'  8)	//�󒍔N����
		AddSql ",'" & RTrim(objRs.Fields("SHITEI_NOUKI_YMD")) & "'"	'  8)	//�w��[���N����
		AddSql ",'9'"	'  1)	//���Ɋ֘Aؽďo��F 0:�������Y�����i���ɊǗ�ؽĂ܂��͓��Ɂ^�I������ؽĂ��� '9:�������Y�����i���ɊǗ�ؽĂ����Ɂ^�I������ؽĂ�������
		AddSql ",'9'"	'  1)	//���ɊǗ�ؽďo��F�u�������Y�����i���ɊǗ�ؽėp�v 0:����Ώ�(�����) 8:����ΏۊO�@9:�����	(0��9)
		AddSql ",'9'"	'  1)	//��������ؽďo��F�u���Ɂ^�I������ؽėp�v�@0:����� 9:�����
		AddSql ",'" & strZaiko & "'"	'  8)	//���ɒI��
		AddSql ",''"	'  8)	//�O�ؑ��E��
		AddSql ",'HM500'"		'  5)	//�ǉ��S����
		AddSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"	' 14)	//�ǉ� ����
		AddSql ",''"	'  5)	//�X�V�S����
		AddSql ",''"	' 14)	//�X�V����  
		AddSql ",''"	'  8)	//�������v���O����
		AddSql ",''"	'  9)	//���e�L�X�g��
		AddSql ",''"	'  7)	//���ѐ���
		AddSql ")"
		Debug strSql
		dim	vRet
		vRet = Execute(strSql)
		select case vRet
		case 0
			Write ":ok:",0
			ZaikoInsert
		case -2147467259	'0x80004005 �d���L�[
			Write ":dup",0
		case else
			Write ":0x" & Hex(vRet),0
		end select
	End Function
	'-----------------------------------------------------------------------
	'List1() ���׃��X�g
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
	'Line1() ���׃��X�g �P�s�\��
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
			Write "��",0
			Write objRs.Fields("DenDt"),9
			WriteLine ""
			Write "��",0
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
	'GroupHead() �O���[�v�w�b�_�[
	'	True:�O���[�v�w�b�_�[
	'  Flase:�p���s
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
	'Item() �i�ڃ}�X�^�[�o�^
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
'		AddSql ",ifnull(i.Hin_Name,'*���o�^*') Hin_Name"
		AddSql "from hmem500 h"
		AddSql "left outer join item i on (i.JGyobu = h.JGyobu and i.NAIGAI='1' and i.HIN_GAI = h.Pn)"
		select case strItem
		case "n","y"	' �ǉ�
			AddSql "where i.HIN_GAI is null"
		case "u"		' �X�V
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
	'ItemInsert() Item�ǉ�
	'-----------------------------------------------------------------------
    Private Function ItemInsert()
		Debug ".ItemInsert()"
		AddSql ""
		AddSql "insert into Item"
		AddSql "("
		AddSql " JGYOBU"				'Char(  1) //���ƕ��敪
		AddSql ",NAIGAI"				'Char(  1) //�����O
		AddSql ",HIN_GAI"				'Char( 20) //�i�ԁi�O���j
		AddSql ",HIN_NAME"				'Char( 40) //�i��
		AddSql ",HIN_NAI"				'Char( 20) //�i�ԁi�����j
		AddSql ",GLICS1_TANA"			'Char( 10) //�O���b�N�X�I�ԂP   2005.05
		AddSql ",GLICS2_TANA"			'Char( 10) //�O���b�N�X�I�ԂQ   2005.05
		AddSql ",GLICS3_TANA"			'Char( 10) //�O���b�N�X�I�ԂR   2005.05
		AddSql ",L_HIN_NAME_E"			'Char( 30) //���i����   �i��
		AddSql ",L_KISHU1"				'Char( 25) //           �@��(1)
		AddSql ",L_KISHU3"				'Char(150) //           �@��(3)(���K�p�@���
		AddSql ",L_URIKIN1"				'Char( 10) //           ���i(1)	//NUMERICSA(10,0)
		AddSql ",L_URIKIN2"				'Char( 10) //           ���i(2)	//NUMERICSA(10,0)
		AddSql ",L_URIKIN3"				'Char( 10) //           ���i(3)	//NUMERICSA(10,0)
		AddSql ",UNIT_BUHIN"			'Char(  1) //�Ưĕ��i�敪       2006.07.28
		AddSql ",NAI_BUHIN"				'Char(  1) //�����������i�敪   2006.07.28
		AddSql ",GAI_BUHIN"				'Char(  1) //�C�O�������i�敪   2006.07.28
		AddSql ",HYO_TANKA"				'Char( 10) //�W���P��   2006.07.28
		AddSql ",KANKYO_KBN"			'Char(  3) //����ދ敪       2010.07.27
		AddSql ",KANKYO_KBN_ST"			'Char(  8) //����ދ敪�K�p�J�n 2010.07.
		AddSql ",KANKYO_KBN_SURYO"		'Char( 10) //����ދ敪����   2010.07.27
		AddSql ",CS_TANTO_CD"			'Char(  8) //CS�S������
		AddSql ",D_MODEL"				'Char(  8) //��\�@��i�ڃR�[�h PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",HINMOKU"				'Char(  8) //�i�ڃR�[�h         PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",K_KEITAI"				'Char( 14) //���`��(14��)     2012.03.13
		AddSql ",INS_TANTO"				'Char(  5) //�ǉ��@�S����
		AddSql ",Ins_DateTime"			'Char( 14) //�ǉ��@����  
		AddSql ",BIKOU20"
		AddSql ",L_PAPER"				' not null   //           ��
		AddSql ",L_PLASTIC"             ' not null   //           �v���X�`�b�N
		AddSql ",L_LABEL"               ' not null   //           �K�p�@������
		AddSql ")"
		AddSql "select top 1"
		AddSql " h.JGyobu"				'//���ƕ��敪
		AddSql ",'1'"					'//�����O
		AddSql ",p.Pn"					'Char( 20) //�i�ԁi�O���j
		AddSql ",p.PnBetsu"				'Char( 40) //�i��
		AddSql ",p.SPn"					'Char( 20) //�i�ԁi�����j
		AddSql ",p.Loc1"				'Char( 10) //�O���b�N�X�I�ԂP   2005.05
		AddSql ",p.Loc2"				'Char( 10) //�O���b�N�X�I�ԂQ   2005.05
		AddSql ",p.Loc3"				'Char( 10) //�O���b�N�X�I�ԂR   2005.05
		AddSql ",RTrim(p.PNameEngA)"	'Char( 30) //���i����   �i��
		AddSql ",p.NaiModel"			'Char( 25) //           �@��(1)
		AddSql ",p.GaiModel"			'Char(150) //           �@��(3)(���K�p�@���
		AddSql ",p.Tanka2"				'Char( 10) //           ���i(1)	//NUMERICSA(10,0)
		AddSql ",p.Tanka3"				'Char( 10) //           ���i(2)	//NUMERICSA(10,0)
		AddSql ",p.Tanka4"				'Char( 10) //           ���i(3)	//NUMERICSA(10,0)
		AddSql ",p.UnitKbn"				'Char(  1) //�Ưĕ��i�敪       2006.07.28
		AddSql ",p.NaiKbn"				'Char(  1) //�����������i�敪   2006.07.28
		AddSql ",p.GaiKbn"				'Char(  1) //�C�O�������i�敪   2006.07.28
		AddSql ",p.HyoTan"				'Char( 10) //�W���P��   2006.07.28
		AddSql ",h.KANKYO_KBN"			'Char(  3) //����ދ敪       2010.07.27
		AddSql ",h.KANKYO_KBN_ST"		'Char(  8) //����ދ敪�K�p�J�n 2010.07.
		AddSql ",h.KANKYO_KBN_SURYO"	'Char( 10) //����ދ敪����   2010.07.27
		AddSql ",p.KobaiTanto"			'Char(  8) //CS�S������
		AddSql ",p.DModel"				'Char(  8) //��\�@��i�ڃR�[�h PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",p.Hinmoku"				'Char(  8) //�i�ڃR�[�h         PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",p.KKeitai"				'Char( 14) //���`��(14��)     2012.03.13
		AddSql ",'HM500'"				'Char(  5) //�ǉ��@�S����
		AddSql ",left(replace(replace(replace(convert(Now(),sql_char),'-',''),':',''),' ',''),14)"	'Char( 14) //�ǉ��@����  
		AddSql ",case p.KobaiTanto"
		AddSql " when 'R101' then '����'"
		AddSql " when 'R102' then '����'"
		AddSql " when 'R103' then '���R'"
		AddSql " when 'R104' then '����'"
		AddSql " when 'R105' then '�쑺'"
		AddSql " when 'R106' then '����'"
		AddSql " else ''"
		AddSql " end"
		AddSql ",'0'"	'//           ��
		AddSql ",'0'"	'//           �v���X�`�b�N
		AddSql ",'0'"	'//           �K�p�@������
		AddSql "from hmem500 h"
		AddSql "inner join Pn5 p on (h.Pn = p.Pn)"
		AddWhere "h.Filename",RTrim(objRs.Fields("Filename"))
		AddWhere "h.Pn",RTrim(objRs.Fields("Pn"))
		Write ":" & Execute(strSql) ,0
	End Function
	'-----------------------------------------------------------------------
	'ItemUpdate() Item�X�V
	'	TORI_SHIIRE_WORK_CTR"	' //�d��ܰ��Z���^�[    
	'	KANKYO_KBN"				' //����ދ敪       
	'	KANKYO_KBN_ST"			' //����ދ敪�K�p�J�n
	'	KANKYO_KBN_SURYO"		' //����ދ敪����   
	'-----------------------------------------------------------------------
    Private Function ItemUpdate()
		Debug ".ItemUpdate()"
		dim	strSet
		strSet = ""
		strSet = SetSql(strSet," �dWC:","TORI_SHIIRE_WORK_CTR",RTrim(objRs.Fields("TORI_SHIIRE_WORK_CTR")),RTrim(objRs.Fields("SHIIRE_WORK_CENTER")))
		strSet = SetSql(strSet," ��:","KANKYO_KBN",RTrim(objRs.Fields("iKANKYO_KBN")),RTrim(objRs.Fields("hKANKYO_KBN")))
		strSet = SetSql(strSet," ���n:","KANKYO_KBN_ST",RTrim(objRs.Fields("iKANKYO_KBN_ST")),RTrim(objRs.Fields("hKANKYO_KBN_ST")))
		strSet = SetSql(strSet," ����:","KANKYO_KBN_SURYO",Trim(objRs.Fields("iKANKYO_KBN_SURYO")),Trim(objRs.Fields("hKANKYO_KBN_SURYO")))
		dim	strMsg
		if RTrim(objRs.Fields("hKANKYO_KBN")) = "LIT" then
'			strMsg = "���`�E���d�r����(" & Trim(objRs.Fields("hKANKYO_KBN_SURYO")) & ")"
'			strSet = SetSql(strSet," ���iMsg:","INSP_MESSAGE",Trim(objRs.Fields("INSP_MESSAGE")),strMsg)
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
				if strSrc = "�P������ ���`�E���d�r����" then
					strDst = strSrc
				end if
			end select
			if strDst <> strSrc then
				Write "��",0
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
	'Zaiko() �݌ɓo�^
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
	'ZaikoInsert() �݌Ƀf�[�^�o�^
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
		AddSql "	Soko_No				"' //�q�ɇ�
		AddSql ",	Retu				"' //�I�ԗ�
		AddSql ",	Ren					"' //�I�ԘA
		AddSql ",	Dan					"' //�I�Ԓi
		AddSql ",	JGYOBU				"' //���ƕ��敪
		AddSql ",	NAIGAI				"' //�����O
		AddSql ",	HIN_GAI				"' //�i�ԁi�O���j
		AddSql ",	GOODS_ON			"' //0:���i���� 1:�����i
		AddSql ",	NYUKA_DT			"' //���ד��t
		AddSql ",	NYUKO_DT			"' //���ɓ��t
		AddSql ",	HIN_NAI				"' //�i�ԁi�����j
		AddSql ",	YUKO_Z_QTY			"' //�L���݌ɐ�
		AddSql ",	LOCK_F				"' //�r���t���O
		AddSql ",	WEL_ID				"' //�g�p�q�@ID
		AddSql ",	PRG_ID				"' //�g�p���v���O����
		AddSql ",	GOODS_YMD			"' //���i�����t
		AddSql ",	SHIIRE_CODE			"' //�d���溰��
		AddSql ",	SHIIRE_TANKA		"' //�d���P��(9(8)V99)
		AddSql ",	KEIJYO_YM			"' //�v��N��
		AddSql ",	GENSANKOKU			"' //���Y����
		AddSql ",	SHIIRE_WORK_CENTER	"' //���ގd����ܰ�����
		AddSql ",	ID_NO2				"' //ID_NO
		AddSql ",	YOSAN_FROM			"' //�\�Z�P�ʁi���j
		AddSql ",	YOSAN_TO			"' //�\�Z�P�ʁi��j
		AddSql ") values ("
		AddSql " '" & mid(strZaiko,1,2) & "'"	' //�q�ɇ�
		AddSql ",'" & mid(strZaiko,3,2) & "'"' //�I�ԗ�
		AddSql ",'" & mid(strZaiko,5,2) & "'"' //�I�ԘA
		AddSql ",'" & mid(strZaiko,7,2) & "'"' //�I�Ԓi
		AddSql ",'" & RTrim(objRs.Fields("JGyobu")) & "'"	' //���ƕ��敪
		AddSql ",'1'"										' //�����O
		AddSql ",'" & RTrim(objRs.Fields("Pn")) & "'"		' //�i�ԁi�O���j
		AddSql ",'1'"										' //0:���i���� 1:�����i
		AddSql ",'" & RTrim(objRs.Fields("DenDt")) & "'"	' //���ד��t
		AddSql ",''"										' //���ɓ��t
		AddSql ",'" & RTrim(objRs.Fields("PnNai")) & "'"	' //�i�ԁi�����j
		AddSql ",'" & RTrim(objRs.Fields("Qty")) & "'"		' //�L���݌ɐ�
		AddSql ",'0'"										' //�r���t���O
		AddSql ",''"										' //�g�p�q�@ID
		AddSql ",''"										' //�g�p���v���O����
		AddSql ",''"										' //���i�����t
		AddSql ",''"										' //�d���溰��
		AddSql ",''"										' //�d���P��(9(8)V99)
		AddSql ",''"										' //�v��N��
		AddSql ",'" & RTrim(objRs.Fields("GENSANKOKU")) & "'"			' //���Y����
		AddSql ",'" & RTrim(objRs.Fields("SHIIRE_WORK_CENTER")) & "'"	' //���ގd����ܰ�����
		AddSql ",'" & RTrim(objRs.Fields("ID_NO")) & "'"	' //ID_NO
		AddSql ",'" & RTrim(objRs.Fields("SyukoCd")) & "'"	' //�U�֌�(�\�Z�P��)
		AddSql ",'" & RTrim(objRs.Fields("SyushiCd")) & "'"	' //�U�֐�(�݌Ɏ��x)
		AddSql ")"
		Debug strSql
		dim	vRet
		vRet = Execute(strSql)
		select case vRet
		case 0
			Write ":ok",0
			ZaikoInsert = 1
		case -2147467259	'0x80004005 �d���L�[
			AddSql ""
			AddSql "update Zaiko"
			if GetOption("","u") = "u" then
				Write ":u",0
				AddSql " set YUKO_Z_QTY = '" & RTrim(objRs.Fields("Qty")) & "'"
				AddSql "   , YOSAN_TO = '" & RTrim(objRs.Fields("SyushiCd")) & "'"
				AddSql " where JGYOBU = '" & RTrim(objRs.Fields("JGyobu")) & "'"
				AddSql "   and NAIGAI = '1'"
				AddSql "   and HIN_GAI = '" & RTrim(objRs.Fields("Pn")) & "'"
				AddSql "   and GOODS_ON = '1'"	' //0:���i���� 1:�����i
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
				AddSql "   and GOODS_ON = '1'"	' //0:���i���� 1:�����i
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
	'Line0() 1�s�\��
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
		AddSql ",if(n.JGyobu is null,'','���� ' + n.JGyobu + ' ' + n.NYUKO_TANABAN) Nyuka"
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
	'Line2() 1�s�\��
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
	'Sql���s
	'-------------------------------------------------------------------
	Private Function Execute(byVal strSql)
		Debug ".Execute():" & strSql
		on error resume next
		Call objDb.Execute(strSql)
		Execute = Err.Number
		select case Execute
		case 0
		case -2147467259	'0x80004005 �d���L�[
		case else
			Wscript.StdErr.WriteLine
			Wscript.StdErr.WriteLine Err.Description
			Wscript.StdErr.WriteLine strSql
		end select
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Sql���s
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
	'������ǉ� strSql
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
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Private Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'�I�v�V�����擾
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
