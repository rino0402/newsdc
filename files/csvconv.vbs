Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "csvconv.vbs [option]"
	Wscript.Echo " /db:newsdc9	�f�[�^�x�[�X"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript csvconv.vbs /db:newsdc9 pop3w9\tmp\�݌�_�I��1.csv"
	Wscript.Echo "cscript csvconv.vbs /db:newsdc9 BoSyukaDet.csv"
End Sub
'-----------------------------------------------------------------------
'BoCnv
'-----------------------------------------------------------------------
Class Csv
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strPathName
	Private	strFileName
	Private	strDT
	'-----------------------------------------------------------------------
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "i"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName	= GetOption("db"	,"newsdc")
		strDT		= year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2)
		set objDB	= nothing
		set objRs	= nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			Debug ".Run():" & strArg
			strPathName = strArg
			strFileName = GetFileName(strPathName)
			Call Conv()
		Next
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Load() �Ǎ�
	'-----------------------------------------------------------------------
    Private Function Conv()
		Debug ".Conv():" & strPathName
		select case CsvType()
		case "BoZaiko"
			Call BoZaiko()
		case "BoZaikoName"
			Call BoZaiko()
		case "BoSyukaDet"
			Call BoSyukaDet()
		end select
	End Function
	'-----------------------------------------------------------------------
	'BoSyukaDet()
	'-----------------------------------------------------------------------
	Private	strYm1
	Private	strYm2
    Private Function BoSyukaDet()
		Debug ".BoSyukaDet()"
		Wscript.StdOut.WriteLine "�t�@�C����:" & strFileName
		Wscript.StdOut.WriteLine "      �`��:" & strCsvType
		BoSyukaDet_Btwn
		Wscript.StdOut.WriteLine "      �N��:" & strYm1 & "-" & strYm2
		BoSyukaDet_Del
		Wscript.StdOut.WriteLine "      �폜:" & RowCount()
		BoSyukaDet_Ins
		Wscript.StdOut.WriteLine "      �ǉ�:" & RowCount()
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_Del()
	'-----------------------------------------------------------------------
    Private Function BoSyukaDet_Btwn()
		Debug ".BoSyukaDet_Btwn()"
		AddSql ""
		AddSql "select"
		AddSql " Min(Left(if(Col=11,Col08,Col06),6)) Ym1"
		AddSql ",Max(Left(if(Col=11,Col08,Col06),6)) Ym2"
		AddSql "from CsvTemp"
		AddSql "where FileName = '" & strFileName & "'"
		AddSql "and Row>1"
		AddSql "and Col01 not like '%�Ǘ��ԍ�'"
'		AddSql "and Col=11"
		CallSql
		strYm1 = ""
		strYm2 = ""
		do while objRs.Eof = False
			strYm1 = RTrim(objRs.Fields("Ym1"))
			strYm2 = RTrim(objRs.Fields("Ym2"))
			exit do
		loop
	End Function
	'-----------------------------------------------------------------------
	'BoSyukaDet_Del()
	'-----------------------------------------------------------------------
    Private Function BoSyukaDet_Del()
		Debug ".BoSyukaDet_Del()"
		AddSql ""
		AddSql "delete from BoSyukaDet where Left(JisekiDt,6) between '" & strYm1 & "' and '" & strYm2 & "'"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoSyukaDet_Ins()
	'-----------------------------------------------------------------------
    Private Function BoSyukaDet_Ins()
		Debug ".BoSyukaDet_Ins()"
		AddSql ""
		AddSql "insert into BoSyukaDet"
		AddSql "("
'		AddSql "No"			'//NO 1
		AddSql " IDNo"		'//�󒍏o�׊Ǘ��ԍ� 700089663241	
		AddSql ",JCode"		'//���Ə�CD 00021184
		AddSql ",Syushi"	'//�݌Ɏ��x 34"
		AddSql ",DenNo"		'//�`�[�ԍ� 032293
		AddSql ",SyukaCd"	'//�o�א�CD 113A
		AddSql ",SyukaNm"	'//�o�א於 MALAYSIA(PM)
		AddSql ",AiteCd"	'//�����CD 
		AddSql ",AiteNm"	'//����於 ���̑�
		AddSql ",ChuKb"		'//�����敪 6:AIR�T�؁i13���j
		AddSql ",JisekiDt"	'//������єN���� 20150401"
		AddSql ",Pn"		'//�i�ڔԍ� A390C7R30WT
		AddSql ",Qty"		'//�o�׎��ѐ� 2
		AddSql ") select distinct "
		AddSql " RTrim(Col01)"						'�󒍏o�׉ߎ�_�󒍏o�׊Ǘ��ԍ�	700093657489
		AddSql ",RTrim(Col02)"						'�󒍏o�׉ߎ�_���Y�Ǘ����Ə�R�[�h	00021529
		AddSql ",RTrim(if(Col=11,Col03,Col04))"		'�󒍏o�׉ߎ�_�݌Ɏ��x�R�[�h	11D
		AddSql ",RTrim(if(Col=11,Col04,Col05))"		'�󒍏o�׉ߎ�_�`�[�ԍ�	027243
		AddSql ",RTrim(if(Col=11,Col05,Col07))"		'�󒍏o�׉ߎ�_���Ӑ�R�[�h(�����CD)	00020162
		AddSql ",RTrim(if(Col=11,Col06,Col08))"		'�󒍏o�׉ߎ�_���Ӑ旪��(����於)	�A�v���C�A���X�Ё@�{��
		AddSql ",RTrim(if(Col=11,Col11,Col13))"		'�󒍏o�׉ߎ�_���������R�[�h	00020162
		AddSql ",''"
		AddSql ",RTrim(if(Col=11,Col07,Col12))"		'�󒍏o�׉ߎ�_�����敪	2
		AddSql ",RTrim(if(Col=11,Col08,Col06))"		'�󒍏o�׉ߎ�_������єN����	20170119
		AddSql ",RTrim(if(Col=11,Col09,Col03))"		'�󒍏o�׉ߎ�_�i�ڔԍ�	ANP300-1530
		AddSql ",Convert(RTrim(if(Col=11,Col10,Col09)),sql_decimal)"	'�󒍏o�׉ߎ�_�o�׎��ѐ�	4
		AddSql "from CsvTemp"
		AddSql "where FileName = '" & strFileName & "'"
		AddSql "and Row > 1"
		AddSql "and Col01 not like '%�Ǘ��ԍ�'"
'		AddSql "and Col = 11"
'1�󒍏o��_�󒍏o�׊Ǘ��ԍ�	700093755635
'2�󒍏o��_���Y�Ǘ����Ə�R�[�h	00023100
'3�󒍏o��_�i�ڔԍ�	AXW22B-7EM0
'4�󒍏o��_�݌Ɏ��x�R�[�h	11D
'5�󒍏o��_�`�[�ԍ�	004781
'6�󒍏o��_������єN����	20170207
'7�󒍏o��_���Ӑ�R�[�h(�����R�[�h)	00020162
'8�󒍏o��_���Ӑ旪��(����於)	�A�v���C�A���X�Ё@�{��
'9�󒍏o��_�o�׎��ѐ�	2
'10�󒍏o��_�q�ɃR�[�h	NAR
'11�󒍏o��_�����敪	2
'12�󒍏o��_���o�Ɏ���敪	20
'13�󒍏o��_���������R�[�h	00020162
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko()
	'-----------------------------------------------------------------------
	Private Function BoZaiko()
		Debug ".BoZaiko()"
		Wscript.StdOut.WriteLine "�t�@�C����:" & strFileName
		Wscript.StdOut.WriteLine "      �`��:" & strCsvType

		Wscript.StdOut.Write "        BoZaiko_Del:"
		BoZaiko_Del
		Wscript.StdOut.WriteLine RowCount()

		select case strCsvType
		case "BoZaiko"
			Wscript.StdOut.Write "        BoZaiko_Ins:"
			BoZaiko_Ins
			Wscript.StdOut.WriteLine RowCount()
		case "BoZaikoName"
			Wscript.StdOut.Write "    BoZaikoName_Ins:"
			BoZaikoName_Ins
			Wscript.StdOut.WriteLine RowCount()
		end select

'		Wscript.StdOut.Write "ZaikoH " & strDT & " Del:"
'		BoZaiko_ZaikoH_Del
'		Wscript.StdOut.WriteLine RowCount()

'		Wscript.StdOut.Write "ZaikoH " & strDT & " Ins:"
'		BoZaiko_ZaikoH
'		Wscript.StdOut.WriteLine RowCount()

'		Wscript.StdOut.Write "            NarTana:"
'		BoZaiko_NarTana()
'		Wscript.StdOut.WriteLine RowCount()
	End Function
	'-----------------------------------------------------------------------
	'RowCount()
	'-----------------------------------------------------------------------
    Private Function RowCount()
		Debug ".RowCount()"
		dim	objRow
		set	objRow = objDb.Execute("select @@rowcount")
		RowCount = objRow.Fields(0)
		set	objRow = Nothing
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_Del()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_Del()
		Debug ".BoZaiko_Del()"
'		Disp "BoZaiko:delete all"
		AddSql ""
		AddSql "delete from BoZaiko"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_Ins()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_Ins()
		Debug ".BoZaiko_Ins()"
'		Disp "BoZaiko:Insert " & strFileName
		AddSql	""
		AddSql	"insert into BoZaiko"
		AddSql	"(Soko"
		AddSql	",JCode"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",SyuShi"
		AddSql	",TanaQty"
		AddSql	",HikiQty"
		AddSql	",SyuShiR"
		AddSql	",SyuShiName"
		AddSql	",Loc1"
		AddSql	") select "
		AddSql	"distinct"
		AddSql	" RTrim(Col01)"				'// �݌Ɏ��x_�q�ɃR�[�h
		AddSql	",RTrim(Col02)"				'// �o�m�q�ɍ݌�_���Ə�R�[�h
		AddSql	",RTrim(Col03)"				'// �o�m�q�ɍ݌�_���Y�Ǘ����Ə�R�[�h
		AddSql	",RTrim(Col04)"				'// �o�m�q�ɍ݌�_�i�ڔԍ�
		AddSql	",RTrim(Col05)"				'// �o�m�q�ɍ݌�_�݌Ɏ��x�R�[�h
		AddSql	",Convert(Col06,Sql_Decimal)"	'// �o�m�q�ɍ݌�_�I�݌ɐ�
		AddSql	",Convert(Col07,Sql_Decimal)"	'// �o�m�q�ɍ݌�_���������\�݌ɐ�
		AddSql	",RTrim(Col08)"				'// �݌Ɏ��x_�݌Ɏ��x������
		AddSql	",Max(RTrim(Col10))"			'// �݌Ɏ��x_�݌Ɏ��x��
		AddSql	",RTrim(Col09)"				'// �I��_�P
		AddSql	"from CsvTemp"  
		AddSql	"where FileName = '" & strFileName & "'"
		AddSql	"and Row > 1"
		AddSql	"and RTrim(Col01) = 'NAR'"
		AddSql	"group by"
		AddSql	"Col01"
		AddSql	",Col02"
		AddSql	",Col03"
		AddSql	",Col04"
		AddSql	",Col05"
		AddSql	",Col06"
		AddSql	",Col07"
		AddSql	",Col08"
		AddSql	",Col09"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaikoName_Ins()
	'-----------------------------------------------------------------------
    Private Function BoZaikoName_Ins()
		Debug ".BoZaikoName_Ins()"
		AddSql	""
		AddSql	"insert into BoZaiko"
		AddSql	"(Soko"
		AddSql	",JCode"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",SyuShi"
		AddSql	",TanaQty"
		AddSql	",HikiQty"
		AddSql	",SyuShiR"
		AddSql	",Loc1"
		AddSql	",SyuShiName"
		AddSql	",PName"
		AddSql	",PNameEng"
		AddSql	") select "
		AddSql	"distinct"
		AddSql	" RTrim(Col01)	Soko"			'PN�q�Ɂi���x)_�q�ɃR�[�h
		AddSql	",RTrim(Col02)	JCode"			'PN�q�ɁiPN�I�j_���Ə�R�[�h
		AddSql	",RTrim(Col03)	ShisanJCode"	'PN�q�ɁiPN�I�j_���Y�Ǘ����Ə�R�[�h
		AddSql	",RTrim(Col04)	Pn"				'PN�q�ɁiPN�I�j_�i�ڔԍ�
		AddSql	",RTrim(Col06)	SyuShi"			'PN�q�ɁiPN�I)_�݌Ɏ��x�R�[�h
		AddSql	",Convert(Col07,Sql_Decimal)	TanaQty"	'PN�q�ɁiPN�I)_�I�݌ɐ��@������
		AddSql	",Convert(Col08,Sql_Decimal)	HikiQty"'	'PN�q�ɁiPN�I)_���������\�݌ɐ�
		AddSql	",RTrim(Col09)	SyuShiR"		'PN�q�ɁiPN�I�j_�݌Ɏ��x������
		AddSql	",RTrim(Col10)	Loc1"			'PN�q�ɁiPN�I�j_���P�[�V�����ԍ��P
		AddSql	",Max(RTrim(Col11))	SyuShiName"	'�݌Ɏ��x_�݌Ɏ��x��'
		AddSql	",Max(RTrim(Col05))	PName"		'�o�m����(JPN)_�i�ږ�
		AddSql	",Max(RTrim(Col12)) PNameEng"	'�o�m����_�i�ڕʖ�(ENG)'
		AddSql	"from CsvTemp"
		AddSql	"where FileName = '" & strFileName & "'"
		AddSql	"and Row > 1"
		AddSql	"group by"
		AddSql	" Soko"
		AddSql	",JCode"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",SyuShi"
		AddSql	",TanaQty"
		AddSql	",HikiQty"
		AddSql	",SyuShiR"
		AddSql	",Loc1"
'		AddSql	",SyuShiName"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_ZaikoH()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_ZaikoH_Del()
		Debug ".BoZaiko_ZaikoH_Del()"
'		Disp "BoZaiko:ZaikoH delete " & strDT
		AddSql	""
		AddSql	"delete from ZaikoH"
		AddSql	"where Kubun = 'Bo'"
'		AddSql	"and DT = left(replace(convert(now(),sql_char),'-',''),8)"
		AddSql	"and DT = '" & strDT & "'"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_ZaikoH()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_ZaikoH()
		Debug ".BoZaiko_ZaikoH()"
'		Disp "BoZaiko:ZaikoH insert " & strDT
		AddSql	""
		AddSql	"insert into ZaikoH"
		AddSql	"(Kubun"
		AddSql	",DT"
		AddSql	",JCode"
		AddSql	",Pn"
		AddSql	",Syushi"
		AddSql	",Qty"
		AddSql	",QtyHiki"
		AddSql	",Loc1"
		AddSql	") select"
		AddSql	"'Bo'"
		AddSql	",'" & strDT & "'"
'		AddSql	",left(replace(convert(now(),sql_char),'-',''),8)"
		AddSql	",ShisanJCode"
		AddSql	",Pn"	'					"�i��"
		AddSql	",SyuShi"	'				"���x"
		AddSql	",TanaQty"	'				"�I�݌ɐ�"
		AddSql	",HikiQty"	'				"�����\�݌ɐ�"
		AddSql	",Loc1"	'					"�I��_�P"
		AddSql	"from BoZaiko"
'		AddSql	"where Soko='NAR'"
		AddSql 	"where (Soko = 'NAR' or (Soko = 'NA2' and Left(Loc1,1) = 'E'))"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'BoZaiko_NarTana()
	'-----------------------------------------------------------------------
    Private Function BoZaiko_NarTana()
		Debug ".BoZaiko_NarTana()"

'		Disp "BoZaiko:NarTana delete in BoZaiko"
		AddSql	""
		AddSql	"delete from NarTana"
		AddSql	"where RTrim(Soko)+RTrim(ShisanJCode)+RTrim(Pn)"
		AddSql	"in (select distinct RTrim(Soko)+RTrim(ShisanJCode)+RTrim(Pn)"
		AddSql	"from BoZaiko"
'		AddSql	"where Soko='NAR'"
		AddSql 	"where (Soko = 'NAR' or (Soko = 'NA2' and Left(Loc1,1) = 'E'))"
		AddSql	"and left(SyuShi,2) in ('11','12','10','41','71','99','15')"
		AddSql	"and Loc1<>''"
		AddSql	")"
		CallSql

'		Disp "BoZaiko:NarTana insert in BoZaiko"
		AddSql	""
		AddSql	"insert into NarTana"
		AddSql	"(Soko"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",Loc1"
		AddSql	",Loc10"
		AddSql	",Loc11"
		AddSql	",Loc12"
		AddSql	",Loc41"
		AddSql	",Loc71"
		AddSql	",Loc99"
		AddSql	",Loc15"
		AddSql	",EntID"
		AddSql	")"
		AddSql	"select"
		AddSql	"distinct"
		AddSql	" Soko"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		AddSql	",Max(Loc1)"
		AddSql	",Max(if(left(SyuShi,2)='10',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='11',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='12',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='41',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='71',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='99',Loc1,''))"
		AddSql	",Max(if(left(SyuShi,2)='15',Loc1,''))"
		AddSql	",'BoZaiko'"
		AddSql	"from BoZaiko"
		AddSql 	"where (Soko = 'NAR' or (Soko = 'NA2' and Left(Loc1,1) = 'E'))"
		AddSql	"and left(SyuShi,2) in ('11','12','10','41','71','99','15')"
		AddSql	"and Loc1<>''"
		AddSql	"group by"
		AddSql	" Soko"
		AddSql	",ShisanJCode"
		AddSql	",Pn"
		CallSql

'		Disp "BoZaiko:NarTana update Loc11"
		AddSql	""
		AddSql	"update NarTana"
		AddSql	"set Loc1=Loc11"
		AddSql	",	UpdID='Loc11'"
		AddSql	",	UpdTm=Now()"
'		AddSql	"where Soko='NAR'"
		AddSql	"where Loc11<>''"
		AddSql	"and Loc1<>Loc11"
		CallSql
	End Function
	'-----------------------------------------------------------------------
	'CsvType()
	'-----------------------------------------------------------------------
	Private	strCsvType
    Private Function CsvType()
		Debug ".CsvType()"
		strCsvType = ""
		AddSql	""
		AddSql	"select"
		AddSql	"y.CsvType cType"
'		AddSql	"from CsvType y"
'		AddSql	"inner join CsvTemp t"
		AddSql	"from CsvTemp t"
		AddSql	"inner join CsvType y"
		AddSql	"on (y.Col = t.Col"
		AddSql	"and y.Col01 = t.Col01"
		AddSql	"and y.Col02 = t.Col02"
		AddSql	"and y.Col03 = t.Col03"
		AddSql	"and y.Col04 = t.Col04"
		AddSql	"and y.Col05 = t.Col05"
		AddSql	"and y.Col06 = t.Col06"
		AddSql	"and y.Col07 = t.Col07"
		AddSql	"and y.Col08 = t.Col08"
		AddSql	"and y.Col09 = t.Col09"
		AddSql	"and y.Col10 = t.Col10"
		AddSql	"and y.Col11 = t.Col11"
		AddSql	"and y.Col12 = t.Col12"
		AddSql	")"
		AddSql	"where t.FileName = '" & strFileName & "'"
		AddSql	"and t.Row = 1"
		Wscript.StdErr.Write strFileName & ":"
		CallSql
		if objRs.Eof = False then
			Debug ".CsvType():" & objRs.Fields("cType")
			strCsvType = RTrim(objRs.Fields("cType"))
		end if
		Wscript.StdErr.WriteLine strCsvType
		CsvType = strCsvType
	End Function
	'-------------------------------------------------------------------
	'�t�@�C����(�p�X������)
	'-------------------------------------------------------------------
	Private Function GetFileName(byVal f)
		dim	objFileSys
		Set objFileSys	= WScript.CreateObject("Scripting.FileSystemObject")

		dim	strFName
		strFName = objFileSys.GetBaseName(f)
		strFName = strFName & "."
		strFName = strFName & objFileSys.GetextensionName(f)
		GetFileName	= strFName

		Set objFileSys	= Nothing
	End Function
	'-------------------------------------------------------------------
	'������ǉ� strSql
	'-------------------------------------------------------------------
	dim	strSql
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
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
	'Sql���s
	'-------------------------------------------------------------------
	Private Function CallSql()
		Debug ".CallSql():" & strSql
'		on error resume next
		Set objRs = objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field��
	'-------------------------------------------------------------------
	Private Function GetFields(byVal strTable)
		Debug ".GetFields():" & strTable
		dim	strFields
		strFields = ""
		dim	objRs
		set objRS = objDB.Execute("select top 1 * from " & strTable)
		dim	objF
		for each objF in objRS.Fields
			if strFields <> "" then
				strFields = strFields & ","
			end if
			strFields = strFields & objF.Name
		next
		set objRs = nothing
		GetFields = strFields
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
	'���b�Z�[�W�\��
	'-----------------------------------------------------------------------
	Private Sub Disp(byVal strMsg)
		Wscript.StdErr.WriteLine strMsg
	End Sub
	'-----------------------------------------------------------------------
	'�I�v�V�����擾
	'-----------------------------------------------------------------------
	Private Function GetOption(byval strName _
					  ,byval strDefault _
					  )
		dim	strValue

		if strName = "" then
			strValue = ""
			if strDefault < WScript.Arguments.UnNamed.Count then
				strValue = WScript.Arguments.UnNamed(strDefault)
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
End Class
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objCsv
	Set objCsv = New Csv
	if objCsv.Init() <> "" then
		call usage()
		exit function
	end if
	call objCsv.Run()
End Function
