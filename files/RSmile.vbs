Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "RSmile.vbs [option]"
	Wscript.Echo " /db:newsdc1	�f�[�^�x�[�X"
	Wscript.Echo " /make      �����f�[�^�쐬(default)"
	Wscript.Echo " /make:test �����f�[�^�쐬:Test�p(�S��)"
	Wscript.Echo " /csv       �����f�[�^�o��"
	Wscript.Echo "Ex."
	Wscript.Echo "sc32//nologo RSmile.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'RSmile CSV�t�@�C���쐬
'2016.10.25 R-smile(SSX)
'2016.10.28 /make:test �����f�[�^�쐬:Test�p(�S��)
'-----------------------------------------------------------------------
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CommandTypeEnum Values ----
Const adCmdUnspecified	= -1	' Unspecified type of command 
Const adCmdText			= 1		' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2		' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4		' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8		' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

Class Rsmile
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strAction	' make/csv
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		strPrgId = GetOption("make"	,"RSmile.vbs")
		set objDB = nothing
		set objRs = nothing
		strAction = "make"
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
		select case strAction
		case "csv"
			Call Csv()
		case "make"
			Call Make()
		end select
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Csv() �����f�[�^�쐬
	'-----------------------------------------------------------------------
    Public Function Csv()
		Debug ".Csv()"
		dim	dtToday
		dtToday = Date()
		strSyukaDt = Year(dtToday) & Right("0" & Month(dtToday),2) & Right("0" & Day(dtToday),2)
		Call SetSql("")
		Call SetSql("select")
		Call SetSql("*")
		Call SetSql("from RSmile")
		Call SetSql("where SyukaDt = '" & strSyukaDt & "'")
		Call SetSql("order by")
		Call SetSql(" Id")
		Debug ".Csv():" & strSql
		set objRs = objDB.Execute(strSql)
		intCnt = 0
		Call CsvLine()
		do while objRs.Eof = False
			intCnt = intCnt + 1
			Call CsvLine()
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'YoteiDt() �\���
	'-------------------------------------------------------------------
	Private Function YoteiDt()
		dim	strDt
		strDt = GetField("SyukaDt")
		dim	dtTmp
		strDt = Left(strDt,4) & "/" & Mid(strDt,5,2) & "/" & Right(strDt,2)
		Debug ".YoteiDt:" & strDt
		dtTmp = CDate(strDt)
		dtTmp = DateAdd("d",1,dtTmp)
		YoteiDt = CStr(dtTmp)
	End Function
	'-------------------------------------------------------------------
	'CsvLine() 1�s�o��
	'-------------------------------------------------------------------
	Private Function CsvLine()
		Debug ".CsvLine()"
		dim	objF
		'1�s�ځF���ږ�
		if intCnt = 0 then
			WScript.StdOut.Write "�͐�"		'1	���͐�R�[�h		���p����	15	15	
			WScript.StdOut.Write ",�d�b�ԍ�"	'2	�d�b�ԍ�	��	���p����	15	15	�n�C�t���t��
			WScript.StdOut.Write ",�X�֔ԍ�"	'3	�X�֔ԍ�		���p����	7	7	�n�C�t���Ȃ�
			WScript.StdOut.Write ",�s���{��"	'4	�s���{��	��	�S�p����	10	20	
			WScript.StdOut.Write ",�s�撬��"	'5	�s�撬��	��	�S�p����	20	40	
			WScript.StdOut.Write ",����"		'6	����	��	�S�p����	30	60	
			WScript.StdOut.Write ",�Ԓn"		'7	�Ԓn�E�r����		�S�p����	30	60	�Ԓn�E�r�����Ȃ�
			WScript.StdOut.Write ",���̂P"	'8	���̂P	��	�S�p����	30	60	
			WScript.StdOut.Write ",���̂Q"	'9	���̂Q		�S�p����	30	60	
'			WScript.StdOut.Write ",�h��"		'10	���͂���h��		�S�p����	2	4	�u�Ȃ��v �A�u�l�v�A�u�䒆�v�A�u�a�v�A�u�����N
'			WScript.StdOut.Write ",��"		'11	��	��	���p����	4	4	�C���|�[�g�\�Ȍ�����3���ł�
'			WScript.StdOut.Write ",�d��"		'12	�d��	��	���p����	7	7	"�����_���͉i�����_�ȉ�3���܂Łj         �O����"
			WScript.StdOut.Write ",�ב��l"	'13	�ב��l�b�c		���p����	6	6	"�����͎��̓f�t�H���g�̉ב��l�b�c���g�p         ���f�t�H���g�̉ב��l�b�c�����ݒ�̏ꍇ�A�K�{����"
											'14	�ב��l�d�b�ԍ�	��	���p����	15	15	�n�C�t���t��
											'15	�ב��l�X�֔ԍ�		���p����	7	7	�n�C�t���Ȃ�
											'16	�ב��l�s���{��	��	�S�p����	10	20	
											'17	�ב��l�s�撬��	��	�S�p����	20	40	
											'18	�ב��l����	��	�S�p����	30	60	
											'19	�ב��l�Ԓn		�S�p����	30	60	�ב��l�̔Ԓn�E�r�����Ȃ�
											'20	��Ж�	��	�S�p����	30	60	
											'21	�ב��l�S���Җ�		�S�p����	30	60	�����͎��̓}�X�^�̓��e���g�p
											'22	���q���܂m��	��	���p����	6	6	
											'23	�����P��		���p����	20	20	99�F���ʃ��C�A�E�g��I�����Ă���ꍇ�́A�����P�ʃt���[���͖������[�U�[�͕R�t�s��
											'24	�`�[�敪	��	�S�p����	2	4	"�����F""����""�A�����F""����""�A      ����F""���"""
			WScript.StdOut.Write ",�q��敪"	'25	�q��敪	��	�S�p����	3	6	"�A���敪���q��̎��̂ݗL��         �}�ցF�󔒁A�`�h�q�F""�`�h�q"""
			WScript.StdOut.Write ",�z�B��"	'26	�z�B�w���		���p����	8	8	�N�����i�N�͐���4���j
'			WScript.StdOut.Write ",�z�B����"	'27	�z�B�w�莞�ԑ�		���p�p��	2	2	
			WScript.StdOut.Write ",�L��1"	'28	�i��	��	�S�p����	30	60	
			WScript.StdOut.Write ",�L��2"	'29	�L���Q		�S�p����	30	60	
			WScript.StdOut.Write ",�L��3"	'30	�L���R		�S�p����	30	60	
			WScript.StdOut.Write ",�L��4"	'31	�L���S		�S�p����	30	60	
			WScript.StdOut.Write ",�L��5"	'32	�L���T		�S�p����	30	60	
			WScript.StdOut.Write ",�o�ד�"	'33	�o�ד�		���p����	8	8	"�N�����i�N�͐���4���j      30����܂Ŏw��\�B      �ߋ����̏ꍇ�͓����̓��t�ɕύX�����"
											'34			���p����	10	10	���ݎg�p���Ă��܂���B
			WScript.StdOut.Write ",�o�הԍ�"	'35	�o�הԍ�		���p����	15	15	
			WScript.StdOut.Write ",�`�[�ԍ�"	'36	�`�[�ԍ�		���p����	1	1	���͂���Ă���ꍇ�A������ɍ̔Ԃ��Ȃ�
			WScript.StdOut.WriteLine
			exit function
		end if
		'����
		RsCsv "�͐�"
		RsCsv "�d�b�ԍ�"	
		RsCsv "�X�֔ԍ�"	
		RsCsv "�s���{��"	
		RsCsv "�s�撬��"	
		RsCsv "����"		
		RsCsv "�Ԓn"		
		RsCsv "���̂P"	
		RsCsv "���̂Q"	
'		RsCsv "�h��"		
'		RsCsv "��"		
'		RsCsv "�d��"		
		RsCsv "�ב��l"	
		RsCsv "�q��敪"	
		RsCsv "�z�B��"	
'		RsCsv "�z�B����"	
		RsCsv "�L��1"	
		RsCsv "�L��2"	
		RsCsv "�L��3"	
		RsCsv "�L��4"	
		RsCsv "�L��5"	
		RsCsv "�o�ד�"	
		RsCsv "�o�הԍ�"	
		RsCsv "�`�[�ԍ�"	
		WScript.StdOut.WriteLine
	End Function
	'-----------------------------------------------------------------------
	'�d�b�ԍ� �n�C�t���ǉ�
	'-----------------------------------------------------------------------
	Private	Function Tel(byVal strTel)
		Debug ".Tel():" & strTel
		Tel = strTel
		if inStr(strTel,"-") > 0 then
			exit function
		end if
		dim	strTel1
		dim	strTel2
		dim	strTel3
		select case len(strTel)
		case 11
			Debug ".Tel():11:" & strTel
			strTel1 = left(strTel,3)
			strTel2 = mid(strTel,4,4)
			strTel3 = right(strTel,4)
			strTel = strTel1 & "-" & strTel2 & "-" & strTel3
			Debug ".Tel():11:" & strTel
		case 10
			Debug ".Tel():10:" & strTel
			select case left(strTel,2)
			case "03","06"
				Debug ".Tel():10-2:" & strTel
				' 03-3456-7890
				strTel1 = left(strTel,2)
				strTel2 = mid(strTel,3,4)
				strTel3 = right(strTel,4)
				strTel = strTel1 & "-" & strTel2 & "-" & strTel3
				Debug ".Tel():10-2:" & strTel
			case else
				select case left(strTel,3)
				case "028","045","046","048","078","080"
					'028-688-8168
					Debug ".Tel():10-3:" & strTel
					strTel1 = left(strTel,3)
					strTel2 = mid(strTel,5,3)
					strTel3 = right(strTel,4)
					strTel = strTel1 & "-" & strTel2 & "-" & strTel3
					Debug ".Tel():10-3:" & strTel
				case else
					select case left(strTel,4)
					case "0258"
						'0258-42-2211
						Debug ".Tel():10-4:" & strTel
						strTel1 = left(strTel,4)
						strTel2 = mid(strTel,5,2)
						strTel3 = right(strTel,4)
						strTel = strTel1 & "-" & strTel2 & "-" & strTel3
						Debug ".Tel():10-4:" & strTel
					case else
					end select
				end select
			end select
		end select
		Tel = strTel
	End Function
	'-----------------------------------------------------------------------
	'R-smile�p(CSV)
	'-----------------------------------------------------------------------
	Private Function RsCsv(byVal strName)
		dim	strComma
		strComma = ","
		dim	strValue
		strValue = ""
		select case strName
		case "�͐�"
			strValue = GetField("SCode")
			strComma = ""
		case "�d�b�ԍ�"	
			strValue = Tel(GetField("STel"))
		case "�X�֔ԍ�"	
			strValue = GetField("SZip")
		case "�s���{��"	
		case "�s�撬��"	
		case "����"		
			strValue = GetField("SAddress")
		case "�Ԓn"		
			strValue = GetField("SBillding")
		case "���̂P"	
			strValue = GetField("SName1")
		case "���̂Q"	
			strValue = GetField("SName2")
'		case "�h��"		
'		case "��"		
'		case "�d��"		
		case "�ב��l"	
			strValue = RsSender()
		case "�q��敪"	
			select case RsSender()
			case 7,8
				strValue = "�`�h�q"
			end select
		case "�z�B��"	
			strValue = YoteiDt()
'		case "�z�B����"	
		case "�L��1"	
			strValue = GetField("SKiji1")
		case "�L��2"	
			strValue = GetField("SKiji2")
		case "�L��3"	
			strValue = GetField("SKiji3")
		case "�L��4"	
			strValue = GetField("SKiji4")
		case "�L��5"	
			strValue = GetField("SKiji5")
		case "�o�ד�"	
			strValue = GetField("SyukaDt")
		case "�o�הԍ�"	
			strValue = GetField("ID") & "0"
		case "�`�[�ԍ�"	
		end select
		WScript.StdOut.Write strComma
		WScript.StdOut.Write Replace(strValue,",",".")
	End Function
	'-----------------------------------------------------------------------
	'RsSender Rs�ב��l
	'	SDC����E�o�Y�@(����)
	'	SDC����F�o�Y�@(����)
	'	SDC����G�o�Y�@(�G�A�[)
	'-----------------------------------------------------------------------
	Private	Function RsSender()
		RsSender = 6
		dim	strAddress
		strAddress = GetField("SAddress")
		if Left(strAddress,2) = "����" then
			RsSender = 7
			exit function
		end if
		if Left(strAddress,3) = "�k�C��" then
			RsSender = 8
			exit function
		end if
	End Function
	'-----------------------------------------------------------------------
	'Make() �����f�[�^�쐬
	'-----------------------------------------------------------------------
	Private	strPrgId
    Public Function Make()
		Debug ".Make()"
		SetSql	""
		if strPrgId = "test" then
			Disp "�e�X�g�f�[�^�쐬"
			Call objDb.Execute("delete from RSmile where EntID = 'test'")
			SetSql	"select"
			SetSql	"distinct"
			SetSql	"Min(IdNo) IdNo"
			SetSql	",Left(Replace(Convert(Now(),sql_char),'-',''),8) SyukaDt"
			SetSql	",d.ChoCode ChoCode"
			SetSql	",d.ChoName ChoName"
			SetSql	",d.ChoAddress ChoAddress"
			SetSql	",d.ChoTel ChoTel"
			SetSql	",d.ChoZip ChoZip"
			SetSql	",'' Id"
			SetSql	"from HtDrctId d"
			SetSql	"where RTrim(ChoCode) <> ''"
			SetSql	"group by"
			SetSql	"SyukaDt"
			SetSql	",ChoCode"
			SetSql	",ChoName"
			SetSql	",ChoCode"
			SetSql	",ChoAddress"
			SetSql	",ChoTel"
			SetSql	",ChoZip"
			SetSql	"order by"
			SetSql	"ChoCode"
		else
			SetSql	"select"
			SetSql	"distinct"
			SetSql	"y.KEY_SYUKA_YMD SyukaDt"
			SetSql	",d.ChoCode ChoCode"
			SetSql	",d.ChoName ChoName"
			SetSql	",d.ChoAddress ChoAddress"
			SetSql	",d.ChoTel ChoTel"
			SetSql	",d.ChoZip ChoZip"
			SetSql	",r.Id Id"
			SetSql	"from y_syuka y"
			SetSql	"inner join HtDrctId d on (d.IDNo = y.KEY_ID_NO)"
			SetSql	"left outer join RSmile r on (y.KEY_SYUKA_YMD = r.SyukaDt and d.ChoCode = r.SCode)"
			SetSql	"order by"
			SetSql	" SyukaDt"
			SetSql	",ChoCode"
		end if
		Debug ".Make():" & strSql
		set objRs = objDB.Execute(strSql)
		do while objRs.Eof = False
			Call MakeData()
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'����悲�Ƃ̌����J�E���g
	'-------------------------------------------------------------------
	Private	strSyukaDt
	Private	strChoCode
	Private	strChoName
	Private	strClientNo
	Private	strSCode
	Private	intCnt
	'-------------------------------------------------------------------
	'MakeData() 1�s�Ǎ�
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData()"
		WScript.StdOut.Write GetField("SyukaDt")
		if strPrgId = "test" then
			WScript.StdOut.Write " " & GetField("IdNo")
		end if
'		WScript.StdOut.Write " " & GetField("Biko1")
		WScript.StdOut.Write " " & Left(GetField("ChoCode") & Space(9),9)
		WScript.StdOut.Write Get_LeftB(GetField("ChoName") & Space(40),40)
'		WScript.StdOut.Write " " & GetField("ChoAddress")
'		WScript.StdOut.Write " " & GetField("ChoTel")
'		WScript.StdOut.Write " " & GetField("ChoZip")
		Call GetKiji()
		Call InsertData()
		WScript.StdOut.WriteLine
	End Function
	'-------------------------------------------------------------------
	'Kiji1-5
	'-------------------------------------------------------------------
	Private	intDenCnt
	Private	strId
	Private	strKiji1
	Private	strKiji2
	Private	strKiji3
	Private	strKiji4
	Private	strKiji5
	Private Function GetKiji()
		Debug ".GetKiji()"
		if strPrgId = "test" then
			strId = GetField("IdNo")
			strKiji1 = String(30,"�e")
			strKiji2 = String(30,"�X")
			strKiji3 = String(30,"�g")
			strKiji4 = String(30,"�f")
			strKiji5 = String(30,"�X")
			exit function
		end if
		SetSql	""
		SetSql	"select"
		SetSql	"distinct"
		SetSql	"y.KEY_SYUKA_YMD SyukaDt"
		SetSql	",y.KEY_ID_NO Id"
		SetSql	",y.KEY_HIN_NO HIN_GAI"
		SetSql	",convert(y.SURYO,sql_decimal) Qty"
		SetSql	",y.Bikou1 Biko1"
		SetSql	"from y_syuka y"
		SetSql	"inner join HtDrctId d on (d.IDNo = y.KEY_ID_NO)"
		SetSql	"where y.KEY_SYUKA_YMD='" & GetField("SyukaDt") & "'"
		SetSql	"and d.ChoCode='" & GetField("ChoCode") & "'"
		SetSql	"order by y.KEY_ID_NO"
		Debug strSql
		dim	objKiji
		set objKiji = objDB.Execute(strSql)
		strId		= ""
		strKiji1	= ""
		strKiji2	= ""
		strKiji3	= ""
		strKiji4	= ""
		strKiji5	= ""
		intDenCnt	= 0
		do while objKiji.EOF = False
			intDenCnt = intDenCnt + 1
			Debug ".GetKiji():" & intDenCnt & " " & objKiji.Fields("SyukaDt") & " " & objKiji.Fields("Id") & " " & objKiji.Fields("HIN_GAI") & " " & objKiji.Fields("Qty")
			select case intDenCnt
			case 1:
				strId = RTrim(objKiji.Fields("Id"))
				strKiji1 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "��"
				strKiji5 = RTrim(objKiji.Fields("Biko1"))
			case 2:
				strKiji2 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "��"
			case 3:
				strKiji3 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "��"
			case 4:
				strKiji4 = objKiji.Fields("HIN_GAI") & objKiji.Fields("Qty") & "��"
			case 5:
				strKiji4 = strKiji4 & " ��"
			case else
			end select
			objKiji.MoveNext
		loop
		strKiji5 = strKiji5 & Space(5) & "�`�[�F" & intDenCnt & "��"
		objKiji.Close
		set objKiji = Nothing
	End Function
	'-------------------------------------------------------------------
	'RSmile Update
	'-------------------------------------------------------------------
	Private	Function Update()
		Debug ".Update()"
		SetSql ""
		SetSql "update RSmile "
		SetSql "set SyukaDt = '" & GetField("SyukaDt") & "'"
		SetSql ",SCode = '" & GetField("ChoCode") & "'"
		SetSql ",STel = '" & GetField("ChoTel") & "'"
		SetSql ",SZip = '" & GetField("ChoZip") & "'"
		SetSql ",SAddress = '" & GetField("ChoAddress") & "'"
'		SetSql ",SBillding"
		SetSql ",SName1 = '" & GetField("ChoName") & "'"
'		SetSql ",SName2"
		SetSql ",SKiji1 = '" & strKiji1 & "'"
		SetSql ",SKiji2 = '" & strKiji2 & "'"
		SetSql ",SKiji3 = '" & strKiji3 & "'"
		SetSql ",SKiji4 = '" & strKiji4 & "'"
		SetSql ",SKiji5 = '" & strKiji5 & "'"
		SetSql ",UpdID = '" & strPrgId & "'"
		SetSql "where Id = '" & GetField("Id") & "'"
		SetSql "and ( SyukaDt <> '" & GetField("SyukaDt") & "'"
		SetSql "or SCode <> '" & GetField("ChoCode") & "'"
		SetSql "or STel <> '" & GetField("ChoTel") & "'"
		SetSql "or SZip <> '" & GetField("ChoZip") & "'"
		SetSql "or SAddress <> '" & GetField("ChoAddress") & "'"
		SetSql "or SName1 <> '" & GetField("ChoName") & "'"
		SetSql "or SKiji1 <> '" & strKiji1 & "'"
		SetSql "or SKiji2 <> '" & strKiji2 & "'"
		SetSql "or SKiji3 <> '" & strKiji3 & "'"
		SetSql "or SKiji4 <> '" & strKiji4 & "'"
		SetSql "or SKiji5 <> '" & strKiji5 & "'"
		SetSql ")"
		on error resume next
		objDb.Execute strSql
		WScript.StdOut.Write ":0x" & Hex(Err.Number) & " " & Err.Description
		on error goto 0
	End Function
	'-------------------------------------------------------------------
	'RSmile insert
	'-------------------------------------------------------------------
	Private Function InsertData()
		Debug ".InsertData()"
		if GetField("Id") <> "" then
			WScript.StdOut.Write ":" & GetField("Id")
			Update
			exit function
		end if
		SetSql ""
		SetSql "insert into RSmile ("
		SetSql "Id"
		SetSql ",SyukaDt"
		SetSql ",SCode"
		SetSql ",STel"
		SetSql ",SZip"
		SetSql ",SAddress"
		SetSql ",SBillding"
		SetSql ",SName1"
		SetSql ",SName2"
		SetSql ",SKiji1"
		SetSql ",SKiji2"
		SetSql ",SKiji3"
		SetSql ",SKiji4"
		SetSql ",SKiji5"
		SetSql ",EntID"
		SetSql ") values ("
		SetSql "'" & strId & "'"
		SetSql ",'" & GetField("SyukaDt") & "'"
		SetSql ",'" & GetField("ChoCode") & "'"
		SetSql ",'" & GetField("ChoTel") & "'"
		SetSql ",'" & GetField("ChoZip") & "'"
		SetSql ",'" & GetField("ChoAddress") & "'"
		SetSql ",''"		'SBillding"
		SetSql ",'" & GetField("ChoName") & "'"
		SetSql ",''"		'SName2"
		SetSql ",'" & strKiji1 & "'"	    'SKiji1"
		SetSql ",'" & strKiji2 & "'"	    'SKiji2"
		SetSql ",'" & strKiji3 & "'"	    'SKiji3"
		SetSql ",'" & strKiji4 & "'"	    'SKiji4"
		SetSql ",'" & strKiji5 & "'"	    'SKiji5"
		SetSql ",'" & strPrgId & "'"
		SetSql ")"
		Debug strSql
		on error resume next
		objDb.Execute strSql
		WScript.StdOut.Write ":0x" & Hex(Err.Number) & " " & Err.Description
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
	'-----------------------------------------------------------------------
	'SQL������ǉ�
	'-----------------------------------------------------------------------
	Private	strSql
	Public Function SetSql(byVal s)
		if s = "" then
			strSql = ""
		else
			if strSql <> "" then
				strSql = strSql & " "
			end if
			strSql = strSql & s
		end if
	End Function
	'-------------------------------------------------------------------
	'Sql���s
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
'		on error resume next
		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field�l
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		on error resume next
		strField = RTrim("" & objRs.Fields(strName))
		if Err.Number <> 0 then
			WScript.Echo "GetField(" & strName & "):Error:0x" & Hex(Err.Number)
			WScript.Quit
		end if
		on error goto 0
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
	End Function
	'-------------------------------------------------------------------
	'Field��
	'-------------------------------------------------------------------
	Public Function GetFields(byVal strTable)
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
	'-------------------------------------------------------------------
	'�S�p�����p
	'-------------------------------------------------------------------
	Private Function zen2han( byVal strVal )
		dim	objBasp
		Set objBasp = CreateObject("Basp21")
		zen2han = objBasp.StrConv( strVal, 8 )
		Set objBasp = Nothing
	End Function
	'-------------------------------------------------------------------
	'LenB()
	'-------------------------------------------------------------------
	Private Function LenB(byVal strVal)
	    Dim i, strChr
	    LenB = 0
	    If Trim(strVal) <> "" Then
	        For i = 1 To Len(strVal)
	            strChr = Mid(strVal, i, 1)
	            '�Q�o�C�g�����́{�Q
	            If (Asc(strChr) And &HFF00) <> 0 Then
	                LenB = LenB + 2
	            Else
	                LenB = LenB + 1
	            End If
	        Next
	    End If
	End Function
	'-------------------------------------------------------------------
	'Get_LeftB()
	'-------------------------------------------------------------------
	Private Function Get_LeftB(byVal a_Str,byVal a_int)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			Get_LeftB = ""
			Exit Function
		End If
		If a_int = 0 Then
			Get_LeftB = ""
			Exit Function
		End If
		For iCount = 1 to Len(a_Str)
			'** Asc�֐��ŕ����R�[�h�擾
			iAscCode = Asc(Mid(a_Str, iCount, 1))
			'** ���p�͕����R�[�h�̒�����2�A�S�p��4(2�ȏ�)�Ƃ��Ĕ��f
			If Len(Hex(iAscCode)) > 2 Then
				iLenCount = iLenCount + 2
			Else
				iLenCount = iLenCount + 1
			End If
			If iLenCount > Cint(a_int) Then
				Exit For
			Else
				iLeftStr = iLeftStr + Mid(a_Str, iCount, 1)
			End If
		Next
		Get_LeftB = iLeftStr
	End Function
	'-----------------------------------------------------------------------
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'���b�Z�[�W�\��
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'�I�v�V�����擾
	'-----------------------------------------------------------------------
	Function GetOption(byval strName _
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
	'-----------------------------------------------------------------------
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			Init = "�I�v�V�����G���[:" & strArg
			Disp Init
			Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "make"
				strAction = "make"
			case "csv"
				strAction = "csv"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
End Class
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objRSmile
	Set objRSmile = New RSmile
	if objRSmile.Init() <> "" then
		call usage()
		exit function
	end if
	call objRSmile.Run()
End Function
