Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "b2data.vbs [option]"
	Wscript.Echo " /db:newsdc1	�f�[�^�x�[�X"
	Wscript.Echo " /make �����f�[�^�쐬(default)"
	Wscript.Echo " /csv  �����f�[�^�o��"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript b2data.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'B2Data
'2016.10.06 �ǉ��F�z�B���ԑы敪
'           0812 �ߑO��
'�@�@�@�@�@ 1416 S5YN30		�Y�@�Q�n�߰󒍊Ǘ�
'�@�@�@�@�@�@�@�@5YE2S7000	�e�N�m�V�X�e���S�R
'2016.10.07 �Z���̌������𕪊� <1770��2F> <ø�WING510> <ø�WING503>
'           ��Ж��𕪊� <�o�d�r�Y�@�V�X�e���i���j�f�B���C�g�@�_��>
'           �R���ȏ�ŕi���Q��<��>������悤�ɏC��
'           �ڋq�Ǘ�No�� <�t�@�C������>-<��> �ɕύX
'2016.10.25 R-smile(SSX)�Ή�
'2017.06.29 �`�[�����^�ː��Z�b�g�������xUp
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

Class B2Data
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
		Call SetSql("from b2data")
		Call SetSql("where SyukaDt = '" & strSyukaDt & "'")
'		Call SetSql("where SyukaDt in ('20180904','20180905')") '
		Call SetSql("order by")
		Call SetSql(" SyukaDt")
		Call SetSql(",ClientNo")
		Call SetSql(",EntTm")
		Call SetSql(",SCode")
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
	'YoteiTime() �z�B���ԑы敪
	'-------------------------------------------------------------------
	'���z�B���ԑы敪
	'�z�B���ԑт��w�肵�܂��B
	'���p4����
	'�^�C���A�c�l�ցA�l�R�|�X�ȊO
	' 0812 : �ߑO��
	' 1214 : 12�`14��
	' 1416 : 14�`16��
	' 1618 : 16�`18��
	' 1820 : 18�`20��
	' 2021 : 20�`21��
	'�^�C���̂�
	' 0010 : �ߑO10���܂�
	' 0017 : �ߌ�5���܂�
	Private	Function YoteiTime()
		select case GetField("SCode")
		'S5YN30		�Y�@�Q�n�߰󒍊Ǘ�
		'5YE2S7000	�e�N�m�V�X�e���S�R
		case "S5YN30" _
			,"5YE2S7000"
			YoteiTime = "1416"
		case else
			YoteiTime = "0812"
		end select
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
			for each objF in objRs.Fields
				WScript.StdOut.Write objF.Name
				WScript.StdOut.Write ","
			next
			WScript.StdOut.Write "�\���"
			WScript.StdOut.Write ",�z�B���ԑ�"
			WScript.StdOut.Write ",RsTel"
			WScript.StdOut.Write ",Rs����"
			WScript.StdOut.Write ",Rs�Ԓn"
			WScript.StdOut.Write ",Rs����1"
			WScript.StdOut.Write ",Rs����2"
			WScript.StdOut.Write ",Rs�ב��l"
			WScript.StdOut.Write ",Rs�z�B�w���"
			WScript.StdOut.Write ",Rs�q��敪"
			WScript.StdOut.Write ",Rs�L��1"
			WScript.StdOut.Write ",Rs�L��2"
			WScript.StdOut.Write ",Rs�L��3"
			WScript.StdOut.Write ",Rs�L��4"
			WScript.StdOut.Write ",Rs�L��5"
			WScript.StdOut.Write ",Rs�o�הԍ�"
			WScript.StdOut.WriteLine
			exit function
		end if
		'����
		for each objF in objRs.Fields
			WScript.StdOut.Write Replace(GetField(objF.Name),",",".")
			WScript.StdOut.Write ","
		next
		WScript.StdOut.Write YoteiDt()
		WScript.StdOut.Write "," & YoteiTime()
		WScript.StdOut.Write "," & RsCsv("RsTel")
		WScript.StdOut.Write "," & RsCsv("Rs����")
		WScript.StdOut.Write "," & RsCsv("Rs�Ԓn")
		WScript.StdOut.Write "," & RsCsv("Rs����1")
		WScript.StdOut.Write "," & RsCsv("Rs����2")
		WScript.StdOut.Write "," & RsCsv("Rs�ב��l")
		WScript.StdOut.Write "," & RsCsv("Rs�z�B�w���")
		WScript.StdOut.Write "," & RsCsv("Rs�q��敪")
		WScript.StdOut.Write "," & RsCsv("Rs�L��1")
		WScript.StdOut.Write "," & RsCsv("Rs�L��2")
		WScript.StdOut.Write "," & RsCsv("Rs�L��3")
		WScript.StdOut.Write "," & RsCsv("Rs�L��4")
		WScript.StdOut.Write "," & RsCsv("Rs�L��5")
		WScript.StdOut.Write "," & RsCsv("Rs�o�הԍ�")
		WScript.StdOut.WriteLine
	End Function
	'-----------------------------------------------------------------------
	'R-smile�p(CSV)
	'-----------------------------------------------------------------------
	Private Function RsCsv(byVal strName)
		dim	strValue
		strValue = ""
		select case strName
		case "RsTel"
			strValue = GetField("STel")
		case "Rs����"
		case "Rs�Ԓn"
		case "Rs����1"	' 60
			strValue = GetField("SCampany1") & GetField("SCampany2") & GetField("SName")
		case "Rs����2"
		case "Rs�ב��l"
			strValue = RsSender()
		case "Rs�z�B�w���"
			strValue = YoteiDt()
		case "Rs�q��敪"
			select case RsSender()
			case 7,8
				strValue = "�`�h�q"
			end select
		case "Rs�L��1"
			strValue = GetField("HinName1")
		case "Rs�L��2"
			strValue = GetField("HinName2")
		case "Rs�L��3"
			strValue = GetField("Kiji")
		case "Rs�L��4"
		case "Rs�L��5"
		case "Rs�o�הԍ�"	'15��
							'20
'			strValue = Replace(GetField("ClientNo"),"-","")
			strValue = GetField("SCode")
		end select
		RsCsv = strValue
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
	'strClientNo
	'-----------------------------------------------------------------------
	Private	intBin
    Public Function B2ClientNo()
		Debug ".B2ClientNo()"
		dim	strToday
		strToday = Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2)
		Call SetSql("")
		Call SetSql("select")
		Call SetSql("distinct")
		Call SetSql("Filename")
		Call SetSql("from HMTAH015_t")
		Call SetSql("where Filename like 'HMTAH015SZZ.dat." & strToday & "-%'")
		Call SetSql("order by")
		Call SetSql(" Filename")
		Debug ".B2ClientNo():" & strSql
		set objRs = objDB.Execute(strSql)
		strClientNo = ""
		intBin = 0
		do while objRs.Eof = false
			intBin = intBin + 1
			strClientNo = GetField("Filename")
			'HMTAH015SZZ.dat.20161007-062300
			strClientNo = Split(strClientNo,".")(2)
			objRs.MoveNext
		loop
		strClientNo = strClientNo & "-" & intBin
		objRs.Close
	End Function

	'-----------------------------------------------------------------------
	'Make() �����f�[�^�쐬
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		Call B2ClientNo()

		Call SetSql("")
		Call SetSql("select")
		Call SetSql("y.KEY_SYUKA_YMD SyukaDt")
		Call SetSql(",y.KEY_HIN_NO Pn")
		Call SetSql(",convert(y.SURYO,SQL_DECIMAL) Qty")
		Call SetSql(",y.bikou1 Biko1")
		Call SetSql(",d.ChoCode ChoCode")
		Call SetSql(",d.ChoName ChoName")
		Call SetSql(",d.ChoAddress ChoAddress")
		Call SetSql(",d.ChoTel ChoTel")
		Call SetSql(",d.ChoZip ChoZip")
		Call SetSql("from y_syuka y")
		Call SetSql("inner join HtDrctId d on (d.IDNo = y.KEY_ID_NO)")
'		Call SetSql("where Aitesaki	=	'00027768'")
'		Call SetSql(  "and ChoCode	<>	''")
'		Call SetSql(  "and Stts		=	'4'")
'		Call SetSql(  "and TMark	<>	'T'")
'		Call SetSql(  "and SyukaDt	=	'20160913'")
		Call SetSql("order by")
		Call SetSql(" SyukaDt")
'		Call SetSql(",Aitesaki")
		Call SetSql(",ChoCode")
		Debug ".Make():" & strSql
		set objRs = objDB.Execute(strSql)
		prvSyukaDt = ""
		prvChoCode = ""
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
	Private	prvSyukaDt
	Private	prvChoCode
	Private	intCnt
	Private	Function Count()
		intCnt = intCnt + 1
		if strSyukaDt <> prvSyukaDt _
		or strChoCode <> prvChoCode then
			intCnt = 1
		end if
		prvSyukaDt = strSyukaDt
		prvChoCode = strChoCode
	End Function
	'-------------------------------------------------------------------
	'MakeData() 1�s�Ǎ�
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData():" & GetField("SyukaDt") & " " & GetField("ChoCode") & " " & GetField("ChoName")
		strSyukaDt = GetField("SyukaDt")
		strChoCode = GetField("ChoCode")
		strChoName = GetField("ChoName")
'		strClientNo	= GetField("Aitesaki")
		strSCode	= strChoCode
		Call Count()
		Call GetB2Data()
		Call SetB2Data()
	End Function
	'-------------------------------------------------------------------
	'B2Data�t�B�[���h�Z�b�g
	'-------------------------------------------------------------------
	Private Function SetB2DataField(byVal strName,byVal strValue)
		dim	intLen
		intLen = objB2DataRs.Fields(strName).DefinedSize
		Debug ".SetB2DataField():" & strName & "(" & intLen & "):" & strValue
		strValue = Get_LeftB(strValue,intLen)
		Debug ".SetB2DataField():" & strName & "(" & intLen & "):" & strValue
		if objB2DataRs.Fields(strName) = strValue then
			exit function
		end if
		objB2DataRs.Fields(strName) = strValue
		objB2DataRs.Fields("UpdID")	= "b2data." & intCnt
	End Function
	'-------------------------------------------------------------------
	'���͂��於�i�����j���p32
	'���͂����ЁE���喼�P	���p50
	'���͂����ЁE���喼�Q	���p50
	'-------------------------------------------------------------------
	Private	strSName
	Private	strSCampany1
	Private	strSCampany2
	Private Function SetB2DataSName()
		strSName = GetField("ChoName")
		strSCampany1 = ""
		strSCampany2 = ""
		dim	intLen
		intLen = LenB(strSName)
		if intLen <= 32 then
			Exit Function
		end if

		select case strSName
		case "�o�d�r�Y�@�V�X�e���i���j�f�B���C�g�@�_��"
			strSCampany2	= "�o�d�r�Y�@�V�X�e���i���j"
			strSName		= "�f�B���C�g�@�_��"
			Exit Function
		end select

		dim	aryWord
		aryWord = ""
		strSName = Replace(strSName,"�@"," ")
		if inStr(strSName," ") > 0 then
			aryWord = Split(strSName," ")
		end if
		if isArray(aryWord) then
			dim	strWord
			for each strWord in aryWord
				strWord = Trim(strWord)
				if strSCampany2 = "" then
					strSCampany2 = strWord
					strSName = ""
				else
					if strSName <> "" then
						strSName = strSName & " "
					end if
					strSName = strSName & strWord
				end if
			next
		end if
		if LenB(strSName) <= 32 then
			Exit Function
		end if
		strSName = zen2han(strSName)
	End Function
	'-------------------------------------------------------------------
	'B2Data�Z������
	'-------------------------------------------------------------------
	Private	strAddress
	Private	strBillding
	Private	Function SetB2DataAddress()
		strAddress	= GetField("ChoAddress")
		strBillding	= ""
		select case strAddress
		case "�Q�n���W�y�S��򒬍�c1-1-1_1770��2F"
			strAddress	= "�Q�n���W�y�S��򒬍�c1-1-1"
			strBillding	= "1770��2F"
		case "�����s��c��{�H�c2����12�\1ø�WING510"
			strAddress	= "�����s��c��{�H�c2����12-1"
			strBillding	= "ø�WING510"
		case "�����s��c��{�H�c2����12�\1ø�WING503"
			strAddress	= "�����s��c��{�H�c2����12-1"
			strBillding	= "ø�WING503"
		end select
	End Function
	'-------------------------------------------------------------------
	'�`�[�����A�ː�
	'-------------------------------------------------------------------
	Private	intDenCnt
	Private	dblSaisu
	Private Function SetB2DataSaisu()
		'�x��
		'Call SetSql("")
		'Call SetSql("select")
		'Call SetSql("Count(*) c")
		'Call SetSql(",Sum(convert(y.SURYO,SQL_DECIMAL) * convert(i.SAI_SU,sql_decimal)) s")
		'Call SetSql("from HtDrctId d")
		'Call SetSql("inner join y_syuka y on (d.IDNo = y.KEY_ID_NO)")
		'Call SetSql("inner join Item i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)")
		'Call SetSql("where y.KEY_SYUKA_YMD='" & GetField("SyukaDt") & "'")
		'Call SetSql("and d.ChoCode='" & GetField("ChoCode") & "'")
		'�x���Ȃ�
		SetSql ""
		SetSql "select"
		SetSql "Count(*) c"
		SetSql ",Sum(convert(y.SURYO,SQL_DECIMAL) * convert(i.SAI_SU,sql_decimal)) s"
		SetSql "from y_syuka y"
		SetSql "inner join Item i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)"
		SetSql "where y.KEY_SYUKA_YMD='" & GetField("SyukaDt") & "'"
		SetSql "and y.KEY_ID_NO in"
		SetSql " (select distinct IDNo from HtDrctId where ChoCode = '" & GetField("ChoCode") & "')"

		dim	objSaisu
		Debug ".SetB2DataSaisu().call"
		set objSaisu = objDB.Execute(strSql)
		Debug ".SetB2DataSaisu().done"
		intDenCnt	= 0
		dblSaisu	= 0
		if objSaisu.EOF = False then
			intDenCnt	= objSaisu.Fields("c")
			dblSaisu	= objSaisu.Fields("s")
		end if
		objSaisu.Close
		set objSaisu = Nothing
	End Function
	'-------------------------------------------------------------------
	'B2Data���R�[�h�Z�b�g
	'-------------------------------------------------------------------
	Private Function SetB2Data()
		Debug ".SetB2Data():" & intCnt
		if intCnt = 1 then
'			Call SetB2DataField("ClientNo"	,GetField("Aitesaki"))
			Call SetB2DataField("STel"		,GetField("ChoTel"))
			Call SetB2DataSName()
			Call SetB2DataField("SName"		,strSName)
			Call SetB2DataField("SCampany1"	,strSCampany1)
			Call SetB2DataField("SCampany2"	,strSCampany2)
			Call SetB2DataAddress()
			Call SetB2DataField("SZip"		,GetField("ChoZip"))
			Call SetB2DataField("SAddress"	,strAddress)
			Call SetB2DataField("SBillding"	,strBillding)
			Call SetB2DataField("HinCode1"	,GetField("Pn"))
			Call SetB2DataField("HinName1"	,GetField("Pn") & " " & GetField("Qty") & "��")
			Call SetB2DataField("HinCode2"	,"")
			Call SetB2DataField("HinName2"	,"")
			Call SetB2DataSaisu()
'			Call SetB2DataField("Kiji"		,GetField("Biko1") & " " & intDenCnt & "��(" & dblSaisu & "��)")
			Call SetB2DataField("Kiji"		,GetField("Biko1") & " �`�[�F" & intDenCnt & "��")
		elseif intCnt = 2 then
			Call SetB2DataField("HinCode2"	,GetField("Pn"))
			Call SetB2DataField("HinName2"	,GetField("Pn") & " " & GetField("Qty") & "��")
		elseif intCnt = 3 then
			Call SetB2DataField("HinName2"	,RTrim(objB2DataRs.Fields("HinName2")) & " ��")
			Debug ".SetB2Data()��:" & objB2DataRs.Fields("HinName2")
		end if
'		Call SetB2DataField("Kiji"		,GetField("ChoCode") & " �`�[�F" & intCnt & "��")
'		Call SetB2DataField("Kiji"		,GetField("Biko1") & " �`�[�F" & intCnt & "��")

		Call DispB2Data()
		Call objB2DataRs.Update()
	End Function
	'-------------------------------------------------------------------
	'B2Data�\��
	'-------------------------------------------------------------------
	dim	strDisp
	Private Function DispB2Data()
		strDisp = ""
		strDisp = strDisp & strSyukaDt
		strDisp = strDisp & " " & strClientNo
		strDisp = strDisp & " " & strChoCode
		strDisp = strDisp & " " & strChoName
		strDisp = strDisp & " " & strSCode
		strDisp = strDisp & ":" & intCnt
		Disp strDisp
	End Function
	'-------------------------------------------------------------------
	'B2Data���R�[�h
	'-------------------------------------------------------------------
	dim	objB2DataRs
	Private Function GetB2Data()
		Debug ".GetB2Data()"
		Call SetSql("")
		Call SetSql("select")
		Call SetSql("*")
		Call SetSql("from B2Data")
		Call SetSql("where SyukaDt	=	'" & strSyukaDt & "'")
'		Call SetSql(  "and ClientNo	=	'" & strClientNo & "'")
		Call SetSql(  "and SCode	=	'" & strSCode & "'")
		Debug ".GetB2Data():" & strSql
		set objB2DataRs = Nothing
		Set objB2DataRs = Wscript.CreateObject("ADODB.Recordset")
		Call objB2DataRs.Open(strSql, objDb, adOpenKeyset, adLockOptimistic)
		if objB2DataRs.Eof = False then
			Exit Function
		end if
		Call objB2DataRs.Close
		Call objB2DataRs.Open("B2Data", objDb, adOpenKeyset, adLockOptimistic,adCmdTableDirect)
		Call objB2DataRs.AddNew
		objB2DataRs.Fields("SyukaDt")	= strSyukaDt
		objB2DataRs.Fields("ClientNo")	= strClientNo
		objB2DataRs.Fields("SCode")		= strSCode
		objB2DataRs.Fields("EntID")		= "b2data.vbs"
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
			WScript.StdErr.WriteLine "GetField(" & strName & "):Error:0x" & Hex(Err.Number)
'			WScript.Quit
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
			Wscript.Echo strMsg
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
	dim	objB2Data
	Set objB2Data = New B2Data
	if objB2Data.Init() <> "" then
		call usage()
		exit function
	end if
	call objB2Data.Run()
End Function
