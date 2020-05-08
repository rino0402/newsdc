Option Explicit
'-----------------------------------------------------------------------
'���C���ďo���C���N���[�h
'-----------------------------------------------------------------------
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
	dim	strScriptPath
	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strFileName = strScriptPath & strFileName
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
	Set fso = Nothing
End Function
Call Include("const.vbs")
Call Include("debug.vbs")
Call Include("excel.vbs")
Call Include("file.vbs")
Call Include("get_b.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "BO�o�׃f�[�^(Excel)�ϊ�"
	Wscript.Echo "BoSyukaX.vbs [option] <�t�@�C����> [�V�[�g��]"
	Wscript.Echo " /db:newsdc9"
	Wscript.Echo " /debug"
	Wscript.Echo "cscript BoSyukaX.vbs ""I:\0SDC_honsya\���ƕ��ʏ��i���o�׋��z�܂Ƃ�\�`�b�@�m�o�k����̏o�׎���\�y�P�Q���x�z�o�׎���_.xls"" /debug"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	dim	strStName
	strStName = ""
	'���O�����I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		elseif strStName = "" then
			strStName = strArg
		end if
	next
	'���O�t���I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "conv"
		case "?"
			strFilename = ""
		case else
			strFilename = ""
		end select
	next
	dim	strConv
	strConv = GetOption("conv","")
	if strConv <> "" then
		dim	objC
		select case strConv
		case "AcCZaiko"
			Set objC = New AcCZaiko
		case "AcYNyuka"
			Set objC = New AcYNyuka
		end select
		'-------------------------------------------------------------------
		'�f�[�^�x�[�X�̏���
		'-------------------------------------------------------------------
		dim	objDb
		Set objDb = OpenAdodb(GetOption("db","newsdc9"))

		Call DispMsg("�o�^��...InitRecord")
		Call objC.InitRecord(objDb)
		Call DispMsg("�o�^��...ContRecord")
		Call objC.ConvRecord(objDb)
		'-------------------------------------------------------------------
		'�f�[�^�x�[�X�̃N���[�Y
		'-------------------------------------------------------------------
		set objDb = CloseAdodb(objDb)
		Set objC = Nothing

		Main = 0
		exit function
	end if
	if strFilename = "" then
		usage()
		Main = 1
		exit Function
	end if
	call Load(strFilename,strStName)
	Main = 0
End Function

Function Load(byVal strFilename,byval strStName)
	'-------------------------------------------------------------------
	'Excel�t�@�C����
	'-------------------------------------------------------------------
	strFilename = GetAbsPath(strFilename)
	Call Debug("Load():" & strFilename & "," & strStName & "")
	if strFileName = "" then
		Call DispMsg("�t�@�C�������w�肵�ĉ�����")
		Exit Function
	end if
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	dim	objXL
	Set objXL = WScript.CreateObject("Excel.Application")
	Call Debug("CreateObject(Excel.Application)")
	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	dim	strPassword
	strPassword = ""
	dim	objBk
	Set objBk = objXL.Workbooks.Open(strFilename,False,True,,strPassword)
	Call Debug("Workbooks.Open=" & objBk.Name)
	'-------------------------------------------------------------------
	'�Ǎ�����
	'-------------------------------------------------------------------
	Dim	objSt
	For each objSt in objBk.Worksheets
		if strStName = "" or strStName = objSt.Name then
			Call LoadXls(objXL,objBk,objSt)
		end if
	Next
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("Load():End")
End Function

Function LoadXls(objXL,objBk,objSt)
	Call Debug("LoadXls():" & objSt.Name & "")
	'-------------------------------------------------------------------
	'�N���X
	'-------------------------------------------------------------------
	dim	objC
	dim	lngMaxRow
	lngMaxRow = -1
	dim i
	for i = 0 to 4
		select case i
		case 0
			Set objC = New AcSyuka
		case 1
			Set objC = New NrSyuka
		case 2
			Set objC = New NrSyFuri
		case 3
			Set objC = New AcCZaiko
		case 4
			Set objC = New AcYNyuka
		end select
		lngMaxRow = objC.CheckHead(objXL,objBk,objSt)
		if lngMaxRow > 0 then
			exit for
		end if
		Set objC = Nothing
	next
	if lngMaxRow <= 0 then
		Exit Function
	end if
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc9"))
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	Call objC.CreateTmp(objDb)
	dim	rsTmp
	set rsTmp = OpenRs(objDb,objC.pTableNameTmp)

	'-------------------------------------------------------------------
	'�Ǎ�
	'-------------------------------------------------------------------
	Call LoadSt(objXL,objBk,objSt,objDb,rsTmp,objC)

	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsTmp = CloseRs(rsTmp)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
	Set objC = Nothing
End Function

Function LoadSt(objXL,objBk,objSt,objDb,rsTmp,objC)
	Call Debug("LoadSt():" & objSt.Name)


	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
	dim	lngRow
	For lngRow = 1 to lngMaxRow
		Call DispMsg("�o�^��..." & objSt.Name & ":" & lngRow & "/" & lngMaxRow)
		Call objC.LoadRow(objXL,objBk,objSt,objDb,rsTmp,lngRow)
	Next

	Call DispMsg("�o�^��...InitRecord")
	Call objC.InitRecord(objDb)
	Call DispMsg("�o�^��...ContRecord")
	Call objC.ConvRecord(objDb)
End Function

'-------------------------------------------------------
' AC���ח\��20150625.xls
'-------------------------------------------------------
Class AcYNyuka
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "AcYNyuka"
        pTableNameTmp	= "AcYNyukaTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("AcYNyuka.CreateTmp()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcYNyukaTmp"
		Call Debug("AcYNyuka.CreateTmp():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcYNyukaTmp using 'AcYNyukaTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = 1 To 65
			c = getColumnName2(i)
			select case c
			case "A"
				strSql = strSql & " x" & c & " Char(60) default '' not null" & vbCrLf
			case else
				strSql = strSql & ",x" & c & " Char(60) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   xA" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcYNyuka.CreateTmp():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	End Sub
    Public Sub InitRecord(objDb)
		Call Debug("AcYNyuka.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcYNyuka"
		Call Debug("AcYNyuka.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcYNyuka using 'AcYNyuka.DAT' with replace (" & vbCrLf
		strSql = strSql & " ID		Char(12) default '' not null" & vbCrLf
		strSql = strSql & ",Stat	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",Wc		Char( 8) default '' not null" & vbCrLf
		strSql = strSql & ",WcName	Char(40) default '' not null" & vbCrLf
		strSql = strSql & ",Pn		Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",Qty		CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",NYDt	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",ORDt	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",SNDt	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   ID" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcYNyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into AcYNyuka "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(xA)"
		strSql = strSql & ",Left(RTrim(xB),1)"
		strSql = strSql & ",RTrim(xG)"
		strSql = strSql & ",RTrim(xH)"
		strSql = strSql & ",RTrim(xJ)"
		strSql = strSql & ",convert(xK,sql_decimal)"
		strSql = strSql & ",RTrim(xV)"
		strSql = strSql & ",RTrim(xW)"
		strSql = strSql & ",RTrim(xX)"
		strSql = strSql & " from AcYNyukaTmp"
		Call Debug("AcYNyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("AcYNyuka.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"Sheet1") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"�����[���Ǘ��ԍ�") then
			Exit Function
		end if
		if CompHead(objSt.Range("B1"),"�i���@�T�[�r�X�f�[�^�i���敪") then
			Exit Function
		end if
		if CompHead(objSt.Range("C1"),"���Ə�R�[�h") then
			Exit Function
		end if
		if CompHead(objSt.Range("D1"),"��v���Ɓ@��v�p���Ə�R�[�h") then
			Exit Function
		end if
		if CompHead(objSt.Range("E1"),"���Y���Ɓ@���Y�Ǘ����Ə�R�[�h") then
			Exit Function
		end if
		if CompHead(objSt.Range("F1"),"���B��敪") then
			Exit Function
		end if
		if CompHead(objSt.Range("G1"),"�d����WC�R�[�h") then
			Exit Function
		end if
		if CompHead(objSt.Range("H1"),"�d����WC����") then
			Exit Function
		end if
		if CompHead(objSt.Range("I1"),"�d����i�ڔԍ�") then
			Exit Function
		end if
		if CompHead(objSt.Range("J1"),"�i�ڔԍ�") then
			Exit Function
		end if
		if CompHead(objSt.Range("K1"),"���o�ɗ\�萔") then
			Exit Function
		end if
		if CompHead(objSt.Range("BM1"),"�X�V��") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("AcYNyuka.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("AcYNyuka.LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 3 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = 1 To 65
			c = GetColumnName2(i)
			Call SetField(rsTmp,objSt,"x" & c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ں��ނ̷� ̨���ނɏd�����鷰�l������܂�(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' AC�Z���^�[�݌�150625.xlsx,
'-------------------------------------------------------
Class AcCZaiko
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "AcCZaiko"
        pTableNameTmp	= "AcCZaikoTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("AcCZaiko.CreateTmp()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcCZaikoTmp"
		Call Debug("AcCZaiko.CreateTmp():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcCZaikoTmp using 'AcCZaikoTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("J")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " " & c & " Char(20) default '' not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & "  ,B" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcCZaiko.CreateTmp():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	End Sub
    Public Sub InitRecord(objDb)
		Call Debug("AcCZaiko.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcCZaiko"
		Call Debug("AcCZaiko.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcCZaiko using 'AcCZaiko.DAT' with replace (" & vbCrLf
		strSql = strSql & " JCode	Char( 8) default '' not null" & vbCrLf
		strSql = strSql & ",Pn 		Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",Tanto	Char(10) default '' not null" & vbCrLf
		strSql = strSql & ",C_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000440_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000441_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000443_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000444_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000446_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",T_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   JCode" & vbCrLf
		strSql = strSql & "  ,Pn" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcCZaiko.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into AcCZaiko "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(A)"
		strSql = strSql & ",RTrim(B)"
		strSql = strSql & ",RTrim(C)"
		strSql = strSql & ",convert(D,sql_decimal)"
		strSql = strSql & ",convert(E,sql_decimal)"
		strSql = strSql & ",convert(F,sql_decimal)"
		strSql = strSql & ",convert(G,sql_decimal)"
		strSql = strSql & ",convert(H,sql_decimal)"
		strSql = strSql & ",convert(I,sql_decimal)"
		strSql = strSql & ",convert(J,sql_decimal)"
		strSql = strSql & " from AcCZaikoTmp"
		Call Debug("AcCZaiko.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("AcCZaiko.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"�݌�") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"���Ə�CD") then
			Exit Function
		end if
		if CompHead(objSt.Range("B1"),"�i�ڔԍ�") then
			Exit Function
		end if
		if CompHead(objSt.Range("C1"),"�w���S����CD") then
			Exit Function
		end if
		if CompHead(objSt.Range("D1"),"�Z���^�[�q��") then
			Exit Function
		end if
		if CompHead(objSt.Range("J1"),"���݌�") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("AcCZaiko.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("AcCZaiko.LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 2 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("J")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ں��ނ̷� ̨���ނɏd�����鷰�l������܂�(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

'-------------------------------------------------------
' Ac�o�׎���
'-------------------------------------------------------
Class AcSyuka
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "BoSyuka"
        pTableNameTmp	= "AcSyukaTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("AcSyuka.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcSyukaTmp"
		Call Debug("AcSyuka.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcSyukaTmp using 'AcSyukaTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("L")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " A UBIGINT  default 0  not null" & vbCrLf
			case "G"
				strSql = strSql & "," & c & " Char(40) default '' not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "Create Unique Index AcSyukaTmp_Key01 On AcSyukaTmp (" & vbCrLf
		strSql = strSql & "   B" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	End Sub
    Public Sub InitRecord(objDb)
'		Call Debug("delete from " & pTableNameTmp)
'		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = "delete from BoSyuka"
		strSql = strSql & " where ShisanJCode in (select distinct RTrim(C) from AcSyukaTmp)"
		strSql = strSql & "   and Left(Dt,6) in (select distinct Left(J,6) from AcSyukaTmp)"
		Call Debug("AcSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "insert into BoSyuka "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(B)"					'// ���x�U�ߎ�_���x�U�֊Ǘ��ԍ�
		strSql = strSql & ",RTrim(C)"					'// ���x�U�ߎ�_���Y�Ǘ����Ə�R�[�h
		strSql = strSql & ",''"							'// ���x�U�ߎ�_���x�U�փR�[�h
		strSql = strSql & ",RTrim(H)"					'// ���x�U�ߎ�_���o�Ɏ���敪
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�݌Ɏ��x������
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�݌Ɏ��x�R�[�h
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�`�[�ԍ�
		strSql = strSql & ",RTrim(E)"					'// ���x�U�ߎ�_�i�ڔԍ�
		strSql = strSql & ",convert(L,sql_decimal)"		'// ���x�U�ߎ�_���o�Ɏ��ѐ�
		strSql = strSql & ",RTrim(J)"					'// ���x�U�ߎ�_���x�U���N����
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�U�֐�݌Ɏ��x������
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�U�֐�݌Ɏ��x�R�[�h
		strSql = strSql & ",''"							'// �݌Ɏ��x_�q�ɃR�[�h
		strSql = strSql & " from AcSyukaTmp"
		Call Debug("AcSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("AcSyuka.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"�O���o�ז���") then
			if CompHead(objSt.Name,"�����o�ז���") then
				Exit Function
			end if
			if objSt.Range("A1") = "" then
				objSt.Range("A1") = "NO"
			end if
		end if
		if CompHead(objSt.Range("A1"),"NO") then
			Exit Function
		end if
		if CompHead(objSt.Range("L1"),"�o�׎��ѐ�") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("AcSyuka.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("AcSyuka.LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 2 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("L")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ں��ނ̷� ̨���ނɏd�����鷰�l������܂�(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

'-------------------------------------------------------
' �ޗǓ��q�� �o�׎���
'-------------------------------------------------------
Class NrSyuka
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "BoSyuka"
        pTableNameTmp	= "NrSyukaTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("NrSyuka.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table NrSyukaTmp"
		Call Debug("NrSyuka.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table NrSyukaTmp using 'NrSyukaTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " A UBIGINT  default 0  not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("NrSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "Create Unique Index NrSyukaTmp_Key01 On NrSyukaTmp (" & vbCrLf
		strSql = strSql & "   C" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("NrSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub InitRecord(objDb)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = "delete from BoSyuka"
		strSql = strSql & " where ShisanJCode in (select distinct RTrim(D) from NrSyukaTmp)"
		strSql = strSql & "   and Left(Dt,6) in (select distinct Left(L,6) from NrSyukaTmp)"
		Call Debug("NrSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "insert into BoSyuka "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(C)"					'// ���x�U�ߎ�_���x�U�֊Ǘ��ԍ�
		strSql = strSql & ",RTrim(D)"					'// ���x�U�ߎ�_���Y�Ǘ����Ə�R�[�h
		strSql = strSql & ",''"							'// ���x�U�ߎ�_���x�U�փR�[�h
		strSql = strSql & ",RTrim(G)"					'// ���x�U�ߎ�_���o�Ɏ���敪
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�݌Ɏ��x������
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�݌Ɏ��x�R�[�h
		strSql = strSql & ",RTrim(F)"					'// ���x�U�ߎ�_�`�[�ԍ�
		strSql = strSql & ",RTrim(E)"					'// ���x�U�ߎ�_�i�ڔԍ�
		strSql = strSql & ",convert(M,sql_decimal)"		'// ���x�U�ߎ�_���o�Ɏ��ѐ�
		strSql = strSql & ",RTrim(L)"					'// ���x�U�ߎ�_���x�U���N����
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�U�֐�݌Ɏ��x������
		strSql = strSql & ",''"							'// ���x�U�ߎ�_�U�֐�݌Ɏ��x�R�[�h
		strSql = strSql & ",RTrim(B)"					'// �݌Ɏ��x_�q�ɃR�[�h
		strSql = strSql & " from NrSyukaTmp"
		Call Debug("NrSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("NrSyuka.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"���q��_�o�׎��і���") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"�����q��_�o�׎��і���") then
			Exit Function
		end if
		if CompHead(objSt.Range("A2"),"NO") then
			Exit Function
		end if
		if CompHead(objSt.Range("M2"),"�o�׎��ѐ�") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("NrSyuka.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 3 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ں��ނ̷� ̨���ނɏd�����鷰�l������܂�(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

'-------------------------------------------------------
' �ޗǓ��q�� ���x�U�֖���
'-------------------------------------------------------
Class NrSyFuri
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "BoSyFuri"
        pTableNameTmp	= "NrSyFuriTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("NrSyFuri.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table NrSyFuriTmp"
		Call Debug("NrSyFuri.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table NrSyFuriTmp using 'NrSyFuriTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " " & c & " Char(20) default '' not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & "  ,I" & vbCrLf
		strSql = strSql & "  ,K" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("NrSyFuri.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
'		strSql = ""
'		strSql = strSql & "Create Unique Index NrSyFuriTmp_Key01 On NrSyFuriTmp (" & vbCrLf
'		strSql = strSql & "   C" & vbCrLf
'		strSql = strSql & ")" & vbCrLf
'		Call Debug("NrSyFuri.InitRecord():" & strSql)
'		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub InitRecord(objDb)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = "delete from BoSyFuri"
		strSql = strSql & " where ShisanJCode in (select distinct RTrim(C) from NrSyFuriTmp)"
		strSql = strSql & "   and Left(Dt,6) in (select distinct Left(Replace(L,'/',''),6) from NrSyFuriTmp)"
		Call Debug("NrSyFuri.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "insert into BoSyFuri "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(A)"					'// �����Ǘ��ԍ�	500380470-40
		strSql = strSql & ",RTrim(B)"					'// �q�ɖ�	�ޗǓ�
		strSql = strSql & ",RTrim(C)"					'// ���Ə�CD	00021184
		strSql = strSql & ",RTrim(D)"					'// �i�ڔԍ�	A0601-1E60S
		strSql = strSql & ",RTrim(E)"					'// �Ɩ���	����:���i��
		strSql = strSql & ",RTrim(F)"					'// ���x�U��CD	2611B
		strSql = strSql & ",RTrim(G)"					'// ���xCD	11B
		strSql = strSql & ",RTrim(H)"					'// ���x����	11N1
		strSql = strSql & ",RTrim(I)"					'// ����敪	45
		strSql = strSql & ",RTrim(J)"					'// �i���敪	7
		strSql = strSql & ",RTrim(K)"					'// �`�[�ԍ�	005628
		strSql = strSql & ",Replace(RTrim(L),'/','')"	'// ���єN����	2015/01/10
		strSql = strSql & ",convert(M,sql_decimal)"		'// ��	2,000
		strSql = strSql & " from NrSyFuriTmp"
		Call Debug("NrSyFuri.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("NrSyFuri.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"���q��_���x�U�֖���") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"�����q��_���x�U�֖���") then
			Exit Function
		end if
		if CompHead(objSt.Range("A2"),"�����Ǘ��ԍ�") then
			Exit Function
		end if
		if CompHead(objSt.Range("M2"),"��") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("NrSyFuri.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("NrSyFuriLoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 3 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ں��ނ̷� ̨���ނɏd�����鷰�l������܂�(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

Function CompHead(byval strV,strTitle)
	Call Debug("CompHead():" & strV & ":" & strTitle)
	if strV = strTitle then
		Call Debug("CompHead():====")
		CompHead = 0
		Exit Function
	end if
	strV = Replace(strV,vbCrLf,"")
	if strV = strTitle then
		CompHead = 0
		Call Debug("CompHead():====CrLf")
		Exit Function
	end if
	strV = Replace(strV,vbLf,"")
	if strV = strTitle then
		CompHead = 0
		Call Debug("CompHead():====Lf")
		Exit Function
	end if
	Call Debug("CompHead():<><>")
	CompHead = 1
End Function

Function SetField(objRs,objSt,byVal strField,byVal strCol,byVal lngRow)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & objSt.Range(strCol & lngRow))
	dim	v
	v = RTrim(objSt.Range(strCol & lngRow))
	select case strField
	case "JDt"
		v = Replace(v,"/","")
	case "SDt","DDt"	'//  I:J ����� '//  K:L �[����
		v = v & "/" & objSt.Range(strCol & lngRow).Offset(0,1)
		if isDate(v) then
			v = CDate(v)
			v = Replace(v,"/","")
		else
			v = ""
		end if
	case "Amount","AmountEHN"
		if isNumeric(v) <> True then
			v = 0
		end if
		v = CCur(v)
	end select
	v = Get_LeftB(v,objRs.Fields(strField).DefinedSize)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & v)
	objRs.Fields(strField) = v
End Function

Function getColumnName2(ByVal pColumn)

    Const cStartChr = 65 'A�̕����R�[�h
    Const cAlphabet = 26 '�A���t�@�x�b�g�̎��

    Dim lColumnNum
    Dim sColumnName

    If pColumn < 1 Then

       getColumnName2 = "??"
       Exit Function

    End If

    Do

       lColumnNum = (pColumn - 1) Mod cAlphabet
       sColumnName = sColumnName & Chr(cStartChr + lColumnNum)

       pColumn = (pColumn - 1) \ cAlphabet

    Loop Until pColumn = 0

    getColumnName2 = StrReverse(sColumnName)

End Function
