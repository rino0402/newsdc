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

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "���ѕ\��(Excel)�f�[�^�ϊ�"
	Wscript.Echo "YOrder.vbs [option] <�t�@�C����> [�V�[�g��]"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo "sc32 YOrder.vbs /debug F:\��������\���ѕ\��.xlsm 201406"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	dim	strSheetname
	strFilename = ""
	strSheetname = ""
	'���O�����I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		else
			strSheetname = strArg
		end if
	next
	'���O�t���I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			strFilename = ""
		case else
			strFilename = ""
		end select
	next
	if strFilename = "" then
		usage()
		Main = 1
		exit Function
	end if
	call LoadYOrder(strFilename,strSheetname)
	Main = 0
End Function

Function LoadYOrder(byVal strFilename,byVal strSheetname)
	Call Debug("LoadYOrder(" & strFilename & "," & strSheetname & ")")
	'-------------------------------------------------------------------
	'Excel�t�@�C����
	'-------------------------------------------------------------------
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
	Call LoadYOrderXls(objXL,objBk,strSheetname)
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadYOrder():End")
End Function

Function LoadYOrderXls(objXL,objBk,byVal strSheetname)
	Call Debug("LoadYOrderXls(" & strSheetname & ")")
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	rsYOrder
	set rsYOrder = OpenRs(objDb,"YOrder")

	Call Debug("LoadYOrderXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadYOrderXls():SheetName=" & strShtName)
		if strSheetname = "" or strSheetname = strShtName then
			Call LoadYOrderXst(objXL,objBk,objSt,objDb,rsYOrder)
		end if
	Next
	
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsYOrder = CloseRs(rsYOrder)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadYOrderXst(objXL,objBk,objSt,objDb,rsYOrder)
	Call Debug("LoadYOrderXst():SheetName=" & objSt.Name)

	dim	strYM
	strYM = chkYM(objSt.Name)
	if strYM = "" then
		Exit Function
	end if

	Call Debug("delete strYM:" & strYM)
	Call ExecuteAdodb(objDb,"delete from YOrder where YM = '" & strYM & "'")

	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"G",lngMaxRow)
	lngMaxRow = excelGetMaxRow(objSt,"H",lngMaxRow)
	dim	lngRow
	For lngRow = 5 to lngMaxRow
		Call LoadYOrderRow(strYM,objXL,objBk,objSt,objDb,rsYOrder,lngRow)
	Next
End Function

Function LoadYOrderRow(byVal strYM,objXL,objBk,objSt,objDb,rsYOrder,byVal lngRow)
	Call Debug("LoadYOrderRow():" & objSt.Name & ":" & lngRow)
	rsYOrder.AddNew
	rsYOrder.Fields("SName")	= objSt.Name					'// �V�[�g��
	rsYOrder.Fields("Row")		= lngRow						'// �s�ԍ�
	rsYOrder.Fields("YM")		= strYM							'//  �����N��(�V�[�g��)
	Call SetField(rsYOrder,objSt,"JDt",			"G" , lngRow)	'//  G �󒍓�
	Call SetField(rsYOrder,objSt,"DenNo",		"H" , lngRow)	'//  H ����No
	Call SetField(rsYOrder,objSt,"SDt",			"I" , lngRow)	'//  I:J �����
	Call SetField(rsYOrder,objSt,"DDt",			"K" , lngRow)	'//  K:L �[����
	Call SetField(rsYOrder,objSt,"OdCd1",		"O" , lngRow)	'//  O ������
	Call SetField(rsYOrder,objSt,"OdCd2",		"P" , lngRow)	'//  P ������
	Call SetField(rsYOrder,objSt,"TkName",		"R" , lngRow)	'//  R ���Ӑ於
	Call SetField(rsYOrder,objSt,"Article1",	"S" , lngRow)	'//  S �H�����E����
	Call SetField(rsYOrder,objSt,"Article2",	"T" , lngRow)	'//  T �E�����������E�}���V�������E���V�X�e��
	Call SetField(rsYOrder,objSt,"Amount",		"V" , lngRow)	'//  V �󒍋��z
	Call SetField(rsYOrder,objSt,"AmountEHN",	"W" , lngRow)	'//  W �󒍋��z(EHN)
	Call SetField(rsYOrder,objSt,"Place",		"X" , lngRow)	'//  X �[�i�ꏊ
	rsYOrder.UpdateBatch
	LoadYOrderRow = lngRow
End Function

Private Function chkYM(byVal strYM)
	dim	a
	if strYM = "����" then
		strYM = Year(Now()) & Right("0" & Month(Now()),2)
	end if
	if Len(strYM) <> 6 then
		strYM = ""
	end if
	if isNumeric(strYM) <> True then
		strYM = ""
	end if
	chkYM = strYM
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
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & v)
	objRs.Fields(strField) = v
End Function
