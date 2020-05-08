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
	Wscript.Echo "JCS MF(Excel)�f�[�^�ϊ�"
	Wscript.Echo "JcsMF.vbs [option] <�t�@�C����> [�V�[�g��]"
	Wscript.Echo " /db:newsdc7"
	Wscript.Echo " /debug"
	Wscript.Echo "sc32 JcsMF.vbs jcs\�l�e.xls ���i�l�e /debug"
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
	call LoadMF(strFilename,strSheetname)
	Main = 0
End Function

Function LoadMF(byVal strFilename,byVal strSheetname)
	'-------------------------------------------------------------------
	'Excel�t�@�C����
	'-------------------------------------------------------------------
	strFilename = GetAbsPath(strFilename)
	Call Debug("LoadMF(" & strFilename & "," & strSheetname & ")")
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
	Call LoadMFXls(objXL,objBk,strSheetname)
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadMF():End")
End Function

Function JcsTableName(byVal strSheetname)
	dim	strTableName
	strTableName = ""
	select case strSheetname
	case "���i�l�e"
		strTableName = "JcsItem"
	end select
	JcsTableName = strTableName
End Function
Function LoadMFXls(objXL,objBk,byVal strSheetname)
	Call Debug("LoadMFXls(" & strSheetname & ")")
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	rsJcsItem
	set rsJcsItem = OpenRs(objDb,JcsTableName(strSheetname))

	Call Debug("LoadMFXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadMFXls():SheetName=" & strShtName)
		if strSheetname = "" or strSheetname = strShtName then
			Call LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem)
		end if
	Next
	
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsJcsItem = CloseRs(rsJcsItem)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem)
	Call Debug("LoadMFXst():SheetName=" & objSt.Name)

	Call Debug("delete from " & JcsTableName(objSt.Name))
	Call ExecuteAdodb(objDb,"delete from " & JcsTableName(objSt.Name))

	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"C",lngMaxRow)
	dim	lngRow
	For lngRow = 3 to lngMaxRow
		Call LoadMFRow(objXL,objBk,objSt,objDb,rsJcsItem,lngRow)
	Next
End Function

Function LoadMFRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
	Call Debug("LoadMFRow():" & objSt.Name & ":" & lngRow)
	Call rsJcsItem.AddNew
	Call SetField(rsJcsItem,objSt,"MazdaPn",		"A" , lngRow)	'// �}�c�_�i�� 
	Call SetField(rsJcsItem,objSt,"NameE",			"B" , lngRow)	'// �i���i�p��j 
	Call SetField(rsJcsItem,objSt,"Pn",				"C" , lngRow)	'// �i�b�r�i�� 
	Call SetField(rsJcsItem,objSt,"NameJ",			"D" , lngRow)	'// �i���i���{��j 
	Call SetField(rsJcsItem,objSt,"Color",			"E" , lngRow)	'// �F 
	Call SetField(rsJcsItem,objSt,"LabelType",		"F" , lngRow)	'// ���x�� MorF
	Call SetField(rsJcsItem,objSt,"LabelCut",		"G" , lngRow)	'//  �ؒf
	Call SetField(rsJcsItem,objSt,"Location",		"H" , lngRow)	'// �݌� 9999-XX
	Call SetField(rsJcsItem,objSt,"SSpec",			"J" , lngRow)	'// ���i�� �d�l��
	Call SetField(rsJcsItem,objSt,"SType",			"K" , lngRow)	'//  �^�C�v
	Call SetField(rsJcsItem,objSt,"GPn",			"L" , lngRow)	'// �O�� �i��
	Call SetField(rsJcsItem,objSt,"GQty",			"M" , lngRow)	'//  ����
	Call SetField(rsJcsItem,objSt,"LastPn",			"N" , lngRow)	'// �ŏI�׎p �i��
	Call SetField(rsJcsItem,objSt,"LastQty",		"O" , lngRow)	'//  ����
	Call SetField(rsJcsItem,objSt,"CheckConf",		"P" , lngRow)	'// �`�F�b�N�m�F 
	Call SetField(rsJcsItem,objSt,"Location1",		"Q" , lngRow)	'// ���P�[�V���� �I��
	Call SetField(rsJcsItem,objSt,"Location2",		"R" , lngRow)	'//  �ʒu
	Call SetField(rsJcsItem,objSt,"ShareNum",		"S" , lngRow)	'// ���p ����
	Call SetField(rsJcsItem,objSt,"ShareNo",		"T" , lngRow)	'//  No.
	Call SetField(rsJcsItem,objSt,"AlterDate",		"Y" , lngRow)	'// �o�^ ���^��
	Call SetField(rsJcsItem,objSt,"AlterPerson",	"Z" , lngRow)	'//  �o�^��
	Call SetField(rsJcsItem,objSt,"LastShipDate",	"AA" , lngRow)	'// �O��o�׎��� ���^��
	Call SetField(rsJcsItem,objSt,"LastShipQty",	"AB" , lngRow)	'//  ����
	Call SetField(rsJcsItem,objSt,"LastShipPn",		"AC" , lngRow)	'//  �i�b�r�i�ԇB
	Call SetField(rsJcsItem,objSt,"CoStockDate",	"AD" , lngRow)	'// �J��z���݌� ���^��
	Call SetField(rsJcsItem,objSt,"CoStockQty",		"AE" , lngRow)	'//  ����
	On Error Resume Next
		Call rsJcsItem.UpdateBatch
		if Err <> 0 Then
			WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
			' ں��ނ̷� ̨���ނɏd�����鷰�l������܂�(Btrieve Error 5)
			if Err.Number <> -2147467259 then
				lngRow = 0
			end if
			Call rsJcsItem.CancelUpdate
		end if
	On Error GoTo 0
	LoadMFRow = lngRow
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

Class JcsItem
	Public pTableName
    Private Sub Class_Initialize
        pTableName = "JcsItem"
    End Sub
End Class
