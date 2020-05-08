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
	Wscript.Echo "cscript JcsMF.vbs jcs\�l�e.xls ���i�l�e /debug"
	Wscript.Echo "cscript JcsMF.vbs jcs\�l�e.xls �^�C�v�l�e /debug"
	Wscript.Echo "cscript JcsMF.vbs jcs\�c�e.xls �ړ��c�e /debug"
	Wscript.Echo "cscript JcsMF.vbs jcs\�c�e.xls �󒍂c�e /debug"
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

Function LoadMFXls(objXL,objBk,byVal strSheetname)
	Call Debug("LoadMFXls(" & strSheetname & ")")
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'JCS�N���X
	'-------------------------------------------------------------------
	dim	objJcs
	select case strSheetname
	case "���i�l�e"
		Set objJcs = New JcsItem
	case "�^�C�v�l�e"
		Set objJcs = New JcsType
 	case "�ړ��c�e"
		Set objJcs = New JcsIdo
 	case "�󒍂c�e"
		Set objJcs = New JcsOrder
	end select
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	rsJcsItem
	set rsJcsItem = OpenRs(objDb,objJcs.pTableNameTmp)

	Call Debug("LoadMFXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadMFXls():SheetName=" & strShtName)
		if strSheetname = "" or strSheetname = strShtName then
			Call LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem,objJcs)
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

Function LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem,objJcs)
	Call Debug("LoadMFXst():SheetName=" & objSt.Name)

	Call objJcs.InitRecord(objDb)

	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"C",lngMaxRow)
	dim	lngRow
	For lngRow = 3 to lngMaxRow
		Call DispMsg("�o�^��..." & objSt.Name & ":" & lngRow & "/" & lngMaxRow)
		Call objJcs.LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,lngRow)
	Next
	Call DispMsg("�e�[�u���R�s�[..." & objJcs.pTableNameTmp & "��" & objJcs.pTableName)
	Call CopyTable(objDb,objJcs.pTableNameTmp,objJcs.pTableName)
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

'-------------------------------------------------------
' ���i�l�e
'-------------------------------------------------------
Class JcsItem
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsItem"
        pTableNameTmp	= "JcsItem_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
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
		Call SetField(rsJcsItem,objSt,"Sagyo",			"X" , lngRow)	'// ��ƕW����
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
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' �^�C�v�l�e
'-------------------------------------------------------
Class JcsType
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsType"
        pTableNameTmp	= "JcsType_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 6 then
			LoadRow = lngRow
			exit function
		end if
		Call rsJcsItem.AddNew
		rsJcsItem("XlsRow") = lngRow							'// Excel�s
		Call SetField(rsJcsItem,objSt,"SType",		"A" , lngRow)	'// �^�C�v
		Call SetField(rsJcsItem,objSt,"SPn1",		"B" , lngRow)	'// ���ޕi�ԂP
		Call SetField(rsJcsItem,objSt,"SQty1",		"C" , lngRow)	'// ���ވ����P
		Call SetField(rsJcsItem,objSt,"SPn2",		"D" , lngRow)	'// ���ޕi�ԂQ
		Call SetField(rsJcsItem,objSt,"SQty2",		"E" , lngRow)	'// ���ވ����Q
		Call SetField(rsJcsItem,objSt,"SPn3",		"F" , lngRow)	'// ���ޕi�ԂR
		Call SetField(rsJcsItem,objSt,"SQty3",		"G" , lngRow)	'// ���ވ����R
		Call SetField(rsJcsItem,objSt,"SPn4",		"H" , lngRow)	'// ���ޕi�ԂS
		Call SetField(rsJcsItem,objSt,"SQty4",		"I" , lngRow)	'// ���ވ����S
                                                                        
		Call SetField(rsJcsItem,objSt,"PUnit",		"K" , lngRow)	'// ����P��
		Call SetField(rsJcsItem,objSt,"MCP1",		"L" , lngRow)	'// �ޗ���P
		Call SetField(rsJcsItem,objSt,"MCP2",		"M" , lngRow)	'// �ޗ���Q

		Call SetField(rsJcsItem,objSt,"PCP1",		"O" , lngRow)	'// ���H��P
		Call SetField(rsJcsItem,objSt,"PCP2",		"P" , lngRow)	'// ���H��Q
                                                                        
		Call SetField(rsJcsItem,objSt,"OCP1",		"R" , lngRow)	'// ���̑��P
		Call SetField(rsJcsItem,objSt,"OCP2",		"S" , lngRow)	'// ���̑��Q
                                                                        
                                                                        
		Call SetField(rsJcsItem,objSt,"AlterDate",	"V" , lngRow)	'// �o�^��
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
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' �ړ��c�e
'-------------------------------------------------------
Class JcsIdo
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsIdo"
        pTableNameTmp	= "JcsIdo_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		Call rsJcsItem.AddNew
		rsJcsItem.Fields("ID") = lngRow									'// �s�ԍ�
		Call SetField(rsJcsItem,objSt,"OrderDt",		"A" , lngRow)	'// �� ����
		Call SetField(rsJcsItem,objSt,"NohinNo",		"B" , lngRow)	'// �[�i�ԍ� 
		Call SetField(rsJcsItem,objSt,"NohinNo2",		"C" , lngRow)	'//  �A
		Call SetField(rsJcsItem,objSt,"Location",		"D" , lngRow)	'// ���i ��
		Call SetField(rsJcsItem,objSt,"DlvDt",			"E" , lngRow)	'// �[�� 
		Call SetField(rsJcsItem,objSt,"MazdaPn",		"F" , lngRow)	'// �}�c�_�i�� 
		Call SetField(rsJcsItem,objSt,"Pn",				"H" , lngRow)	'// q 
		Call SetField(rsJcsItem,objSt,"Qty",			"I" , lngRow)	'// �w���� 
		Call SetField(rsJcsItem,objSt,"L46",			"K" , lngRow)	'// L46 
		Call SetField(rsJcsItem,objSt,"Dt",				"L" , lngRow)	'// ���� 
		Call SetField(rsJcsItem,objSt,"DestCode",		"M" , lngRow)	'// �󕥐�or���� CODE
		Call SetField(rsJcsItem,objSt,"DestName",		"N" , lngRow)	'//  �󕥐於or����
		Call SetField(rsJcsItem,objSt,"IQty",			"O" , lngRow)	'// �ړ����� ����
		Call SetField(rsJcsItem,objSt,"OQty",			"P" , lngRow)	'//  �o��
		Call SetField(rsJcsItem,objSt,"SSpec",			"Q" , lngRow)	'// ���i�� �m
		Call SetField(rsJcsItem,objSt,"SType",			"R" , lngRow)	'//  �^�C�v
		Call SetField(rsJcsItem,objSt,"SPrice",			"S" , lngRow)	'//  �P��
		Call SetField(rsJcsItem,objSt,"SAmont",			"T" , lngRow)	'//  ���z
		Call SetField(rsJcsItem,objSt,"GPn",			"U" , lngRow)	'// �O�� �i��
		Call SetField(rsJcsItem,objSt,"GQty",			"V" , lngRow)	'//  ����
		Call SetField(rsJcsItem,objSt,"LastPn",			"W" , lngRow)	'// �ŏI�׎p �i��
		Call SetField(rsJcsItem,objSt,"LastQty",		"X" , lngRow)	'//  ����
		Call rsJcsItem.UpdateBatch
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' �󒍂c�e
'-------------------------------------------------------
Class JcsOrder
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsOrder"
        pTableNameTmp	= "JcsOrder_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 11 then
			LoadRow = lngRow
			exit function
		end if
		Call rsJcsItem.AddNew
		rsJcsItem.Fields("XlsRow")	= lngRow
		Call SetField(rsJcsItem,objSt,	"OrderDt",	"A", lngRow)	'//	�󒍓�
		Call SetField(rsJcsItem,objSt,	"NohinNo",	"B", lngRow)	'//	�[�i�ԍ��@
		Call SetField(rsJcsItem,objSt,	"NohinNo2",	"C", lngRow)	'//	�[�i�ԍ��A
		Call SetField(rsJcsItem,objSt,	"PartWH",	"D", lngRow)	'//	���i��
		Call SetField(rsJcsItem,objSt,	"DlvDt",	"E", lngRow)	'//	�[��
		Call SetField(rsJcsItem,objSt,	"MazdaPn",	"F", lngRow)	'//	�}�c�_�i��
							
		Call SetField(rsJcsItem,objSt,	"Pn",		"H", lngRow)	'//	JCS�i��
		Call SetField(rsJcsItem,objSt,	"Qty",		"I", lngRow)	'//	�w����
							
		Call SetField(rsJcsItem,objSt,	"Location",	"K", lngRow)	'//	�݌ɏꏊ
		Call SetField(rsJcsItem,objSt,	"MPn",		"L", lngRow)	'//	�}�X�^�[�o�^JCS�i��
		Call SetField(rsJcsItem,objSt,	"InfoDlvDt","M", lngRow)	'//	�[�i��� �[��
		Call SetField(rsJcsItem,objSt,	"InfoTotal","N", lngRow)	'//	�ތ^
		Call SetField(rsJcsItem,objSt,	"InfoZan",	"O", lngRow)	'//	�c
		Call SetField(rsJcsItem,objSt,	"NG",		"P", lngRow)	'//	NG
		Call SetField(rsJcsItem,objSt,	"Biko",		"Q", lngRow)	'//	���l
		Call SetField(rsJcsItem,objSt,	"No",		"R", lngRow)	'//	No
		Call rsJcsItem.UpdateBatch
		LoadRow = lngRow
	End Function
End Class
