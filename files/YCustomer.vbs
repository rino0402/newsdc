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
Call Include("get_b.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
Function GetCD()
	Dim objWshShell
	'�@WScript.Shell�I�u�W�F�N�g�̍쐬
	Set objWshShell = CreateObject("WScript.Shell")
	'�J�����g�f�B���N�g����\��
	dim	strCD
	strCD = objWshShell.CurrentDirectory
	Set objWshShell = Nothing
	GetCD = strCD
End Function

Function GetAbsPath(byVal strPath)
	Dim objFileSys
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	strPath = objFileSys.GetAbsolutePathName(strPath)
	Set objFileSys = Nothing
	GetAbsPath = strPath
End Function

Function GetDate2(byVal v)
	dim	strDate
	strDate = ""
	if isDate(v) then
		strDate = year(v) & Right(00 & month(v), 2) & Right(00 & day(v), 2)
	end if
	GetDate2 = strDate
End Function

Function GetScriptPath()
	GetScriptPath = Left(WScript.ScriptFullName,Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
End Function

Function GetFileName(byVal strFullName)
	dim	strFileName
	strFileName = strFullName
	dim	c
	for each c in split(strFileName,"\")
		Call Debug("GetFileName():" & c)
		if c <> "" then
			strFileName = c
		end if
	next
	GetFileName = strFileName
End Function

Function GetTab(ByVal s)
    Dim r
	r = Split(s,vbTab)
	GetTab = r
End Function

Function GetTrim(byval c)
	if left(c,1) = """" then
		if right(c,1) = """" then
			c = Right(c,Len(c) -1 )
			c = Left(c,Len(c) -1 )
		end if
	end if
	GetTrim = c
End Function

'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "�퐶���Ӑ�}�X�^�["
	Wscript.Echo "YCustomer.vbs [option] <�t�@�C����>"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo "ex."
	Wscript.Echo "sc32 YCustomer.vbs /db:fhd /debug ""F:\it\���Ӑ惊�X�g.xlsx"""
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	'���O�����I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		else
			strFilename = ""
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
	call LoadYCustomer(strFilename)
	Main = 0
End Function

Function LoadYCustomer(byVal strFilename)
	Call Debug("LoadYCustomer(" & strFilename & ")")
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
	Call LoadYCustomerXls(objXL,objBk)
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadYCustomer():End")
End Function

Function LoadYCustomerXls(objXL,objBk)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	rsYCustomer
	set rsYCustomer = OpenRs(objDb,"YCustomer")

	Call Debug("LoadYCustomerXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadYCustomerXls():SheetName=" & strShtName)
		Call LoadYCustomerXst(objXL,objBk,objSt,objDb,rsYCustomer)
	Next
	
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsYCustomer = CloseRs(rsYCustomer)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadYCustomerXst(objXL,objBk,objSt,objDb,rsYCustomer)
	Call Debug("LoadYCustomerXst():SheetName=" & objSt.Name)

	Call Debug("delete from YCustomer")
	Call ExecuteAdodb(objDb,"delete from YCustomer")

	dim	lngMaxRow
	lngMaxRow = excelGetMaxRow(objSt,"B",0)
	dim	lngRow
	For lngRow = 6 to lngMaxRow
		if LoadYCustomerRow(objXL,objBk,objSt,objDb,rsYCustomer,lngRow) = 0 then
			Exit For
		end if
	Next
End Function

Function LoadYCustomerRow(objXL,objBk,objSt,objDb,rsYCustomer,byVal lngRow)
	Call Debug("LoadYCustomerRow():" & objSt.Name & ":" & lngRow)
	dim	strB
	strB = objSt.Range("B" & lngRow)
	if strB = "" then
		LoadYCustomerRow = 0
		exit function
	end if
	rsYCustomer.AddNew
	Call SetField(rsYCustomer,objSt,"Code",		"B" , lngRow)	'		//  �R�[�h
	Call SetField(rsYCustomer,objSt,"Name",		"C" , lngRow)	'		//	����
	Call SetField(rsYCustomer,objSt,"NameK",	"D" , lngRow)	'		//	�t���K�i
	Call SetField(rsYCustomer,objSt,"NameR",	"E" , lngRow)	'		//	����
	Call SetField(rsYCustomer,objSt,"Zip",		"F" , lngRow)	'		//	�X�֔ԍ�
	Call SetField(rsYCustomer,objSt,"Address1",	"G" , lngRow)	'		//	�Z���P
	Call SetField(rsYCustomer,objSt,"Address2",	"H" , lngRow)	'		//	�Z���Q
	Call SetField(rsYCustomer,objSt,"Section",	"I" , lngRow)	'		//	������
	Call SetField(rsYCustomer,objSt,"Post",		"J" , lngRow)	'		//	��E��
	Call SetField(rsYCustomer,objSt,"Person",	"K" , lngRow)	'		//	���S����
	Call SetField(rsYCustomer,objSt,"Prefix",	"L" , lngRow)	'		//	�h��
	Call SetField(rsYCustomer,objSt,"Tel",		"M" , lngRow)	'		//	TEL
	Call SetField(rsYCustomer,objSt,"Fax",		"N" , lngRow)	'		//	FAX
	Call SetField(rsYCustomer,objSt,"Mail",		"O" , lngRow)	'		//	���[���A�h���X
	Call SetField(rsYCustomer,objSt,"Url",		"P" , lngRow)	'		//	�z�[���y�[�W
	Call SetField(rsYCustomer,objSt,"Tanto",	"Q" , lngRow)	'		//	�S����
	Call SetField(rsYCustomer,objSt,"TantoName","R" , lngRow)	'		//	�S���Җ�
	Call SetField(rsYCustomer,objSt,"TKbn",		"S" , lngRow)	'		//	����敪
	Call SetField(rsYCustomer,objSt,"TkS",		"T" , lngRow)	'		//	�P�����
	Call SetField(rsYCustomer,objSt,"Rate",		"U" , lngRow)	'		//	�|��
	Call SetField(rsYCustomer,objSt,"Bill1",	"V" , lngRow)	'		//	������
	Call SetField(rsYCustomer,objSt,"Bill2",	"W" , lngRow)	'		//	
	Call SetField(rsYCustomer,objSt,"SGrp1",	"X" , lngRow)	'		//	���O���[�v
	Call SetField(rsYCustomer,objSt,"SGrp2",	"Y" , lngRow)	'		//	
	Call SetField(rsYCustomer,objSt,"s1",		"Z" , lngRow)	'		//	���z�[������
	Call SetField(rsYCustomer,objSt,"s2",		"AA" , lngRow)	'		//	�Œ[������
	Call SetField(rsYCustomer,objSt,"s3",		"AB" , lngRow)	'		//	�œ]��
	Call SetField(rsYCustomer,objSt,"s4",		"AC" , lngRow)	'		//	�^�M���x�z
	Call SetField(rsYCustomer,objSt,"s5",		"AD" , lngRow)	'		//	���|�c��
	Call SetField(rsYCustomer,objSt,"s6",		"AE" , lngRow)	'		//	������@
	Call SetField(rsYCustomer,objSt,"s7",		"AH" , lngRow)	'		//	����T�C�N��
	Call SetField(rsYCustomer,objSt,"s8",		"AI" , lngRow)	'		//	�����
	Call SetField(rsYCustomer,objSt,"s9",		"AJ" , lngRow)	'		//	�萔�����S
	Call SetField(rsYCustomer,objSt,"s10",		"AK" , lngRow)	'		//	�T�C�g
	Call SetField(rsYCustomer,objSt,"s11",		"AL" , lngRow)	'		//	�w�蔄��`�[
	Call SetField(rsYCustomer,objSt,"s12",		"AM" , lngRow)	'		//	�w�萿����
	Call SetField(rsYCustomer,objSt,"s13",		"AN" , lngRow)	'		//	�������x��
	Call SetField(rsYCustomer,objSt,"s14",		"AO" , lngRow)	'		//	��ƃR�[�h
	Call SetField(rsYCustomer,objSt,"Cate1",	"AP" , lngRow)	'		//	���ނP
	Call SetField(rsYCustomer,objSt,"CateNate1","AQ" , lngRow)	'		//	���ނP����
	Call SetField(rsYCustomer,objSt,"Cate2",	"AR" , lngRow)	'		//	���ނQ
	Call SetField(rsYCustomer,objSt,"CateNate2","AS" , lngRow)	'		//	���ނQ����
	Call SetField(rsYCustomer,objSt,"Cate3",	"AT" , lngRow)	'		//	���ނR
	Call SetField(rsYCustomer,objSt,"CateNate3","AU" , lngRow)	'		//	���ނR����
	Call SetField(rsYCustomer,objSt,"Memo",		"AV" , lngRow)	'		//	������
	Call SetField(rsYCustomer,objSt,"s19",		"AW" , lngRow)	'		//	�Q�ƕ\��
	Call SetField(rsYCustomer,objSt,"s20",		"AX" , lngRow)	'		//	�X�V��
	Call SetField(rsYCustomer,objSt,"s21",		"AY" , lngRow)	'		//	���������Z
	rsYCustomer.UpdateBatch
	LoadYCustomerRow = lngRow
End Function

Function SetField(objRs,objSt,byVal strField,byVal strCol,byVal lngRow)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & objSt.Range(strCol & lngRow))
	dim	v
	v = RTrim(objSt.Range(strCol & lngRow))
	select case strField
	case "SDt"
		v = Replace(v,"/","")
	end select
	dim	dsize
	dsize = objRs.Fields(strField).DefinedSize
	v = Get_LeftB(v,dsize)
	Call Debug("SetField():" & lngRow & ":" & strField & "(" & dsize & ")=" & v)
	objRs.Fields(strField) = v
End Function
