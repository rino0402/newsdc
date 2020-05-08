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
	Wscript.Echo "ac Excel�f�[�^�ϊ�"
	Wscript.Echo "AcConv.vbs [option] <�t�@�C����> [�V�[�g��]"
	Wscript.Echo " /db:newsdc9"
	Wscript.Echo " /debug"
	Wscript.Echo "cscript AcConv.vbs AC�Z���^�[�݌�150625.xlsx /debug"
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
	call LoadExcel(strFilename,strSheetname)
	Main = 0
End Function

Function LoadExcel(byVal strFilename,byVal strSheetname)
	'-------------------------------------------------------------------
	'Excel�t�@�C����
	'-------------------------------------------------------------------
	strFilename = GetAbsPath(strFilename)
	Call Debug("LoadExcel():" & strFilename & "," & strSheetname)
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
	Call LoadBk(objXL,objBk,strSheetname)
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadExcel():End")
End Function

Function LoadBk(objXL,objBk,byVal strSheetname)
	Call Debug("LoadBk():" & strSheetname)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc9"))

	Call Debug("LoadBk():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadBk():SheetName=" & strShtName)
		if strSheetname = "" or strSheetname = strShtName then
'			Call LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem,objJcs)
		end if
	Next
	
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function
