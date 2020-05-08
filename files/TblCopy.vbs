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

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "JCS MF(Excel)�f�[�^�ϊ�"
	Wscript.Echo "TblCopy.vbs [option] <�R�s�[��> <�R�s�[��>"
	Wscript.Echo " /db:newsdc7"
	Wscript.Echo " /debug"
	Wscript.Echo "sc32 TblCopy.vbs /db:newsdc7 JcsItem_Tmp JcsItem"
	Wscript.Echo "sc32 TblCopy.vbs /db:newsdc7 JcsIdo_Tmp JcsIdo"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strSrc
	dim	strDst
	strSrc = ""
	strDst = ""
	'���O�����I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.UnNamed
		if strSrc = "" then
			strSrc = strArg
		elseif strDst = "" then
			strDst = strArg
		else
			usage()
			Main = 1
			exit Function
		end if
	next
	'���O�t���I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			strSrc = ""
		case else
			strSrc = ""
		end select
	next
	if strSrc = "" then
		usage()
		Main = 1
		exit Function
	end if
	call TblCopy(strSrc,strDst)
	Main = 0
End Function

Function TblCopy(byVal strSrc,byVal strDst)
	Call Debug("TblCopy(" & strSrc & "," & strDst & ")")
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'�e�[�u���R�s�[
	'-------------------------------------------------------------------
	Call DispMsg("�e�[�u���R�s�[:" & strSrc & "��" & strDst)
	Call CopyTable(objDb,strSrc,strDst)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

