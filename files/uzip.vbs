Option Explicit

Dim objSA
Dim objBefore
Dim objAfter

Const FOF_SILENT 			= &H04 	'�i���_�C�A���O��\�����Ȃ��B
Const FOF_RENAMEONCOLLISION = &H08 	'�t�@�C����t�H���_�����d������Ƃ��́u�R�s�[ �` �v�̂悤�ȃt�@�C�����Ƀ��l�[������B
Const FOF_NOCONFIRMATION 	= &H10 	'�㏑���m�F�_�C�A���O��\�����Ȃ��i[���ׂď㏑��]�Ɠ����j�B
Const FOF_ALLOWUNDO 		= &H40 	'����̎������i[�ҏW]-[���ɖ߂�]��{ctrl}+{z}�j��L���ɂ���B
Const FOF_FILESONLY 		= &H80 	'���C���h�J�[�h���w�肳�ꂽ�ꍇ�̂ݎ��s����B
Const FOF_SIMPLEPROGRESS 	= &H100 '�i���_�C�A���O�͕\�����邪�t�@�C�����͕\�����Ȃ��B
Const FOF_NOCONFIRMMKDIR 	= &H200 '�t�H���_�쐬�m�F�_�C�A���O��\�����Ȃ��i�����ō쐬�j�B
Const FOF_NOERRORUI 		= &H400 '�R�s�[��ړ����ł��Ȃ������ꍇ�̎��s���G���[�𔭐������Ȃ��B�������A�Ώۂ̃t�@�C�����΂��ď����𑱂���킯�ł͂Ȃ����Ƃɒ��ӁB
Const FOF_NORECURSION 		= &H1000 '�T�u�t�H���_���̃t�@�C���̓R�s�[���Ȃ��i�������A�t�H���_�͍쐬�����j�B

dim	sa
Set sa = WScript.CreateObject("Shell.Application")

dim	arg
For Each arg In WScript.Arguments
	dim	src
	Set src = sa.NameSpace(arg)
'	src.ParentFolder.CopyHere(src.Items)
	Wscript.Echo arg
	dim	itm
	For Each itm in src.Items
		Wscript.Echo itm.Name
		Call XDeleteFile(itm.Name)
		Call src.ParentFolder.CopyHere(itm,FOF_NOCONFIRMATION)
	Next
'	Wscript.Echo "FOF=0x" & Hex(FOF_NOCONFIRMATION)
Next

Function XDeleteFile(byVal txtFilename)
	Dim objFileSys
	Dim strScriptPath
	Dim strDeleteFrom
	
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	
	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	
	strDeleteFrom = objFileSys.BuildPath(strScriptPath, "\" & txtFilename)
	
	WScript.echo "DeleteFile:" & strDeleteFrom

	if objFileSys.FileExists(strDeleteFrom) = True Then
		objFileSys.DeleteFile strDeleteFrom, True
	End if
	
	Set objFileSys = Nothing
End Function


