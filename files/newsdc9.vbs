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
Call Include("BoConv_sub.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "BO�f�[�^�ϊ�(���[��������M)"
	Wscript.Echo "newsdc9.vbs [option]"
	Wscript.Echo "/save	��M���[�����T�[�o�[�Ɏc��(default)"
	Wscript.Echo "/savd	��M���[�����폜����"
	Wscript.Echo "/load	<filename>"
	Wscript.Echo "Ex."
	Wscript.Echo "newsdc9.vbs"
	Wscript.Echo "newsdc9.vbs /load temp\�I���ް�.csv"

	dim	c
	for each c in FileList("temp","path")
		Call DispMsg("FileList():" & c)
	next

End Sub

'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	dim	strOption
	strOption = ""
	dim	strDb
	strDb = "newsdc9"
'	For Each strArg In WScript.Arguments.UnNamed
'    	select case strArg
'		case else
'			usage()
'			Main = 1
'			exit Function
'		end select
'	Next
	For Each strArg In WScript.Arguments
'		Call DispMsg(strArg)
		select case strOption
		case "/load"
			strFilename = strArg
			strOption = ""
		case else
			strArg = lcase(strArg)
	    	select case Split(strArg,":")(0)
			case "/db"
				strDb = Split(strArg,":")(1)
			case "/debug"
			case "/save"
			case "/savd"
			case "/load"
				strOption = strArg
			case else
				Call DispMsg("unknown:" & strArg )
				usage()
				Main = 1
				exit Function
			end select
		end select
	Next
	if strOption <> "" then
		Call DispMsg("option:" & strOption )
		usage()
		Main = 1
		exit Function
	end if
	if strFilename <> "" then
		dim	strMsg
		strMsg = Load(strDb,strFilename)
		Call DispMsg(strMsg)
	else
		call RcvMail()
	end if

	Main = 0
End Function

'-----------------------------------------------------------------------
'RcvMail()�̃I�v�V����
'-----------------------------------------------------------------------
Private Function RcvMailOpt()
	RcvMailOpt = "SAVE 1-10"
	if WScript.Arguments.Named.Exists("savd") then
		RcvMailOpt = "SAVD 1-10"
	end if
End Function

'-----------------------------------------------------------------------
'���[����M
'-----------------------------------------------------------------------
Private Function RcvMail()
	Call Debug("RcvMail()")
	' ���[������MAPI�̐錾
	dim	bobj
	Set bobj = CreateObject("Basp21")

	dim	svname
	dim	user
	dim	pass
	dim	dirname
	dim	strDb

	strDb = "newsdctest"

	svname	= "ns"						' POP3�T�[�o�}�V����
	user	= "newsdc9"					' ���[���{�b�N�X��
	pass	= "123daa@Z"				' �p�X���[�h
	dirname = "rcvtemp"					' �ۑ��f�B���N�g����

'       SAVE n[-n2] .... n�Ԗڂ̃��[������M���܂�
'                        n2���w�肷���n2�Ԗڂ܂ł̃��[������M���܂��B
'       SAVD n[-n2] ... n�Ԗڂ̃��[������M���A�T�[�o�̃��[���{�b�N�X����
'                   �폜���܂�
'                   n2���w�肷���n2�Ԗڂ܂ł̃��[������M���č폜���܂�
	Call DispMsg("���[����M��:" & RcvMailOpt())
	dim	outarray
	outarray = bobj.RcvMail(svname,user,pass,RcvMailOpt(),dirname)
	if IsArray(outarray) then	' OK ?
		dim	file
		for each file in outarray
			dim	strSubject
			dim	strSubject1
			dim	strSubject2
			dim	strFrom
			dim	strBody
			strSubject	= ""
			strSubject1	= ""
			strSubject2	= ""
			strFrom		= ""
			strBody		= ""
			strBody = strBody & "BO�f�[�^�ϊ����[������M���܂����B" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "���Y�t�t�@�C��" & vbCrlf
			dim	array2
			array2 = bobj.ReadMail(file,"subject:from:date:",">" & dirname)
			if IsArray(array2) then	' OK ?
				dim	data
				for each data in array2
					'1�s�ڂ�\��
					Call DispMsg(Split(data,vbCrLf)(0))
					dim	strHead
					strHead = Split(data,":")(0)
					select case strHead
					case "Subject"
									strSubject = ""
									if Len(data) > 9 then
										strSubject = Right(data,Len(data) - 9)
									end if
									strSubject1 = "�J�n:BO�f�[�^�ϊ� " & strSubject
									if instr(trim(lcase(strSubject)),"bodata") > 0 then
										strDb = "newsdc9"
									end if
					case "From"
									strFrom	= Right(data,Len(data) - 6)
									if instr(trim(lcase(strSubject)),"bodata") > 0 then
										if inStr(strFrom,"akagi.seisuke@kk.jp.panasonic.com") > 0 then
											strFrom = "kitamura.ryuichi@kk.jp.panasonic.com"
											strFrom = strFrom & vbTab & "cc"
											strFrom = strFrom & vbTab & "kitamura@kk-sdc.co.jp"
											strFrom = strFrom & vbTab & "system@kk-sdc.co.jp"
										elseif inStr(strFrom,"kubo@kk-sdc.co.jp") > 0 then
											strFrom = strFrom & vbTab & "cc"
											strFrom = strFrom & vbTab & "kitamura.ryuichi@kk.jp.panasonic.com"
											strFrom = strFrom & vbTab & "kitamura@kk-sdc.co.jp"
											strFrom = strFrom & vbTab & "kishimi@kk-sdc.co.jp"
										else
											strFrom = strFrom & vbTab & "cc"
											strFrom = strFrom & vbTab & "kubo@kk-sdc.co.jp"
										end if
									else
										strFrom = strFrom & vbTab & "cc"
										strFrom = strFrom & vbTab & "kubo@kk-sdc.co.jp"
									end if
					case "Body"
					case "File"
						Call SaveFileEx(bobj,data)
						strBody = strBody & "�@" & Right(data,Len(data) - 6) & vbCrlf
					end select
				next
			end if
			strBody = strBody & vbCrlf
			strBody = strBody & "���W�J�t�@�C��" & vbCrlf
			strBody = strBody & FileList("temp","list")

			strBody = strBody & vbCrlf
			strBody = strBody & "BO�f�[�^�ϊ��������J�n���܂��B" & vbCrlf
			strBody = strBody & "�����I����Ɋ����ʒm���[���𑗐M���܂��B" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "db=" & GetOption("db",strDb) & vbCrlf

			Call TanaMake()
			dim	strFile
			strFile = "\\mint\newsdc\tana.csv"
			strFile = strFile & vbTab & "\\mint\newsdc\tananew.csv"

			'�ԐM
			Call DispMsg("SendMail:" & svname & ":" & strFrom & ":" & user & ":" & strSubject)
			dim strMsg
			strMsg = bobj.SendMail(svname,strFrom,"newsdc9@kk-sdc.co.jp", strSubject1,strBody,strFile)
			Call DispMsg(strMsg)
			'�ϊ�����
			strSubject2 = "����:BO�f�[�^�ϊ� " & strSubject
			strBody = ""
			strBody = strBody & "BO�f�[�^�ϊ��������������܂����B" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "db=" & GetOption("db",strDb) & vbCrlf
			dim f
			for each f in FileList("temp","path")
				Call DispMsg("�ϊ�����:" & f)
				strMsg = Load(strDb,f)
				strBody = strBody & strMsg & vbCrlf
				Call DeleteFile(f)
			next
			strMsg = bobj.SendMail(svname,strFrom,"newsdc9@kk-sdc.co.jp", strSubject2,strBody,"")
			Call DispMsg(strMsg)
		next
	end if
End Function

Function TanaMake()
	Dim objIE

	'IE�I�u�W�F�N�g���쐬���܂�
	Set objIE = CreateObject("InternetExplorer.Application")

	'�E�B���h�E�̑傫����ύX���܂�
	objIE.Width = 800
	objIE.Height = 600

	'�\���ʒu��ύX���܂�
	objIE.Left = 0
	objIE.Top = 0

	'�X�e�[�^�X�o�[�ƃc�[���o�[���\���ɂ��܂�
	objIE.Statusbar = False
	objIE.ToolBar = False

	'�C���^�[�l�b�g�G�N�X�v���[����ʂ�\�����܂�
	objIE.Visible = True

	'�@�w�肵��URL��\�����܂�
	objIE.Navigate "http://mint/newsdc/tanamake.php"

	'�A�y�[�W�̓ǂݍ��݂��I���܂ŃR�R�ŃO���O�����
	Do Until objIE.Busy = False
	   '�󃋁[�v���Ɩ��ʂ�CPU���g���̂�250�~���b�̃C���^�[�o����u��
	   WScript.sleep(250)
	Loop

	'�X�e�[�^�X�o�[�ƃc�[���o�[��\�����܂�
	objIE.Statusbar = True
	objIE.ToolBar = True

	WScript.Sleep 5000

	objIE.Quit
	Set objIE = Nothing
End Function

'-----------------------------------------------------------------------
'�Y�t�t�@�C���ۑ���LZH�W�J
'-----------------------------------------------------------------------
Private Function SaveFileEx(bobj,byVal m)
	Call Debug("SaveFileEx(" & m & ")")
	dim	strFilename
	strFilename = Right(m,Len(m) - 6)
	Call Debug("strFilename=" & strFilename)

	'rc = bobj.Execute("cmd.exe /c c:\lha.exe l basp21.lzh",1,stdout)
	dim cmd
	cmd = "cmd.exe /c lha32 e " & strFilename & " temp\"
	Call DispMsg(cmd)
	dim	rc
	dim	stdout
	rc = bobj.Execute(cmd,1,stdout)
	Call DispMsg(stdout)
End Function
'-----------------------------------------------------------------------
'�t�@�C���ꗗ
'-----------------------------------------------------------------------
Private Function FileList(byval strPath,byval strRcv)
	Dim objFileSys
	Dim strScriptPath
	Dim strTargetPath
	Dim objFolder
	Dim objItem

	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

	strTargetPath = objFileSys.BuildPath(strScriptPath, strPath)

	Set objFolder = objFileSys.GetFolder(strTargetPath)

	Call Debug("�t�@�C�����ꗗ")
	dim	aryFile()
	dim	i
	i = 0
	dim	strList
	strList = ""
	For Each objItem In objFolder.Files
	    Call Debug(objItem.Name)
	    Call Debug(objItem.Path)
	    Call Debug(objItem.DateLastModified)
		strList = strList & TmFormat(objItem.DateLastModified)
		strList = strList & " " & NumFormat(objItem.Size)
		strList = strList & " " & objItem.Name
		strList = strList & vbCrlf

		Redim Preserve aryFile(i)
		aryFile(i) = objItem.Path
		i = i + 1
	Next

	Call Debug("�t�@�C�����F" & objFolder.Files.Count)

	Set objFolder = Nothing
	Set objFileSys = Nothing
	if strRcv = "list" then
		FileList = strList
	else
		FileList = aryFile
	end if
End Function

Private Function NumFormat(v)
	NumFormat = Right(Space(12) & FormatNumber(v,0,,-1),12)
End Function

Private Function TmFormat(v)
	dim	dt
	dim	tm
	dt = Split(v," ")(0)
	tm = Split(v," ")(1)
	TmFormat = dt & Right("  " & tm,9)
End Function

Function DeleteFile(byVal path)
	Dim objFileSys
	Dim strScriptPath
	Dim strDeleteFrom
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
'	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
'	strDeleteFrom = objFileSys.BuildPath(strScriptPath, "\backup\TestData.csv")
	objFileSys.DeleteFile path, True
'	WScript.echo "BackUp����TestData.csv���폜���܂����B"
	Set objFileSys = Nothing
End Function
