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
Call Include("file.vbs")
Call Include("debug.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "NPL���[���A�g"
	Wscript.Echo "nplrecv.vbs [option]"
	Wscript.Echo "/save	��M���[�����T�[�o�[�Ɏc��(default)"
	Wscript.Echo "/savd	��M���[�����폜����"
	Wscript.Echo "/load	<filename>"
	Wscript.Echo "Ex."
	Wscript.Echo "nplrecv.vbs"
	Wscript.Echo "nplrecv.vbs /load temp"

	dim	c
	for each c in FileList("npltemp\","path")
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
'		strMsg = Load(strDb,strFilename)
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
	RcvMailOpt = "SAVE 1"
	if WScript.Arguments.Named.Exists("savd") then
		RcvMailOpt = "SAVD 1"
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

	strDb = "newsdc9"

	svname	= "ns"						' POP3�T�[�o�}�V����
	user	= "npl"						' ���[���{�b�N�X��
	pass	= "123daaa@Z"				' �p�X���[�h
	dirname = "nplrecv"					' �ۑ��f�B���N�g����

	Call DispMsg("RcvMail():���[����M��:" & RcvMailOpt())
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
			dim	aryFile()
			redim	aryFile(0)
			strSubject	= ""
			strSubject1	= ""
			strSubject2	= ""
			strFrom		= ""
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
									strSubject = Right(data,Len(data) - 9)
					case "From"
									strFrom	= lcase(Right(data,Len(data) - 6))
					case "Body"
					case "File"
									if UBound(aryFile) > 0 then
										ReDim Preserve aryFile(UBound(aryFile) + 1)
									end if
									aryFile(UBound(aryFile)) = Right(data,Len(data) - 6)
					end select
				next
			end if

			strSubject1 = "�J�n:BO�f�[�^�ϊ� " & strSubject
			if CheckFrom(strFrom,"jp.panasonic.com") > 0 then
				strFrom = strFrom & vbTab & "cc"
				strFrom = strFrom & vbTab & "system@kk-sdc.co.jp"
			elseif CheckFrom(strFrom,"kubo@kk-sdc.co.jp") > 0 then
			else
				strFrom = strFrom & vbTab & "cc"
				strFrom = strFrom & vbTab & "kubo@kk-sdc.co.jp"
			end if
			strBody		= ""
			strBody = strBody & "BO�f�[�^�ϊ����[������M���܂����B" & vbCrlf
			strBody = strBody & "BO�f�[�^�ϊ��������J�n���܂��B" & vbCrlf
			strBody = strBody & "�����I����Ɋ����ʒm���[���𑗐M���܂��B" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "���Y�t�t�@�C��" & vbCrlf
			dim	strFile
			for each strFile in aryFile
				strBody = strBody & strFile & vbCrlf
			next
			strBody = strBody & vbCrlf
			strBody = strBody & "db=" & GetOption("db",strDb) & vbCrlf
			'---------------------------------------------------
			'�ԐM(��t)
			'---------------------------------------------------
			Call DispMsg("SendMail:" & svname & ":" & strFrom & ":" & user & ":" & strSubject)
			dim strMsg
			strMsg = bobj.SendMail(svname,strFrom,"npl@kk-sdc.co.jp", strSubject1,strBody,strFile)
			Call DispMsg(strMsg)
			'---------------------------------------------------
			'�ϊ�����
			'---------------------------------------------------
			strSubject2 = "����:BO�f�[�^�ϊ� " & strSubject
			strBody = ""
			strBody = strBody & "BO�f�[�^�ϊ��������������܂����B" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "db=" & GetOption("db",strDb) & vbCrlf
			for each strFile in aryFile
'			for each f in FileList("npltemp","path")
				Call DispMsg("�ϊ�����:" & strFile)
				strMsg = ConvCmd(bobj,strFile)
				strBody = strBody & strMsg & vbCrlf
'				Call DeleteFile(strFile)
			next
			Call Debug("RcvMail(): svname:" & svname)
			Call Debug("RcvMail():     To:" & strFrom)
			Call Debug("RcvMail():Subject:" & strSubject2)
			Call Debug("RcvMail():   Body:" & strBody)
			strMsg = bobj.SendMail(svname,strFrom,"npl@kk-sdc.co.jp", strSubject2,strBody,"s_hirei_s.xlsm")
			Call DispMsg(strMsg)
		next
	end if
End Function

Function CheckFrom(byVal strFrom,byVal strAddress)
	CheckFrom = 1
End Function
'-----------------------------------------------------------------------
'�Y�t�t�@�C���ۑ���LZH�W�J
'-----------------------------------------------------------------------
Private Function ConvCmd(bobj,byVal strFile)
	Call Debug("ConvCmd():" & strFile)
'	dim	strFilename
'	strFilename = Right(strFile,Len(strFile) - 6)
	Call Debug("ConvCmd():strFile=" & strFile)

	'rc = bobj.Execute("cmd.exe /c c:\lha.exe l basp21.lzh",1,stdout)
	dim cmd
'	cmd = "cmd.exe /c lha32 e " & strFilename & " nplrecv\"
	cmd = "cmd.exe /c cscript /nologo BoSyukaX.vbs " & strFile
	Call DispMsg(cmd)
	dim	rc
	dim	stdout
	stdout = ""
	rc = bobj.Execute(cmd,1,stdout)
	Call DispMsg(stdout)
	ConvCmd = rc & " = " & cmd & vbCrLf & stdout

	cmd = "cmd.exe /c cscript /nologo s_hirei.vbs s_hirei.xlsm s_hirei_s.xlsm"
	Call DispMsg(cmd)
	stdout = ""
	rc = bobj.Execute(cmd,1,stdout)
	Call DispMsg(stdout)
	ConvCmd = ConvCmd & vbCrLf & rc & " = " & cmd & vbCrLf & stdout
End Function
