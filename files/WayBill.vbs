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
Call Include("get_b.vbs")
Call Include("csv.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "���� ����CSV�ϊ�"
	Wscript.Echo "WayBill.vbs [option] <�t�@�C����>"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo " /delete"
	Wscript.Echo "ex."
	Wscript.Echo "sc32 WayBill.vbs /db:fhd /debug SYUKKA.CSV"
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
		case "delete"
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
	call LoadWayBill(strFilename)
	Main = 0
End Function

Function LoadWayBill(byVal strFilename)
	Call Debug("LoadWayBill(" & strFilename & ")")
	'-------------------------------------------------------------------
	'�t�@�C����
	'-------------------------------------------------------------------
	if strFileName = "" then
		Call DispMsg("�t�@�C�������w�肵�ĉ�����")
		Exit Function
	end if
	'-------------------------------------------------------------------
	'FileSystemObject�̏���
	'-------------------------------------------------------------------
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Call Debug("CreateObject(Scripting.FileSystemObject)")
	'-------------------------------------------------------------------
	'�t�@�C���I�[�v��
	'-------------------------------------------------------------------
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	Call Debug("OpenTextFile()=" & strFilename)
	'-------------------------------------------------------------------
	'�Ǎ�����
	'-------------------------------------------------------------------
	Call LoadWayBillCsv(objFile)
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objFile.Close()
	set objFile = Nothing
	set objFSO = Nothing
	Call Debug("LoadWayBill():End")
End Function

Function LoadWayBillCsv(objFile)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	Call Debug("OpenAdodb(" & objDb.ConnectionString & ")")
	Call Debug(GetProperties(objDb))
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	rsWayBill
	set rsWayBill = OpenRs(objDb,"WayBill")
	Call Debug("OpenRs()")
	Call Debug(GetProperties(rsWayBill))

	Call Debug("LoadWayBillCsv()")
	Call LoadWayBillCsv1(objFile,objDb,rsWayBill)
	
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsWayBill = CloseRs(rsWayBill)
	Call Debug("CloseRs()")
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
	Call Debug("CloseAdodb()")
End Function

Function LoadWayBillCsv1(objFile,objDb,rsWayBill)
	Call Debug("LoadWayBillCsv1()")

	if WScript.Arguments.Named.Exists("delete") then
		Call Debug("delete from WayBill")
		Call ExecuteAdodb(objDb,"delete from WayBill")
	end if

	dim	lngRow
	lngRow = 0
	do while ( objFile.AtEndOfStream = False )
		lngRow = lngRow + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
'		Call Debug(strBuff)
		if lngRow > 1 then
			if LoadWayBillRow(objDb,rsWayBill,strBuff) = 0 then
				Exit do
			end if
		end if
	loop
End Function

Function LoadWayBillRow(objDb,rsWayBill,byVal strBuff)
	Call Debug("LoadWayBillRow():" & strBuff)

	rsWayBill.AddNew
	dim		aryBuff
	aryBuff = GetCSV(strBuff)
	dim	i
	i = 1
	dim	f
	for each f in rsWayBill.Fields
		dim	dsize
		dsize = f.DefinedSize
		dim	a
		a = aryBuff(i)
		Call Debug(f.Name & "(" & i & ")" & a & "(" & dsize & ")")
		if left(a,1) = "'" then
			a = right(a,len(a) - 1)
		end if
		f.Value = Get_LeftB(a,dsize)
		Call Debug(f.Name & "(" & i & ")" & f & "(" & dsize & ")")
		i = i + 1
'		Call Debug(GetProperties(f))
	next
	on error resume next
		Call rsWayBill.UpdateBatch
		select case DispErr(Err)
		case &h80004005
			Call rsWayBill.CancelUpdate
			Call Debug("����d�o�^��")
		case 0
		case else
			Call rsWayBill.CancelUpdate
			LoadWayBillRow = 0
			Exit Function
		end select
	on error goto 0
	LoadWayBillRow = i - 1
End Function
