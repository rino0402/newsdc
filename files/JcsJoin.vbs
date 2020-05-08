Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "JcsJoin.vbs [option]"
	Wscript.Echo " /db:newsdc7 �f�[�^�x�[�X"
	Wscript.Echo " /s:10000    �J�n�s"
	Wscript.Echo " /l:100      �ǂݍ��ލs��"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript JcsJoin.vbs /db:newsdc7 Jcs\l130859.csv"
	Wscript.Echo "cscript JcsJoin.vbs /db:newsdc7 Jcs\l131010.csv"
	Wscript.Echo "cscript JcsJoin.vbs /db:newsdc7 Jcs\l131021.csv"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objJcsJoin
	Set objJcsJoin = New JcsJoin
	if objJcsJoin.Init() <> "" then
		call usage()
		exit function
	end if
	call objJcsJoin.Run()
End Function
'-----------------------------------------------------------------------
'JcsJoin
'-----------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Class JcsJoin
	Private	strDBName
	Private	objDB
	Private	objRs
	Public	strJGYOBU
	Private	strAction
	Private	strFileName
	'-----------------------------------------------------------------------
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'���b�Z�[�W�\��
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'�I�v�V�����擾
	'-----------------------------------------------------------------------
	Function GetOption(byval strName _
					  ,byval strDefault _
					  )
		dim	strValue

		if strName = "" then
			strValue = ""
			if strDefault < WScript.Arguments.UnNamed.Count then
				strValue = WScript.Arguments.UnNamed(strDefault)
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
	'-----------------------------------------------------------------------
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		strAction = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		if WScript.Arguments.UnNamed.Count = 0 then
			Init = "�t�@�C�����w��"
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "s"
			case "l"
			case "debug"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'CheckFunction
	'-----------------------------------------------------------------------
	Private Function CheckFunction(byval strA)
		Debug ".CheckFunction():" & strA
		CheckFunction = False
		if strAction = "" then
			exit function
		end if
		if WScript.Arguments.Named.Exists(strA) then
			exit function
		end if
		CheckFunction = True
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			strFileName = strArg
			Call Load()
		Next
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Load() �Ǎ�
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load():" & strFileName
			Call OpenCsv()
			Call LoadCsv()
			Call CloseCsv()
	End Function
	'-------------------------------------------------------------------
	'��΃p�X
	'-------------------------------------------------------------------
	Private Function GetAbsPath(byVal strPath)
		Dim objFileSys
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		strPath = objFileSys.GetAbsolutePathName(strPath)
		Set objFileSys = Nothing
		GetAbsPath = strPath
	End Function
	'-------------------------------------------------------------------
	' csv
	'-------------------------------------------------------------------
	Private	objFSO
	Private	objFile
	'-------------------------------------------------------------------
	' csv Open
	'-------------------------------------------------------------------
	Private Function OpenCsv()
		Debug ".OpenCsv():" & strFileName
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	End Function
	'-------------------------------------------------------------------
	' csv Load
	'-------------------------------------------------------------------
	Private	strBuff
	Private	aryBuff
	Private bStop
	Private	lngRow
	Private	objJoin
	Private Function LoadCsv()
		Debug ".LoadCsv()"
		lngRow = 0
		bStop = False
		do while ( objFile.AtEndOfStream = False )
			strBuff = objFile.ReadLine()
'			Debug strBuff
			aryBuff = GetCSV(strBuff)
			lngRow = lngRow + 1
			if lngRow = 1 then
				Call HeaderCheck()
				if objJoin is Nothing then
					Disp "�f�[�^�`���F�s��"
					exit do
				end if
			else
				Call AddRecord()
			end if
			if bStop then
				exit do
			end if
		loop
	End Function
	'-------------------------------------------------------------------
	' HeaderCheck
	'-------------------------------------------------------------------
	Private Function HeaderCheck()
		Debug ".HeaderCheck()"
		' Join����f�[�^
		Set objJoin = New JoinHikitori
		HeaderCheck = objJoin.Init(me,strBuff)
		if HeaderCheck then
			exit function
		end if
'		Set objJoin = Nothing
	End Function
	'-------------------------------------------------------------------
	' AddRecord
	'-------------------------------------------------------------------
	Private Function AddRecord()
		Debug ".AddRecord()"
		if objJoin is nothing then
			exit function
		end if
		objJoin.AddRecord(aryBuff)
	End Function
	'-------------------------------------------------------------------
	' csv Close
	'-------------------------------------------------------------------
	Private Function CloseCsv()
		Debug ".CloseCsv()"
		objFile.Close
		set objFile = nothing
		set objFSO = nothing
	End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'Sql���s
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
'		on error resume next
		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field��
	'-------------------------------------------------------------------
	Public Function GetFields(byVal strTable)
		Debug ".GetFields():" & strTable
		dim	strFields
		strFields = ""
		dim	objRs
		set objRS = objDB.Execute("select top 1 * from " & strTable)
		dim	objF
		for each objF in objRS.Fields
			if strFields <> "" then
				strFields = strFields & ","
			end if
			strFields = strFields & objF.Name
		next
		set objRs = nothing
		GetFields = strFields
	End Function
End Class

'Jcs Join���� 
Const cnsHikitori = "��������,�������敪,�A��,�ڋq�[�i�ԍ�,�ڋq���i�ԍ�,�ڋq�[����,�ڋq�[�����V�t�g,�ڋq�[������,�����ԍ�,��Ǝw���ԍ�,���i�ԍ�,����敪,�\��,���i����,���_,�[����,�[�����V�t�g,�[������,�����,����萔��,�]�[��,���P�[�V����,�����ԍ�,�[���ύX��,�[���ύX����,��z�ԍ�,��z���[�敪,���s�N����,���s����,SDC�݌ɐ���"
Class JoinHikitori
	'��������,�������敪,�A��,�ڋq�[�i�ԍ�,�ڋq���i�ԍ�,�ڋq�[����,�ڋq�[�����V�t�g,�ڋq�[������,�����ԍ�,��Ǝw���ԍ�,���i�ԍ�,����敪,�\��,���i����,���_,�[����,�[�����V�t�g,�[������,�����,����萔��,�]�[��,���P�[�V����,�����ԍ�,�[���ύX��,�[���ύX����,��z�ԍ�,��z���[�敪,���s�N����,���s����,SDC�݌ɐ��� 
	Private	objParent
	Public Function Init(oParent,byVal strBuff)
		oParent.Debug ".Init()"
		oParent.Debug strBuff
		oParent.Debug cnsHikitori
		Init = false
		set objParent = oParent
		if Trim(strBuff) = Trim(cnsHikitori) then
			Init = true
			exit function
		end if
	End Function
	Public Function AddRecord(byVal aryBuff)
		objParent.Debug ".AddRecord():" & aryBuff(1)
'		objParent.Debug objParent.strBuff
	End Function
End Class

Private Function GetCSV(ByVal s)
    Const One = 1
    ReDim r(0)

    Const sUndef = 11 ' ���m��(�J���}���_�u���N�H�[�e�[�V�������u�X�y�[�X�ȊO�̕����v��҂��)
    Const sQuot = 22 ' �_�u���N�H�[�e�[�V�����ň͂܂ꂽ���Ƃ��J�n���Ă��܂������(�_�u���N�H�[�e�[�V��������т��̌�̃J���}�҂�)
    Const sPlain = 33 ' �_�u���N�H�[�e�[�V�����Ȃ��̂��Ƃ��J�n���Ă��܂������(�J���}�҂�)
    Const sTerm = 44 ' �_�u���N�H�[�e�[�V�����ň͂܂ꂽ���Ƃ��I�����Ă��܂������(�J���}�҂�)
    Const sEsc = 55 ' �_�u���N�H�[�e�[�V�����ň͂܂ꂽ���Ƃ��J�n���Ă��܂�����ԂŁA���_�u���N�H�[�e�[�V�������o��������ԁB
    Dim w
    w = sUndef

    Dim a
    a = ""
    Dim i
    For i = 0 To Len(s) - One + 1
        Dim c
        c = Mid(s, i + One, 1)
        If c = """" Then
            If w = sUndef Then
                a = ""
                w = sQuot
            ElseIf w = sQuot Then
                w = sEsc
            ElseIf w = sPlain Then ' �G���[
                ReDim r(0)
                Exit For
            ElseIf w = sTerm Then ' �G���[
                ReDim r(0)
                Exit For
            ElseIf w = sEsc Then
                a = a & c
                w = sQuot
            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
            End If
        ElseIf c = "," Then
            If w = sUndef Then
                ReDim Preserve r(UBound(r) + 1)
                r(UBound(r)) = ""
            ElseIf w = sQuot Then
                a = a & c
            ElseIf w = sPlain Then
                ReDim Preserve r(UBound(r) + 1)
                r(UBound(r)) = RTrim(a)
                a = ""
                w = sUndef
            ElseIf w = sTerm Then
                a = ""
                w = sUndef
            ElseIf w = sEsc Then
                ReDim Preserve r(UBound(r) + 1)
                r(UBound(r)) = a
                a = ""
                w = sUndef
            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
            End If
        ElseIf c = " " Then
            If w = sUndef Then
                ' do nothing.
            ElseIf w = sQuot Then
                a = a & c
            ElseIf w = sPlain Then
                a = a & c
            ElseIf w = sTerm Then
                ' do nothing
            ElseIf w = sEsc Then
                ReDim Preserve r(UBound(r) + 1)
                r(UBound(r)) = a
                a = ""
                w = sTerm
            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
            End If
        ElseIf c = "" Then ' �ŏI���[�v�̂�
            If w = sUndef Then
                ReDim Preserve r(UBound(r) + 1)
                r(UBound(r)) = ""
            ElseIf w = sQuot Then
                ReDim r(0)
                Exit For
            ElseIf w = sPlain Then
                ReDim Preserve r(UBound(r) + 1)
                r(UBound(r)) = RTrim(a)
                a = ""
                w = sUndef
            ElseIf w = sTerm Then
                ' do nothing
            ElseIf w = sEsc Then
                ReDim Preserve r(UBound(r) + 1)
                r(UBound(r)) = RTrim(a)
                a = ""
                w = sUndef
            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
            End If
        Else
            If w = sUndef Then
                a = a & c
                w = sPlain
            ElseIf w = sQuot Then
                a = a & c
            ElseIf w = sPlain Then
                a = a & c
            ElseIf w = sTerm Then
                ReDim r(0)
                Exit For
            ElseIf w = sEsc Then
                ReDim r(0)
                Exit For
            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
            End If
        End If
    Next

    GetCSV = r
End Function
