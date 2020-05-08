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
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "BO�݌Ƀf�[�^"
	Wscript.Echo "BoZaiko.vbs [option]"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	For Each strArg In WScript.Arguments.UnNamed
    	select case strArg
		case else
			if strFilename <> "" then
				usage()
				Main = 1
				exit Function
			end if
			strFilename = strArg
		end select
	Next
	For Each strArg In WScript.Arguments.Named
    	select case lcase(strArg)
		case "db"
		case "list"
		case "load"
		case "top"
		case else
			usage()
			Main = 1
			exit Function
		end select
	Next
	select case GetFunction()
	case "list"
		Call List()
	case "load"
		Call Load(strFilename)
	end select
	Main = 0
End Function

Private Function GetFunction()
	GetFunction = "list"
	if WScript.Arguments.Named.Exists("load") then
		GetFunction = "load"
	end if
End Function

Private Function Load(byval strFilename)
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	dim	objRs
	set objRs = OpenRs(objDb,"BoZaiko")
	Call ExecuteAdodb(objDb,"delete from BoZaiko")
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	dim	cnt
	cnt = 0
	do while ( objFile.AtEndOfStream = False )
		cnt = cnt + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		Call DispMsg(strBuff)
		if cnt > 1 then
			dim	aryBuff
			aryBuff = GetTab(strBuff)
			objRs.AddNew
			dim	i
			i = 0
			dim	c
			for each c in (aryBuff)
				c = GetTrim(c)
				Call DispMsg(i & ":" & c)
				select case i
				case else
				end select
				objRs.Fields(i) = c
				i = i + 1
			next
			Call objRs.UpdateBatch
		end if
	loop
	objFile.Close
	set objFile = nothing
	set objFSO = nothing

	set objRs = CloseRs(objRs)
	set objDb = nothing
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

Private Function List()
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	Call DispMsg("objDb.Execute(" & strSql & ")")
	set rsList = objDb.Execute(strSql)

	do while rsList.Eof = False
		Call DispMsg("��" _
			 & " " & rsList.Fields("Soko") _
			 & " " & rsList.Fields("JCode") _
			 & " " & rsList.Fields("ShisanJCode") _
			 & " " & rsList.Fields("Pn") _
			 & " " & rsList.Fields("SyuShi") _
			 & " " & rsList.Fields("TanaQty") _
			 & " " & rsList.Fields("HikiQty") _
			 & " " & rsList.Fields("SyuShiR") _
			 & " " & rsList.Fields("SyuShiName") _
					)
		Call rsList.MoveNext
	loop

	Call DispMsg("CloseAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = Nothing
End Function

Private Function makeSql()
	dim	strSql
	dim	strTop
	strTop = GetOption("top","")
	if strTop <> "" then
		strTop = " top " & strTop
	end if
	strSql = "select" & strTop
	strSql = strSql & " *"
	strSql = strSql & " from BoZaiko"
	makeSql = strSql
End Function

Function GetTab(ByVal s)
    Dim r
	r = Split(s,vbTab)
	GetTab = r
End Function

Function GetCSV(ByVal s)
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
