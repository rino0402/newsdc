Private Function CheckHead(byval aryHead())
	Call DispMsg("CheckHead()")
	dim	strTable
	strTable = ""
	select case aryHead(0)
	case """PN����_���Y�Ǘ����Ə�R�[�h"""
		strTable = "BoTana"
	case """�݌Ɏ��x_�q�ɃR�[�h"""
		strTable = "BoZaiko"
	case """PN�q�Ɂi���x)_�q�ɃR�[�h"""
		strTable = "BoZaiko"
	end select
	Call DispMsg("CheckHead():" & aryHead(0) & ":" & strTable)
	CheckHead = strTable
End Function

Function Load(byVal strDb,byval strFilename)
	dim	strRet
	strRet = ""
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db",strDb) & ")")
	set objDb = OpenAdodb(GetOption("db",strDb))
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	dim	cnt
	cnt = 0
	dim	objRs
	dim	lngOk
	dim	lngDp
	dim	lngNk
	lngOk = 0
	lngDp = 0
	lngNg = 0
	do while ( objFile.AtEndOfStream = False )
		cnt = cnt + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		Call DispMsg("�Ǎ���..." & cnt & " OK:" & lngOk & " DP:" & lngDp)
		Call Debug(strBuff)
		dim	aryBuff
		aryBuff = GetTab(strBuff)
		if cnt = 1 then
			dim	aryTop
			aryTop = aryBuff
			dim	strTable
			strTable = CheckHead(aryBuff)
			if strTable = "" then
				strRet = strRet & "Error:�Ή����Ă��Ȃ��t�@�C���ł��B" & vbCrlf
				exit do
			end if
			call DispMsg("OpenRs(" & strTable & ")")
			set objRs = OpenRs(objDb,strTable)
			call DispMsg("ExecuteAdodb(" & strTable & ")")
			objDb.CommandTimeout = 0
			Call ExecuteAdodb(objDb,"delete from " & strTable)
		else
			on error resume next
			objRs.AddNew
			on error goto 0
			dim	i
			i = 0
			dim	c
			for each c in (aryBuff)
				c = GetTrim(c)
				Call Debug(i & ":" & objRs.Fields(i).Name & ":" & c)
				if strTable = "BoZaiko" then
					select case i
					case 8
						objRs.Fields(9) = c
					case 9
						objRs.Fields(8) = c
					case else
						objRs.Fields(i) = c
					end select
				else
					objRs.Fields(i) = c
				end if
				i = i + 1
			next
			on error resume next
			Call objRs.UpdateBatch
			select case Err.Number
			case &h80004005
				Call objRs.CancelUpdate
				lngDp = lngDp + 1
				Call DispMsg("����d�o�^��")
			case 0
				lngOk = lngOk + 1
			case else
				strRet = strRet & strBuff & vbCrlf
				strRet = strRet & "�G���[�����F" & cnt & vbCrlf
				strRet = strRet & "Error.Number:0x" & Hex(Err.Number) & vbCrlf
				strRet = strRet & "Error.Description:" & Err.Description & vbCrlf
				lngNg = lngNg + 1
				Call objRs.CancelUpdate
				Exit Do
			end select
			on error goto 0
		end if
	loop
	objFile.Close
	set objFile = nothing
	set objFSO = nothing

	set objRs = CloseRs(objRs)
	set objDb = nothing
	strRet = strRet & "�e�[�u�����F" & strTable & vbCrlf
	strRet = strRet & "�@�@�@����F" & Right(Space(10) & formatnumber(lngOk,0,,-1),10) & "��" & vbCrlf
	strRet = strRet & "�@�@�@�d���F" & Right(Space(10) & formatnumber(lngDp,0,,-1),10) & "��" & vbCrlf
	strRet = strRet & "�@�@�G���[�F" & Right(Space(10) & formatnumber(lngNg,0,,-1),10) & "��" & vbCrlf
	Load = strRet
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

Private Function makeSql()
	dim	strSql
	dim	strTop
	strTop = GetOption("top","")
	if strTop <> "" then
		strTop = " top " & strTop
	end if
	strSql = "select" & strTop
	strSql = strSql & " *"
	strSql = strSql & " from BoTana"
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
