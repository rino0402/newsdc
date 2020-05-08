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
	Wscript.Echo "Active�����f�[�^"
	Wscript.Echo "a_order.vbs [option]"
	Wscript.Echo " /list"
	Wscript.Echo " /debug"
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
		case "debug"
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
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	dim	objRs
	set objRs = OpenRs(objDb,"a_order")
	Call ExecuteAdodb(objDb,"delete from a_order where jCode = ''")
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	dim	cnt
	cnt = 0
	dim	cntAdd
	cntAdd = 0
	Dim		aryJCode()
	ReDim	aryJCode(0)
	do while ( objFile.AtEndOfStream = False )
		cnt = cnt + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		Call Debug(strBuff)
		dim		aryTitle
		dim		aryBuff
		if cnt = 1 then
			aryTitle = GetCSV(strBuff)
		elseif cnt > 1 then
			cntAdd = cntAdd + 1
			aryBuff = GetCSV(strBuff)
			objRs.AddNew
			dim	i
			i = 0
			dim	c
			for i = LBound(aryBuff) to UBound(aryBuff)
				c = aryBuff(i)
				dim	strFName
				strFName = GetFName(aryTitle(i))
				Call Debug(i & ":" & strFName & ":" & c)
				if strFName <> "" then
					select case strFName
					case "JCode"
						if chkJCode(aryJCode,c) <> "" then
							Call Debug("delete jCode=" & c )
							Call ExecuteAdodb(objDb,"delete from a_order where jCode = '" & c & "'")
							Call Debug("aryJCode(" & UBound(aryJCode) & ")=" & aryJCode(UBound(aryJCode)))
			                ReDim Preserve aryJCode(UBound(aryJCode) + 1)
            			    aryJCode(UBound(aryJCode)) = c
							Call Debug("aryJCode(" & UBound(aryJCode) & ")=" & aryJCode(UBound(aryJCode)))
						end if
					case "strFName" _
						,"SrvDtSts" _
						,"ChuKbn"
						c = Left(c,1)
'					case 5,8,14,15,18
'						c = Left(c,1)
'					case 20,21,22
'						c = Replace(c,"/","")
					end select
					objRs.Fields(strFName) = c
				end if
			next
			Call objRs.UpdateBatch
		end if
	loop
	objFile.Close
	set objFile = nothing
	set objFSO = nothing

	set objRs = CloseRs(objRs)
	set objDb = nothing
	Call DispMsg("�Ǎ������F" & cnt)
	Call DispMsg("�o�^�����F" & cntAdd)
End Function

Private Function chkJCode(byVal aryJCode(),byVal strJCode)
	dim	a
'	for each a in aryJCode
	dim	i
	Call Debug("chkJCode:" & LBound(aryJCode) & " to " & UBound(aryJCode))
	for i = LBound(aryJCode) to UBound(aryJCode)
		a = aryJCode(i)
		Call Debug("chkJCode:" & a & "=" & strJCode)
		if a = strJCode then
			strJCode = ""
			exit for
		end if
	next
	chkJCode = strJCode
End Function

Private Function GetFName(byval strTitle)
	dim	strFName
	strFName = ""
						'strFName = "BTKbn"		' ���i����敪
	select case strTitle
	case "����敪"
	case "�T�[�r�X�f�[�^�i���敪"
						strFName = "SrvDtSts"	' �T�[�r�X�f�[�^�i���敪
	case "���Y�Ǘ����Ə�R�[�h"
						strFName = "JCode"		' ���Y���Ɓ@���Y�Ǘ����Ə�R�[�h
	case "�i�ڔԍ�"
						strFName = "Pn"			' �o�וi�ڔԍ�
	case "�O���[�o���i�ڔԍ�"
	case "�T�[�r�X�i�ڔԍ�"
	case "��t�i�ڔԍ�"
						strFName = "PnRcv"		' �󒍕i�ڔԍ�
	case "�����R�[�h"
	case "����於"
	case "����"
'						strFName = "QtyRcv"		' �󒍎��ѐ�
						strFName = "QtySnd"		' �󒍎��ѐ�
	case "�P��"
						strFName = "Price"		' �P���@���ےP��    9999999.0000
	case "���ۋ��z"
						strFName = "Amount"		' ���ۋ��z
	case "�I�[�_�[No."
						strFName = "OrderNo"	' �I�[�_�[NO
	case "ITEM-No."
	case "�`�[�ԍ�"
						strFName = "DenNo"		' �`�[�ԍ�
	case "ID-No."
						strFName = "IDNo"		' ID-NO
	case "�݌Ɏ��x������"
						strFName = "ZSyushiRk"	' �݌Ɏ��x������
	case "�݌Ɏ��x�R�[�h"
	case "���Y�Ǘ��݌Ɏ��x�R�[�h"
	case "�⏕�݌Ɏ��x�R�[�h"
	case "���[�敪"
						strFName = "CHKbn"		' ���[�敪
	case "�l���敪"
						strFName = "NSKbn"		' �l���敪
	case "�ԕi�敪"
	case "���ѓ�(�\���)"
						strFName = "SalesDt"	' ����\��N���� yyyymmdd
	case "�󔭒��N����"
						strFName = "RcvDt"		' �󒍔N����
	case "�o�ɔN����"
						strFName = "PckDt"		' �o�ɗ\��N����
	case "�o�הN����"
						strFName = "SndDt"		' �o�ח\��N����
	case "�����N����"
	case "�o�׎w��N����"
	case "�w��[���N����"
						strFName = "DlvDt"		' �w��[�����@�w��[���N����
	case "�[���񓚔N����"
						strFName = "AnsDt"		' �[���񓚓��@�[���񓚔N����
	case "�󒍏o�ׁE�̔��敪"
	case "�󒍏o�ׁE������R�[�h"
	case "�󒍏o�ׁE�����敪"
						strFName = "ChuKbn"		' �����敪
	end select
	GetFName = strFName
End Function

Private Function List()
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	Call DispMsg("objDb.Execute(" & strSql & ")")
	set rsList = objDb.Execute(strSql)

	do while rsList.Eof = False
		Call DispMsg("��" _
			 & " " & rsList.Fields("IDNo") _
			 & " " & rsList.Fields("JCode") _
			 & " " & rsList.Fields("Pn") _
			 & " " & rsList.Fields("PnRcv") _
			 & " " & rsList.Fields("BTKbn") _
			 & " " & rsList.Fields("TKCode") _
			 & " " & rsList.Fields("ChokuCode") _
			 & " " & rsList.Fields("SrvDtSts") _
					)
		Call rsList.MoveNext
	loop

	Call Debug("CloseAdodb(" & GetOption("db","newsdc") & ")")
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
	strSql = strSql & " from a_order"
	makeSql = strSql
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
