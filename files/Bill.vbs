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
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "�o�׏��i����p�����f�[�^"
	Wscript.Echo "Bill.vbs [option]"
	Wscript.Echo " /db:<database>"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
	Wscript.Echo " /debug"
	Wscript.Echo "Ex."
	Wscript.Echo "Bill.vbs /db:newsdc-ono /load ""\\hs1\sec\ppsc\PPSC��o�������ߋ���\201204����\����\04PPSC�������iIHCS���j.xlsx"""
	Wscript.Echo "----"
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
		case "jgyobu"
		case "load"
		case "top"
		case "debug"
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
	case "usage"
		Call usage()
	end select
	Main = 0
End Function

Private Function GetFunction()
	GetFunction = "usage"
	if WScript.Arguments.Named.Exists("load") then
		GetFunction = "load"
	elseif WScript.Arguments.Named.Exists("list") then
		GetFunction = "list"
	end if
End Function

'-------------------------------------------------------------------
'�����f�[�^(Excel)��Bill
'-------------------------------------------------------------------
Private Function Load(byval strFilename)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X����
	'-------------------------------------------------------------------
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	'-------------------------------------------------------------------
	'�o�^�p���R�[�h�Z�b�g����
	'-------------------------------------------------------------------
	dim	objRs
	set objRs = OpenRs(objDb,"Bill")
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	dim	objXL
	dim	objBk
	Call Debug("CreateObject(Excel.Application)")
	Set objXL = WScript.CreateObject("Excel.Application")
'	objXL.Application.Visible = True
	Call Debug("Workbooks.Open()" & strFilename)
	Set objBk = objXL.Workbooks.Open(strFilename,False,True)
	dim	strJGyobu
	strJGyobu = GetJGyobu(objBk)
	if strJGyobu = "" then
		call DispMsg("���ƕ��s��")
	else
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"�D����")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"�F�����o��")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"�@�A�B�C�o�ז���")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"�A�B�C�D�`�b���i��")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"�A�B���i���H��")
	end if
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	Call objBk.Close(False)
	set objBk = Nothing
	set objXL = Nothing
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̌㏈��
	'-------------------------------------------------------------------
	set objRs = CloseRs(objRs)
	set objDb = nothing
End Function

Function GetJGyobu(objBk)
	dim	strJGyobu
	strJGyobu = GetOption("jgyobu","")
	if strJGyobu <> "" then
		GetJGyobu = strJGyobu
		exit function
	end if
	dim	objSt
	for each objSt in objBk.Worksheets
		dim	strName
		select case objSt.Name
'		case "��������"
'			strName = Trim(objSt.Range("D9"))
'			select case strName
'			case "�G�A�R��"
'				strJGyobu = "A"
'			case "�①��"
'				strJGyobu = "R"
'			end select
'			Call Debug("GetJGyobu():" & strName & ":" & strJGyobu)
'			exit for
		case "��������","�������׃t�H�[�� (2)","�������׃t�H�[��(2)","�������ׁi��o�p�j","�������� (2)"
			strName = Trim(objSt.Range("C8"))
			if strName = "" then
				strName = Trim(objSt.Range("D8"))
				if strName = "" or strName = "���ꕨ���Z���^�[" then
					' �������� �G�A�R��/�①��
					strName = Trim(objSt.Range("D9"))
				end if
			end if
			select case strName
			case "����p�[�c�Z���^�[�@�iIH���ݸ�˰�����j"
				strJGyobu = "D"
			case "����p�[�c�Z���^�[�@�i�r���[�e�B�E���r���OBU���j","����p�[�c�Z���^�[�@�i�r���[�e�B�E���r���O���j","����p�[�c�Z���^�[�@�i�����������j"
				strJGyobu = "5"
			case "����p�[�c�Z���^�[�@�i���ы@�핪�j"
				strJGyobu = "4"
			case "����p�[�c�Z���^�["
				strJGyobu = "7"
			case "�G�A�R��"
				strJGyobu = "A"
			case "�①��"
				strJGyobu = "R"
			end select
			Call Debug("GetJGyobu():" & strName & ":" & strJGyobu)
			exit for
		end select
	next
	GetJGyobu = strJGyobu
End Function

Function GetBillSheet(objBk,byVal strSheetName)
	set GetBillSheet = Nothing
	dim	objSt
	for each objSt in objBk.Worksheets
		if objSt.Name = strSheetName then
			set GetBillSheet = objSt
			exit for
		end if
		select case strSheetName
		case "�@�A�B�C�o�ז���"
			if objSt.Name = "�B�C�D�E�o�ז���" then
				set GetBillSheet = objSt
				exit for
			end if
		case "�F�����o��"
			if objSt.Name = "�I�����o��" then
				set GetBillSheet = objSt
				exit for
			end if
		case "�D����"
			if objSt.Name = "�F����" then
				set GetBillSheet = objSt
				exit for
			end if
		end select
	next
End Function


Function LoadBill(objDb,objRs,objBk,byVal strJGyobu,byVal strSheetName)

	Call Debug("LoadBill(" & strJGyobu & "," & strSheetName & ")")
	dim	objSt
	set objSt = GetBillSheet(objBk,strSheetName)
	if objSt is Nothing then
		Call DispMsg("LoadBill():" & strSheetName & "�F" & "�w��V�[�g������܂���.")
		exit function
	end if
	dim	strBillDt
	strBillDt = ""
	dim	strYM
	strYM = ""
	dim	strKBN
	select case strSheetName
	case "�@�A�B�C�o�ז���","�B�C�D�E�o�ז���"
		strKBN = "A"
		strBillDt = GetDt(RTrim(objSt.Range("C4")))
	case "�F�����o��","�I�����o��"
		strKBN = "B"
		strBillDt = GetDt(RTrim(objSt.Range("C3")))
		if strBillDt = "" then
			strBillDt = GetDt(RTrim(objSt.Range("C4")))
		end if
	case "�D����","�F����"
		strKBN = "C"
		strBillDt = GetDt(RTrim(objSt.Range("B4")))
	case "�A�B�C�D�`�b���i��"
		strKBN = "D"
		strBillDt = GetDt(RTrim(objSt.Range("M3")))
	case "�A�B���i���H��"
		strKBN = "E"
		strBillDt = GetDt(RTrim(objSt.Range("H2")))
	case else
		Exit Function
	end select
	strYM = GetYM(strBillDt)
	if strYM <> "" then
		Call DispMsg(strSheetName & "�F" & strYM)
		'-------------------------------------------------------------------
		'�N���f�[�^�폜
		'-------------------------------------------------------------------
		dim	strSql
		strSql = "delete from Bill" _
			   & " where JGyobu = '" & strJGyobu & "'" _
			   & "   and BillDt = '" & strBillDt & "'" _
			   & "   and YM = '" & strYM & "'" _
			   & "   and KBN = '" &	strKBN & "'"
		Call Debug("�폜:" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	end if
	'-------------------------------------------------------------------
	'Excel�ŏI�s
	'-------------------------------------------------------------------
	Const xlUp = -4162
	dim	lngRowTop
	dim	lngRowMax
	select case strKBN
	case "D"
		lngRowTop = 11
		lngRowMax = objSt.Range("B65536").End(xlUp).Row
	case "E"
		lngRowTop = 11
		lngRowMax = objSt.Range("A65536").End(xlUp).Row
	case else
		lngRowTop = 4
		lngRowMax = objSt.Range("A65536").End(xlUp).Row
	end select

	dim	cntAdd
	cntAdd = 0

	'-------------------------------------------------------------------
	'���[�v�F3�`�ŏI�s
	'-------------------------------------------------------------------
	dim	lngRow
	for lngRow = lngRowTop to lngRowMax
		Call Debug(strJGyobu & " " & strYM & " " & strBillDt & " " & lngRow & _
					" " & RTrim(objSt.Range("A" & lngRow)) & _
					" " & RTrim(objSt.Range("B" & lngRow)) & _
					" " & RTrim(objSt.Range("C" & lngRow)) & _
					" " & RTrim(objSt.Range("D" & lngRow)) & _
					" " & RTrim(objSt.Range("E" & lngRow)) & _
					" " & RTrim(objSt.Range("F" & lngRow)) _
					)
		dim	lngNo
		if strKBN = "D" then
			lngNo = GetNumValue(RTrim(objSt.Range("B" & lngRow)))
		else
			lngNo = GetNumValue(RTrim(objSt.Range("A" & lngRow)))
		end if
		if lngNo = 0 then
			Call Debug("Exit for:lngNo = 0")
			exit for
		end if
		cntAdd = cntAdd + 1
		objRs.AddNew
		objRs.Fields("JGyobu") 		= strJGyobu			'// ���ƕ�
		objRs.Fields("BillDt")		= strBillDt			'// ������
		objRs.Fields("YM") 			= strYM				'// �����N��
		objRs.Fields("KBN") 		= strKBN			'// �����敪
														'// 1:PPSC�o��
														'// 2:PPSC�����o��
		objRs.Fields("No") 			= lngNo				'// ����������No
		select case strKBN
		case "A"	'�@�A�B�C�o�ז���
			objRs.Fields("IdNo") 		= RTrim(objSt.Range("B" & lngRow))		'// ID-No
			objRs.Fields("Dt") 			= Replace(RTrim(objSt.Range("C" & lngRow)),"/","")		'// �o�ד�
			objRs.Fields("DenNo") 		= RTrim(objSt.Range("D" & lngRow))		'// �`�[�ԍ�
			objRs.Fields("SyukaCd")		= RTrim(objSt.Range("E" & lngRow))		'// �o�א�
			objRs.Fields("SyukaNm")		= RTrim(objSt.Range("F" & lngRow))		'// �o�א於
			objRs.Fields("Pn") 			= RTrim(objSt.Range("G" & lngRow))		'// �i��
			objRs.Fields("PnName") 		= RTrim(objSt.Range("H" & lngRow))		'// �i��
			objRs.Fields("Qty") 		= RTrim(objSt.Range("I" & lngRow))		'// �o�א�
			objRs.Fields("Pick") 		= RTrim(objSt.Range("J" & lngRow))		'// �o�ɍH��
			objRs.Fields("Ship") 		= RTrim(objSt.Range("K" & lngRow))		'// �o�׍H��
			objRs.Fields("AnyKbn") 		= RTrim(objSt.Range("M" & lngRow))		'// �敪
			objRs.Fields("KoryoPrc")	= RTrim(objSt.Range("N" & lngRow))		'// ���H���P��
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("O" & lngRow))		'// ���H��
			objRs.Fields("HakoPrc") 	= RTrim(objSt.Range("P" & lngRow))		'// ������P��
			objRs.Fields("Hako") 		= RTrim(objSt.Range("Q" & lngRow))		'// ������
		case "B"	'�F�����o��
			objRs.Fields("IdNo") 		= RTrim(objSt.Range("B" & lngRow))		'// ID-No
			objRs.Fields("Dt") 			= Replace(RTrim(objSt.Range("C" & lngRow)),"/","")		'// �o�ד�
			objRs.Fields("DenNo") 		= RTrim(objSt.Range("D" & lngRow))		'// �`�[�ԍ�
			objRs.Fields("SyukaCd")		= RTrim(objSt.Range("E" & lngRow))		'// �o�א�
			objRs.Fields("SyukaNm")		= RTrim(objSt.Range("F" & lngRow))		'// �o�א於
			objRs.Fields("Pn") 			= RTrim(objSt.Range("G" & lngRow))		'// �i��
			objRs.Fields("PnName") 		= ""									'// �i��
			objRs.Fields("Qty") 		= RTrim(objSt.Range("I" & lngRow))		'// �o�א�
			objRs.Fields("Pick") 		= RTrim(objSt.Range("K" & lngRow))		'// �o�ɍH��
			objRs.Fields("Ship") 		= 0										'// �o�׍H��
			dim	strCol
			strCol = ""
'			if objSt.Range("P2") = "��or��" then
'				strCol = "P"
'			elseif objSt.Range("Q2") = "��or��" then
'				strCol = "Q"
'			else
'			end if
			if strCol <> "" then
				objRs.Fields("AnyKbn") 		= RTrim(objSt.Range(strCol & lngRow))	'// �敪
			end if
			objRs.Fields("KoryoPrc")	= RTrim(objSt.Range("L" & lngRow))		'// ���H���P��
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("M" & lngRow))		'// ���H��
			objRs.Fields("HakoPrc") 	= RTrim(objSt.Range("N" & lngRow))		'// ������P��
			objRs.Fields("Hako") 		= RTrim(objSt.Range("O" & lngRow))		'// ������
		case "C"	'�D����
			objRs.Fields("Dt") 			= Replace(RTrim(objSt.Range("B" & lngRow)),"/","")		'���ɓ�
			objRs.Fields("DenNo") 		= RTrim(objSt.Range("C" & lngRow))						'�`��
			objRs.Fields("SyukaCd")		= RTrim(objSt.Range("D" & lngRow))						'�����
			objRs.Fields("Pn") 			= RTrim(objSt.Range("E" & lngRow))		'�i��
			objRs.Fields("PnName") 		= RTrim(objSt.Range("F" & lngRow))		'�i��
			objRs.Fields("Qty") 		= RTrim(objSt.Range("G" & lngRow))		'����
																				'�I��
			objRs.Fields("AnyKbn") 		= Get_LeftB(RTrim(objSt.Range("I" & lngRow)),10)		'���ɋ敪
																				'���ɍH�� �P��
			objRs.Fields("Pick") 		= RTrim(objSt.Range("K" & lngRow))		'���ɍH�� ���z
		case "D"	'�A�B�C�D�`�b���i��
			objRs.Fields("Dt") 			= GetDt(objSt.Range("C" & lngRow))		'�����
			objRs.Fields("Pn") 			= Get_LeftB(RTrim(objSt.Range("D" & lngRow)),20)		'�i��
			objRs.Fields("PnName") 		= RTrim(objSt.Range("E" & lngRow))		'�i��
			objRs.Fields("Qty") 		= RTrim(objSt.Range("F" & lngRow))		'����
			objRs.Fields("KoryoPrc") 	= RTrim(objSt.Range("G" & lngRow))		'�H����
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("H" & lngRow))		'�H�����z
			objRs.Fields("HakoPrc") 	= RTrim(objSt.Range("I" & lngRow))		'���し
			objRs.Fields("Hako") 		= RTrim(objSt.Range("J" & lngRow))		'������z
			objRs.Fields("GaisoPrc") 	= RTrim(objSt.Range("K" & lngRow))		'�O����
			objRs.Fields("Gaiso") 		= RTrim(objSt.Range("L" & lngRow))		'�O�����z
			objRs.Fields("FutaiPrc") 	= RTrim(objSt.Range("M" & lngRow))		'�t�с�
			objRs.Fields("Futai") 		= RTrim(objSt.Range("N" & lngRow))		'�t�ы��z
		case "E"	'�A�B���i���H��
			objRs.Fields("Dt") 			= GetDt(objSt.Range("B" & lngRow))		'�����
			objRs.Fields("Pn") 			= RTrim(objSt.Range("C" & lngRow))		'�i��
			objRs.Fields("PnName") 		= RTrim(objSt.Range("D" & lngRow))		'�i��
			objRs.Fields("Qty") 		= RTrim(objSt.Range("E" & lngRow))		'����
			objRs.Fields("KoryoPrc") 	= RTrim(objSt.Range("F" & lngRow))		'�H����
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("G" & lngRow))		'�H�����z
			objRs.Fields("FutaiPrc") 	= RTrim(objSt.Range("H" & lngRow))		'�t����Ɓ�
			objRs.Fields("Futai") 		= RTrim(objSt.Range("I" & lngRow))		'�t����Ƌ��z
		end select
		objRs.UpdateBatch
	next
	dim	strStat
	strStat = "head"

	Call DispMsg("  �Ǎ������F" & lngRow)
	Call DispMsg("  �o�^�����F" & cntAdd)

End Function

Function GetDt(byVal strDt)
	strDt = RTrim(strDt)
	if inStr(strDt,"/") > 0 then
		strDt = Replace(strDt,"/","")
	end if
	GetDt = left(strDt,8)
End Function

Function GetNumValue(strV)
	dim	dblV
' for debug
'	Wscript.Echo "GetNumValue(" & len(rtrim(strV)) & " " & rtrim(strV) & ")"
' for debug
	dblV = 0
	if isnumeric(strV) = True then
		dblV = cdbl(strV)
	end if
	GetNumValue = dblV
End Function

Private Function GetYM(byVal strDt)
	dim	strYM
	dim	iY
	dim	iM
	dim	iD
	Call Debug("GetYM(" & strDt & ")")
	if inStr(strDt,"/") > 0then
		iY = CInt(Split(strDt,"/")(0))
		iM = CInt(Split(strDt,"/")(1))
		iD = CInt(Split(strDt,"/")(2))
	else
		iY = CInt(Left(strDt,4))
		iM = CInt(Mid(strDt,5,2))
		iD = CInt(Right(strDt,2))
	end if
	if iD > 20 Then
		iM = iM + 1
	end if
	if iM > 12 Then
		iY = iY + 1
		iM = 1
	end if
	strYM = iY & Right("0" & iM,2)
	GetYM = strYM
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
		' DT 	JGYOBU 	NAIGAI 	HIN_GAI 	SyukaCnt 	SyukaQty
		Call DispMsg("��" _
			 & " " & rsList.Fields("DT") _
			 & " " & rsList.Fields("JGYOBU") _
			 & " " & rsList.Fields("NAIGAI") _
			 & " " & rsList.Fields("HIN_GAI") _
			 & " " & rsList.Fields("SyukaCnt") _
			 & " " & rsList.Fields("SyukaQty") _
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
	strSql = strSql & " from MonthlyQty"
	makeSql = strSql
End Function

