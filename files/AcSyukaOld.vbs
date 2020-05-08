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
	Wscript.Echo "�G�A�R�� �o�׃f�[�^ 201309-201004"
	Wscript.Echo "AcSyuka.vbs [option]"
	Wscript.Echo " /db:<database>"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
	Wscript.Echo " /debug"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript AcSyuka.vbs /db:newsdc4 /load ""I:\0SDC_honsya\���ƕ��ʏ��i���o�׋��z�܂Ƃ�\AC NPL����̏o�׎���\201309AC�o�׎���.xlsx"""
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
'�①�ɏo�׃f�[�^(Excel)�ϊ���MonthlyQty
'-------------------------------------------------------------------
Private Function Load(byval strFilename)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X����
	'-------------------------------------------------------------------
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc4") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc4"))
	'-------------------------------------------------------------------
	'�o�^�p���R�[�h�Z�b�g����
	'-------------------------------------------------------------------
	dim	objRs
	set objRs = OpenRs(objDb,"BoSyuka")
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	dim	objXL
	dim	objBk
	dim	objSt
	Call Debug("CreateObject(Excel.Application)")
	Set objXL = WScript.CreateObject("Excel.Application")
'	objXL.Application.Visible = True
	Call Debug("Workbooks.Open()" & strFilename)
	Set objBk = objXL.Workbooks.Open(strFilename,False,True)
	set objSt = objBk.ActiveSheet
	Call Debug("objSt.Name=" & objSt.Name)
	'-------------------------------------------------------------------
	'Excel�ŏI�s
	'-------------------------------------------------------------------
	Const xlUp = -4162
	dim	lngRowMax
	lngRowMax = objSt.Range("B65536").End(xlUp).Row

	dim	cntAdd
	cntAdd = 0

	'-------------------------------------------------------------------
	'�N���擾
	'-------------------------------------------------------------------
	dim	rngYM
	set rngYM = objSt.Range("BF2")
	dim	strCol
	strCol = ""
	do while strCol <> "K"
		'-------------------------------------------------------------------
		'Excel�񖼎擾
		'-------------------------------------------------------------------
		strCol = Split(rngYM.Address,"$")(1)

		dim	strYM
		strYM = GetYM(rngYM)
		Call DispMsg(strCol & ":�N��:" & strYM)
		if strYM <> "" then
			'-------------------------------------------------------------------
			'�N���f�[�^�폜
			'-------------------------------------------------------------------
			dim	strSql
			strSql = "delete from BoSyuka"
			strSql = strSql & " where ShisanJCode='00025800'"
			strSql = strSql & "   and DT like '" & strYM & "%'"
			Call DispMsg("�폜:" & strSql)
			Call ExecuteAdodb(objDb,strSql)
			'-------------------------------------------------------------------
			'���[�v�F3�`�ŏI�s
			'-------------------------------------------------------------------
			dim lngRow
			for lngRow = 3 to lngRowMax
				'-------------------------------------------------------------------
				'A�F���Y�Ǘ����Ə�
				'-------------------------------------------------------------------
				dim	strJCode
				strJCode = RTrim(objSt.Range("A" & lngRow))
				'-------------------------------------------------------------------
				'C�F�i��
				'-------------------------------------------------------------------
				dim	strPn
				strPn = RTrim(objSt.Range("C" & lngRow))
				'-------------------------------------------------------------------
				'�o�א�
				'-------------------------------------------------------------------
				dim	strQty
				strQty = RTrim(objSt.Range(strCol & lngRow))
				'-------------------------------------------------------------------
				'IdNo
				'-------------------------------------------------------------------
				dim	strIdNo
				strIdNo = strYM & Right("00000" & lngRow,5)
				'-------------------------------------------------------------------
				'���R�[�h�ǉ�
				'-------------------------------------------------------------------
				Call DispMsg("�N��:" & strYM & ":" & strCol & lngRow & ":" & strPn & " " & strQty)
				if strQty <> "" then
					if strPn <> RTrim(objSt.Range("C" & lngRow - 1)) then
						cntAdd = cntAdd + 1
						objRs.AddNew
						objRs.Fields("IdNo") = strIdNo
						objRs.Fields("ShisanJCode") = strJCode
						objRs.Fields("Dt") = strYM & "01"
						objRs.Fields("Pn") = strPn
						objRs.Fields("Qty") = strQty
						objRs.UpdateBatch
					end if
				end if
			next
		end if
		set rngYM = rngYM.Offset(0,-1)
	loop
	dim	strStat
	strStat = "head"

	Call DispMsg("�Ǎ������F" & lngRow)
	Call DispMsg("�o�^�����F" & cntAdd)
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

Private Function GetYM(rngYM)
	GetYM = ""
	dim	strYM
	strYM = rngYM
	if Len(strYM) <> 6 then
		exit function
	end if
	if isNumeric(strYM) = false then
		exit function
	end if
	GetYM = strYM
End Function

Private Function List()
End Function
