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
	Wscript.Echo "BO�݌Ƀf�[�^ �①�ɗp"
	Wscript.Echo "BoZaikoRf.vbs [option]"
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
	set objRs = OpenRs(objDb,"BoZaikoRf")
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

	dim	cnt
	cnt = 0
	dim	cntAdd
	cntAdd = 0

	Const xlUp = -4162
	dim	lngRowMax
	lngRowMax = objSt.Range("B65536").End(xlUp).Row
	dim	strJCode
	dim	strShisanJCode
	strJCode 		= ""
	strShisanJCode	= ""
	dim	strStat
	strStat = "head"
	dim lngRow
	for lngRow = 1 to lngRowMax
		dim	strB
		strB = objSt.Range("B" & lngRow)
		Call Debug(lngRow & ":" & strStat & ":" & strB)
		select case strStat
		case "head"
			select case strB
			case ""
			case "���ÃZ���^�q�ɍ݌�"
				strStat = "title"
				strJCode		= "00021259"
				strShisanJCode	= "00021259"
			case "�o�o�r�b�ޗǍ݌Ɍ���"
				strStat = "title"
				strJCode		= "00036003"
				strShisanJCode	= "00021259"
			case else
				Call DispMsg("�w�b�_�[�G���[�F" & strB)
			end select
		case "title"
			select case strB
			case ""
			case "�i�ڔԍ�"
				dim	strDeleteSql
				strDeleteSql = "delete from BoZaikoRf where jCode = '" & strJCode & "' and ShisanJCode = '" & strShisanJCode & "'"
				Call Debug(strDeleteSql)
				Call ExecuteAdodb(objDb,strDeleteSql)
				strStat = "value"
			case else
				Call DispMsg("���ږ��G���[�F" & strB)
			end select
		case "value"
			'	 JCode			Char( 8) default '' not null	// ���Ə�R�[�h
			'	,ShisanJCode	Char( 8) default '' not null	// ���Y�Ǘ����Ə�R�[�h
			'	,Pn				Char(20) default '' not null	// �i�ڔԍ�
			'	,PName			Char(40) default '' not null	// �i�ږ�
			'	,DModel			Char(20) default '' not null	// ��\�@��
			'	,HikiQty		CURRENCY default 0  not null	// ���������\�݌ɐ�
			'	,Hinmoku		Char(10) default '' not null	// �o�m����_�i�ڃR�[�h�Q
			'	,SyuShi			Char( 8) default '' not null	// �݌Ɏ��x�R�[�h
			cntAdd = cntAdd + 1
			objRs.Addnew
			objRs.Fields("JCode")		= strJCode
			objRs.Fields("ShisanJCode")	= strShisanJCode
			objRs.Fields("Pn")			= strB
			objRs.Fields("PName")		= RTrim(objSt.Range("C" & lngRow))
			objRs.Fields("DModel")		= RTrim(objSt.Range("D" & lngRow))
			objRs.Fields("HikiQty")		= RTrim(objSt.Range("E" & lngRow))
			objRs.Fields("Hinmoku")		= RTrim(objSt.Range("F" & lngRow))
			objRs.Fields("SyuShi")		= RTrim(objSt.Range("G" & lngRow))
			objRs.UpdateBatch
		end select
	next
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
	strSql = strSql & " from BoZaikoRf"
	makeSql = strSql
End Function

