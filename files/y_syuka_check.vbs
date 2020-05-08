Option Explicit
' 2014.04.30 �o�ד����{���ɕύX
' 2014.08.25 �o�Ɏc ��\��
' 2016.10.01 �Q�ƃe�[�u���ύX g_syuka �� HMTAH015
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
WScript.Quit(lngRet)
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "�o�׊����`�F�b�N"
	Wscript.Echo "y_syuka_check.vbs [option]"
	Wscript.Echo " /db:newsdc"
	Wscript.Echo " /debug"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	'���O�����I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.UnNamed
		call usage()
		Main = -1
		exit Function
	next
	'���O�t���I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			call usage()
			Main = -1
			exit Function
		case else
			call usage()
			Main = -1
			exit Function
		end select
	next
'	call YSyukaCheck()
	Main = YSyukaCheck()
End Function

Function YSyukaCheck()
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc"))
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	strSql
	strSql = GetSql()
	dim	rsYSyuka
	set rsYSyuka = objDb.Execute(strSql)

	strBuff = "                      ���v"
	strBuff = "12345678 12345678 12345678 123456 123456"
	strBuff = "yyyymmdd 2        1        999999 999999 999999 999999 999999 999999"
	strBuff = "�o�ד�   �����敪 �����敪   ���� �o�Ɏc ���i�c ���M�c ���юc   ����"
	Wscript.Echo strBuff

	dim	lngCnt
	lngCnt = 0
	dim	lngPZan
	lngPZan = 0
	dim	lngKZan
	lngKZan = 0
	dim	lngSZan
	lngSZan = 0
	dim	lngJZan
	lngJZan = 0
	dim	lngQty
	lngQty = 0
	Do While Not rsYSyuka.EOF
		lngCnt	= lngCnt	+ CLng(rsYSyuka.Fields("����"))
		lngPZan = lngPZan	+ CLng(rsYSyuka.Fields("�o�Ɏc"))
		lngKZan = lngKZan	+ CLng(rsYSyuka.Fields("���i�c"))
		lngSZan = lngSZan	+ CLng(rsYSyuka.Fields("���M�c"))
		lngJZan = lngJZan	+ CLng(rsYSyuka.Fields("���юc"))
		lngQty	= lngQty	+ CLng(rsYSyuka.Fields("����"))
		dim	strBuff
		strBuff = ""
		strBuff = strBuff & left(trim(rsYSyuka.Fields("�o�ד�")) & "        ",8)
		strBuff = strBuff & " " & Get_LeftB(trim(rsYSyuka.Fields("�����敪")) & "        ",8)
		strBuff = strBuff & " " & Get_LeftB(trim(rsYSyuka.Fields("�����敪")) & "        ",8)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("����")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("�o�Ɏc")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("���i�c")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("���M�c")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("���юc")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("����")),6)
		Wscript.Echo strBuff
		rsYSyuka.movenext
	Loop
	strBuff = "                      ���v"
	strBuff = strBuff & right("       " & lngCnt,7)
	strBuff = strBuff & right("       " & lngPZan,7)
	strBuff = strBuff & right("       " & lngKZan,7)
	strBuff = strBuff & right("       " & lngSZan,7)
	strBuff = strBuff & right("       " & lngJZan,7)
	strBuff = strBuff & right("       " & lngQty,7)
	Wscript.Echo strBuff
	Wscript.Echo ""
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsYSyuka = CloseRs(rsYSyuka)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
	'-------------------------------------------------------------------
	' ���^�[���l�Z�b�g
	'-------------------------------------------------------------------
	if lngCnt = 0 then
		' �o�ח\��O��(�x��)
		YSyukaCheck = -100
	else
		' �o�׎��јA�g�̎c��
		YSyukaCheck = lngJZan
	end if
End Function

Function GetSql()
	dim	sqlStr
	sqlStr = "select"
	sqlStr = sqlStr & " y.KEY_SYUKA_YMD									""�o�ד�"""
	sqlStr = sqlStr & ",y.KEY_CYU_KBN									""�����敪"""
	sqlStr = sqlStr & ",y.CHOKU_KBN + if(y.CHOKU_KBN='1',' ����','')	""�����敪"""
	sqlStr = sqlStr & ",count(*)										""����"""
	sqlStr = sqlStr & ",sum(if(RTrim(y.KAN_KBN)='0',1,0))				""�o�Ɏc"""
	sqlStr = sqlStr & ",sum(if(RTrim(y.KENPIN_TANTO_CODE)='',1,0))		""���i�c"""
	sqlStr = sqlStr & ",sum(if(y.KEY_CYU_KBN = 'E' or RTrim(y.LK_SEQ_NO)<>'',0,1))				""���M�c"""
	sqlStr = sqlStr & ",sum(if(g.IDno is null,0,1))						""���юc"""
	sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_NUMERIC))				""����"""
	sqlStr = sqlStr & " from y_syuka y"
	sqlStr = sqlStr & " left outer join HMTAH015 g on (y.KEY_ID_NO = g.IDno)"
	sqlStr = sqlStr & " where y.JGYOBA = '00036003'"
	sqlStr = sqlStr & "   and y.DATA_KBN in ('1','3')"
	sqlStr = sqlStr & "   and y.KEY_SYUKA_YMD <= replace(convert(curdate(),SQL_CHAR),'-','')"
	sqlStr = sqlStr & " group by ""�o�ד�"",""�����敪"",""�����敪"""
	sqlStr = sqlStr & " order by ""�o�ד�"",""�����敪"",""�����敪"""
	GetSql = sqlStr
End Function

Function Get_LeftB(byval a_Str,byval a_int)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc�֐��ŕ����R�[�h�擾
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** ���p�͕����R�[�h�̒�����2�A�S�p��4(2�ȏ�)�Ƃ��Ĕ��f
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
		If iLenCount > Cint(a_int) Then
			Exit For
		Else
			iLeftStr = iLeftStr + Mid(a_Str, iCount, 1)
		End If
	Next
	Get_LeftB = iLeftStr
End Function
