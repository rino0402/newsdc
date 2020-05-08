Option Explicit
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
End Function
Call Include("const.vbs")

Call Main()

function usage()
'    Wscript.Echo "���R�ʉ^���n�f�[�^�o��(2010.03.01) �d�ʂ���̎��A�ː�=000 ���Z�b�g"
'    Wscript.Echo "���R�ʉ^���n�f�[�^�o��(2010.03.16) 1���ڃL�����Z���̑Ή�"
'    Wscript.Echo "���R�ʉ^���n�f�[�^�o��(2010.04.05) TelNo�Ή�(2�ւ��) ���X�֔ԍ��̓f�[�^���ڂɂȂ�"
'    Wscript.Echo "���R�ʉ^���n�f�[�^�o��(2010.06.11) IDNo�ɏo�ד���ǉ�"
'    Wscript.Echo "���R�ʉ^���n�f�[�^�o��(2011.10.05) �ː����Œ�1�ɂȂ�悤�ɕύX"
    Wscript.Echo "���R�ʉ^���n�f�[�^�o��(2019.10.25) �����F310% �����O"
	Wscript.Echo "y_syuka_h.vbs [option] <yyyymmdd>"
	Wscript.Echo "               -del : del_syuka_h ���Q��(�f�t�H���g y_syuka_h)"
	Wscript.Echo "               -b1  : 1��"
	Wscript.Echo "               -b2  : 2��"
	Wscript.Echo "               -b3  : 3��"
	Wscript.Echo "               -label : �׎D���x���f�[�^�o��"
    Wscript.Echo "               -?"
end function

Sub Main()
	dim	db
	dim	dbName
	dim	strSql
	dim	rsList
	dim	strFilename
	dim	i
	dim	strBuff
	dim	objFSO
	dim	objFile
	dim	objLog
	dim	strFind
	dim	strMsg
	dim	strUpdMsg
	dim	lngCnt			' ����󌏐�
	dim	lngQty			' ����
	dim	lngSai			' �ː�
	dim	lngWait			' �d��
	dim	lngQty100		' ���� 100�ȏ�̌���
	dim	strDt
	dim	strNinushi
	dim	strBukasyo
	dim	strIdNo
	dim	strHDt
	dim	strONo
	dim	strNoS
	dim	strNoE
	dim	strHKbn
	dim	strMKbn
	dim	strQty
	dim	strSai
	dim	strWait
	dim	strHoken
	dim	strAddress1
	dim	strAddress2
	dim	strName1
	dim	strName2
	dim	strTel
	dim	strKiji1
	dim	strKiji2
	dim	strKiji3
	dim	strKiji4
	dim	strKiji5
	dim	strYobi
	dim	strYSyukaH
	dim	strBin
	dim	strLabel
	dim	strWork
	dim	strErr

	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const adSearchForward = 1
	' ObjectStateEnum
	' �I�u�W�F�N�g���J���Ă��邩���Ă��邩�A�f�[�^ �\�[�X�ɐڑ������A
	' �R�}���h�����s�����A�܂��̓f�[�^���擾�����ǂ�����\���܂��B
	Const	adStateClosed		= 0 ' �I�u�W�F�N�g�����Ă��邱�Ƃ������܂��B 
	Const	adStateOpen			= 1 ' �I�u�W�F�N�g���J���Ă��邱�Ƃ������܂��B 
	Const	adStateConnecting	= 2 ' �I�u�W�F�N�g���ڑ����Ă��邱�Ƃ������܂��B 
	Const	adStateExecuting	= 4 ' �I�u�W�F�N�g���R�}���h�����s���ł��邱�Ƃ������܂��B 
	Const	adStateFetching		= 8 ' �I�u�W�F�N�g�̍s���擾����Ă��邱�Ƃ������܂��B 

	strYSyukaH	= "y_syuka_h"
	strDt		= ""
	strBin		= ""
	strLabel	= ""
	strWork		= ""
	strErr		= ""
	for i = 0 to WScript.Arguments.count - 1
	    select case lcase(WScript.Arguments(i))
	    case "-del"
			strYSyukaH = "del_syuka_h"
	    case "-b1"
			strBin = "01"
	    case "-b2"
			strBin = "02"
	    case "-b3"
			strBin = "03"
	    case "-label"
			strLabel = "label"
		case "-work"
			strWork		= "work"
		case "-err"
			strErr		= "err"
	    case "-?"
			usage()
			Wscript.Quit
	    case else
			if strDt = "" then
				strDt = WScript.Arguments(i)
			else
				usage()
				Wscript.Quit
			end if
	    end select
	next
	if strDt = "" then
		usage()
		Wscript.Quit
	end if
	Wscript.Echo "y_syuka_h.vbs " & strDt & " " & strYSyukaH & " " & strBin

	' �f�[�^�x�[�XOpen
	dbName = "newsdc"

	Set db = Wscript.CreateObject("ADODB.Connection")
	Wscript.Echo "open db : " & dbName
	db.open dbName

	' PN�e�[�u��Open
	Set rsList = Wscript.CreateObject("ADODB.Recordset")
	strSql = "select * from " & strYSyukaH
'	strSql = strSql & " where SYUKA_YMD like '" & strDt & "'"
	if strYSyukaH = "del_syuka_h" then
		strSql = strSql & " where length(rtrim(OKURI_NO)) = 11"
		strSql = strSql & " and SYUKA_YMD = '" & strDt & "'"
		strSql = strSql & " and left(KENPIN_NOW,8) <> '" & strDt & "'"
	else
		strSql = strSql & " where length(rtrim(OKURI_NO)) = 11"
	end if
'	strSql = strSql & " and SEQ_NO = '1'"
	strSql = strSql & " and CANCEL_F = ''"
	strSql = strSql & " and UNSOU_KAISHA = '���R�ʉ^'"
	strSql = strSql & " and OKURI_NO not like '310%'"
	if strErr = "" then
		strSql = strSql & " and (convert(SAI_SU,sql_numeric) > 0"
		strSql = strSql & "  or  convert(JURYO,sql_numeric) > 0)"
	else
		strSql = strSql & " and (convert(SAI_SU,sql_numeric) = 0"
		strSql = strSql & "  and convert(JURYO,sql_numeric) = 0)"
	end if
	if strBin <> "" then
		strSql = strSql & " and INS_BIN = '" & strBin & "'"
	end if
	if strWork = "" then
		strSql = strSql & " and OKURI_NO not in"
		strSql = strSql & "  (select distinct OKURI_NO from y_syuka_h"
		strSql = strSql &   " where SYUKA_YMD = '" & strDt & "'"
		strSql = strSql &   " and length(rtrim(OKURI_NO)) = 11"
		strSql = strSql &   " and CANCEL_F = ''"
		strSql = strSql &   " and UNSOU_KAISHA = '���R�ʉ^'"
		strSql = strSql &   " and (convert(SAI_SU,sql_numeric) = 0"
		strSql = strSql &   " and  convert(JURYO,sql_numeric) = 0)"
		strSql = strSql &   ")"
	end if
	strSql = strSql & " order by OKURI_NO,ID_NO"
	rsList.Open strSql, db, adOpenForwardOnly, adLockBatchOptimistic
	strONo 		= ""
	lngCnt 		= 0		' ����󌏐�
	lngQty		= 0		' ����
	lngSai		= 0		' �ː�
	lngWait	 	= 0		' �d��
	lngQty100	= 0		' ���� 100�ȏ�̌���
	do while ( rsList.Eof = False )
		if rtrim(strONo) <> rtrim(rsList.Fields("OKURI_NO")) then
			strNinushi		= "072874606S"
			strBukasyo		= "      "
'			strHDt			= Get_Buff(right(rsList.Fields("SYUKA_YMD"),6),6)
'			strHDt			= Get_Buff(right(strDt,6),6)
			strHDt			= Get_Buff(right(left(rsList.Fields("KENPIN_NOW"),8),6),6)
			strIdNo			= Get_Buff(strHDt & rsList.Fields("ID_NO"),20)
'			strHDt			= Get_Buff(right(rsList.Fields("SYUKA_YMD"),6),6)
			strONo			= Get_Buff(rsList.Fields("OKURI_NO"),11)
			strNoS			= Get_Buff(Left(RTrim(rsList.Fields("OKURI_NO")),11) & "01",13)
			strNoE			= Get_Buff(Left(RTrim(rsList.Fields("OKURI_NO")),11) & Right(RTrim(rsList.Fields("KUTI_SU")),2),13)
			strHKbn			= "1"
			strMKbn			= "1"
			strQty			= Get_Buff(Right(RTrim(rsList.Fields("KUTI_SU")),3),3)
'			strSai			= Get_Buff(Right("000"&round(cdbl("0"&RTrim(rsList.Fields("SAI_SU"))),0),3),3)	' "000"
			strSai			= Get_Buff(Right("000"&GetSaisu(rsList.Fields("SAI_SU")),3),3)			' "000"
			strWait			= Get_Buff(Right("0000"&round(cdbl("0"&RTrim(rsList.Fields("JURYO"))),0),4),4)	' "0000"
			if strWait <> "0000" then
				strSai		= "000"
			end if
			strHoken		= "0000"
			strAddress1		= Get_BuffZ(RTrim(rsList.Fields("JYUSHO")),80)
			strAddress2		= ""
'			strAddress1		= Get_BuffZ("�׎�l�Z���P",40)
'			strAddress2		= Get_BuffZ("�׎�l�Z���Q",40)
			strName1		= Get_BuffZ(rsList.Fields("OKURISAKI"),40)
			strName2		= ""
			if rsList.Fields("OKURISAKI") <> rsList.Fields("MUKE_NAME") then
				strName2 = rsList.Fields("MUKE_NAME")
			end if
			strName2		= Get_BuffZ(strName2,40)				' MUKE_NAME			Char(40)	
'			strTel			= Get_Buff("00-0000-0000",15)
			strTel			= Get_Buff(rsList.Fields("TEL_No"),15)
			strKiji1		= Get_BuffZ(rsList.Fields("BIKOU"),200)	' BIKOU				Char(100)
			strKiji2		= ""
			if rsList.Fields("SYUKA_YMD") <> strDt then
				strKiji1	= Get_BuffZ("�o�ד��ύX�F" & rsList.Fields("SYUKA_YMD"),40)
				strKiji2	= Get_BuffZ(rsList.Fields("BIKOU"),160)	' BIKOU				Char(100)
			end if
			strKiji3		= ""
			strKiji4		= ""
			strKiji5		= ""
'			strKiji2		= Get_BuffZ("�L�����Q",40)
'			strKiji3		= Get_BuffZ("�L�����R",40)
'			strKiji4		= Get_BuffZ("�L�����S",40)
'			strKiji5		= Get_BuffZ("�L�����T",40)
			strYobi			= Get_BuffZ("",40)
			if strLabel <> "" then
				Wscript.Echo  "JOB"
				WScript.Echo "DEF MK=1,DK=8,MD=1,PW=384,PH=344,XO=8,UM=8"
				WScript.Echo "START"
				WScript.Echo "FONT TP=3,CS=0"
				WScript.Echo "TEXT X=33,Y=0,L=1,NS=12,NE=2,NZ=0"
				WScript.Echo strONo & "01"
				WScript.Echo "TEXT X=275,Y=0,L=1,NS=1,NE=3,NZ=1"
				WScript.Echo "001/" & GetQty(strQty," ")
				WScript.Echo "BCD TP=6,X=0,Y=22,HT=40,HR=0,NS=12,NE=2,NZ=0"
				WScript.Echo strONo & "01"
				WScript.Echo "FONT TP=7,CS=0,LG=36,WD=18,LS=0"
				WScript.Echo "TEXT X=574,Y=65,L=1"
				WScript.Echo "���X:000"
				WScript.Echo "TEXT X=0,Y=65,L=7"
				WScript.Echo strTel
				WScript.Echo Get_LeftB(strAddress1,40)
				WScript.Echo Get_MidB(strAddress1,41,40)
				WScript.Echo strName1
				WScript.Echo strName2
				WScript.Echo "                          20" & Get_MidB(strHDt,1,2) & "�N" & Get_MidB(strHDt,3,2) & "��" & Get_MidB(strHDt,5,2) & "��"
				WScript.Echo "        (��)�G�X�f�B�[�V�B�[�@���ޗ��ʂb"
				WScript.Echo "QTY P=" & GetQty(strQty,"")
				WScript.Echo "END"
				WScript.Echo "JOBE"
			else
				strMsg = ""
				strMsg = strMsg & strNinushi
				strMsg = strMsg & strBukasyo	
				strMsg = strMsg & strIdNo		
				strMsg = strMsg & strHDt		
				strMsg = strMsg & strONo		
				strMsg = strMsg & strNoS		
				strMsg = strMsg & strNoE		
				strMsg = strMsg & strHKbn		
				strMsg = strMsg & strMKbn		
				strMsg = strMsg & strQty		
				strMsg = strMsg & strSai		
				strMsg = strMsg & strWait		
				strMsg = strMsg & strHoken		
				strMsg = strMsg & strAddress1	
				strMsg = strMsg & strAddress2	
				strMsg = strMsg & strName1		
				strMsg = strMsg & strName2		
				strMsg = strMsg & strTel		
				strMsg = strMsg & strKiji1		
				strMsg = strMsg & strKiji2		
				strMsg = strMsg & strKiji3		
				strMsg = strMsg & strKiji4		
				strMsg = strMsg & strKiji5		
				strMsg = strMsg & strYobi		
				Wscript.Echo strMsg
			end if
			lngCnt	= lngCnt  + 1			' ����󌏐�
			lngQty	= lngQty  + clng(strQty)		' ����
			lngSai	= lngSai  + clng(strSai)	' �ː�
			lngWait	= lngWait + clng(strWait)	' �d��
			if clng(strQty) >= 100 then
				lngQty100	= lngQty100  + 1		' ����(>100)
			end if
		end if
		rsList.movenext
	loop
	Wscript.Echo "�o�ד��F" & strDt
	Wscript.Echo "�����F" & right("        " & lngCnt ,6)
	Wscript.Echo "  �����F" & right("        " & lngQty ,6)
	Wscript.Echo "  �ː��F" & right("        " & lngSai ,6)
	Wscript.Echo "  �d�ʁF" & right("        " & lngWait,6)
	Wscript.Echo "������100�̌����F" & lngQty100 & " ��"

	' �e�[�u��Close
	Wscript.Echo "close table : " & strYSyukaH
	rsList.Close

	' DBClose
	Wscript.Echo "close db : " & dbName
	db.Close
	set db = nothing
End Sub

Function GetTm(t)
	GetTm = year(t) & right("0" & month(t),2) & right("0" & day(t),2) & right("0" & hour(t),2)& right("0" & minute(t),2)
End Function

Function Get_Buff(a_Str,a_int)
	dim	strRet

	strRet = a_Str & space(a_int)
	strRet = Get_LeftB(strRet,a_int)
	Get_Buff = strRet
End Function

Function Get_BuffZ(a_Str,a_int)
	dim	strRet

	strRet = StrConvWide(rtrim(a_Str)) & string(a_int,"�@")
	strRet = Get_LeftB(strRet,a_int)
	Get_BuffZ = strRet
End Function

Function Get_LeftB(a_Str, a_int)
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

Function Get_MidB(a_Str,s_int, a_int)
	Dim iCount, iAscCode, iLenCount, iMidStr
	iLenCount = 0
	iMidStr = ""
	If Len(a_Str) = 0 Then
		Get_MidB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_MidB = ""
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
		if iLenCount >= s_int then
			If iLenCount > Cint(s_int) + Cint(a_int) - 1 Then
				Exit For
			Else
				iMidStr = iMidStr + Mid(a_Str, iCount, 1)
			End If
		end if
	Next
	Get_MidB = iMidStr
End Function

Function Get_LenB(a_Str)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LenB = 0
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
	Next
	Get_LenB = iLenCount
End Function


Function SetField(rsPn,strFieldName,strValue,strTitle,strUpdMsg)
	if rtrim(rsPn.Fields(strFieldName)) <> rtrim(strValue) then
		if strUpdMsg <> "-" then
			strUpdMsg = strUpdMsg & rsPn.Fields(strFieldName) & " ��" & strTitle & vbNewLine
			strUpdMsg = strUpdMsg & strValue & " ���ύX" & vbNewLine
		end if
		rsPn.Fields(strFieldName) = strValue
	end if
	SetField = strUpdMsg
End Function


'**********************************************************************************
' Script�֐����@  : DBCS_Convert(�ϊ����镶����) Ver1.0
' Script�֐����A  : SBCS_Convert(�ϊ����镶����) Ver1.0
' Script�֐����B  : SBCS_DBCS_Check(�`�F�b�N����P����) Ver1.0
' �@�\�T�v  : �@�����񒆂̔��p������S�p�ɕϊ����܂�
'           : �A�����񒆂̑S�p�����𔼊p�ɕϊ����܂�
'           : �B������S�p�����p�����肵�܂�
' Made By   : Copyright(C) 2008 T.Tokunaga All right reserved
'           : ���̃v���O�����͓��{�����쌠�@����э��ۏ��ɂ��ی삳��Ă��܂��B
'           : ���̃v���O������]�ڂ���ꍇ�͒��쌠���L�҂̋����K�v�ƂȂ�܂��
'**********************************************************************************

'���_�E�����_�����i���p�j
'�޷޸޹޺޻޼޽޾޿������������������������������߳�
Public CNDakutenSBCS
CNDakutenSBCS = "�޷޸޹޺޻޼޽޾޿������������������������������߳�"

'���_�E�����_�����i�S�p�j
'�K�M�O�Q�S�U�W�Y�[�]�_�a�d�f�h�o�p�r�s�u�v�x�y�{�|��
Public CNDakutenDBCS
CNDakutenDBCS = "�K�M�O�Q�S�U�W�Y�[�]�_�a�d�f�h�o�p�r�s�u�v�x�y�{�|��"

'���p����
' !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~���������������������������������������������������������������
Public CNConvSBCS
CNConvSBCS = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~���������������������������������������������������������������"

'�S�p����
'�@�I�h���������f�i�j���{�C�|�D�^�O�P�Q�R�S�T�U�V�W�X�F�G�������H���`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w�x�y�m���n�O�Q�e�����������������������������������������������������o�b�p�`�B�u�v�A�E���@�B�D�F�H�������b�[�A�C�E�G�I�J�L�N�P�R�T�V�X�Z�\�^�`�c�e�g�i�j�k�l�m�n�q�t�w�z�}�~���������������������������J�K
Public CNConvDBCS
CNConvDBCS = "�@�I�h���������f�i�j���{�C�|�D�^�O�P�Q�R�S�T�U�V�W�X�F�G�������H���`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w�x�y�m���n�O�Q�e�����������������������������������������������������o�b�p�`�B�u�v�A�E���@�B�D�F�H�������b�[�A�C�E�G�I�J�L�N�P�R�T�V�X�Z�\�^�`�c�e�g�i�j�k�l�m�n�q�t�w�z�}�~���������������������������J�K"

'*************************************
'�֐��Ăяo����
'
'Check_after = DBCS_Convert(CNDakutenSBCS)
'
'MsgBox Check_after
'
'Check_after = DBCS_Convert(CNConvSBCS)
'
'MsgBox Check_after
'
'Check_after = SBCS_Convert(CNDakutenDBCS)
'
'MsgBox Check_after
'
'Check_after = SBCS_Convert(CNConvDBCS)
'
'MsgBox Check_after

'*************************************

Private Function DBCS_Convert(Check_String)
dim	DBCS_Convert_Temp_Data
dim	DBCS_Convert_j
dim	DBCS_Convert_jj
dim	DBCS_Convert_AscChk_Data
dim	DBCS_Convert_Sarch_SBCS
dim	DBCS_Convert_Conv_Data

  If  Len(Check_String) > 0 Then
    DBCS_Convert_Temp_Data = ""
    For DBCS_Convert_j = 1 To Len(Check_String)
      If SBCS_DBCS_Check(Mid(Check_String,DBCS_Convert_j,1)) = 1 Then
        DBCS_Convert_Temp_Data = DBCS_Convert_Temp_Data & Mid(Check_String,DBCS_Convert_j,1)
      Else
        DBCS_Convert_jj = DBCS_Convert_j + 1
        If DBCS_Convert_jj <= Len(Check_String) Then
          DBCS_Convert_AscChk_Data = Mid(Check_String,DBCS_Convert_jj,1)
          If Mid(Check_String,DBCS_Convert_jj,1) <> "�" And Mid(Check_String,DBCS_Convert_jj,1) <> "�" Then
            DBCS_Convert_AscChk_Data = ""
          Else 
            DBCS_Convert_AscChk_Data = Mid(Check_String,DBCS_Convert_j,1)&Mid(Check_String,DBCS_Convert_jj,1)
          End If
        Else
          DBCS_Convert_AscChk_Data = ""
        End If
        If DBCS_Convert_AscChk_Data = "" Then
          DBCS_Convert_Sarch_SBCS = InStr(1,CNConvSBCS,Mid(Check_String,DBCS_Convert_j,1),vbBinaryCompare)
          If DBCS_Convert_Sarch_SBCS = "" Or  DBCS_Convert_Sarch_SBCS = 0 Then
            DBCS_Convert_Conv_Data = "�@"	' Mid(Check_String,DBCS_Convert_j,1)
          Else
            DBCS_Convert_Conv_Data = Mid(CNConvDBCS,DBCS_Convert_Sarch_SBCS,1)
          End If
        Else
          DBCS_Convert_Sarch_SBCS = InStr(1,CNDakutenSBCS,DBCS_Convert_AscChk_Data,vbBinaryCompare)
          If DBCS_Convert_Sarch_SBCS = "" Or  DBCS_Convert_Sarch_SBCS = 0 Then
            DBCS_Convert_Sarch_SBCS = InStr(1,CNConvSBCS,Mid(Check_String,DBCS_Convert_j,1),vbBinaryCompare)
            If DBCS_Convert_Sarch_SBCS = "" Or  DBCS_Convert_Sarch_SBCS = 0 Then
              DBCS_Convert_Conv_Data = "�A"
            Else
              DBCS_Convert_Conv_Data = Mid(CNConvDBCS,DBCS_Convert_Sarch_SBCS,1)
            End If
          Else
            DBCS_Convert_Conv_Data = Mid(CNDakutenDBCS,(DBCS_Convert_Sarch_SBCS + 1 ) / 2, 1)
            DBCS_Convert_j = DBCS_Convert_j + 1
          End If
        End If
        DBCS_Convert_Temp_Data = DBCS_Convert_Temp_Data & DBCS_Convert_Conv_Data
      End If
    Next
    DBCS_Convert = DBCS_Convert_Temp_Data
  Else
    DBCS_Convert = Check_String
  End If

End Function

Private Function SBCS_Convert(Check_String)
	dim	SBCS_Convert_Temp_Data
	dim	SBCS_Convert_j
	dim	SBCS_Convert_Sarch_DBCS
	dim	SBCS_Convert_Conv_Data

  If  Len(Check_String) > 0 Then
    SBCS_Convert_Temp_Data = ""
    For SBCS_Convert_j = 1 To Len(Check_String)
      If SBCS_DBCS_Check(Mid(Check_String,SBCS_Convert_j,1)) = 0 Then
        SBCS_Convert_Temp_Data = SBCS_Convert_Temp_Data & Mid(Check_String,SBCS_Convert_j,1)
      Else
        SBCS_Convert_Sarch_DBCS = InStr(1,CNDakutenDBCS,Mid(Check_String,SBCS_Convert_j,1),vbBinaryCompare)
        If SBCS_Convert_Sarch_DBCS = "" Or  SBCS_Convert_Sarch_DBCS = 0 Then
          SBCS_Convert_Sarch_DBCS = InStr(1,CNConvDBCS,Mid(Check_String,SBCS_Convert_j,1),vbBinaryCompare)
          If SBCS_Convert_Sarch_DBCS = "" Or  SBCS_Convert_Sarch_DBCS = 0 Then
            SBCS_Convert_Conv_Data = "?"
          Else
            SBCS_Convert_Conv_Data = Mid(CNConvSBCS,SBCS_Convert_Sarch_DBCS,1)
          End If
        Else
          SBCS_Convert_Conv_Data = Mid(CNDakutenSBCS,(SBCS_Convert_Sarch_DBCS * 2) - 1, 2)
        End If
        SBCS_Convert_Temp_Data = SBCS_Convert_Temp_Data & SBCS_Convert_Conv_Data
      End If
    Next
    SBCS_Convert = SBCS_Convert_Temp_Data
  Else
    SBCS_Convert = Check_String
  End If
End Function

Private Function SBCS_DBCS_Check(Check_Word)
  If ASC(Check_Word) > 255 Or ASC(Check_Word) < 0 Then
     'DBCS(�S�p)�ƔF��
    SBCS_DBCS_Check = 1
  Else
     'SBCS(���p)�ƔF��
    SBCS_DBCS_Check = 0
  End If
End Function

'#################################################################
'# StrConv Clone For VBScript
'#  author: Yasuhiro Matsumoto
'#  url: http://www.ac.cyberhome.ne.jp/~mattn/cgi-bin/blosxom.cgi
'#  mailto: mattn.jp@gmai.com
'#  see: http://www.ac.cyberhome.ne.jp/~mattn/AcrobatASP/1.html
'#################################################################

'***************************************************
' StrConvUpperCase
'---------------------------------------------------
' �p�r : StrConv(sInp,vbUpperCase) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvUpperCase(sInp)
	StrConvUpperCase = UCase(sInp)
End Function

'***************************************************
' StrConvLowerCase
'---------------------------------------------------
' �p�r : StrConv(sInp,vbLowerCase) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvLowerCase(sInp)
	StrConvLowerCase = LCase(sInp)
End Function

'***************************************************
' StrConvProperCase
'---------------------------------------------------
' �p�r : StrConv(sInp,vbProperCase) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvProperCase(sInp)
	Dim nPos
	Dim nSpc

	nPos = 1
	Do While InStr(nPos, sInp, " ", 1) <> 0
		nSpc = InStr(nPos, sInp, " ", 1)
		StrConvProperCase = StrConvProperCase & UCase(Mid(sInp, nPos, 1))
		StrConvProperCase = StrConvProperCase & LCase(Mid(sInp, nPos + 1, nSpc - nPos))
		nPos = nSpc + 1
	Loop

	StrConvProperCase = StrConvProperCase & UCase(Mid(sInp, nPos, 1))
	StrConvProperCase = StrConvProperCase & LCase(Mid(sInp, nPos + 1))
	StrConvProperCase = StrConvProperCase
End Function

'***************************************************
' StrConvWide
'---------------------------------------------------
' �p�r : StrConv(s,vbWide) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvWide(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("��", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case " "
			StrConvWide = StrConvWide & "�@"
		Case "!"
			StrConvWide = StrConvWide & "�I"
		Case """"
			StrConvWide = StrConvWide & "�W"
		Case "#"
			StrConvWide = StrConvWide & "��"
		Case "$"
			StrConvWide = StrConvWide & "��"
		Case "%"
			StrConvWide = StrConvWide & "��"
		Case "&"
			StrConvWide = StrConvWide & "��"
		Case "'"
			StrConvWide = StrConvWide & "�V"
		Case "("
			StrConvWide = StrConvWide & "�i"
		Case ")"
			StrConvWide = StrConvWide & "�j"
		Case "*"
			StrConvWide = StrConvWide & "��"
		Case "+"
			StrConvWide = StrConvWide & "�{"
		Case ","
			StrConvWide = StrConvWide & "�C"
		Case "-"
			StrConvWide = StrConvWide & "�|"
		Case "."
			StrConvWide = StrConvWide & "�D"
		Case "/"
			StrConvWide = StrConvWide & "�^"
		Case "0"
			StrConvWide = StrConvWide & "�O"
		Case "1"
			StrConvWide = StrConvWide & "�P"
		Case "2"
			StrConvWide = StrConvWide & "�Q"
		Case "3"
			StrConvWide = StrConvWide & "�R"
		Case "4"
			StrConvWide = StrConvWide & "�S"
		Case "5"
			StrConvWide = StrConvWide & "�T"
		Case "6"
			StrConvWide = StrConvWide & "�U"
		Case "7"
			StrConvWide = StrConvWide & "�V"
		Case "8"
			StrConvWide = StrConvWide & "�W"
		Case "9"
			StrConvWide = StrConvWide & "�X"
		Case ":"
			StrConvWide = StrConvWide & "�F"
		Case ";"
			StrConvWide = StrConvWide & "�G"
		Case "<"
			StrConvWide = StrConvWide & "��"
		Case "="
			StrConvWide = StrConvWide & "��"
		Case ">"
			StrConvWide = StrConvWide & "��"
		Case "?"
			StrConvWide = StrConvWide & "�H"
		Case "@"
			StrConvWide = StrConvWide & "��"
		Case "A"
			StrConvWide = StrConvWide & "�`"
		Case "B"
			StrConvWide = StrConvWide & "�a"
		Case "C"
			StrConvWide = StrConvWide & "�b"
		Case "D"
			StrConvWide = StrConvWide & "�c"
		Case "E"
			StrConvWide = StrConvWide & "�d"
		Case "F"
			StrConvWide = StrConvWide & "�e"
		Case "G"
			StrConvWide = StrConvWide & "�f"
		Case "H"
			StrConvWide = StrConvWide & "�g"
		Case "I"
			StrConvWide = StrConvWide & "�h"
		Case "J"
			StrConvWide = StrConvWide & "�i"
		Case "K"
			StrConvWide = StrConvWide & "�j"
		Case "L"
			StrConvWide = StrConvWide & "�k"
		Case "M"
			StrConvWide = StrConvWide & "�l"
		Case "N"
			StrConvWide = StrConvWide & "�m"
		Case "O"
			StrConvWide = StrConvWide & "�n"
		Case "P"
			StrConvWide = StrConvWide & "�o"
		Case "Q"
			StrConvWide = StrConvWide & "�p"
		Case "R"
			StrConvWide = StrConvWide & "�q"
		Case "S"
			StrConvWide = StrConvWide & "�r"
		Case "T"
			StrConvWide = StrConvWide & "�s"
		Case "U"
			StrConvWide = StrConvWide & "�t"
		Case "V"
			StrConvWide = StrConvWide & "�u"
		Case "W"
			StrConvWide = StrConvWide & "�v"
		Case "X"
			StrConvWide = StrConvWide & "�w"
		Case "Y"
			StrConvWide = StrConvWide & "�x"
		Case "Z"
			StrConvWide = StrConvWide & "�y"
		Case "["
			StrConvWide = StrConvWide & "�m"
		Case "]"
			StrConvWide = StrConvWide & "�n"
		Case "^"
			StrConvWide = StrConvWide & "�O"
		Case "_"
			StrConvWide = StrConvWide & "�Q"
		Case "`"
			StrConvWide = StrConvWide & "�M"
		Case "a"
			StrConvWide = StrConvWide & "��"
		Case "b"
			StrConvWide = StrConvWide & "��"
		Case "c"
			StrConvWide = StrConvWide & "��"
		Case "d"
			StrConvWide = StrConvWide & "��"
		Case "e"
			StrConvWide = StrConvWide & "��"
		Case "f"
			StrConvWide = StrConvWide & "��"
		Case "g"
			StrConvWide = StrConvWide & "��"
		Case "h"
			StrConvWide = StrConvWide & "��"
		Case "i"
			StrConvWide = StrConvWide & "��"
		Case "j"
			StrConvWide = StrConvWide & "��"
		Case "k"
			StrConvWide = StrConvWide & "��"
		Case "l"
			StrConvWide = StrConvWide & "��"
		Case "m"
			StrConvWide = StrConvWide & "��"
		Case "n"
			StrConvWide = StrConvWide & "��"
		Case "o"
			StrConvWide = StrConvWide & "��"
		Case "p"
			StrConvWide = StrConvWide & "��"
		Case "q"
			StrConvWide = StrConvWide & "��"
		Case "r"
			StrConvWide = StrConvWide & "��"
		Case "s"
			StrConvWide = StrConvWide & "��"
		Case "t"
			StrConvWide = StrConvWide & "��"
		Case "u"
			StrConvWide = StrConvWide & "��"
		Case "v"
			StrConvWide = StrConvWide & "��"
		Case "w"
			StrConvWide = StrConvWide & "��"
		Case "x"
			StrConvWide = StrConvWide & "��"
		Case "y"
			StrConvWide = StrConvWide & "��"
		Case "z"
			StrConvWide = StrConvWide & "��"
		Case "{"
			StrConvWide = StrConvWide & "�o"
		Case "|"
			StrConvWide = StrConvWide & "�b"
		Case "}"
			StrConvWide = StrConvWide & "�p"
		Case "~"
			StrConvWide = StrConvWide & "�`"
		Case "�"
			StrConvWide = StrConvWide & "�B"
		Case "�"
			StrConvWide = StrConvWide & "�u"
		Case "�"
			StrConvWide = StrConvWide & "�v"
		Case "�"
			StrConvWide = StrConvWide & "�A"
		Case "�"
			StrConvWide = StrConvWide & "�E"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�@"
		Case "�"
			StrConvWide = StrConvWide & "�B"
		Case "�"
			StrConvWide = StrConvWide & "�D"
		Case "�"
			StrConvWide = StrConvWide & "�F"
		Case "�"
			StrConvWide = StrConvWide & "�H"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�b"
		Case "�"
			StrConvWide = StrConvWide & "�["
		Case "�"
			StrConvWide = StrConvWide & "�A"
		Case "�"
			StrConvWide = StrConvWide & "�C"
		Case "�"
			StrConvWide = StrConvWide & "�E"
		Case "��"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�G"
		Case "�"
			StrConvWide = StrConvWide & "�I"
		Case "�"
			StrConvWide = StrConvWide & "�J"
		Case "��"
			StrConvWide = StrConvWide & "�K"
		Case "�"
			StrConvWide = StrConvWide & "�L"
		Case "��"
			StrConvWide = StrConvWide & "�M"
		Case "�"
			StrConvWide = StrConvWide & "�N"
		Case "��"
			StrConvWide = StrConvWide & "�O"
		Case "�"
			StrConvWide = StrConvWide & "�P"
		Case "��"
			StrConvWide = StrConvWide & "�Q"
		Case "�"
			StrConvWide = StrConvWide & "�R"
		Case "��"
			StrConvWide = StrConvWide & "�S"
		Case "�"
			StrConvWide = StrConvWide & "�T"
		Case "��"
			StrConvWide = StrConvWide & "�U"
		Case "�"
			StrConvWide = StrConvWide & "�V"
		Case "��"
			StrConvWide = StrConvWide & "�W"
		Case "�"
			StrConvWide = StrConvWide & "�X"
		Case "��"
			StrConvWide = StrConvWide & "�Y"
		Case "�"
			StrConvWide = StrConvWide & "�Z"
		Case "��"
			StrConvWide = StrConvWide & "�["
		Case "�"
			StrConvWide = StrConvWide & "�\"
		Case "��"
			StrConvWide = StrConvWide & "�]"
		Case "�"
			StrConvWide = StrConvWide & "�^"
		Case "��"
			StrConvWide = StrConvWide & "�_"
		Case "�"
			StrConvWide = StrConvWide & "�`"
		Case "��"
			StrConvWide = StrConvWide & "�a"
		Case "�"
			StrConvWide = StrConvWide & "�c"
		Case "��"
			StrConvWide = StrConvWide & "�d"
		Case "�"
			StrConvWide = StrConvWide & "�e"
		Case "��"
			StrConvWide = StrConvWide & "�f"
		Case "�"
			StrConvWide = StrConvWide & "�g"
		Case "��"
			StrConvWide = StrConvWide & "�h"
		Case "�"
			StrConvWide = StrConvWide & "�i"
		Case "�"
			StrConvWide = StrConvWide & "�j"
		Case "�"
			StrConvWide = StrConvWide & "�k"
		Case "�"
			StrConvWide = StrConvWide & "�l"
		Case "�"
			StrConvWide = StrConvWide & "�m"
		Case "�"
			StrConvWide = StrConvWide & "�n"
		Case "��"
			StrConvWide = StrConvWide & "�o"
		Case "��"
			StrConvWide = StrConvWide & "�p"
		Case "�"
			StrConvWide = StrConvWide & "�q"
		Case "��"
			StrConvWide = StrConvWide & "�r"
		Case "��"
			StrConvWide = StrConvWide & "�s"
		Case "�"
			StrConvWide = StrConvWide & "�t"
		Case "��"
			StrConvWide = StrConvWide & "�u"
		Case "��"
			StrConvWide = StrConvWide & "�v"
		Case "�"
			StrConvWide = StrConvWide & "�w"
		Case "��"
			StrConvWide = StrConvWide & "�x"
		Case "��"
			StrConvWide = StrConvWide & "�y"
		Case "�"
			StrConvWide = StrConvWide & "�z"
		Case "��"
			StrConvWide = StrConvWide & "�{"
		Case "��"
			StrConvWide = StrConvWide & "�|"
		Case "�"
			StrConvWide = StrConvWide & "�}"
		Case "�"
			StrConvWide = StrConvWide & "�~"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�J"
		Case "�"
			StrConvWide = StrConvWide & "�K"
		Case " "
			StrConvWide = StrConvWide & "�@"
		Case "!"
			StrConvWide = StrConvWide & "�I"
		Case """"
			StrConvWide = StrConvWide & "�W"
		Case "#"
			StrConvWide = StrConvWide & "��"
		Case "$"
			StrConvWide = StrConvWide & "��"
		Case "%"
			StrConvWide = StrConvWide & "��"
		Case "&"
			StrConvWide = StrConvWide & "��"
		Case "'"
			StrConvWide = StrConvWide & "�V"
		Case "("
			StrConvWide = StrConvWide & "�i"
		Case ")"
			StrConvWide = StrConvWide & "�j"
		Case "*"
			StrConvWide = StrConvWide & "��"
		Case "+"
			StrConvWide = StrConvWide & "�{"
		Case ","
			StrConvWide = StrConvWide & "�C"
		Case "-"
			StrConvWide = StrConvWide & "�|"
		Case "."
			StrConvWide = StrConvWide & "�D"
		Case "/"
			StrConvWide = StrConvWide & "�^"
		Case "0"
			StrConvWide = StrConvWide & "�O"
		Case "1"
			StrConvWide = StrConvWide & "�P"
		Case "2"
			StrConvWide = StrConvWide & "�Q"
		Case "3"
			StrConvWide = StrConvWide & "�R"
		Case "4"
			StrConvWide = StrConvWide & "�S"
		Case "5"
			StrConvWide = StrConvWide & "�T"
		Case "6"
			StrConvWide = StrConvWide & "�U"
		Case "7"
			StrConvWide = StrConvWide & "�V"
		Case "8"
			StrConvWide = StrConvWide & "�W"
		Case "9"
			StrConvWide = StrConvWide & "�X"
		Case ":"
			StrConvWide = StrConvWide & "�F"
		Case ";"
			StrConvWide = StrConvWide & "�G"
		Case "<"
			StrConvWide = StrConvWide & "��"
		Case "="
			StrConvWide = StrConvWide & "��"
		Case ">"
			StrConvWide = StrConvWide & "��"
		Case "?"
			StrConvWide = StrConvWide & "�H"
		Case "@"
			StrConvWide = StrConvWide & "��"
		Case "A"
			StrConvWide = StrConvWide & "�`"
		Case "B"
			StrConvWide = StrConvWide & "�a"
		Case "C"
			StrConvWide = StrConvWide & "�b"
		Case "D"
			StrConvWide = StrConvWide & "�c"
		Case "E"
			StrConvWide = StrConvWide & "�d"
		Case "F"
			StrConvWide = StrConvWide & "�e"
		Case "G"
			StrConvWide = StrConvWide & "�f"
		Case "H"
			StrConvWide = StrConvWide & "�g"
		Case "I"
			StrConvWide = StrConvWide & "�h"
		Case "J"
			StrConvWide = StrConvWide & "�i"
		Case "K"
			StrConvWide = StrConvWide & "�j"
		Case "L"
			StrConvWide = StrConvWide & "�k"
		Case "M"
			StrConvWide = StrConvWide & "�l"
		Case "N"
			StrConvWide = StrConvWide & "�m"
		Case "O"
			StrConvWide = StrConvWide & "�n"
		Case "P"
			StrConvWide = StrConvWide & "�o"
		Case "Q"
			StrConvWide = StrConvWide & "�p"
		Case "R"
			StrConvWide = StrConvWide & "�q"
		Case "S"
			StrConvWide = StrConvWide & "�r"
		Case "T"
			StrConvWide = StrConvWide & "�s"
		Case "U"
			StrConvWide = StrConvWide & "�t"
		Case "V"
			StrConvWide = StrConvWide & "�u"
		Case "W"
			StrConvWide = StrConvWide & "�v"
		Case "X"
			StrConvWide = StrConvWide & "�w"
		Case "Y"
			StrConvWide = StrConvWide & "�x"
		Case "Z"
			StrConvWide = StrConvWide & "�y"
		Case "["
			StrConvWide = StrConvWide & "�m"
		Case "]"
			StrConvWide = StrConvWide & "�n"
		Case "^"
			StrConvWide = StrConvWide & "�O"
		Case "_"
			StrConvWide = StrConvWide & "�Q"
		Case "`"
			StrConvWide = StrConvWide & "�M"
		Case "a"
			StrConvWide = StrConvWide & "��"
		Case "b"
			StrConvWide = StrConvWide & "��"
		Case "c"
			StrConvWide = StrConvWide & "��"
		Case "d"
			StrConvWide = StrConvWide & "��"
		Case "e"
			StrConvWide = StrConvWide & "��"
		Case "f"
			StrConvWide = StrConvWide & "��"
		Case "g"
			StrConvWide = StrConvWide & "��"
		Case "h"
			StrConvWide = StrConvWide & "��"
		Case "i"
			StrConvWide = StrConvWide & "��"
		Case "j"
			StrConvWide = StrConvWide & "��"
		Case "k"
			StrConvWide = StrConvWide & "��"
		Case "l"
			StrConvWide = StrConvWide & "��"
		Case "m"
			StrConvWide = StrConvWide & "��"
		Case "n"
			StrConvWide = StrConvWide & "��"
		Case "o"
			StrConvWide = StrConvWide & "��"
		Case "p"
			StrConvWide = StrConvWide & "��"
		Case "q"
			StrConvWide = StrConvWide & "��"
		Case "r"
			StrConvWide = StrConvWide & "��"
		Case "s"
			StrConvWide = StrConvWide & "��"
		Case "t"
			StrConvWide = StrConvWide & "��"
		Case "u"
			StrConvWide = StrConvWide & "��"
		Case "v"
			StrConvWide = StrConvWide & "��"
		Case "w"
			StrConvWide = StrConvWide & "��"
		Case "x"
			StrConvWide = StrConvWide & "��"
		Case "y"
			StrConvWide = StrConvWide & "��"
		Case "z"
			StrConvWide = StrConvWide & "��"
		Case "{"
			StrConvWide = StrConvWide & "�o"
		Case "|"
			StrConvWide = StrConvWide & "�b"
		Case "}"
			StrConvWide = StrConvWide & "�p"
		Case "~"
			StrConvWide = StrConvWide & "�`"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "���"
			StrConvWide = StrConvWide & "��"
		Case "�E�"
			StrConvWide = StrConvWide & "��"
		Case "�J�"
			StrConvWide = StrConvWide & "�K"
		Case "�L�"
			StrConvWide = StrConvWide & "�M"
		Case "�N�"
			StrConvWide = StrConvWide & "�O"
		Case "�P�"
			StrConvWide = StrConvWide & "�Q"
		Case "�R�"
			StrConvWide = StrConvWide & "�S"
		Case "�T�"
			StrConvWide = StrConvWide & "�U"
		Case "�V�"
			StrConvWide = StrConvWide & "�W"
		Case "�X�"
			StrConvWide = StrConvWide & "�Y"
		Case "�Z�"
			StrConvWide = StrConvWide & "�["
		Case "�\�"
			StrConvWide = StrConvWide & "�]"
		Case "�^�"
			StrConvWide = StrConvWide & "�_"
		Case "�`�"
			StrConvWide = StrConvWide & "�a"
		Case "�c�"
			StrConvWide = StrConvWide & "�d"
		Case "�e�"
			StrConvWide = StrConvWide & "�f"
		Case "�g�"
			StrConvWide = StrConvWide & "�h"
		Case "�n�"
			StrConvWide = StrConvWide & "�o"
		Case "�n�"
			StrConvWide = StrConvWide & "�p"
		Case "�q�"
			StrConvWide = StrConvWide & "�r"
		Case "�q�"
			StrConvWide = StrConvWide & "�s"
		Case "�t�"
			StrConvWide = StrConvWide & "�u"
		Case "�t�"
			StrConvWide = StrConvWide & "�v"
		Case "�w�"
			StrConvWide = StrConvWide & "�x"
		Case "�w�"
			StrConvWide = StrConvWide & "�y"
		Case "�z�"
			StrConvWide = StrConvWide & "�{"
		Case "�z�"
			StrConvWide = StrConvWide & "�|"
		Case "�"
			StrConvWide = StrConvWide & "�B"
		Case "�"
			StrConvWide = StrConvWide & "�u"
		Case "�"
			StrConvWide = StrConvWide & "�v"
		Case "�"
			StrConvWide = StrConvWide & "�A"
		Case "�"
			StrConvWide = StrConvWide & "�E"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�@"
		Case "�"
			StrConvWide = StrConvWide & "�B"
		Case "�"
			StrConvWide = StrConvWide & "�D"
		Case "�"
			StrConvWide = StrConvWide & "�F"
		Case "�"
			StrConvWide = StrConvWide & "�H"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�b"
		Case "�"
			StrConvWide = StrConvWide & "�["
		Case "�"
			StrConvWide = StrConvWide & "�A"
		Case "�"
			StrConvWide = StrConvWide & "�C"
		Case "�"
			StrConvWide = StrConvWide & "�E"
		Case "��"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�G"
		Case "�"
			StrConvWide = StrConvWide & "�I"
		Case "�"
			StrConvWide = StrConvWide & "�J"
		Case "��"
			StrConvWide = StrConvWide & "�K"
		Case "�"
			StrConvWide = StrConvWide & "�L"
		Case "��"
			StrConvWide = StrConvWide & "�M"
		Case "�"
			StrConvWide = StrConvWide & "�N"
		Case "��"
			StrConvWide = StrConvWide & "�O"
		Case "�"
			StrConvWide = StrConvWide & "�P"
		Case "��"
			StrConvWide = StrConvWide & "�Q"
		Case "�"
			StrConvWide = StrConvWide & "�R"
		Case "��"
			StrConvWide = StrConvWide & "�S"
		Case "�"
			StrConvWide = StrConvWide & "�T"
		Case "��"
			StrConvWide = StrConvWide & "�U"
		Case "�"
			StrConvWide = StrConvWide & "�V"
		Case "��"
			StrConvWide = StrConvWide & "�W"
		Case "�"
			StrConvWide = StrConvWide & "�X"
		Case "��"
			StrConvWide = StrConvWide & "�Y"
		Case "�"
			StrConvWide = StrConvWide & "�Z"
		Case "��"
			StrConvWide = StrConvWide & "�["
		Case "�"
			StrConvWide = StrConvWide & "�\"
		Case "��"
			StrConvWide = StrConvWide & "�]"
		Case "�"
			StrConvWide = StrConvWide & "�^"
		Case "��"
			StrConvWide = StrConvWide & "�_"
		Case "�"
			StrConvWide = StrConvWide & "�`"
		Case "��"
			StrConvWide = StrConvWide & "�a"
		Case "�"
			StrConvWide = StrConvWide & "�c"
		Case "��"
			StrConvWide = StrConvWide & "�d"
		Case "�"
			StrConvWide = StrConvWide & "�e"
		Case "��"
			StrConvWide = StrConvWide & "�f"
		Case "�"
			StrConvWide = StrConvWide & "�g"
		Case "��"
			StrConvWide = StrConvWide & "�h"
		Case "�"
			StrConvWide = StrConvWide & "�i"
		Case "�"
			StrConvWide = StrConvWide & "�j"
		Case "�"
			StrConvWide = StrConvWide & "�k"
		Case "�"
			StrConvWide = StrConvWide & "�l"
		Case "�"
			StrConvWide = StrConvWide & "�m"
		Case "�"
			StrConvWide = StrConvWide & "�n"
		Case "��"
			StrConvWide = StrConvWide & "�o"
		Case "��"
			StrConvWide = StrConvWide & "�p"
		Case "�"
			StrConvWide = StrConvWide & "�q"
		Case "��"
			StrConvWide = StrConvWide & "�r"
		Case "��"
			StrConvWide = StrConvWide & "�s"
		Case "�"
			StrConvWide = StrConvWide & "�t"
		Case "��"
			StrConvWide = StrConvWide & "�u"
		Case "��"
			StrConvWide = StrConvWide & "�v"
		Case "�"
			StrConvWide = StrConvWide & "�w"
		Case "��"
			StrConvWide = StrConvWide & "�x"
		Case "��"
			StrConvWide = StrConvWide & "�y"
		Case "�"
			StrConvWide = StrConvWide & "�z"
		Case "��"
			StrConvWide = StrConvWide & "�{"
		Case "��"
			StrConvWide = StrConvWide & "�|"
		Case "�"
			StrConvWide = StrConvWide & "�}"
		Case "�"
			StrConvWide = StrConvWide & "�~"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "��"
		Case "�"
			StrConvWide = StrConvWide & "�J"
		Case "�"
			StrConvWide = StrConvWide & "�K"
		Case Else
			StrConvWide = StrConvWide & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvNarrow
'---------------------------------------------------
' �p�r : StrConv(s,vbNarrow) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvNarrow(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("��", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�@"
			StrConvNarrow = StrConvNarrow & " "
		Case "�A"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�B"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�C"
			StrConvNarrow = StrConvNarrow & ","
		Case "�D"
			StrConvNarrow = StrConvNarrow & "."
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�F"
			StrConvNarrow = StrConvNarrow & ":"
		Case "�G"
			StrConvNarrow = StrConvNarrow & ";"
		Case "�H"
			StrConvNarrow = StrConvNarrow & "?"
		Case "�I"
			StrConvNarrow = StrConvNarrow & "!"
		Case "�J"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�K"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�M"
			StrConvNarrow = StrConvNarrow & "`"
		Case "�O"
			StrConvNarrow = StrConvNarrow & "^"
		Case "�Q"
			StrConvNarrow = StrConvNarrow & "_"
		Case "�["
			StrConvNarrow = StrConvNarrow & "�"
		Case "�^"
			StrConvNarrow = StrConvNarrow & "/"
		Case "�`"
			StrConvNarrow = StrConvNarrow & "~"
		Case "�b"
			StrConvNarrow = StrConvNarrow & "|"
		Case "�e"
			StrConvNarrow = StrConvNarrow & "'"
		Case "�f"
			StrConvNarrow = StrConvNarrow & "'"
		Case "�g"
			StrConvNarrow = StrConvNarrow & """"
		Case "�h"
			StrConvNarrow = StrConvNarrow & """"
		Case "�i"
			StrConvNarrow = StrConvNarrow & "("
		Case "�j"
			StrConvNarrow = StrConvNarrow & ")"
		Case "�m"
			StrConvNarrow = StrConvNarrow & "["
		Case "�n"
			StrConvNarrow = StrConvNarrow & "]"
		Case "�o"
			StrConvNarrow = StrConvNarrow & "{"
		Case "�p"
			StrConvNarrow = StrConvNarrow & "}"
		Case "�u"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�v"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�{"
			StrConvNarrow = StrConvNarrow & "+"
		Case "�|"
			StrConvNarrow = StrConvNarrow & "-"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "="
		Case "��"
			StrConvNarrow = StrConvNarrow & "<"
		Case "��"
			StrConvNarrow = StrConvNarrow & ">"
		Case "��"
			StrConvNarrow = StrConvNarrow & "\"
		Case "��"
			StrConvNarrow = StrConvNarrow & "$"
		Case "��"
			StrConvNarrow = StrConvNarrow & "%"
		Case "��"
			StrConvNarrow = StrConvNarrow & "#"
		Case "��"
			StrConvNarrow = StrConvNarrow & "&"
		Case "��"
			StrConvNarrow = StrConvNarrow & "*"
		Case "��"
			StrConvNarrow = StrConvNarrow & "@"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�O"
			StrConvNarrow = StrConvNarrow & "0"
		Case "�P"
			StrConvNarrow = StrConvNarrow & "1"
		Case "�Q"
			StrConvNarrow = StrConvNarrow & "2"
		Case "�R"
			StrConvNarrow = StrConvNarrow & "3"
		Case "�S"
			StrConvNarrow = StrConvNarrow & "4"
		Case "�T"
			StrConvNarrow = StrConvNarrow & "5"
		Case "�U"
			StrConvNarrow = StrConvNarrow & "6"
		Case "�V"
			StrConvNarrow = StrConvNarrow & "7"
		Case "�W"
			StrConvNarrow = StrConvNarrow & "8"
		Case "�X"
			StrConvNarrow = StrConvNarrow & "9"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�`"
			StrConvNarrow = StrConvNarrow & "A"
		Case "�a"
			StrConvNarrow = StrConvNarrow & "B"
		Case "�b"
			StrConvNarrow = StrConvNarrow & "C"
		Case "�c"
			StrConvNarrow = StrConvNarrow & "D"
		Case "�d"
			StrConvNarrow = StrConvNarrow & "E"
		Case "�e"
			StrConvNarrow = StrConvNarrow & "F"
		Case "�f"
			StrConvNarrow = StrConvNarrow & "G"
		Case "�g"
			StrConvNarrow = StrConvNarrow & "H"
		Case "�h"
			StrConvNarrow = StrConvNarrow & "I"
		Case "�i"
			StrConvNarrow = StrConvNarrow & "J"
		Case "�j"
			StrConvNarrow = StrConvNarrow & "K"
		Case "�k"
			StrConvNarrow = StrConvNarrow & "L"
		Case "�l"
			StrConvNarrow = StrConvNarrow & "M"
		Case "�m"
			StrConvNarrow = StrConvNarrow & "N"
		Case "�n"
			StrConvNarrow = StrConvNarrow & "O"
		Case "�o"
			StrConvNarrow = StrConvNarrow & "P"
		Case "�p"
			StrConvNarrow = StrConvNarrow & "Q"
		Case "�q"
			StrConvNarrow = StrConvNarrow & "R"
		Case "�r"
			StrConvNarrow = StrConvNarrow & "S"
		Case "�s"
			StrConvNarrow = StrConvNarrow & "T"
		Case "�t"
			StrConvNarrow = StrConvNarrow & "U"
		Case "�u"
			StrConvNarrow = StrConvNarrow & "V"
		Case "�v"
			StrConvNarrow = StrConvNarrow & "W"
		Case "�w"
			StrConvNarrow = StrConvNarrow & "X"
		Case "�x"
			StrConvNarrow = StrConvNarrow & "Y"
		Case "�y"
			StrConvNarrow = StrConvNarrow & "Z"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "a"
		Case "��"
			StrConvNarrow = StrConvNarrow & "b"
		Case "��"
			StrConvNarrow = StrConvNarrow & "c"
		Case "��"
			StrConvNarrow = StrConvNarrow & "d"
		Case "��"
			StrConvNarrow = StrConvNarrow & "e"
		Case "��"
			StrConvNarrow = StrConvNarrow & "f"
		Case "��"
			StrConvNarrow = StrConvNarrow & "g"
		Case "��"
			StrConvNarrow = StrConvNarrow & "h"
		Case "��"
			StrConvNarrow = StrConvNarrow & "i"
		Case "��"
			StrConvNarrow = StrConvNarrow & "j"
		Case "��"
			StrConvNarrow = StrConvNarrow & "k"
		Case "��"
			StrConvNarrow = StrConvNarrow & "l"
		Case "��"
			StrConvNarrow = StrConvNarrow & "m"
		Case "��"
			StrConvNarrow = StrConvNarrow & "n"
		Case "��"
			StrConvNarrow = StrConvNarrow & "o"
		Case "��"
			StrConvNarrow = StrConvNarrow & "p"
		Case "��"
			StrConvNarrow = StrConvNarrow & "q"
		Case "��"
			StrConvNarrow = StrConvNarrow & "r"
		Case "��"
			StrConvNarrow = StrConvNarrow & "s"
		Case "��"
			StrConvNarrow = StrConvNarrow & "t"
		Case "��"
			StrConvNarrow = StrConvNarrow & "u"
		Case "��"
			StrConvNarrow = StrConvNarrow & "v"
		Case "��"
			StrConvNarrow = StrConvNarrow & "w"
		Case "��"
			StrConvNarrow = StrConvNarrow & "x"
		Case "��"
			StrConvNarrow = StrConvNarrow & "y"
		Case "��"
			StrConvNarrow = StrConvNarrow & "z"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�@"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�A"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�B"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�C"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�D"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�F"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�G"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�H"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�I"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�J"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�K"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�L"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�M"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�N"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�O"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�P"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�Q"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�R"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�S"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�T"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�U"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�V"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�W"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�X"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�Y"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�Z"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�["
			StrConvNarrow = StrConvNarrow & "��"
		Case "�\"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�]"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�^"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�_"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�`"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�a"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�b"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�c"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�d"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�e"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�f"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�g"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�h"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�i"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�j"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�k"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�l"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�m"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�n"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�o"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�p"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�q"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�r"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�s"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�t"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�u"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�v"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�w"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�x"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�y"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�z"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�{"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�|"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�}"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�~"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "�"
		Case "��"
			StrConvNarrow = StrConvNarrow & "��"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�U"
			StrConvNarrow = StrConvNarrow & "|"
		Case "�V"
			StrConvNarrow = StrConvNarrow & "'"
		Case "�W"
			StrConvNarrow = StrConvNarrow & """"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case "�U"
			StrConvNarrow = StrConvNarrow & "|"
		Case "�V"
			StrConvNarrow = StrConvNarrow & "'"
		Case "�W"
			StrConvNarrow = StrConvNarrow & """"
		Case "�E"
			StrConvNarrow = StrConvNarrow & "�"
		Case Else
			StrConvNarrow = StrConvNarrow & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvKatakana
'---------------------------------------------------
' �p�r : StrConv(s,vbKatakana) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvKatakana(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("��", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case "�T"
			StrConvKatakana = StrConvKatakana & "�R"
		Case "�U"
			StrConvKatakana = StrConvKatakana & "�S"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�@"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�A"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�B"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�C"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�D"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�E"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�F"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�G"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�H"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�I"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�J"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�K"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�L"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�M"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�N"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�O"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�P"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�Q"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�R"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�S"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�T"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�U"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�V"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�W"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�X"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�Y"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�Z"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�["
		Case "��"
			StrConvKatakana = StrConvKatakana & "�\"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�]"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�^"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�_"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�`"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�a"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�b"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�c"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�d"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�e"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�f"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�g"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�h"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�i"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�j"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�k"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�l"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�m"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�n"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�o"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�p"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�q"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�r"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�s"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�t"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�u"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�v"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�w"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�x"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�y"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�z"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�{"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�|"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�}"
		Case "��"
			StrConvKatakana = StrConvKatakana & "�~"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case "��"
			StrConvKatakana = StrConvKatakana & "��"
		Case Else
			StrConvKatakana = StrConvKatakana & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvHiragana
'---------------------------------------------------
' �p�r : StrConv(s,vbHiragana) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvHiragana(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("��", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case "�R"
			StrConvHiragana = StrConvHiragana & "�T"
		Case "�S"
			StrConvHiragana = StrConvHiragana & "�U"
		Case "�@"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�A"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�B"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�C"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�D"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�E"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�F"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�G"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�H"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�I"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�J"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�K"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�L"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�M"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�N"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�O"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�P"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�Q"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�R"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�S"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�T"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�U"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�V"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�W"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�X"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�Y"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�Z"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�["
			StrConvHiragana = StrConvHiragana & "��"
		Case "�\"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�]"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�^"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�_"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�`"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�a"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�b"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�c"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�d"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�e"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�f"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�g"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�h"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�i"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�j"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�k"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�l"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�m"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�n"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�o"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�p"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�q"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�r"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�s"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�t"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�u"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�v"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�w"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�x"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�y"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�z"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�{"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�|"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�}"
			StrConvHiragana = StrConvHiragana & "��"
		Case "�~"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case "��"
			StrConvHiragana = StrConvHiragana & "��"
		Case Else
			StrConvHiragana = StrConvHiragana & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvUnicode
'---------------------------------------------------
' �p�r : StrConv(s,vbUnicode) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvUnicode(sInp)
	Dim nCnt
	Dim nLen
	Dim nAsc
	Dim nChr
	nLen = LenB(sInp)
	For nCnt = 1 To nLen
		nAsc = AscB(MidB(sInp, nCnt, 1))
		If (&h81 <= nAsc And nAsc <= &h9F) Or (&hE0 <= nAsc And nAsc <= &hEF) Then
			nChr = nAsc * 256 + AscB(MidB(sInp, nCnt+1, 1))
			StrConvUnicode = StrConvUnicode & Chr(nChr)
			nCnt = nCnt + 1
		Else
			StrConvUnicode = StrConvUnicode & Chr(AscB(MidB(sInp, nCnt, 1)))
		End If
	Next
End Function

'***************************************************
' StrConvFromUnicode
'---------------------------------------------------
' �p�r : StrConv(s,vbFromUnicode) �̃N���[��
' ���� : �ϊ����镶����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConvFromUnicode(sInp)
	Dim nCnt
	Dim nLen
	Dim nAsc
	nLen = Len(sInp)
	For nCnt = 1 to nLen
		nAsc = Asc(Mid(sInp, nCnt, 1))
		If nAsc And &hFF00 Then
			StrConvFromUnicode = StrConvFromUnicode & ChrB(Int(nAsc / 256) And &hFF)
			StrConvFromUnicode = StrConvFromUnicode & ChrB(nAsc And &hFF)
		Else
			StrConvFromUnicode = StrConvFromUnicode & ChrB(nAsc)
		End If
	Next
End Function

'***************************************************
' StrConv ���g�p����萔�S
'***************************************************
' Enum VbStrConv
Const vbUpperCase=1
Const vbLowerCase=2
Const vbProperCase=3
Const vbWide=4
Const vbNarrow=8
Const vbKatakana=16
Const vbHiragana=32
Const vbUnicode = 64
Const vbFromUnicode = 128

'***************************************************
' StrConv
'---------------------------------------------------
' ���� : �ϊ����镶����,�ϊ�����
' �ߒl : �ϊ����ꂽ������
'***************************************************
Function StrConv(sInp,eCnv)
	StrConv = sInp
	' Cnv �ɑ΂��ď�����U�蕪��
	If eCnv And vbUpperCase Then
		StrConv = StrConvUpperCase(StrConv)
	End If
	If eCnv And vbLowerCase Then
		StrConv = StrConvLowerCase(StrConv)
	End If
	If eCnv = vbProperCase Then
		StrConv = StrConvProperCase(StrConv)
	End If
	If eCnv And vbWide Then
		StrConv = StrConvWide(StrConv)
	End If
	If eCnv And vbNarrow Then
		StrConv = StrConvNarrow(StrConv)
	End If
	If eCnv And vbKatakana Then
		StrConv = StrConvKatakana(StrConv)
	End If
	If eCnv And vbHiragana Then
		StrConv = StrConvHiragana(StrConv)
	End If
	If eCnv And vbUnicode Then
		StrConv = StrConvUnicode(StrConv)
	End If
	If eCnv And vbFromUnicode Then
		StrConv = StrConvFromUnicode(StrConv)
	End If
End Function

Function GetQty(strB,strMode)
	dim	nLen
	dim	nCnt
	dim	nChr
	dim	strRet

	strRet = ""
	nLen = Len(strB)
	For nCnt = 1 to nLen
		nChr = Mid(strB, nCnt, 1)
		if nChr = "0" then
			nChr = strMode
		else
		end if
		strRet = strRet & nChr
	Next
	GetQty = strRet
End Function

Function GetSaisu(strSaisu)
	dim	dblSaisu
	dim	lngSaisu
	
	lngSaisu = 0
	dblSaisu	= cdbl(Rtrim(strSaisu))
	if dblSaisu > 0 then
		lngSaisu = round(dblSaisu,0)
		if lngSaisu = 0 then
			lngSaisu = 1
		end if
	end if
	GetSaisu = lngSaisu
End Function

