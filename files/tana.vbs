Option Explicit
' ITEM �I�ԃ`�F�b�N
' 2009.11.09 �V�K�쐬

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly	= 0
Const adOpenKeyset		= 1
Const adOpenDynamic		= 2
Const adOpenStatic		= 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

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

function LogOutput(fl,msg)
	Wscript.Echo	msg
	fl.WriteLine	msg
end function

function GetDateTime(dt)
	dim	tmpYYYYMMDD
	dim	tmpHHMMSS
	'/// �N���� �쐬
	tmpYYYYMMDD = year(dt) & Right(00 & month(dt), 2) & Right(00 & day(dt), 2)
	'/// ���� �쐬   
	tmpHHMMSS   = Right(00 & hour(dt), 2) & Right(00 & minute(dt), 2) & Right(00 & second(dt), 2)
	'/// ����   
	GetDateTime = tmpYYYYMMDD & tmpHHMMSS
end function

function usage()
    Wscript.Echo "�I�ԃ`�F�b�N(2009.11.09)"
    Wscript.Echo "tana.vbs [option] <JGyobu> <JCode> [PN]"
    Wscript.Echo " JGyobu:���ƕ��R�[�h"
    Wscript.Echo "  JCode:���Ə�R�[�h"
    Wscript.Echo "     PN:�i��"
    Wscript.Echo " -t : Test���[�h"
    Wscript.Echo " -?"
end function

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

Function SetField(rsItem,strFieldName,strValue,strTitle,strUpdMsg)
	dim	strItemValue
	dim	strPnValue

	strItemValue = rtrim(rsItem.Fields(strFieldName))
	strPnValue	 = rtrim(strValue)
	select case strFieldName
	case "HIN_NAME"
		select case rsItem.Fields("JGYOBU")
		case "7"
			if strItemValue <> "" then
				strPnValue = strItemValue
			end if
		end select
	case "L_URIKIN1", _
		 "L_URIKIN2", _
		 "L_URIKIN3"
		strItemValue	= CStr(GetNumValue(strItemValue))
		strPnValue		= CStr(fix(GetNumValue(strPnValue)))
	case "HYO_TANKA"
		strItemValue	= CStr(GetNumValue(strItemValue))
		strPnValue		= CStr(GetNumValue(strPnValue))
	case "GENSANKOKU"
		if strPnValue = "" then
			select case rsItem.Fields("JGYOBU")
			case "4","A","R"
			case "D","7","1"
				strPnValue = "JAPAN"
			end select
		end if
	case "GLICS2_TANA", _
		 "GLICS3_TANA"
		select case rsItem.Fields("JGYOBU")
		case "4","D"
			if strItemValue <> "" then
				strPnValue = strItemValue
			end if
		end select
	end select
	if strItemValue <> strPnValue then
		strUpdMsg = strUpdMsg & strTitle & "(�ύX�O)��" & strItemValue & vbNewLine
		strUpdMsg = strUpdMsg & strTitle & "(�ύX��)��" & strPnValue   & vbNewLine
'		strUpdMsg = strUpdMsg & strFieldName & "(" & rsItem.Fields(strFieldName).Type & "," & rsItem.Fields(strFieldName).DefinedSize & "," & rsItem.Fields(strFieldName).ActualSize & ")" & vbNewLine
		if Get_LenB(strPnValue) > rsItem.Fields(strFieldName).DefinedSize then
			strUpdMsg = strUpdMsg & "DefinedSize:" & rsItem.Fields(strFieldName).DefinedSize & vbNewLine
			strPnValue = Get_LeftB(strPnValue,rsItem.Fields(strFieldName).DefinedSize)
		end if
'		strUpdMsg = strUpdMsg & "Err.Number:" & Err.Number & vbNewLine
On Error Resume Next
		rsItem.Fields(strFieldName) = strPnValue
		if Err.Number <> 0 then
			strUpdMsg = strUpdMsg & "Err.Number:" & Err.Number & vbNewLine
			strUpdMsg = strUpdMsg & "(" & rsItem.Fields(strFieldName) & ")" & vbNewLine
			strUpdMsg = strUpdMsg & "Get_LenB(" & Get_LenB(rsItem.Fields(strFieldName)) & ")" & vbNewLine
			Err.Clear
		end if
		rsItem.Fields("UPD_TANTO") = "pn2it"
		rsItem.Fields("UPD_DATETIME") = GetDateTime(now())
	end if
	SetField = strUpdMsg
End Function

dim	db
dim	dbName
dim	sqlStr
dim	rsList
dim	rsPn
dim	strBuff
dim	i
dim	Fs
dim	logFile
dim	strJGyobu
dim	strJCode
dim	strPn
dim	flgTest
dim	strUpdMsg
dim	lngCnt
dim	lngUpd
dim	lngErr

flgTest = False
strJGyobu = ""
strJCode = ""
strPn	 = ""
for i = 0 to WScript.Arguments.count - 1
    select case ucase(WScript.Arguments(i))
    case "-T"
		flgTest = True
    case "-?"
		usage()
		Wscript.Quit
    case else
		if strJGyobu = "" then
		        strJGyobu = WScript.Arguments(i)
		elseif strJCode = "" then
		        strJCode = WScript.Arguments(i)
		elseif strPn = "" then
		        strPn = WScript.Arguments(i)
		else
			usage()
			Wscript.Quit
		end if
    end select
next
if strJGyobu = "" then
	usage()
	Wscript.Quit
end if

Wscript.Echo "tana.vbs"
if flgTest = True then
	Wscript.Echo "TEST���[�h"
end if

dbName = "newsdc"
Set db = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbName
db.open dbName

sqlStr = "select "
sqlStr = sqlStr & " JGYOBU"
sqlStr = sqlStr & ",NAIGAI"
sqlStr = sqlStr & ",HIN_GAI"
sqlStr = sqlStr & ",ST_SET_DT"
if strJGyobu = "7" then
	sqlStr = sqlStr & ",if(ST_SOKO >= '01' and ST_SOKO <= '26',char(convert(ST_SOKO,SQL_NUMERIC)+64),ST_SOKO) + ST_RETU + ST_REN + ST_DAN as ST_TANA"
else
	sqlStr = sqlStr & ",ST_SOKO + ST_RETU + ST_REN + ST_DAN as ST_TANA"
end if
if strJCode <> "" then
	sqlStr = sqlStr & ",pn2.Loc1 as G_TANA1"
else
	sqlStr = sqlStr & ",GLICS1_TANA as G_TANA1"
end if

sqlStr = sqlStr & " from item"
if strJCode <> "" then
	sqlStr = sqlStr & " inner join pn2 on (HIN_GAI = pn2.Pn and pn2.ShisanJCode = '" & strJCode & "')"
end if
sqlStr = sqlStr & " where jgyobu = '" & strJGyobu & "'"
sqlStr = sqlStr & "   and naigai = '1'"
sqlStr = sqlStr & "   and ST_SOKO <> ''"
if strJGyobu = "7" then
	if strJCode <> "" then
		sqlStr = sqlStr & "   and (if(JGYOBU = '7' and ST_SOKO >= '01' and ST_SOKO <= '26',char(convert(ST_SOKO,SQL_NUMERIC)+64),ST_SOKO) + ST_RETU + ST_REN + ST_DAN) <> pn2.Loc1"
	else
		sqlStr = sqlStr & "   and (if(JGYOBU = '7' and ST_SOKO >= '01' and ST_SOKO <= '26',char(convert(ST_SOKO,SQL_NUMERIC)+64),ST_SOKO) + ST_RETU + ST_REN + ST_DAN) <> GLICS1_TANA"
	end if
else
	if strJCode <> "" then
		sqlStr = sqlStr & "   and (ST_SOKO + ST_RETU + ST_REN + ST_DAN) <> pn2.Loc1"
	else
		sqlStr = sqlStr & "   and (ST_SOKO + ST_RETU + ST_REN + ST_DAN) <> GLICS1_TANA"
	end if
end if
if strPn <> "" then
	sqlStr = sqlStr & "   and hin_gai = '" & strPn & "'"
end if


' �R�}���h�^�C���A�E�g�ύX:0 �^�C���A�E�g�Ȃ�
Wscript.Echo "db.CommandTimeout : " & db.CommandTimeout
db.CommandTimeout = 0
Wscript.Echo "db.CommandTimeout : " & db.CommandTimeout
Wscript.Echo "sql : " & sqlStr

' ���R�[�h�Z�b�g�I�u�W�F�N�g�쐬
' set rsList = db.Execute(sqlStr)
Set rsList = Wscript.CreateObject("ADODB.Recordset")

' �J�[�\�����P�[�V�����Z�b�g:UseClient
Wscript.Echo "rsList.CursorLocation : " & rsList.CursorLocation
'rsList.CursorLocation = adUseClient
Wscript.Echo "rsList.CursorLocation : " & rsList.CursorLocation

'rsList.Open sqlStr, db, adOpenDynamic, adLockOptimistic
'rsList.Open sqlStr, db, adOpenForwardOnly, adLockBatchOptimistic
'rsList.Open sqlStr, db, adOpenForwardOnly, adLockReadOnly
rsList.Open sqlStr, db, adOpenStatic, adLockReadOnly

Wscript.Echo "sql : ����"

'On Error Resume Next
lngCnt = 0
lngUpd = 0
lngErr = 0
Do While Not rsList.EOF
	strBuff = rsList.Fields("JGYOBU")
	strBuff = strBuff & " " & rsList.Fields("NAIGAI")
	strBuff = strBuff & " " & rsList.Fields("HIN_GAI")
	strBuff = strBuff & " " & rsList.Fields("ST_SET_DT")
	strBuff = strBuff & " " & rsList.Fields("ST_TANA")
	strBuff = strBuff & " " & rsList.Fields("G_TANA1")
'
	Wscript.Echo strBuff
	rsList.movenext
Loop

Wscript.Echo ""
rsList.close
Wscript.Echo "close db : " & dbName
db.Close
set db = nothing
Wscript.Echo "end"
