'-------------------------------
'const.vbs
'newsdc\files�p
'-------------------------------
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CommandTypeEnum Values ----
Const adCmdUnspecified	= -1	' Unspecified type of command 
Const adCmdText			= 1		' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2		' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4		' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8		' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

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

Function makeMsg(byval sVal,byval iLen)
	sVal = RTrim(sVal)
	if iLen > 0 then
		sVal = Right(space(iLen) & sVal,iLen)
	else
		iLen = iLen * -1
		sVal = Left(sVal & space(iLen),iLen)
	end if
	makeMsg = sVal
End Function

Function GetDate(dt)
	'/// �N���� �쐬
	GetDate = year(dt) & Right(00 & month(dt), 2) & Right(00 & day(dt), 2)
End Function

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

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub
'-----------------------------------------------------------------------
'�f�[�^�x�[�X�I�[�v��
'-----------------------------------------------------------------------
Function OpenAdodb(byval strDbName)
	dim	objDb
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	Call objDb.Open(strDbName)
	Set OpenAdodb = objDb
End Function
'-----------------------------------------------------------------------
'�f�[�^�x�[�X�N���[�Y
'-----------------------------------------------------------------------
Function CloseAdodb(byval objDb)
	objDb.Close
	set CloseAdodb = Nothing
End Function
'-----------------------------------------------------------------------
'�I�v�V�����`�F�b�N
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
'�t�B�[���h�l
'-----------------------------------------------------------------------
Function GetFieldValue(byval objRs _
		      ,byval strName _
		      )
	dim	v
'	Debug "GetFieldValue(" & strName & "):Type=" & objRs.Fields(strName).Type
	On Error Resume Next
		v = objRs.Fields(strName)
		if Err.Number <> 0 then
			strMsg = strMsg & " GetFieldValue() Error:" & Hex(Err.Number) & " " & Err.Description
		end if
	On Error Goto 0
	
	select case objRs.Fields(strName).Type
	case 6
		if isnull(v) then
			v = 0
		end if
		if v = "" then
			v = 0
		end if
	case else
		if isnull(v) then
			v = ""
		end if
	end select
	GetFieldValue = Rtrim(v)
End Function
'-----------------------------------------------------------------------
'���R�[�h�Z�b�g�I�[�v��
'-----------------------------------------------------------------------
Function OpenRs(objDb,byval strTableName)
	dim	objRs
	Set objRs = Wscript.CreateObject("ADODB.Recordset")
	if strTableName <> "" then
		objRs.Open strTableName, objDb, adOpenKeyset, adLockOptimistic,adCmdTableDirect
	end if
	set OpenRs = objRs
End Function
Function UpdateOpenRs(objDb,objRs,byval strSql)
	if objRs.State <> adStateClosed then
		objRs.Close
	end if
	objRs.Open strSql, objDb, adOpenKeyset, adLockOptimistic
	UpdateOpenRs = objRs.EOF
End Function
'-----------------------------------------------------------------------
'���R�[�h�Z�b�gExecute
'-----------------------------------------------------------------------
Function ExecuteAdodb(byval objDb,byval strSql)
	set ExecuteAdodb = objDb.Execute(strSql)
End Function
'-----------------------------------------------------------------------
'���R�[�h�Z�b�g�N���[�Y
'-----------------------------------------------------------------------
Function CloseRs(byval objRs)
	if objRs.State <> adStateClosed then
		objRs.Close
	end if
	set CloseRs = Nothing
End Function
'-----------------------------------------------------------------------
'�e�[�u���R�s�[
'-----------------------------------------------------------------------
Function CopyTable(objDb,byVal strSrc,byVal strDst)
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	rsSrc
	set rsSrc = ExecuteAdodb(objDb,"select top 1 * from " & strSrc)
	'-------------------------------------------------------------------
	'insert SQL���쐬
	'-------------------------------------------------------------------
	dim	strSql
	strSql = ""
	strSql = StrCrLf(strSql,"insert into " & strDst)
	dim	strDlm
	strDlm = "("
	dim	objF
	for each objF in (rsSrc.Fields)
		strSql = StrCrLf(strSql,strDlm & objF.Name)
		strDlm = ","
	next
	strSql = StrCrLf(strSql,")")
	strSql = StrCrLf(strSql,"select * from " & strSrc)
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsSrc = CloseRs(rsSrc)
	'-------------------------------------------------------------------
	'�폜SQL���s
	'-------------------------------------------------------------------
	Call ExecuteAdodb(objDb,"delete from " & strDst)
	Call ExecuteAdodb(objDb,strSql)
End Function
'-----------------------------------------------
' ������A��CrLf
'-----------------------------------------------
Function StrCrLf(byVal strDst,byVal strAdd)
	StrCrLf = strDst & strAdd & vbCrLf
End Function
