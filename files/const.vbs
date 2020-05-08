'-------------------------------
'const.vbs
'newsdc\files用
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
' オブジェクトを開いているか閉じているか、データ ソースに接続中か、
' コマンドを実行中か、またはデータを取得中かどうかを表します。
Const	adStateClosed		= 0 ' オブジェクトが閉じていることを示します。 
Const	adStateOpen			= 1 ' オブジェクトが開いていることを示します。 
Const	adStateConnecting	= 2 ' オブジェクトが接続していることを示します。 
Const	adStateExecuting	= 4 ' オブジェクトがコマンドを実行中であることを示します。 
Const	adStateFetching		= 8 ' オブジェクトの行が取得されていることを示します。 

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
	'/// 年月日 作成
	GetDate = year(dt) & Right(00 & month(dt), 2) & Right(00 & day(dt), 2)
End Function

function GetDateTime(dt)
	dim	tmpYYYYMMDD
	dim	tmpHHMMSS
	'/// 年月日 作成
	tmpYYYYMMDD = year(dt) & Right(00 & month(dt), 2) & Right(00 & day(dt), 2)
	'/// 時分 作成   
	tmpHHMMSS   = Right(00 & hour(dt), 2) & Right(00 & minute(dt), 2) & Right(00 & second(dt), 2)
	'/// 合成   
	GetDateTime = tmpYYYYMMDD & tmpHHMMSS
end function

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub
'-----------------------------------------------------------------------
'データベースオープン
'-----------------------------------------------------------------------
Function OpenAdodb(byval strDbName)
	dim	objDb
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	Call objDb.Open(strDbName)
	Set OpenAdodb = objDb
End Function
'-----------------------------------------------------------------------
'データベースクローズ
'-----------------------------------------------------------------------
Function CloseAdodb(byval objDb)
	objDb.Close
	set CloseAdodb = Nothing
End Function
'-----------------------------------------------------------------------
'オプションチェック
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
'フィールド値
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
'レコードセットオープン
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
'レコードセットExecute
'-----------------------------------------------------------------------
Function ExecuteAdodb(byval objDb,byval strSql)
	set ExecuteAdodb = objDb.Execute(strSql)
End Function
'-----------------------------------------------------------------------
'レコードセットクローズ
'-----------------------------------------------------------------------
Function CloseRs(byval objRs)
	if objRs.State <> adStateClosed then
		objRs.Close
	end if
	set CloseRs = Nothing
End Function
'-----------------------------------------------------------------------
'テーブルコピー
'-----------------------------------------------------------------------
Function CopyTable(objDb,byVal strSrc,byVal strDst)
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsSrc
	set rsSrc = ExecuteAdodb(objDb,"select top 1 * from " & strSrc)
	'-------------------------------------------------------------------
	'insert SQL文作成
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
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsSrc = CloseRs(rsSrc)
	'-------------------------------------------------------------------
	'削除SQL実行
	'-------------------------------------------------------------------
	Call ExecuteAdodb(objDb,"delete from " & strDst)
	Call ExecuteAdodb(objDb,strSql)
End Function
'-----------------------------------------------
' 文字列連結CrLf
'-----------------------------------------------
Function StrCrLf(byVal strDst,byVal strAdd)
	StrCrLf = strDst & strAdd & vbCrLf
End Function
