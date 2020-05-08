Option Explicit
' 積水注文データ処理
' 2011.12.22 大阪PC対応
Call Main()
WScript.Quit 0
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
Const adCmdText			= 1	' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2	' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4	' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8	' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

function usage()
	Wscript.Echo "積水注文データ処理"
	Wscript.Echo "y_syuka_tei.vbs <database> [option]"
	Wscript.Echo " -?"
end function

Sub Main()
	dim	i
	dim	strArg
	dim	strDbName

	strDbName = ""
	for i = 0 to WScript.Arguments.count - 1
		strArg = WScript.Arguments(i)
	    select case strArg
		case "-?"
			call usage()
			exit sub
		case else
			if strDbName = "" then
				strDbName = strArg
			else
				usage()
				exit sub
			end if
		end select
	next
	if strDbName = "" then
		usage()
		exit sub
	end if
	call ySyukaRecober(strDbName)
End Sub

Sub ySyukaRecober(strDbName)
	dim	objDb
	dim	rsTei
	dim	rsTeiW
	dim	strSql
	dim	strTeiId

	call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	call DispMsg("Open:" & strDbName)
	objDb.Open strDbName
	call DispMsg(objDb.ConnectionString)

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsTei = Wscript.CreateObject("ADODB.Recordset")

	strSql = GetOpenSql()
	call DispMsg("open:" & strSql)
	rsTei.Open strSql, objDb, adOpenKeyset, adLockBatchOptimistic

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsTeiW = Wscript.CreateObject("ADODB.Recordset")

	Do While Not rsTei.EOF
		call DispMsg("  " & _
					 rsTei.Fields("SND_YMD") & " " & _
					 rsTei.Fields("SND_HMS") & " " & _
					 rsTei.Fields("TOK_CD") & " " & _
					 rsTei.Fields("CHO_CD") & " " & _
					 rsTei.Fields("TEI_LABELID") _
					)
		strSql = GetTeiW(rsTei.Fields("TEI_LABELID"))
		rsTeiW.Open strSql, objDb, adOpenKeyset, adLockBatchOptimistic
		if rsTeiW.EOF = False then
			strTeiId = GetNewTeiId(rsTeiW.Fields("TEI_LABELID"))
			call DispMsg("W " & _
						 rsTeiW.Fields("SND_YMD") & " " & _
						 rsTeiW.Fields("SND_HMS") & " " & _
						 rsTeiW.Fields("TOK_CD") & " " & _
						 rsTeiW.Fields("CHO_CD") & " " & _
						 strTeiId & " " & _
						 rsTeiW.Fields("TEI_LABELID") _
						)
			rsTeiW.Fields("TEI_LABELID") = strTeiId
			rsTeiW.UpdateBatch
		end if
		rsTeiW.Close

		rsTei.MoveNext
	Loop

	rsTei.Close
	Set	rsTei = Nothing
	Set	rsTeiW = Nothing

	call DispMsg("Close:" & strDbName)
	objDb.Close
	set objDb = Nothing
End Sub
Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub

Function GetNewTeiId(byval strTeiId)
	dim	strC
	strTeiId = rtrim(strTeiId)
	strC = mid(strTeiId,8,1)
	strC = chr(asc(strC) + 1)
	strTeiId = left(strTeiId,7) & strC & right(strTeiId,2)
	GetNewTeiId = strTeiId
End Function

Function GetTeiW(byval strTeiId)
	dim	strSql

	strSql = "select *"
	strSql = strSql & " from y_syuka_tei"
	strSql = strSql & " where (SND_YMD <> '20131002'"
	strSql = strSql & "     or SND_HMS <> '141142')"
	strSql = strSql & "    and TEI_LABELID = '" & rtrim(strTeiId) & "'"
	
	GetTeiW = strSql
End Function


Function GetOpenSql()
	dim	strSql

	strSql = "select *"
	strSql = strSql & " from y_syuka_tei"
	strSql = strSql & " where SND_YMD = '20131002'"
	strSql = strSql & "   and SND_HMS = '141142'"
	strSql = strSql & "    and TEI_LABELID in ('XYG1573Q01')"

	
	GetOpenSql = strSql
End Function
