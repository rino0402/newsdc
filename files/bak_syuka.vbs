Option Explicit
' bak_syuka 出荷予定データ復旧
' 2011.05.20 新規作成
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
Const adCmdText		= 1	' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable	= 2	' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4	' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown	= 8	' Default. Unknown type of command 
Const adCmdFile		= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

function usage()
	Wscript.Echo "出荷予定データ復旧処理"
	Wscript.Echo "bak_syuka.vbs <database> [option]"
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
	dim	rsBak
	dim	rsDel
	dim	fldBak
	dim	strSql
	dim	strMsg

	call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	
	call DispMsg("Open:" & strDbName)
	objDb.Open strDbName
	call DispMsg(objDb.ConnectionString)

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsDel = Wscript.CreateObject("ADODB.Recordset")
	rsDel.MaxRecords = 1
	rsDel.CursorLocation = adUseServer
	call DispMsg("open:del_syuka")
	rsDel.Open "del_syuka", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

	strSql = GetOpenSql()
	call DispMsg("Execute:" & strSql)
	set rsBak = objDb.Execute(strSql)
	call DispMsg("Execute:End")


	Do While Not rsBak.EOF
'		call DispMsg(rsBak.Fields("KEY_ID_NO"))
		rsDel.Addnew
		for each fldBak in rsBak.Fields
			' call DispMsg(fldBak.Name & ":" & fldBak)
			rsDel.Fields(fldBak.Name) = fldBak
		next
		rsDel.Fields("UPD_NOW") = "20110520000000"
		On Error Resume Next
		rsDel.UpdateBatch
		if Err.Number = 0 then
			strMsg = "Ok"
		else
			strMsg = "Err:" & Err.Number & " " & Err.Description
		end if
		call DispMsg(rsBak.Fields("KEY_ID_NO:") & strMsg)
		Err.Clear
		On Error Goto 0
		rsBak.MoveNext
	Loop

	call DispMsg("Close:del_syuka")
	rsDel.Close
	set rsDel = Nothing
	
	call DispMsg("Close:bak_syuka")
	rsBak.Close
	set rsBak = Nothing

	call DispMsg("Close:" & strDbName)
	objDb.Close
	set objDb = Nothing
End Sub

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub

Function GetOpenSql()
	dim	strSql

	strSql = "select *"
	strSql = strSql & " from bak_syuka"
	strSql = strSql & " where KEY_SYUKA_YMD = '20110518'"
	strSql = strSql & "   and KEY_ID_NO not in"
	strSql = strSql & "       (select KEY_ID_NO from del_syuka where KEY_SYUKA_YMD = '20110518')"
	strSql = strSql & "   and KEY_ID_NO not in"
	strSql = strSql & "       (select KEY_ID_NO from   y_syuka where KEY_SYUKA_YMD = '20110518')"

	GetOpenSql = strSql
End Function
