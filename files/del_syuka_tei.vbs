Option Explicit
' del_syuka �o�ח\��f�[�^����
' 2011.05.20 �V�K�쐬
' 2011.07.21 ���PC�Ή�
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
	Wscript.Echo "�o�ח\��f�[�^��������"
	Wscript.Echo "del_syuka_tei.vbs <database> [option]"
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
	dim	strTest

	strTest = ""

	call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	
	call DispMsg("Open:" & strDbName)
	objDb.Open strDbName
	call DispMsg(objDb.ConnectionString)

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsDel = Wscript.CreateObject("ADODB.Recordset")
	rsDel.MaxRecords = 1
	rsDel.CursorLocation = adUseServer
	call DispMsg("open:y_syuka_tei")
	rsDel.Open "y_syuka_tei", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

	strSql = GetOpenSql()
	call DispMsg("Execute:" & strSql)
	set rsBak = objDb.Execute(strSql)
	call DispMsg("Execute:End")

	Do While Not rsBak.EOF
'debug		call DispMsg(rsBak.Fields("KEY_ID_NO"))
		if strTest = "" then
			rsDel.Addnew
		end if
		for each fldBak in rsBak.Fields
'debug			call DispMsg(fldBak.Name & ":" & fldBak)
			if strTest = "" then
				rsDel.Fields(fldBak.Name) = fldBak
			end if
		next
		On Error Resume Next
		if strTest = "" then
			rsDel.UpdateBatch
			if Err.Number = 0 then
				strMsg = "Ok"
			else
				strMsg = "Err:" & Err.Number & " " & Err.Description
			end if
		else
			strMsg = "Test"
		end if
		Err.Clear
		On Error Goto 0
		call DispMsg(rsBak.Fields("SND_YMD") & " " & rsBak.Fields("SND_HMS") & " "  & rsBak.Fields("TOK_CD") & " " & rsBak.Fields("CHO_CD") & ":" & strMsg)
		rsBak.MoveNext
	Loop

	call DispMsg("Close:y_syuka_tei")
	rsDel.Close
	set rsDel = Nothing
	
	call DispMsg("Close:del_syuka_tei")
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
	strSql = strSql & " from del_syuka_tei"
	strSql = strSql & " where SND_YMD = '20120528'"
	strSql = strSql & "   and TOK_CD = '7868HB'"
	strSql = strSql & "   and CHO_CD = '051'"
	strSql = strSql & "   and CHU_CD = 'XFJ0746'"
	GetOpenSql = strSql
End Function
