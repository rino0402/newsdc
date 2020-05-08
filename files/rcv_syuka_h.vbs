Option Explicit
' del_syuka_h 出荷予定データ復旧
' 2011.07.13 新規作成
' 2011.07.21 
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
	Wscript.Echo "出荷予定データ復旧処理"
	Wscript.Echo "del_syuka.vbs [option] <database> <出荷日> <便> <送先コード>"
	Wscript.Echo " -?"
	Wscript.Echo " -move  : del_syuka_h から y_syuka_h へ移動する"
	Wscript.Echo " -disp  : 表示のみ(デフォルト)"
	Wscript.Echo " -debug : デバッグメッセージ表示"
end function

Sub Main()
	dim	i
	dim	strArg
	dim	strDbName
	dim	strSYUKA_YMD
	dim	strINS_BIN
	dim	strOKURISAKI_CD
	dim	strMove
	dim	strDebug

	strDbName		= ""
	strSYUKA_YMD	= ""
	strINS_BIN		= ""
	strOKURISAKI_CD	= ""
	strMove			= "disp"
	strDebug		= ""

	for i = 0 to WScript.Arguments.count - 1
		strArg = WScript.Arguments(i)
	    select case left(strArg,1)
		case "-"
		    select case strArg
			case "-?"
				call usage()
				exit sub
			case "-disp"
				strMove		= "disp"
			case "-move"
				strMove		= "move"
			case "-debug"
				strDebug	= "debug"
			case else
				call usage()
				exit sub
			end select
		case else
			if strDbName = "" then
				strDbName = strArg
			elseif strSYUKA_YMD = "" then
				strSYUKA_YMD = strArg
			elseif strINS_BIN = "" then
				strINS_BIN = strArg
			elseif strOKURISAKI_CD = "" then
				strOKURISAKI_CD = strArg
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
	if strSYUKA_YMD = "" then
		usage()
		exit sub
	end if
	if strINS_BIN = "" then
		usage()
		exit sub
	end if
	if strOKURISAKI_CD = "" then
		usage()
		exit sub
	end if
	call delSyukaRecober(strDbName,strSYUKA_YMD,strINS_BIN,strOKURISAKI_CD,strMove,strDebug)
End Sub

Sub delSyukaRecober(strDbName,strSYUKA_YMD,strINS_BIN,strOKURISAKI_CD,strMove,strDebug)
	dim	objDb
	dim	rsY_Syuka_H
	dim	rsDel_Syuka_H
	dim	fldDel_Syuka_H
	dim	strSql
	dim	strMsg
	dim	rsY_Syuka
	dim	rsDel_Syuka
	dim	fldDel_Syuka
	dim	strKEY_SYUKA_YMD
	dim	strKEY_ID_NO
	dim	i

	call DispMsg("CreateObject(ADODB.Connection)",strDebug)
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	
	call DispMsg("Open:" & strDbName,strDebug)
	objDb.Open strDbName
	call DispMsg(objDb.ConnectionString,strDebug)

	call DispMsg("CreateObject(ADODB.Recordset)",strDebug)
	Set rsY_Syuka_H = Wscript.CreateObject("ADODB.Recordset")
	rsY_Syuka_H.MaxRecords = 1
	rsY_Syuka_H.CursorLocation = adUseServer
	call DispMsg("open:y_syuka_h",strDebug)
	rsY_Syuka_H.Open "y_syuka_h", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

	strSql = GetOpenSql(strSYUKA_YMD,strINS_BIN,strOKURISAKI_CD)
	call DispMsg("Execute:" & strSql,strDebug)
	set rsDel_Syuka_H = objDb.Execute(strSql)
	call DispMsg("Execute:End",strDebug)

	call DispMsg("CreateObject(ADODB.Recordset)",strDebug)
	Set rsY_Syuka = Wscript.CreateObject("ADODB.Recordset")
	rsY_Syuka.MaxRecords = 1
	rsY_Syuka.CursorLocation = adUseServer
	call DispMsg("open:y_syuka",strDebug)
	rsY_Syuka.Open "y_syuka", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

	Do While Not rsDel_Syuka_H.EOF
'		call DispMsg(rsBak.Fields("KEY_ID_NO"))
		if strMove = "move" then
			rsY_Syuka_H.Addnew
		end if
		for each fldDel_Syuka_H in rsDel_Syuka_H.Fields
			call DispMsg(fldDel_Syuka_H.Name & ":" & fldDel_Syuka_H,strDebug)
			if strMove = "move" then
				rsY_Syuka_H.Fields(fldDel_Syuka_H.Name) = fldDel_Syuka_H
			else
			end if
		next
		On Error Resume Next
		if strMove = "move" then
			rsY_Syuka_H.UpdateBatch
			if Err.Number = 0 then
				strMsg = "y_syuka_h:Ok"
			else
				strMsg = "y_syuka_h:Err:" & Err.Number & " " & Err.Description
			end if
		else
			strMsg = "y_syuka_h:Disp"
		end if
		Err.Clear
		On Error Goto 0
		if strMove = "move" then
			call DispMsg(rsY_Syuka_H.Fields("SYUKA_YMD") & " " & rsY_Syuka_H.Fields("INS_BIN") & " " & rsY_Syuka_H.Fields("OKURISAKI_CD") & " " & rsY_Syuka_H.Fields("ID_NO") & ":" & strMsg,"info")
		else
			call DispMsg(rsDel_Syuka_H.Fields("SYUKA_YMD") & " " & rsDel_Syuka_H.Fields("INS_BIN") & " " & rsDel_Syuka_H.Fields("OKURISAKI_CD") & " " & rsDel_Syuka_H.Fields("ID_NO") & ":" & strMsg,"info")
		end if
		strKEY_SYUKA_YMD	= rtrim(rsDel_Syuka_H.Fields("SYUKA_YMD"))
		strKEY_ID_NO		= rtrim(rsDel_Syuka_H.Fields("ID_NO"))
		strSql = GetDelSyukaSql(strKEY_SYUKA_YMD,strKEY_ID_NO)
		call DispMsg("Execute:" & strSql,strDebug)
		set rsDel_Syuka = objDb.Execute(strSql)
		call DispMsg("Execute:End",strDebug)
		if rsDel_Syuka.EOF then
			call DispMsg(strKEY_SYUKA_YMD & " " & strKEY_ID_NO & ":not found","error")
		else
			i = 0
			do while not rsDel_Syuka.EOF
				i = i + 1
				if i = 1 then
					call DispMsg(strKEY_SYUKA_YMD & " " & strKEY_ID_NO & ":OK",strDebug)
					if strMove = "move" then
						rsY_Syuka.Addnew
					end if
					for each fldDel_Syuka in rsDel_Syuka.Fields
						call DispMsg(fldDel_Syuka.Name & ":" & fldDel_Syuka,strDebug)
						if strMove = "move" then
							rsY_Syuka.Fields(fldDel_Syuka.Name) = fldDel_Syuka
						end if
					next
					On Error Resume Next
					if strMove = "move" then
						rsY_Syuka.Update
						if Err.Number = 0 then
							strMsg = "y_syuka  :Ok"
						else
							strMsg = "y_syuka  :Err:" & Err.Number & " " & Err.Description
						end if
					else
						strMsg = "y_syuka  :Disp"
					end if
					Err.Clear
					On Error Goto 0
					if strMove = "move" then
						call DispMsg(rsY_Syuka.Fields("KEY_SYUKA_YMD") & "    " & rsY_Syuka.Fields("KEY_MUKE_CODE") & "  " & rsY_Syuka.Fields("KEY_ID_NO") & ":" & strMsg,"info")
					else
						call DispMsg(rsDel_Syuka.Fields("KEY_SYUKA_YMD") & "    " & rsDel_Syuka.Fields("KEY_MUKE_CODE") & "  " & rsDel_Syuka.Fields("KEY_ID_NO") & ":" & strMsg,"info")
					end if
				else
					call DispMsg(strKEY_SYUKA_YMD & " " & strKEY_ID_NO & ":" & i,"error")
				end if
				rsDel_Syuka.MoveNext
			loop
		end if
		rsDel_Syuka.Close

		rsDel_Syuka_H.MoveNext
	Loop

	call DispMsg("Close:del_syuka_h",strDebug)
	rsDel_Syuka_H.Close
	set rsDel_Syuka_H = Nothing
	
	call DispMsg("Close:y_syuka_h",strDebug)
	rsY_Syuka_H.Close
	set rsY_Syuka_H = Nothing

	call DispMsg("Close:y_syuka",strDebug)
	rsY_Syuka.Close
	set rsY_Syuka = Nothing

	call DispMsg("Close:" & strDbName,strDebug)
	objDb.Close
	set objDb = Nothing
End Sub

Sub DispMsg(strMsg,strM)
	select case strM
	case ""
	case "debug"
		Wscript.Echo strM & " " & strMsg
	case "info"
		Wscript.Echo strMsg
	case else
		Wscript.Echo strM & " " & strMsg
	end select
End Sub

Function GetOpenSql(strSYUKA_YMD,strINS_BIN,strOKURISAKI_CD)
	dim	strSql

	strSql = "select *"
	strSql = strSql & " from del_syuka_h"
	strSql = strSql & " where SYUKA_YMD = '"	& strSYUKA_YMD		& "'"
	strSql = strSql & "   and INS_BIN = '"		& strINS_BIN		& "'"
	strSql = strSql & "   and OKURISAKI_CD = '"	& strOKURISAKI_CD	& "'"

	GetOpenSql = strSql
End Function

Function GetDelSyukaSql(strKEY_SYUKA_YMD,strKEY_ID_NO)
	dim	strSql

	strSql = "select *"
	strSql = strSql & " from del_syuka"
	strSql = strSql & " where KEY_SYUKA_YMD = '"	& strKEY_SYUKA_YMD	& "'"
	strSql = strSql & "   and KEY_ID_NO = '"		& strKEY_ID_NO		& "'"

	GetDelSyukaSql = strSql
End Function

