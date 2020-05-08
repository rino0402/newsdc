Option Explicit
' ITEM.DAT サイズセット
' 2010.06.07 新規作成

function usage()
	Wscript.Echo "ITEMサイズ設定(2010.06.07)"
	Wscript.Echo "itemsize.vbs <database> [<仕向先>] [option]"
	Wscript.Echo " -?"
end function

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

dim	dbObj
dim	dbName
dim	sqlStr
dim	rsList
dim	rsPCompoK
dim	f		' Field Object

dim	strBuff

dim	lngCnt		' 登録件数
dim	i
dim	lngCntTest	' TEST登録件数
dim	strJgyobu
dim	strAdd
dim	strHinGai
dim	strAction
dim	strOption
dim	strArg
dim	strShimuke
dim	strMode		' -list , -update
dim	strSaisu
dim	strSizeW
dim	strSizeD
dim	strSizeH
dim	lngRecordQty
dim	lngTimeout

dbName		= ""
strShimuke	= ""
strMode		= "-list"

for i = 0 to WScript.Arguments.count - 1
	strArg = WScript.Arguments(i)
    select case strArg
	case "-?"
		call usage()
		WScript.Quit
	case "-list"
		strMode = strArg
	case "-update"
		strMode = strArg
	case "-sql"
		strMode = strArg
	case else
		if dbName = "" then
			dbName = strArg
		elseif strShimuke = "" then
			strShimuke = strArg
		else
			usage()
			WScript.Quit
		end if
	end select
next

if dbName = "" then
	call usage()
	WScript.Quit
end if

Wscript.Echo "open db : " & dbName
Set dbObj = Wscript.CreateObject("ADODB.Connection")
dbObj.open dbName

sqlStr = "select *"
sqlStr = sqlStr & " from item"
sqlStr = sqlStr & " where jgyobu = 'S'" 
sqlStr = sqlStr & "   and naigai = '1'"
sqlStr = sqlStr & "   and convert(SAI_SU,sql_numeric) <> 0"

Wscript.Echo "sql : " & sqlStr
set rsList = dbObj.Execute(sqlStr)
Wscript.Echo "sql : 完了"

'On Error Resume Next
lngCnt	= 0
Do While Not rsList.EOF
	lngCnt	= lngCnt + 1
	strHinGai = rtrim(rsList.Fields("HIN_GAI"))
	strSaisu = rtrim(rsList.Fields("SAI_SU"))
	strSizeW = rtrim(rsList.Fields("D_SIZE_W"))
	strSizeD = rtrim(rsList.Fields("D_SIZE_D"))
	strSizeH = rtrim(rsList.Fields("D_SIZE_H"))
	strBuff = "  " & rsList.Fields("JGYOBU")
	strBuff = strBuff & " " & rsList.Fields("NAIGAI")
	strBuff = strBuff & " " & rsList.Fields("HIN_GAI")
	strBuff = strBuff & " " & rsList.Fields("SAI_SU")
	strBuff = strBuff & " " & rsList.Fields("D_SIZE_W")
	strBuff = strBuff & " " & rsList.Fields("D_SIZE_D")
	strBuff = strBuff & " " & rsList.Fields("D_SIZE_H")
	Wscript.Echo strBuff
	if strShimuke <> "" then
		select case strMode
		case "-list"
			sqlStr = "select *"
			sqlStr = sqlStr & " from item"
		case "-update","-sql"
			sqlStr = "update item"
			sqlStr = sqlStr & " set SAI_SU = '" & strSaisu & "'"
			sqlStr = sqlStr & "   , D_SIZE_W = '" & strSizeW & "'"
			sqlStr = sqlStr & "   , D_SIZE_D = '" & strSizeD & "'"
			sqlStr = sqlStr & "   , D_SIZE_H = '" & strSizeH & "'"
		end select
		sqlStr = sqlStr & " where JGYOBU in ("
		sqlStr = sqlStr & 		" select distinct OPTION1 From P_CODE where DATA_KBN = '04' and C_Code = '" & strShimuke & "'"
		sqlStr = sqlStr & 		")"
		sqlStr = sqlStr & "   and NAIGAI = '1'"
		sqlStr = sqlStr & "   and HIN_GAI in ("
		sqlStr = sqlStr & 		" select distinct hin_gai"
		sqlStr = sqlStr & 		" from p_compo_k"
		sqlStr = sqlStr &		" where shimuke_code = '" & strShimuke & "'" 
		sqlStr = sqlStr &		"   and DATA_KBN = '1'"
		sqlStr = sqlStr &		"   and SEQNO = '010'"
		sqlStr = sqlStr &		"   and KO_JGYOBU = 'S'"
		sqlStr = sqlStr &		"   and KO_NAIGAI = '1'"
		sqlStr = sqlStr &		"   and KO_HIN_GAI = '" & strHinGai & "'"
		sqlStr = sqlStr & 		")"
		if strMode = "-list" then
			set rsPCompoK = dbObj.Execute(sqlStr,lngRecordQty)
			Do While Not rsPCompoK.EOF
				strBuff = "→" & rsPCompoK.Fields("JGYOBU")
				strBuff = strBuff & " " & rsPCompoK.Fields("NAIGAI")
				strBuff = strBuff & " " & rsPCompoK.Fields("HIN_GAI")
				strBuff = strBuff & " " & rsPCompoK.Fields("SAI_SU")
				strBuff = strBuff & " " & rsPCompoK.Fields("D_SIZE_W")
				strBuff = strBuff & " " & rsPCompoK.Fields("D_SIZE_D")
				strBuff = strBuff & " " & rsPCompoK.Fields("D_SIZE_H")
				Wscript.Echo strBuff
				rsPCompoK.movenext
			loop
		else
			if strMode = "-update" then
				lngTimeout = dbObj.CommandTimeout
				dbObj.CommandTimeout = 0
				call dbObj.Execute(sqlStr,lngRecordQty)
				Wscript.Echo "更新→ " & lngRecordQty & " 件"
				dbObj.CommandTimeout = lngTimeout
			else
				Wscript.Echo sqlStr & " #"
			end if
		end if
	end if
	rsList.movenext
Loop
rsList.close

Wscript.Echo "close db : " & dbName
dbObj.Close
set dbObj = nothing

Wscript.Echo "end"
