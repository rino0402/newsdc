Option Explicit
' ITEM.DAT ibZ[WZbg
' 2010.07.29 VKì¬

function usage()
	Wscript.Echo "ITEM ibZ[WZbg(2010.07.29)"
	Wscript.Echo "item-insp-message.vbs <database> [option]"
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

dim	lngCnt		' o^
dim	i
dim	lngCntTest	' TESTo^
dim	strJgyobu
dim	strAdd
dim	strHinGai
dim	strAction
dim	strOption
dim	strArg
dim	strShimuke
dim	strMode		' -list , -update
dim	strInspMessage
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

sqlStr = "select "
sqlStr = sqlStr & " JGYOBU"
sqlStr = sqlStr & ",NAIGAI"
sqlStr = sqlStr & ",HIN_GAI"
sqlStr = sqlStr & ",INSP_MESSAGE"
sqlStr = sqlStr & ",KANKYO_KBN"
sqlStr = sqlStr & ",KANKYO_KBN_ST"
sqlStr = sqlStr & ",KANKYO_KBN_SURYO"
sqlStr = sqlStr & " from item"
sqlStr = sqlStr & " where KANKYO_KBN <> ''" 

Wscript.Echo "sql : " & sqlStr
set rsList = dbObj.Execute(sqlStr)
Wscript.Echo "sql : ®¹"

'On Error Resume Next
lngCnt	= 0
Do While Not rsList.EOF
	lngCnt	= lngCnt + 1
	strBuff = rsList.Fields("JGYOBU")
	strBuff = strBuff & " " & rsList.Fields("NAIGAI")
	strBuff = strBuff & " " & rsList.Fields("HIN_GAI")
	strBuff = strBuff & " " & rsList.Fields("KANKYO_KBN")
	strBuff = strBuff & " " & rsList.Fields("KANKYO_KBN_ST")
	strBuff = strBuff & " " & rsList.Fields("KANKYO_KBN_SURYO")
	strBuff = strBuff & " " & rsList.Fields("INSP_MESSAGE")
	strInspMessage = ""
	if rsList.Fields("KANKYO_KBN") = "LIT" then
		strInspMessage = "`EdrÚ(" & trim(rsList.Fields("KANKYO_KBN_SURYO")) & ")"
		if rsList.Fields("INSP_MESSAGE") <> strInspMessage then
			strBuff = strBuff & "¨" & strInspMessage
		end if
	end if
	Wscript.Echo strBuff
	select case strMode
	case "-update","-sql"
		sqlStr = "update item"
		sqlStr = sqlStr & " set KANKYO_KBN = '" 		& rsList.Fields("KANKYO_KBN") 		& "'"
		sqlStr = sqlStr & "   , KANKYO_KBN_ST = '" 		& rsList.Fields("KANKYO_KBN_ST") 	& "'"
		sqlStr = sqlStr & "   , KANKYO_KBN_SURYO = '" 	& rsList.Fields("KANKYO_KBN_SURYO") & "'"
	end select
	sqlStr = sqlStr & " where JGYOBU = '" & rsList.Fields("JGYOBU") & "'"
	sqlStr = sqlStr & "   and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
	sqlStr = sqlStr & "   and HIN_GAI = '" & rsList.Fields("HIN_GAI") & "'"
	if strMode = "-list" then
	else
		if strMode = "-update" then
			call dbObj.Execute(sqlStr,lngRecordQty)
			Wscript.Echo "XV¨ " & lngRecordQty & " "
		else
			Wscript.Echo sqlStr & " #"
		end if
	end if
	rsList.movenext
Loop
rsList.close

Wscript.Echo "close db : " & dbName
dbObj.Close
set dbObj = nothing

Wscript.Echo "end"
