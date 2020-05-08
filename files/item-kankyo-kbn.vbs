Option Explicit
' ITEM.DAT ä¬ã´ãÊï™ÉZÉbÉg
' 2010.06.07 êVãKçÏê¨

function usage()
	Wscript.Echo "ITEMä¬ã´ãÊï™ÉZÉbÉg(2010.07.27)"
	Wscript.Echo "item-kanryo-kbn.vbs <database> [option]"
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

dim	lngCnt		' ìoò^åèêî
dim	i
dim	lngCntTest	' TESTìoò^åèêî
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

sqlStr = "select distinct"
sqlStr = sqlStr & " JGYOBU"
sqlStr = sqlStr & ",NAIGAI"
sqlStr = sqlStr & ",HIN_NO"
sqlStr = sqlStr & ",KANKYO_KBN"
sqlStr = sqlStr & ",KANKYO_KBN_ST"
sqlStr = sqlStr & ",KANKYO_KBN_SURYO"
sqlStr = sqlStr & " from y_glics"
sqlStr = sqlStr & " where jgyobu = '4'" 
sqlStr = sqlStr & "   and naigai = '1'"
sqlStr = sqlStr & "   and KANKYO_KBN <> ''"
sqlStr = sqlStr & "   and SYUKA_YMD > '20100701'"

Wscript.Echo "sql : " & sqlStr
set rsList = dbObj.Execute(sqlStr)
Wscript.Echo "sql : äÆóπ"

'On Error Resume Next
lngCnt	= 0
Do While Not rsList.EOF
	lngCnt	= lngCnt + 1
	strBuff = "  " & rsList.Fields("JGYOBU")
	strBuff = strBuff & " " & rsList.Fields("NAIGAI")
	strBuff = strBuff & " " & rsList.Fields("HIN_NO")
	strBuff = strBuff & " " & rsList.Fields("KANKYO_KBN")
	strBuff = strBuff & " " & rsList.Fields("KANKYO_KBN_ST")
	strBuff = strBuff & " " & rsList.Fields("KANKYO_KBN_SURYO")
	Wscript.Echo strBuff
	select case strMode
	case "-list"
		sqlStr = "select *"
		sqlStr = sqlStr & " from item"
	case "-update","-sql"
		sqlStr = "update item"
		sqlStr = sqlStr & " set KANKYO_KBN = '" 		& rsList.Fields("KANKYO_KBN") 		& "'"
		sqlStr = sqlStr & "   , KANKYO_KBN_ST = '" 		& rsList.Fields("KANKYO_KBN_ST") 	& "'"
		sqlStr = sqlStr & "   , KANKYO_KBN_SURYO = '" 	& rsList.Fields("KANKYO_KBN_SURYO") & "'"
	end select
	sqlStr = sqlStr & " where JGYOBU = '" & rsList.Fields("JGYOBU") & "'"
	sqlStr = sqlStr & "   and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
	sqlStr = sqlStr & "   and HIN_GAI = '" & rsList.Fields("HIN_NO") & "'"
	if strMode = "-list" then
		set rsPCompoK = dbObj.Execute(sqlStr,lngRecordQty)
		Do While Not rsPCompoK.EOF
			strBuff = "Å®" & rsPCompoK.Fields("JGYOBU")
			strBuff = strBuff & " " & rsPCompoK.Fields("NAIGAI")
			strBuff = strBuff & " " & rsPCompoK.Fields("HIN_GAI")
			strBuff = strBuff & " " & rsPCompoK.Fields("KANKYO_KBN")
			strBuff = strBuff & " " & rsPCompoK.Fields("KANKYO_KBN_ST")
			strBuff = strBuff & " " & rsPCompoK.Fields("KANKYO_KBN_SURYO")
			Wscript.Echo strBuff
			rsPCompoK.movenext
		loop
	else
		if strMode = "-update" then
			call dbObj.Execute(sqlStr,lngRecordQty)
			Wscript.Echo "çXêVÅ® " & lngRecordQty & " åè"
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
