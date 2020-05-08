Option Explicit
' ITEM.DAT 検品メッセージセット
' 2010.07.29 新規作成

function usage()
	Wscript.Echo "ITEM 検品メッセージセット(2010.07.29)"
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
dim	rsGensan
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
dim	strInspMessage
dim	strSizeW
dim	strSizeD
dim	strSizeH
dim	lngRecordQty
dim	lngTimeout
dim	strGensankoku

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
sqlStr = sqlStr & ",HIN_GAI"
sqlStr = sqlStr & ",gensankoku"
sqlStr = sqlStr & " from zaiko"
sqlStr = sqlStr & " where JGYOBU = '4'" 
sqlStr = sqlStr & "   and NAIGAI = '1'" 
sqlStr = sqlStr & "   and gensankoku = ''" 

Wscript.Echo "sql : " & sqlStr
set rsList = dbObj.Execute(sqlStr)
Wscript.Echo "sql : 完了"

'On Error Resume Next
lngCnt	= 0
Do While Not rsList.EOF
	lngCnt	= lngCnt + 1
	strBuff = rsList.Fields("JGYOBU")
	strBuff = strBuff & " " & rsList.Fields("NAIGAI")
	strBuff = strBuff & " " & rsList.Fields("HIN_GAI")
	strBuff = strBuff & " " & rsList.Fields("gensankoku")

	sqlStr = "select "
	sqlStr = sqlStr & " MadeInCode,CountryName2"
	sqlStr = sqlStr & " from Pn3"
	sqlStr = sqlStr & " left outer join country on (MadeInCode = CountryCode)" 
	sqlStr = sqlStr & " where ShisanJCode = '00023410'" 
	sqlStr = sqlStr & "   and Pn3.Pn  = '" & rtrim(rsList.Fields("HIN_GAI")) & "'" 

	set rsGensan = dbObj.Execute(sqlStr)
	strGensankoku = ""
	if rsGensan.eof = false then
		strGensankoku = rtrim(rsGensan.Fields("CountryName2"))
		strBuff = strBuff & " " & rsGensan.Fields("CountryName2")
		strBuff = strBuff & " " & rsGensan.Fields("MadeInCode")
	else
		strBuff = strBuff & " ???"
	end if
	Wscript.Echo strBuff

	select case strMode
	case "-update","-sql"
		sqlStr = "update zaiko"
		sqlStr = sqlStr & " set gensankoku = '" 		& strGensankoku & "'"
	end select
	sqlStr = sqlStr & " where JGYOBU = '" & rsList.Fields("JGYOBU") & "'"
	sqlStr = sqlStr & "   and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
	sqlStr = sqlStr & "   and HIN_GAI = '" & rtrim(rsList.Fields("HIN_GAI")) & "'"
	sqlStr = sqlStr & "   and gensankoku = ''" 
	if strMode = "-list" then
	else
		if strMode = "-update" then
			call dbObj.Execute(sqlStr,lngRecordQty)
			Wscript.Echo "更新→ " & lngRecordQty & " 件"
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
