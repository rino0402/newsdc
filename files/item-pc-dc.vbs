Option Explicit
' ITEM.DAT移管プログラム
' 2010.03.24 新規作成
' 2011.07.15 item-pc-dc.vbs

Const adStateClosed = 0

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

dim	dbSrc		' コピー元 DB Object
dim	dbSrcName	' コピー元 Name
dim	sqlStr
dim	rsSrcList
dim	f		' Field Object

dim	dbDst		' コピー先 DB Object
dim	dbDstName	' コピー先 DB Name
dim	rsDstItem	' コピー先 レコードセット(ITEM)
dim	rsFind		' レコードセット(ITEM)

dim	strBuff
dim	Fs
dim	logFile

dim	lngCnt		' 登録件数
dim	i
dim	lngCntTest	' TEST登録件数
dim	strJgyobu
dim	strAdd
dim	strHinGai
dim	strAction
dim	strOption
dim	strArg
dim	strMode

lngCntTest	= 0
strJgyobu 	= "R"
strHinGai	= ""
strOption	= ""
strMode		= "conv"
for i = 0 to WScript.Arguments.count - 1
	strArg = WScript.Arguments(i)
	select case strOption
	case "-pn"
		strHinGai 	= strArg
		strOption	= ""
	case else
		strArg = lcase(strArg)
		strOption	= ""
	    	select case strArg
	    	case "-test"
	    		lngCntTest	= 10
	    	case "-pn"
			strOption = strArg
	    	case "-check"
			strMode		= "check"
		case else
	        	Wscript.Echo "ITEM移管(2011.07.15)"
	        	Wscript.Echo "item-pc-dc.vbs [option]"
	        	Wscript.Echo " -?"
	        	Wscript.Echo " -test : 10件だけ登録"
	        	Wscript.Echo GetDateTime(now())
			WScript.Quit
		end select
	end select
next

Wscript.Echo "item-pc-dc.vbs "
Wscript.Echo "ITEM移管(滋賀PC→滋賀物流)"

dbSrcName	= "newsdc-sig"
dbDstName	= "newsdc"

Set dbSrc = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbSrcName
dbSrc.open dbSrcName

Set dbDst = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbDstName
dbDst.open dbDstName

sqlStr = "select *"
sqlStr = sqlStr & " from item"
sqlStr = sqlStr & " where jgyobu = '" & strJgyobu & "'" 
sqlStr = sqlStr & "   and naigai = '1'"
if strHinGai <> "" then
	sqlStr = sqlStr & "   and hin_gai like '" & strHinGai & "'"
end if
sqlStr = sqlStr & " order by Ins_DateTime desc"

Wscript.Echo "db.CommandTimeout : " & dbSrc.CommandTimeout
dbSrc.CommandTimeout = 0
Wscript.Echo "db.CommandTimeout : " & dbSrc.CommandTimeout
Wscript.Echo "sql : " & sqlStr

' set rsList = db.Execute(sqlStr)
Set rsSrcList = Wscript.CreateObject("ADODB.Recordset")
'rsList.Open sqlStr, db, adOpenDynamic, adLockOptimistic
rsSrcList.Open sqlStr, dbSrc, adOpenForwardOnly, adLockReadOnly

Wscript.Echo "sql : 完了"


' コピー先 ITEMオープン
Wscript.Echo "コピー先ITEM : Open"
Set rsDstItem = Wscript.CreateObject("ADODB.Recordset")
rsDstItem.MaxRecords = 1
rsDstItem.CursorLocation = adUseServer

lngCnt	= 0
Do While Not rsSrcList.EOF
	lngCnt	= lngCnt + 1
	if lngCntTest > 0 then
		if lngCnt > 10 then
			exit do
		end if
	end if
	strAdd = ""
	strHinGai = rtrim(rsSrcList.Fields("HIN_GAI"))
	strBuff = rsSrcList.Fields("JGYOBU")
	strBuff = strBuff & " " & rsSrcList.Fields("NAIGAI")
	strBuff = strBuff & " " & strHinGai
	strBuff = strBuff & " " & rsSrcList.Fields("HIN_NAME")
	strBuff = strBuff & " " & rsSrcList.Fields("G_SYUSHI")

	strAction = ""
	sqlStr = "select *"
	sqlStr = sqlStr & " from item"
	sqlStr = sqlStr & " where jgyobu = '" & rsSrcList.Fields("JGYOBU") & "'" 
	sqlStr = sqlStr & "   and naigai = '" & rsSrcList.Fields("NAIGAI") & "'" 
	sqlStr = sqlStr & "   and HIN_GAI = '" & strHinGai & "'"

	if strMode = "check" then
		set rsDstItem = dbDst.Execute(sqlStr)
		if rsDstItem.EOF = true then
			Wscript.Echo lngCnt & " " & strBuff & ":not found"
		else
			if rtrim(rsDstItem.Fields("SE_TANKA_MEMO")) <> rtrim(rsSrcList.Fields("SE_TANKA_MEMO")) then
				if rtrim(rsDstItem.Fields("UPD_DATETIME")) <= rtrim(rsSrcList.Fields("UPD_DATETIME")) then
					strBuff = rsSrcList.Fields("JGYOBU")
					strBuff = strBuff & " " & rsSrcList.Fields("NAIGAI")
					strBuff = strBuff & " " & strHinGai
					strBuff = strBuff & " " & rsSrcList.Fields("UPD_DATETIME")
					strBuff = strBuff & " " & rsSrcList.Fields("SE_TANKA_MEMO")
					Wscript.Echo lngCnt & " " & strBuff
					strBuff = rsDstItem.Fields("JGYOBU")
					strBuff = strBuff & " " & rsDstItem.Fields("NAIGAI")
					strBuff = strBuff & " " & strHinGai
					strBuff = strBuff & " " & rsDstItem.Fields("UPD_DATETIME")
					strBuff = strBuff & " " & rsDstItem.Fields("SE_TANKA_MEMO")
					Wscript.Echo lngCnt & " " & strBuff
					sqlStr = "update item"
					sqlStr = sqlStr & " set SE_TANKA_MEMO = '" & rtrim(rsSrcList.Fields("SE_TANKA_MEMO")) & "'"
					sqlStr = sqlStr & " where jgyobu = '" & rsSrcList.Fields("JGYOBU") & "'" 
					sqlStr = sqlStr & "   and naigai = '" & rsSrcList.Fields("NAIGAI") & "'" 
					sqlStr = sqlStr & "   and HIN_GAI = '" & strHinGai & "'"
					call dbDst.Execute(sqlStr)
				end if
			end if

'			for each f in rsSrcList.Fields
'				case "SE_TANKA_MEMO"
'
'				if rtrim(rsDstItem.Fields(f.Name)) <> rtrim(f) then
'					if strBuff <> "" then
'						Wscript.Echo lngCnt & " " & strBuff & ":diff"
'						strBuff = ""
'					end if
'					Wscript.Echo f.Name & "=" & rtrim(f)
'					Wscript.Echo f.Name & "=" & rtrim(rsDstItem.Fields(f.Name))
'				end if
'			next
'			if strBuff <> "" then
'				Wscript.Echo lngCnt & " " & strBuff & ":ok"
'			end if
		end if
	else
		if rsDstItem.state <> adStateClosed then
			rsDstItem.Close
		end if
		rsDstItem.Open sqlStr, dbDst, adOpenKeyset, adLockBatchOptimistic
		if rsDstItem.EOF = true then
			if rsDstItem.state <> adStateClosed then
				rsDstItem.Close
			end if
			rsDstItem.Open "ITEM", dbDst, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
			rsDstItem.Addnew
			rsDstItem.Fields("JGYOBU") = rsSrcList.Fields("JGYOBU")
			rsDstItem.Fields("NAIGAI") = rsSrcList.Fields("NAIGAI")
			rsDstItem.Fields("HIN_GAI") = strHinGai
			strAction = "Add"
		else
			strAction = "Upd"
		end if
		for each f in rsSrcList.Fields
	'		Wscript.Echo f.Name & "=" & rtrim(f) & "(" & asc(f) & ")"
			select case ucase(f.Name)
			case "JGYOBU"
			case "NAIGAI"
			case "HIN_GAI"
			case "SE_TANKA_MEMO"
			case else
				rsDstItem.Fields(f.Name) = f.value
			end select
		next
		On Error Resume Next
		rsDstItem.UpdateBatch
		if Err.Number = 0 then
			strBuff = strAction & "  Ok:" & strBuff
			strBuff = lngCnt & " " & strBuff
			Wscript.Echo strBuff
		else
			strBuff = strAction & " Err:" & strBuff
			strBuff = lngCnt & " " & strBuff
			Wscript.Echo strBuff
			Wscript.Echo "Err.Number:" & Err.Number & " " & Err.Description
		end if
		Err.Clear
		On Error Goto 0
	end if
	rsSrcList.movenext
Loop

Wscript.Echo ""
rsSrcList.close

Wscript.Echo "close db : " & dbSrcName
dbSrc.Close
set dbSrc = nothing

Wscript.Echo "close rsDstItem"
rsDstItem.Close
set rsDstItem = nothing

Wscript.Echo "close db : " & dbDstName
dbDst.Close
set dbDst = nothing

Wscript.Echo "end"
