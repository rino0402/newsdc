Option Explicit
' ITEM.DAT移管プログラム
' 2010.03.24 新規作成

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

dim	db
dim	dbName
dim	sqlStr
dim	rsList
dim	rsSrcItem
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

lngCntTest	= 0
strJgyobu 	= "R"
strHinGai	= ""
strOption	= ""
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
	    case "-s"
			strJgyobu 	= "S"
	    case "-pn"
			strOption = strArg
	    case else
	        Wscript.Echo "ITEM移管(2010.03.24)"
	        Wscript.Echo "item.vbs [option]"
	        Wscript.Echo " -?"
	        Wscript.Echo " -test : 10件だけ登録"
	        Wscript.Echo GetDateTime(now())
			WScript.Quit
	    end select
    end select
next

Wscript.Echo "item.vbs "
Wscript.Echo "ITEM移管(草津SC→滋賀PC)"

dbName		= "newsdc-kst"
dbDstName	= "newsdc-sig"

Set db = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbName
db.open dbName

Set dbDst = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbDstName
dbDst.open dbDstName

'select distinct ko_jgyobu,ko_naigai,ko_hin_gai
' from P_COMPO_K
' where shimuke_code = '04'
'   and data_kbn <> '0'
'   and ko_hin_gai not in (select distinct hin_gai from item)

sqlStr = "select distinct ko_jgyobu,ko_naigai,ko_hin_gai"
sqlStr = sqlStr & " from P_COMPO_K"
sqlStr = sqlStr & " where shimuke_code = '04'" 
sqlStr = sqlStr & "   and data_kbn <> '0'"
sqlStr = sqlStr & "   and ko_hin_gai not in (select distinct hin_gai from item)"

Wscript.Echo "sql : " & sqlStr

Set rsList = Wscript.CreateObject("ADODB.Recordset")
dbDst.CommandTimeout = 0
rsList.Open sqlStr, dbDst, adOpenForwardOnly, adLockReadOnly

Wscript.Echo "sql : 完了"


' コピー先 ITEMオープン
Wscript.Echo "コピー先ITEM : Open"
Set rsDstItem = Wscript.CreateObject("ADODB.Recordset")
rsDstItem.MaxRecords = 1
rsDstItem.CursorLocation = adUseServer
rsDstItem.Open "ITEM", dbDst, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

On Error Resume Next
lngCnt	= 0
Do While Not rsList.EOF
	lngCnt	= lngCnt + 1
	if lngCntTest > 0 then
		if lngCnt > 10 then
			exit do
		end if
	end if
	strHinGai = rsList.Fields("KO_HIN_GAI")

	sqlStr = "select *"
	sqlStr = sqlStr & " from item"
	sqlStr = sqlStr & " where jgyobu = 'S'" 
	sqlStr = sqlStr & "   and naigai = '1'"
	sqlStr = sqlStr & "   and HIN_GAI = '" & strHinGai & "'"

	set rsFind = db.Execute(sqlStr)
	if rsFind.Eof = False then
		strAction = "Add"
		strBuff = rsFind.Fields("JGYOBU")
		strBuff = strBuff & " " & rsFind.Fields("NAIGAI")
		strBuff = strBuff & " " & rsFind.Fields("HIN_GAI")
		strBuff = strBuff & " " & rsFind.Fields("G_SYUSHI")
		strBuff = strBuff & " " & rsFind.Fields("S_KOUSU")
		strBuff = strBuff & " " & rtrim(rsFind.Fields("HIN_NAME"))
	else
		strAction = "---"
		strBuff = rsList.Fields("KO_JGYOBU")
		strBuff = strBuff & " " & rsList.Fields("KO_NAIGAI")
		strBuff = strBuff & " " & strHinGai
	end if

	Err.Clear

	if strAction = "Add" then
		rsDstItem.Addnew
		for each f in rsFind.Fields
			select case ucase(f.Name)
			case "G_SYUSHI"
				if strJgyobu = "S" then
					select case f
					case "130"
						rsDstItem.Fields(f.Name) = "710"
					case "140"
						rsDstItem.Fields(f.Name) = "720"
					case "202"
						rsDstItem.Fields(f.Name) = "730"
					case "220"
						rsDstItem.Fields(f.Name) = "740"
					case "240"
						rsDstItem.Fields(f.Name) = "750"
					case else
						rsDstItem.Fields(f.Name) = f
					end select
				end if
			case "HIN_GAI"
				rsDstItem.Fields(f.Name) = strHinGai
			case "UPD_TANTO"
				rsDstItem.Fields(f.Name) = "KS2SG"
			case "UPD_DATETIME"
				rsDstItem.Fields(f.Name) = "20100405000000"
			case else
				rsDstItem.Fields(f.Name) = f
			end select
		next
		rsDstItem.UpdateBatch
	end if
	if Err.Number = 0 then
		strBuff = strAction & "  Ok:" & strBuff
		Wscript.Echo strBuff
	else
		strBuff = strAction & " Err:" & strBuff
		Wscript.Echo strBuff
		Wscript.Echo "Err.Number:" & Err.Number & " " & Err.Description
	end if
	Err.Clear
	rsList.movenext
Loop

Wscript.Echo ""
rsList.close

Wscript.Echo "close db : " & dbName
db.Close
set db = nothing

Wscript.Echo "close rsDstItem"
rsDstItem.Close
set rsDstItem = nothing

Wscript.Echo "close db : " & dbDstName
dbDst.Close
set dbDst = nothing

Wscript.Echo "end"
