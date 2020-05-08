Option Explicit
' P_CLASS.DAT移管プログラム
' 2010.03.31 新規作成

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

dim	dbSrc
dim	dbSrcName
dim	rsSrcPClass
dim	f		' Field Object
dim	sqlStr

dim	dbDst		' コピー先 DB Object
dim	dbDstName	' コピー先 DB Name
dim	rsDstPClass	' コピー先 レコードセット(ITEM)
dim	rsFind		' レコードセット(ITEM)

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

lngCntTest	= 0
strJgyobu 	= "R"
strHinGai	= ""
strOption	= ""
for i = 0 to WScript.Arguments.count - 1
	strArg = WScript.Arguments(i)
    select case strOption
'    case "-pn"
'		strHinGai 	= strArg
'		strOption	= ""
    case else
		strArg = lcase(strArg)
		strOption	= ""
	    select case strArg
	    case "-test"
	        lngCntTest	= 10
'	    case "-s"
'			strJgyobu 	= "S"
'	    case "-pn"
'			strOption = strArg
	    case else
	        Wscript.Echo "P_CLASS移管(2010.03.31)"
	        Wscript.Echo "p_class.vbs [option]"
	        Wscript.Echo " -?"
	        Wscript.Echo " -test : 10件だけ登録"
			WScript.Quit
	    end select
    end select
next

Wscript.Echo "p_class.vbs "
Wscript.Echo "P_CLASS移管(草津SC→滋賀PC)"

dbSrcName		= "newsdc-kst"
dbDstName	= "newsdc-sig"

Set dbSrc = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbSrcName
dbSrc.open dbSrcName

Set dbDst = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbDstName
dbDst.open dbDstName

sqlStr = "select *"
sqlStr = sqlStr & " from p_class"
sqlStr = sqlStr & " where SHIMUKE_CODE = '02'" 

Wscript.Echo "sql : " & sqlStr
Set rsSrcPClass = Wscript.CreateObject("ADODB.Recordset")
rsSrcPClass.Open sqlStr, dbSrc, adOpenForwardOnly, adLockReadOnly
Wscript.Echo "sql : 完了"

' コピー先 ITEMオープン
Wscript.Echo "コピー先P_CLASS : Open"
Set rsDstPClass = Wscript.CreateObject("ADODB.Recordset")
rsDstPClass.MaxRecords = 1
rsDstPClass.CursorLocation = adUseServer
rsDstPClass.Open "P_CLASS", dbDst, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

On Error Resume Next
lngCnt	= 0
Do While Not rsSrcPClass.EOF
	lngCnt	= lngCnt + 1
	if lngCntTest > 0 then
		if lngCnt > 10 then
			exit do
		end if
	end if
	strAdd = ""

	strBuff = rsSrcPClass.Fields("SHIMUKE_CODE")
	strBuff = strBuff & " " & rsSrcPClass.Fields("CLASS_CODE")
	strBuff = strBuff & " " & rsSrcPClass.Fields("CLASS_NAME")

	strAction = ""
	rsDstPClass.Addnew
	strAction = "Add"
	Err.Clear

	if strAction = "Add" then
		for each f in rsSrcPClass.Fields
			select case ucase(f.Name)
			case "SHIMUKE_CODE"
				select case f
					case "02"
						rsDstPClass.Fields(f.Name) = "04"
					case else
						rsDstPClass.Fields(f.Name) = "99"
				end select
			case else
				rsDstPClass.Fields(f.Name) = f
			end select
		next
		rsDstPClass.UpdateBatch
	end if
	if Err.Number = 0 then
		strBuff = strAction & "  Ok:" & strBuff
		Wscript.Echo strBuff
	else
		strBuff = strAction & " Err:" & strBuff
		Wscript.Echo strBuff
		Wscript.Echo "Err.Number:" & Err.Number
	end if
	Err.Clear
	rsSrcPClass.movenext
Loop

Wscript.Echo ""
rsSrcPClass.close

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
