Option Explicit
' P_SHORDER移管プログラム
' 2010.04.01 新規作成
' select * from p_shorder
'  where G_SYUSHI in ('130','140','202','220','240')
'    and ORDER_DT >= '20100201'
'    and KAN_F <> '1'
' delete From P_SHORDER as o where UPD_DATETIME = '20100401000000'

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

dim	dbSrc			' コピー元 DB Object
dim	dbSrcName		' コピー元 DB Name
dim	rsSrcPSHOrder	' コピー元 レコードセット
dim	f				' Field Object

dim	sqlStr			' SQL

dim	dbDst			' コピー先 DB Object
dim	dbDstName		' コピー先 DB Name
dim	rsDstPSHOrder	' コピー先 レコードセット
dim	rsFind			' 検索用 レコードセット
dim	strHinGai

dim	strBuff

dim	lngCnt			' 登録件数
dim	i
dim	lngCntTest		' TEST登録件数

lngCntTest	= 0
for i = 0 to WScript.Arguments.count - 1
    select case lcase(WScript.Arguments(i))
    case "-test"
        lngCntTest	= 10
    case else
        Wscript.Echo "P_SHORDER移管(2010.03.30)"
        Wscript.Echo "P_SHORDER.vbs [option]"
        Wscript.Echo " -?"
        Wscript.Echo " -test : 10件だけ登録"
		WScript.Quit
    end select
next

Wscript.Echo "P_SHORDER.vbs "
Wscript.Echo "P_SHORDER移管(草津SC→滋賀PC)"

dbSrcName	= "newsdc-sig"
dbDstName	= "newsdc"

Set dbSrc = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbSrcName
dbSrc.open dbSrcName

Set dbDst = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbDstName
dbDst.open dbDstName

sqlStr = "select *"
sqlStr = sqlStr & " from p_shukeire"
sqlStr = sqlStr & " where KEIJYO_YM = '201107'"
sqlStr = sqlStr & "   and ORDER_NO in (select ORDER_NO from P_SHORDER where G_SYUSHI in ('710','720','730','740','750'))"

Wscript.Echo "sql : " & sqlStr

Set rsSrcPSHOrder = Wscript.CreateObject("ADODB.Recordset")
rsSrcPSHOrder.Open sqlStr, dbSrc, adOpenForwardOnly, adLockReadOnly

Wscript.Echo "sql : 完了"

' コピー先 テーブルオープン
Wscript.Echo "コピー先 : Open p_shukeire"
Set rsDstPSHOrder = Wscript.CreateObject("ADODB.Recordset")
rsDstPSHOrder.MaxRecords = 1
rsDstPSHOrder.CursorLocation = adUseServer
rsDstPSHOrder.Open "p_shukeire", dbDst, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

On Error Resume Next
lngCnt	= 0
Do While Not rsSrcPSHOrder.EOF
	lngCnt	= lngCnt + 1
	if lngCntTest > 0 then
		if lngCnt > 10 then
			exit do
		end if
	end if

	strBuff = "SRC"
	strBuff = strBuff & " " & rsSrcPSHOrder.Fields("ORDER_NO")
	Wscript.Echo strBuff

	rsDstPSHOrder.Addnew
	for each f in rsSrcPSHOrder.Fields
		select case ucase(f.Name)
		case "ORDER_NO"
			select case left(f,1)
			case "0"
				rsDstPSHOrder.Fields(f.Name) = "2" & right(f,4)
			case "1"
				rsDstPSHOrder.Fields(f.Name) = "3" & right(f,4)
			case else
				rsDstPSHOrder.Fields(f.Name) = "4" & right(f,4)
			end select
		case else
			rsDstPSHOrder.Fields(f.Name) = f
		end select

	next
	rsDstPSHOrder.UpdateBatch

	rsSrcPSHOrder.movenext
Loop

Wscript.Echo ""
rsSrcPSHOrder.close

Wscript.Echo "close db : " & dbSrcName
dbSrc.Close
set dbSrc = nothing

Wscript.Echo "close table"
rsDstPSHOrder.Close

Wscript.Echo "close db : " & dbDstName
dbDst.Close
set dbDst = nothing

Wscript.Echo "end"
