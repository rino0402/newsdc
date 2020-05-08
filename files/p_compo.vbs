Option Explicit
' P_COMPO.DAT移管プログラム
' 2010.03.24 新規作成
' select SHIMUKE_CODE,JGYOBU,NAIGAI,count(*),max(UPD_DATETIME)
'  from p_compo
'  group by SHIMUKE_CODE,JGYOBU,NAIGAI

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

dim	dbSrc		' コピー元 DB Object
dim	dbSrcName	' コピー元 DB Name
dim	rsSrcPCompo	' コピー元 レコードセット(P_COMPO)
dim	f		' Field Object

dim	sqlStr		' SQL

dim	dbDst		' コピー先 DB Object
dim	dbDstName	' コピー先 DB Name
dim	rsDstPCompo	' コピー先 レコードセット(P_COMPO)

dim	strBuff
dim	Fs
dim	logFile

dim	lngCnt		' 登録件数
dim	i
dim	lngCntTest	' TEST登録件数
dim	strTable

lngCntTest	= 0
strTable	= "p_compo"
for i = 0 to WScript.Arguments.count - 1
    select case lcase(WScript.Arguments(i))
    case "-test"
        lngCntTest	= 10
    case "-k"
        strTable	= "p_compo_k"
    case else
        Wscript.Echo "P_COMPO移管(2010.03.24)"
        Wscript.Echo "p_compo.vbs [option]"
        Wscript.Echo " -?"
        Wscript.Echo " -test : 10件だけ登録"
	WScript.Quit
    end select
next

Wscript.Echo "p_compo.vbs "
Wscript.Echo "P_COMPO移管(草津SC→滋賀PC)"

dbSrcName	= "newsdc-sig"
dbDstName	= "newsdc"

Set dbSrc = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbSrcName
dbSrc.open dbSrcName

Set dbDst = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbDstName
dbDst.open dbDstName

sqlStr = "select *"
sqlStr = sqlStr & " from " & strTable
sqlStr = sqlStr & " where SHIMUKE_CODE = '04'"
sqlStr = sqlStr & "   and jgyobu = 'R'"
sqlStr = sqlStr & "   and naigai = '1'"
if strTable = "p_compo" then
	sqlStr = sqlStr & "   and DATA_KBN = '0'"
else
	sqlStr = sqlStr & "   and DATA_KBN <> '0'"
end if

Wscript.Echo "db.CommandTimeout : " & dbSrc.CommandTimeout
dbSrc.CommandTimeout = 0
Wscript.Echo "db.CommandTimeout : " & dbSrc.CommandTimeout
Wscript.Echo "sql : " & sqlStr

Set rsSrcPCompo = Wscript.CreateObject("ADODB.Recordset")
rsSrcPCompo.Open sqlStr, dbSrc, adOpenForwardOnly, adLockReadOnly

Wscript.Echo "sql : 完了"

' コピー先 P_COMPOオープン
Wscript.Echo "コピー先P_COMPO : Open " & strTable
Set rsDstPCompo = Wscript.CreateObject("ADODB.Recordset")
rsDstPCompo.MaxRecords = 1
rsDstPCompo.CursorLocation = adUseServer
rsDstPCompo.Open strTable, dbDst, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

' ログファイルオープン

On Error Resume Next
lngCnt	= 0
Do While Not rsSrcPCompo.EOF
	lngCnt	= lngCnt + 1
	if lngCntTest > 0 then
		if lngCnt > 10 then
			exit do
		end if
	end if
	strBuff = "INS"
	strBuff = strBuff & " " & rsSrcPCompo.Fields("SHIMUKE_CODE")
	strBuff = strBuff & " " & rsSrcPCompo.Fields("JGYOBU")
	strBuff = strBuff & " " & rsSrcPCompo.Fields("NAIGAI")
	strBuff = strBuff & " " & rsSrcPCompo.Fields("HIN_GAI")
	strBuff = strBuff & " " & rsSrcPCompo.Fields("DATA_KBN")
	strBuff = strBuff & " " & rsSrcPCompo.Fields("SEQNO")
	if strTable = "p_compo" then
		strBuff = strBuff & " " & rsSrcPCompo.Fields("CLASS_CODE")
	else
		strBuff = strBuff & " " & rsSrcPCompo.Fields("KO_HIN_GAI")
	end if

	Wscript.Echo strBuff
	rsDstPCompo.Addnew
	for each f in rsSrcPCompo.Fields
'		Wscript.Echo f.Name,f
		rsDstPCompo.Fields(f.Name) = f
		if ucase(f.Name) = "SHIMUKE_CODE" then
			rsDstPCompo.Fields(f.Name) = "02"
		else
			rsDstPCompo.Fields(f.Name) = f
		end if
	next
	rsDstPCompo.UpdateBatch
	rsSrcPCompo.movenext
Loop

Wscript.Echo ""
rsSrcPCompo.close

Wscript.Echo "close db : " & dbSrcName
dbSrc.Close
set dbSrc = nothing

Wscript.Echo "close rsDstItem"
rsDstPCompo.Close
set rsDstPCompo = nothing

Wscript.Echo "close db : " & dbDstName
dbDst.Close
set dbDst = nothing

Wscript.Echo "end"
