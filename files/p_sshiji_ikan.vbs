Option Explicit
' P_SSHIJI.DAT�ڊǃv���O����
' 2010.03.30 �V�K�쐬
' select SHIMUKE_CODE,JGYOBU,NAIGAI,count(*),max(UPD_DATETIME)
'  from p_compo
'  group by SHIMUKE_CODE,JGYOBU,NAIGAI
'select left(PRINT_DATETIME,8),count(*) from p_sshiji_o
' where SHIMUKE_CODE = '02'
'   and KAN_F <> '1'
' group by left(PRINT_DATETIME,8)

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

dim	dbSrc		' �R�s�[�� DB Object
dim	dbSrcName	' �R�s�[�� DB Name
dim	rsSrcPSShijiO	' �R�s�[�� ���R�[�h�Z�b�g(P_SSHIJI_O)
dim	rsSrcPSShijiK	' �R�s�[�� ���R�[�h�Z�b�g(P_SSHIJI_K)
dim	f		' Field Object

dim	sqlStr		' SQL

dim	dbDst		' �R�s�[�� DB Object
dim	dbDstName	' �R�s�[�� DB Name
dim	rsDstPSShijiO	' �R�s�[�� ���R�[�h�Z�b�g(P_SSHIJI_O)
dim	rsDstPSShijiK	' �R�s�[�� ���R�[�h�Z�b�g(P_SSHIJI_K)

dim	strBuff
dim	Fs
dim	logFile

dim	lngCnt		' �o�^����
dim	i
dim	lngCntTest	' TEST�o�^����
dim	strTable

lngCntTest	= 0
strTable	= "p_sshiji_o"
for i = 0 to WScript.Arguments.count - 1
    select case lcase(WScript.Arguments(i))
    case "-test"
        lngCntTest	= 10
'    case "-k"
'        strTable	= "p_compo_k"
    case else
        Wscript.Echo "P_SSHIJI�ڊ�(2010.03.30)"
        Wscript.Echo "p_sshiji.vbs [option]"
        Wscript.Echo " -?"
        Wscript.Echo " -test : 10�������o�^"
	WScript.Quit
    end select
next

Wscript.Echo "p_sshiji.vbs "
Wscript.Echo "P_SSHIJI�ڊ�(����SC������PC)"

dbSrcName	= "newsdc-sig"
dbDstName	= "newsdc"

Set dbSrc = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbSrcName
dbSrc.open dbSrcName

Set dbDst = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbDstName
dbDst.open dbDstName

sqlStr = "select *"
sqlStr = sqlStr & " from p_sshiji_o"
sqlStr = sqlStr & " where SHIMUKE_CODE = '04'"

Wscript.Echo "db.CommandTimeout : " & dbSrc.CommandTimeout
dbSrc.CommandTimeout = 0
Wscript.Echo "db.CommandTimeout : " & dbSrc.CommandTimeout
Wscript.Echo "sql : " & sqlStr

Set rsSrcPSShijiO = Wscript.CreateObject("ADODB.Recordset")
rsSrcPSShijiO.Open sqlStr, dbSrc, adOpenForwardOnly, adLockReadOnly

Set rsSrcPSShijiK = Wscript.CreateObject("ADODB.Recordset")

Wscript.Echo "sql : ����"

' �R�s�[�� p_sshiji_o�I�[�v��
Wscript.Echo "�R�s�[�� : Open p_sshiji_o"
Set rsDstPSShijiO = Wscript.CreateObject("ADODB.Recordset")
rsDstPSShijiO.MaxRecords = 1
rsDstPSShijiO.CursorLocation = adUseServer
rsDstPSShijiO.Open "p_sshiji_o", dbDst, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

' �R�s�[�� p_sshiji_k�I�[�v��
Wscript.Echo "�R�s�[�� : Open p_sshiji_k"
Set rsDstPSShijiK = Wscript.CreateObject("ADODB.Recordset")
rsDstPSShijiK.MaxRecords = 1
rsDstPSShijiK.CursorLocation = adUseServer
rsDstPSShijiK.Open "p_sshiji_k", dbDst, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

' ���O�t�@�C���I�[�v��

On Error Resume Next
lngCnt	= 0
Do While Not rsSrcPSShijiO.EOF
	lngCnt	= lngCnt + 1
	if lngCntTest > 0 then
		if lngCnt > 10 then
			exit do
		end if
	end if
	strBuff = "SRC"
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("SHIMUKE_CODE")
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("JGYOBU")
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("NAIGAI")
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("HIN_GAI")
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("SHIJI_NO")

	Wscript.Echo strBuff
	strBuff = "DST"
	strBuff = strBuff & " " & "02"
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("JGYOBU")
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("NAIGAI")
	strBuff = strBuff & " " & rsSrcPSShijiO.Fields("HIN_GAI")
	strBuff = strBuff & " " & "9" & right(rsSrcPSShijiO.Fields("SHIJI_NO"),7)
	Wscript.Echo strBuff

	rsDstPSShijiO.Addnew
	for each f in rsSrcPSShijiO.Fields
		rsDstPSShijiO.Fields(f.Name) = f
		if ucase(f.Name) = "SHIMUKE_CODE" then
			rsDstPSShijiO.Fields(f.Name) = "02"
		elseif ucase(f.Name) = "SHIJI_NO" then
			rsDstPSShijiO.Fields(f.Name) = "9" & right(rsSrcPSShijiO.Fields("SHIJI_NO"),7)
		else
			rsDstPSShijiO.Fields(f.Name) = f
		end if
	next
	rsDstPSShijiO.UpdateBatch

	sqlStr = "select *"
	sqlStr = sqlStr & " from p_sshiji_k"
	sqlStr = sqlStr & " where SHIJI_NO = '" & rtrim(rsSrcPSShijiO.Fields("SHIJI_NO")) & "'"
	if rsSrcPSShijiK.state <> adStateClosed then
		rsSrcPSShijiK.Close
	end if
	rsSrcPSShijiK.Open sqlStr, dbSrc, adOpenForwardOnly, adLockReadOnly
	Do While Not rsSrcPSShijiK.EOF
		strBuff = "  K"
		strBuff = strBuff & " " & rsSrcPSShijiK.Fields("SHIJI_NO")
		strBuff = strBuff & " " & rsSrcPSShijiK.Fields("DATA_KBN")
		strBuff = strBuff & " " & rsSrcPSShijiK.Fields("SEQNO")

		Wscript.Echo strBuff
		rsDstPSShijiK.Addnew
		for each f in rsSrcPSShijiK.Fields
			rsDstPSShijiK.Fields(f.Name) = f
			if ucase(f.Name) = "SHIMUKE_CODE" then
				rsDstPSShijiK.Fields(f.Name) = "02"
			elseif ucase(f.Name) = "SHIJI_NO" then
				rsDstPSShijiK.Fields(f.Name) = "9" & right(rsSrcPSShijiK.Fields("SHIJI_NO"),7)
			else
				rsDstPSShijiK.Fields(f.Name) = f
			end if
		next
		rsDstPSShijiK.UpdateBatch
		rsSrcPSShijiK.movenext
	loop

	rsSrcPSShijiO.movenext
Loop

Wscript.Echo ""
rsSrcPSShijiO.close
rsSrcPSShijiK.close

Wscript.Echo "close db : " & dbSrcName
dbSrc.Close
set dbSrc = nothing

Wscript.Echo "close rsDstItem"
rsDstPSShijiO.Close
rsDstPSShijiK.Close

Wscript.Echo "close db : " & dbDstName
dbDst.Close
set dbDst = nothing

Wscript.Echo "end"
