Option Explicit
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
End Function
Call Include("const.vbs")

Call Main()

Function usage()
    Wscript.Echo "del_syuka_ret.vbs [/db:DbName] [/test] [/skip:num] [/limit:num] [/delete]"
    Wscript.Echo "<例>"
    Wscript.Echo "del_syuka_ret.vbs /db:newsdc-ono"
End Function

Function Main()
	dim	strArg
	dim	objDb

	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		call usage()
		Exit Function
	next
	'名前付きオプションチェック
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "limit"
		case "test"
		case "skip"
		case "insert"
		case "move"
		case "y_syuka"
		case "del_syuka"
		case else
			call usage()
			exit function
		end select
	next
	'データベースオープン
	set objDb = OpenDb()
	' テーブルレコードセット
	dim	rsTable
	dim	rsQuery

	Set rsTable = Wscript.CreateObject("ADODB.Recordset")
	EchoMsg "    テーブルオープン:" & GetTable()
	rsTable.Open GetTable(), objDb, adOpenKeyset, adLockOptimistic,adCmdTableDirect

	Set rsQuery = Wscript.CreateObject("ADODB.Recordset")
	EchoMsg "    クエリーオープン:" & GetSql()
	rsQuery.Open GetSql(), objDb, adOpenForwardOnly, adLockOptimistic

	dim	lngCnt
	lngCnt = 0
	Do While rsQuery.EOF = False
		lngCnt	= lngCnt + 1
		if LimitCheck(lngCnt) then
			exit do
		end if
		dim	strMsg
		dim	strId
		strMsg = Right(space(10) & lngCnt,10) _
			  & " " & rsQuery.Fields("JGYOBU") _
			  & " " & rsQuery.Fields("NAIGAI") _
			  & " " & rsQuery.Fields("key_hin_no") _
			  & " " & rsQuery.Fields("key_syuka_ymd") _
			  & " " & rsQuery.Fields("LK_SEQ_NO") _
			  & " " & rsQuery.Fields("KAN_KBN") _
			  & " " & rsQuery.Fields("JITU_SURYO") _
			  & " " & rsQuery.Fields("UPD_NOW") _
			  & " " & rsQuery.Fields("key_id_no") _
			  & ""
		if strId = "" then
			strMsg = strMsg & " データなし"
			if WScript.Arguments.Named.Exists("move") then
				strMsg = strMsg & " " & strId & DoMove(rsTable,rsQuery)
			end if
		else
			strMsg = strMsg & " データあり " & strId
		end if
		
		EchoMsg strMsg
		rsQuery.MoveNext
	Loop

	' テーブルClose
	rsTable.Close
	set rsTable = Nothing
	' クエリ-Close
	rsQuery.Close
	set rsQuery = Nothing
'	rsRead1.Close
	'データベースクローズ
	set objDb = CloseDb(objDb)
End Function

Function DoMove(byval rsTable,byval rsQuery)
	dim	strMsg
	strMsg = ""
	On Error Resume Next
		rsTable.AddNew
		dim	f
		for each f in (rsQuery.Fields)
			rsTable.Fields(f.name) = f
		next
'		rsTable.Fields("UPD_NOW") = "20120421000000"
		rsTable.UpdateBatch
		if Err.Number = 0 then
			' 削除
			rsQuery.Delete
			if Err.Number = 0 then
				strMsg = strMsg & " 削除OK"
			else
				strMsg = strMsg & " 削除Err:" & Hex(Err.Number) & " " & Err.Description
				rsQuery.CancelBatch
			end if
			Err.Clear
		else
			strMsg = strMsg & " 追加Err:" & Hex(Err.Number) & " " & Err.Description
			rsTable.CancelBatch
		end if
		Err.Clear
	On Error Goto 0
	DoMove = strMsg
End Function

Function GetTable()
	dim	strTableName
	strTableName = "y_syuka"
	GetTable = strTableName
End Function

Function GetSql()
	dim	strSql
	strSql = "select * " _
		   & " from del_syuka" _
		   & " where key_syuka_ymd = '20120426'" _
		   & ""

'		   & "   and key_hin_no in (select distinct hin_gai from item where jgyobu = '5' and naigai = '1')" _
'		   & "   and key_syuka_ymd > '20110320'"
	GetSql = strSql
End Function

Function OpenDb()
	dim	objDb
	dim	strDbName

	strDbName = GetOption("Db","newsdc-ono")

	Set objDb = Wscript.CreateObject("ADODB.Connection")

	EchoMsg "データベースオープン:" & strDbName
	objDb.Open strDbName

	Set OpenDb = objDb
End Function

Function CloseDb(byval objDb)

	EchoMsg "データベースクローズ:" & objDb.Properties("Data Source")
	objDb.Close

	set objDb = Nothing

	Set CloseDb = objDb
End Function

Sub EchoMsg(byval strMsg)
	Wscript.Echo strMsg
End Sub

Function GetOption(byval strName _
				  ,byval strDefault _
				  )
	dim	strValue

	strValue = strDefault
	if WScript.Arguments.Named.Exists(strName) then
		strValue = WScript.Arguments.Named(strName)
	end if
	GetOption = strValue
End Function

Function LimitCheck(byval lngCnt)
	dim	lngLimit
	dim	bLimit

	bLimit = False
	lngLimit = CLng(GetOption("limit",0))
	if lngLimit > 0 then
		if lngCnt > lngLimit then
			bLimit = True
		end if
	end if
	LimitCheck = bLimit
End Function
