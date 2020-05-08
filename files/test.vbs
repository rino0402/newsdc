Option Explicit
dim db
dim	dbName
dim	sqlStr
dim	rsList
dim	strSoko

strSoko = WScript.Arguments(0)
Wscript.Echo "test.vbs " & strSoko

dbName = "newsdc"

Set db = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbName
db.open dbName

	sqlStr = "select"
	sqlStr = sqlStr & " Soko as ""倉庫"""
	sqlStr = sqlStr & ",SyukaDt as ""出荷日"""
	sqlStr = sqlStr & ",IDNo as ""ID"""
	sqlStr = sqlStr & ",left(Pn,12) as ""品番"""
	sqlStr = sqlStr & ",left(PName,20) as ""品名"""
	sqlStr = sqlStr & ",CONVERT(Qty,SQL_NUMERIC) as ""数量"""
	sqlStr = sqlStr & ",DenNo as ""伝票No"""
	sqlStr = sqlStr & ",rtrim(Aitesaki) + ' ' + rtrim(AitesakiName) as ""相手先"""
'	sqlStr = sqlStr & ",Syushi as ""収支"""
'	sqlStr = sqlStr & ",SSyushi as ""事業場"""
	sqlStr = sqlStr & ",rtrim(CyuKbn) + ' ' + rtrim(CyuName) as ""注文区分"""
	sqlStr = sqlStr & ",if(ChokuKbn='1','直送','') as ""直送区分"""
	sqlStr = sqlStr & ",rtrim(Biko1) as ""備考1"""
	sqlStr = sqlStr & ",rtrim(Biko2) as ""備考2"""
	sqlStr = sqlStr & " from g_syuka"
	sqlStr = sqlStr & " where Soko = '" & strSoko & "'"
	sqlStr = sqlStr & " and SyukaDT < REPLACE(convert(CURDATE(),SQL_CHAR),'-','')"
	sqlStr = sqlStr & " and Syushi <> '02'"
	sqlStr = sqlStr & " order by ""出荷日"",ID"

Wscript.Echo "sql : " & sqlStr
dim	i
dim	strBuff
dim	objFSO
dim	objOutput
Const ForReading = 1, ForWriting = 2, ForAppending = 8

	set rsList = db.Execute(sqlStr)
	Wscript.Echo ""

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objOutput = objFSO.OpenTextFile("test.txt", ForWriting, True)

	Do While Not rsList.EOF
		strBuff = ""
		For i=0 To rsList.Fields.Count-1
			strBuff = strBuff & " " & rsList.Fields(i)
		next
		Wscript.Echo left(strBuff,79)
		objOutput.WriteLine strBuff
		rsList.movenext
	Loop
	Wscript.Echo ""
	objOutput.Close 

Wscript.Echo "close db : " & dbName
db.Close
set db = nothing
Wscript.Echo "end"
