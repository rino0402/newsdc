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
	sqlStr = sqlStr & " Soko as ""�q��"""
	sqlStr = sqlStr & ",SyukaDt as ""�o�ד�"""
	sqlStr = sqlStr & ",IDNo as ""ID"""
	sqlStr = sqlStr & ",left(Pn,12) as ""�i��"""
	sqlStr = sqlStr & ",left(PName,20) as ""�i��"""
	sqlStr = sqlStr & ",CONVERT(Qty,SQL_NUMERIC) as ""����"""
	sqlStr = sqlStr & ",DenNo as ""�`�[No"""
	sqlStr = sqlStr & ",rtrim(Aitesaki) + ' ' + rtrim(AitesakiName) as ""�����"""
'	sqlStr = sqlStr & ",Syushi as ""���x"""
'	sqlStr = sqlStr & ",SSyushi as ""���Ə�"""
	sqlStr = sqlStr & ",rtrim(CyuKbn) + ' ' + rtrim(CyuName) as ""�����敪"""
	sqlStr = sqlStr & ",if(ChokuKbn='1','����','') as ""�����敪"""
	sqlStr = sqlStr & ",rtrim(Biko1) as ""���l1"""
	sqlStr = sqlStr & ",rtrim(Biko2) as ""���l2"""
	sqlStr = sqlStr & " from g_syuka"
	sqlStr = sqlStr & " where Soko = '" & strSoko & "'"
	sqlStr = sqlStr & " and SyukaDT < REPLACE(convert(CURDATE(),SQL_CHAR),'-','')"
	sqlStr = sqlStr & " and Syushi <> '02'"
	sqlStr = sqlStr & " order by ""�o�ד�"",ID"

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
