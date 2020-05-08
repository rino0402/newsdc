Option Explicit
'-----------------------------------------------------------------------
'メイン呼出＆インクルード
'-----------------------------------------------------------------------
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
	dim	strScriptPath
	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strFileName = strScriptPath & strFileName
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
	Set fso = Nothing
End Function
Call Include("const.vbs")
Call Include("debug.vbs")
Call Include("excel.vbs")
Call Include("get_b.vbs")
Call Include("csv.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
Function GetCD()
	Dim objWshShell
	'①WScript.Shellオブジェクトの作成
	Set objWshShell = CreateObject("WScript.Shell")
	'カレントディレクトリを表示
	dim	strCD
	strCD = objWshShell.CurrentDirectory
	Set objWshShell = Nothing
	GetCD = strCD
End Function

Function GetAbsPath(byVal strPath)
	Dim objFileSys
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	strPath = objFileSys.GetAbsolutePathName(strPath)
	Set objFileSys = Nothing
	GetAbsPath = strPath
End Function

Function GetDate2(byVal v)
	dim	strDate
	strDate = ""
	if isDate(v) then
		strDate = year(v) & Right(00 & month(v), 2) & Right(00 & day(v), 2)
	end if
	GetDate2 = strDate
End Function

Function GetScriptPath()
	GetScriptPath = Left(WScript.ScriptFullName,Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
End Function

Function GetFileName(byVal strFullName)
	dim	strFileName
	strFileName = strFullName
	dim	c
	for each c in split(strFileName,"\")
		Call Debug("GetFileName():" & c)
		if c <> "" then
			strFileName = c
		end if
	next
	GetFileName = strFileName
End Function

Function GetTab(ByVal s)
    Dim r
	r = Split(s,vbTab)
	GetTab = r
End Function

Function GetTrim(byval c)
	if left(c,1) = """" then
		if right(c,1) = """" then
			c = Right(c,Len(c) -1 )
			c = Left(c,Len(c) -1 )
		end if
	end if
	GetTrim = c
End Function

'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "NI顧客深耕日報 商談情報"
	Wscript.Echo "NiSales.vbs [option] <ファイル名>"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo "ex."
	Wscript.Echo "sc32 NiSales.vbs /db:fhd /debug ""C:\Users\kubo\Downloads\商談情報 (12).csv"""
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		else
			strFilename = ""
		end if
	next
	'名前付きオプションチェック
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			strFilename = ""
		case else
			strFilename = ""
		end select
	next
	if strFilename = "" then
		usage()
		Main = 1
		exit Function
	end if
	call LoadNiSales(strFilename)
	Main = 0
End Function

Function LoadNiSales(byVal strFilename)
	Call Debug("LoadNiSales(" & strFilename & ")")
	'-------------------------------------------------------------------
	'ファイル名
	'-------------------------------------------------------------------
	if strFileName = "" then
		Call DispMsg("ファイル名を指定して下さい")
		Exit Function
	end if
	'-------------------------------------------------------------------
	'FileSystemObjectの準備
	'-------------------------------------------------------------------
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Call Debug("CreateObject(Scripting.FileSystemObject)")
	'-------------------------------------------------------------------
	'ファイルオープン
	'-------------------------------------------------------------------
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	Call Debug("OpenTextFile()=" & strFilename)
	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Call LoadNiSalesCsv(objFile)
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objFile.Close()
	set objFile = Nothing
	set objFSO = Nothing
	Call Debug("LoadNiSales():End")
End Function

Function LoadNiSalesCsv(objFile)
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsNiSales
	set rsNiSales = OpenRs(objDb,"NiSales")

	Call Debug("LoadNiSalesCsv()")
	Call LoadNiSalesCsv1(objFile,objDb,rsNiSales)
	
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsNiSales = CloseRs(rsNiSales)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadNiSalesCsv1(objFile,objDb,rsNiSales)
	Call Debug("LoadNiSalesCsv1()")

	Call Debug("delete from NiSales")
	Call ExecuteAdodb(objDb,"delete from NiSales")

	dim	lngRow
	lngRow = 0
	do while ( objFile.AtEndOfStream = False )
		lngRow = lngRow + 1
		dim	strBuff
		strBuff = CsvReadLine(objFile)
		if lngRow > 1 then
			if LoadNiSalesRow(objDb,rsNiSales,strBuff) = 0 then
				Exit do
			end if
		end if
	loop
End Function

Function CsvReadLine(objFile)
	dim	strBuff
	strBuff = ""
	dim	strLineLast
	strLineLast = ""
	do while (True)
		if objFile.AtEndOfStream = True Then
			exit do
		end if
		strBuff = strBuff & objFile.ReadLine()
		strLineLast = Right(strBuff,1)
		Call Debug("CsvReadLine():(" & strLineLast & ")")
		if strLineLast = """" then
			exit do
		end if
		strBuff = strBuff & vbLF
	loop
	CsvReadLine = strBuff
End Function

Function LoadNiSalesRow(objDb,rsNiSales,byVal strBuff)
	Call Debug("LoadNiSalesRow():" & strBuff)

	rsNiSales.AddNew
	dim		aryBuff
	aryBuff = GetCSV(strBuff)
	dim	i
	i = -1
	dim	a
	for each a in aryBuff
		if i >= 0 then
			Call Debug("(" & i & "):" & a)
			Call Debug(rsNiSales.Fields(i).Name & "(" & i & "):" & a)
			dim	dsize
			dsize = rsNiSales.Fields(i).DefinedSize
			rsNiSales.Fields(i) = Get_LeftB(a,dsize)
		end if
		i = i + 1
	next
	rsNiSales.UpdateBatch
	LoadNiSalesRow = i
End Function

Function SetField(objRs,objSt,byVal strField,byVal strCol,byVal lngRow)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & objSt.Range(strCol & lngRow))
	dim	v
	v = RTrim(objSt.Range(strCol & lngRow))
	select case strField
	case "SDt"
		v = Replace(v,"/","")
	end select
	dim	dsize
	dsize = objRs.Fields(strField).DefinedSize
	v = Get_LeftB(v,dsize)
	Call Debug("SetField():" & lngRow & ":" & strField & "(" & dsize & ")=" & v)
	objRs.Fields(strField) = v
End Function
