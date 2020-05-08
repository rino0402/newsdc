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
	Wscript.Echo "NIコラボ顧客プロフィール"
	Wscript.Echo "NiCustomer.vbs [option] <ファイル名>"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo "ex."
	Wscript.Echo "sc32 NiCustomer.vbs /db:fhd /debug ""C:\Users\kubo\Downloads\顧客プロフィール (1).csv"""
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
	call LoadNiCustomer(strFilename)
	Main = 0
End Function

Function LoadNiCustomer(byVal strFilename)
	Call Debug("LoadNiCustomer(" & strFilename & ")")
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
	Call LoadNiCustomerCsv(objFile)
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objFile.Close()
	set objFile = Nothing
	set objFSO = Nothing
	Call Debug("LoadNiCustomer():End")
End Function

Function LoadNiCustomerCsv(objFile)
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsNiCustomer
	set rsNiCustomer = OpenRs(objDb,"NiCustomer")

	Call Debug("LoadNiCustomerCsv()")
	Call LoadNiCustomerCsv1(objFile,objDb,rsNiCustomer)
	
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsNiCustomer = CloseRs(rsNiCustomer)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadNiCustomerCsv1(objFile,objDb,rsNiCustomer)
	Call Debug("LoadNiCustomerCsv1()")

	Call Debug("delete from NiCustomer")
	Call ExecuteAdodb(objDb,"delete from NiCustomer")

	dim	lngRow
	lngRow = 0
	do while ( objFile.AtEndOfStream = False )
		lngRow = lngRow + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		if lngRow > 1 then
			if LoadNiCustomerRow(objDb,rsNiCustomer,strBuff) = 0 then
				Exit do
			end if
		end if
	loop
End Function

Function LoadNiCustomerRow(objDb,rsNiCustomer,byVal strBuff)
	Call Debug("LoadNiCustomerRow():" & strBuff)

	rsNiCustomer.AddNew
	dim		aryBuff
	aryBuff = GetCSV(strBuff)
	dim	i
	i = -1
	dim	a
	for each a in aryBuff
		if i >= 0 then
			select case rsNiCustomer.Fields(i).Name
			case "Code"
				Call ToHalf(a)
			end select
			Call Debug(rsNiCustomer.Fields(i).Name & "(" & i & "):" & a)
			dim	dsize
			dsize = rsNiCustomer.Fields(i).DefinedSize
			rsNiCustomer.Fields(i) = Get_LeftB(a,dsize)
		end if
		i = i + 1
	next
	rsNiCustomer.UpdateBatch
	LoadNiCustomerRow = i
End Function

Function SetField(objRs,objSt,byVal strField,byVal strCol,byVal lngRow)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & objSt.Range(strCol & lngRow))
	dim	v
	v = RTrim(objSt.Range(strCol & lngRow))
	select case strField
	case "SDt"
		v = Replace(v,"/","")
	case "Code"
		Call ToHalf(v)
	end select
	dim	dsize
	dsize = objRs.Fields(strField).DefinedSize
	v = Get_LeftB(v,dsize)
	Call Debug("SetField():" & lngRow & ":" & strField & "(" & dsize & ")=" & v)
	objRs.Fields(strField) = v
End Function

Sub ToHalf( ByRef strText )
	Const conToHalf = 33311
	Dim strBuf
	Dim lngHex, lngUcLower, lngUcUpper, lngLcLower, lngLcUpper
	Dim i
	
	lngUcLower = CLng("&h" & Hex(Asc("０")))
	lngUcUpper = CLng("&h" & Hex(Asc("Ｚ")))
	
	lngLcLower = CLng("&h" & Hex(Asc("ａ")))
	lngLcUpper = CLng("&h" & Hex(Asc("ｚ")))
	
	strBuf = ""
	
	For i = 1 To Len(strText)
		lngHex = CLng( "&h" & Hex(Asc( Mid(strText, i, 1) )) )
		If lngHex >= lngUcLower And lngHex <= lngUcUpper Then
			lngHex = lngHex - conToHalf
		ElseIf lngHex >= lngLcLower And lngHex <= lngLcUpper Then
			lngHex = lngHex - conToHalf - 1
		ElseIF lngHex = 33148 Then' "－"
			lngHex = 45' "-"
		ElseIf lngHex = 33172 Then' "＃"
			lngHex = 35' "#"
		ElseIF lngHex = 33088 Then' "　"
			lngHex = 32' " "
		End If
		strBuf = strBuf & Chr(lngHex)
	Next
	strText = strBuf
End Sub
