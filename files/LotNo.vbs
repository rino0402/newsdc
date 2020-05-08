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
	Wscript.Echo "床暖パネル製造ロット番号管理データ"
	Wscript.Echo "LotNo.vbs [option]"
	Wscript.Echo " /Load:ファイル名"
	Wscript.Echo " /db:fhd"
	Wscript.Echo "CurrentDirectory=" & GetCD()
	Wscript.Echo "GetAbsPath()=" & GetAbsPath("LotNo.vbs")
	Wscript.Echo "WScript.Path=" & WScript.Path
	Wscript.Echo "WScript.ScriptFullName=" & WScript.ScriptFullName
	Wscript.Echo "WScript.ScriptName=" & WScript.ScriptName
	Wscript.Echo "GetScriptPath()=" & GetScriptPath()
	Wscript.Echo "GetFileName()=" & GetFileName(WScript.ScriptFullName)
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	'名前無しオプションチェック
	select case WScript.Arguments.UnNamed.Count
	case 0
	case else
		usage()
		Main = 1
		exit Function
	end select
	'名前付きオプションチェック
	dim	strArg
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "load"
		case "debug"
		case "?"
			call usage()
			exit function
		case else
			call usage()
			exit function
		end select
	next
	call LoadLotNo()
	Main = 0
End Function

Function LoadLotNo()
	Call Debug("LoadLotNo()")
	'-------------------------------------------------------------------
	'Excelファイル名
	'-------------------------------------------------------------------
	dim	strFileName
	strFileName = GetAbsPath(GetOption("load",""))
	Call Debug("strFileName=" & strFileName)
	if strFileName = "" then
		Call DispMsg("ファイル名を指定して下さい")
		Exit Function
	end if
	select case LCase(Right(strFileName,3))
	case "xls"
		Call LoadLotNoExcel(strFileName)
	case "txt"
		Call LoadLotNoText(strFileName)
	case else
	end select
End Function

Function LoadLotNoText(byVal strFileName)
	Call Debug("LoadLotNoText(" & strFileName & ")")

	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","fhd") & ")")
	set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	objRs
	set objRs = OpenRs(objDb,"LotNo")
	'-------------------------------------------------------------------
	'テキストファイルオープン
	'-------------------------------------------------------------------
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	dim	strBookName
	strBookName = GetFileName(strFilename)
	dim	cnt
	cnt = 0
	do while ( objFile.AtEndOfStream = False )
		cnt = cnt + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		Call DispMsg(strBuff)
		dim	aryBuff
		aryBuff = GetTab(strBuff)
		if cnt = 1 then
			dim	aryTop
			aryTop = aryBuff
		else
			Dim	strNo
			Dim	strModel
			Dim	strPLotNo
			Dim	strQty
			Dim	strEDt
			strNo		= ""
			strModel	= ""
			strPLotNo	= ""
			strQty		= ""
			strEDt		= ""
			dim	i
			i = 0
			dim	c
			for each c in (aryBuff)
				c = GetTrim(c)
				Call DispMsg(i & ":" & c)
				select case i
				case 0
					strNo = c
				case 1
					strModel = c
				case 2
					strPLotNo = c
				case 3
					strQty = c
				case 4
					strEDt = c
				case else
				end select
				i = i + 1
			next
			if AddLotNo(objDb,objRs,strNo,strModel,strPLotNo,strQty,strEDt,strBookName) <> "" Then
				exit do
			end if
		end if
	loop
	objFile.Close
	set objFile = nothing
	set objFSO = nothing

	set objRs = CloseRs(objRs)
	set objDb = nothing
End Function

Function LoadLotNoExcel(byVal strFileName)
		'-------------------------------------------------------------------
		'Excelの準備
		'-------------------------------------------------------------------
		dim	objXL
		Set objXL = WScript.CreateObject("Excel.Application.15")
		Call Debug("CreateObject(Excel.Application)")
		'-------------------------------------------------------------------
		'Excel ファイルオープン
		'-------------------------------------------------------------------
		dim	strPassword
		strPassword = ""
		dim	objBk
		Set objBk = objXL.Workbooks.Open(strFilename,False,True,,strPassword)
		Call Debug("Workbooks.Open=" & objBk.Name)
		'-------------------------------------------------------------------
		'読込処理
		'-------------------------------------------------------------------
		Call LoadLotNoXls(objXL,objBk)
		'-------------------------------------------------------------------
		'Excelの後処理
		'-------------------------------------------------------------------
		call objBk.Close(False)
		set objXL = Nothing
		Call Debug("LoadLotNo():End")
End Function

Function LoadLotNoXls(objXL,objBk)
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc-osk"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsLotNo
	set rsLotNo = OpenRs(objDb,"LotNo")

	Call Debug("LoadLotNoXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadLotNoXls():SheetName=" & strShtName)
		Call LoadLotNoSheet(objXL,objBk,objSt,objDb,rsLotNo)
	Next
	
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsLotNo = CloseRs(rsLotNo)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadLotNoSheet(objXL,objBk,objSt,objDb,rsLotNo)
	Call Debug("LoadLotNoSheet():SheetName=" & objSt.Name)
	dim	lngRow
	For lngRow = 1 to 65536
		Dim	strNo
		Dim	strModel
		Dim	strPLotNo
		Dim	strQty
		dim	strEDt
		strNo = objSt.Range("A" & lngRow)
		strModel = objSt.Range("B" & lngRow)
		strPLotNo = objSt.Range("C" & lngRow)
		strQty = objSt.Range("D" & lngRow)
		strEDt = GetDate2(objSt.Range("E" & lngRow))
		Call Debug("LoadLotNoSheet():" & strNo & " " & strModel & " " & strPLotNo & " " & strQty & " " & strEDt)
		If strNo = "" then
			Exit For
		End if
		if AddLotNo(objDb,rsLotNo,strNo,strModel,strPLotNo,strQty,strEDt,objBk.Name) <> "" Then
			Exit For
		End if
	Next
End Function

Function AddLotNo(objDb,rsLotNo,byVal strNo,byVal strModel,byVal strPLotNo,byVal strQty,byVal strEDt,byVal strBookName)
	dim	strEntID
	strEntID = "LotNo.vbs"
	dim	strEntDtm
	strEntDtm = GetDateTime(Now())
	Call Debug("AddLotNo():" & strNo & " " & strModel & " " & strPLotNo & " " & strQty & " " & strEDt & " " & strBookName)
	if strEDt = "" then
		if inStr(strBookName,"_") > 0 then
			dim	strS
			strS = Split(strBookName,"_")(1)
			if strS <> "" then
				strEDt = strS
			end if
		end if
	end if
	if ucase(strNo) = "NO" Then
		AddLotNo = ""
		Exit Function
	End if
	if FindLotNo(objDb,rsLotNo,strModel,strPLotNo) Then
		Call Debug("AddLotNo():rsLotNo.AddNew")
		if rsLotNo.State <> adStateClosed then
			rsLotNo.Close
		end if
		rsLotNo.Open "LotNo", objDb, adOpenKeyset, adLockOptimistic,adCmdTableDirect
		rsLotNo.AddNew
		rsLotNo.Fields("EntID") 	= strEntID
		rsLotNo.Fields("EntDtm") 	= strEntDtm
	else
		rsLotNo.Fields("UpdID") 	= strEntID
		rsLotNo.Fields("UpdDtm") 	= strEntDtm
	end if

	if strEDt = "" then
		
	end if

	rsLotNo.Fields("Model") 	= strModel
	rsLotNo.Fields("PLotNo") 	= strPLotNo
	rsLotNo.Fields("IQty") 		= Right("000000" & strQty,6)
	rsLotNo.Fields("MemoNo") 	= strNo
	rsLotNo.Fields("EDt") 		= strEDt
	rsLotNo.Fields("EntFN") 	= strBookName
	rsLotNo.UpdateBatch
	if Err.Number <> 0 then
		Call DispErr(Err)
	end if
	Call Debug("AddLotNo():rsLotNo.Status=" & rsLotNo.Status)
'	Call Debug("AddLotNo():rsLotNo.DataSource =" & rsLotNo.DataSource )
	AddLotNo = ""
End Function

Function FindLotNo(objDb,rsLotNo,strModel,strPLotNo)
	dim	strSql
	strSql = "select * from LotNo"
	strSql = makeWhere(strSql,"Model"	,strModel	,"")
	strSql = makeWhere(strSql,"PLotNo"	,strPLotNo	,"")
	FindLotNo = UpdateOpenRs(objDb,rsLotNo,strSql)
End Function

