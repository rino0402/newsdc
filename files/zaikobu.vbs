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
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "Glics在庫データ"
	Wscript.Echo "zaikobu.vbs [option]"
	Wscript.Echo " /Load:ファイル名"
	Wscript.Echo "CurrentDirectory=" & GetCD()
	Wscript.Echo "GetAbsPath()=" & GetAbsPath("zaikobu.vbs")
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
	call LoadGS()
	Main = 0
End Function

Function LoadGS()
	Call Debug("LoadGS()")
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
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	dim	objXL
	Set objXL = WScript.CreateObject("Excel.Application")
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
	Call LoadGSXls(objXL,objBk)
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadLotNo():End")
End Function

Function LoadGSXls(objXL,objBk)
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsZaiko
	set rsZaiko = OpenRs(objDb,"ZaikoBu")

	Call Debug("LoadLotNoXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadGSXls():SheetName=" & strShtName)
		Call LoadGSSheet(objXL,objBk,objSt,objDb,rsZaiko)
	Next
	
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsZaiko = CloseRs(rsZaiko)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadGSSheet(objXL,objBk,objSt,objDb,rsZaiko)
	Call Debug("LoadGSSheet():SheetName=" & objSt.Name)
	dim	lngRow
	Dim	strStat
	strStat = ""
	For lngRow = 1 to 65536
		dim	strA
		dim	strB
		strA = objSt.Range("A" & lngRow)
		strB = objSt.Range("B" & lngRow)
		Call Debug("LoadGSSheet():" & lngRow & ":" & strA & " " & strB)
		select case strStat
		case ""
			if strB <> "" then
				strStat = "data"
			end if
		case "data"
			if strA = "" then
				strStat = "end"
				exit for
			end if
			Call LoadSGRow(objSt,objDb,rsZaiko,lngRow)
		end select
'		strNo = objSt.Range("A" & lngRow)
'		strModel = objSt.Range("B" & lngRow)
'		strPLotNo = objSt.Range("C" & lngRow)
'		strQty = objSt.Range("D" & lngRow)
'		strEDt = GetDate2(objSt.Range("E" & lngRow))
'		Call Debug("LoadLotNoSheet():" & strNo & " " & strModel & " " & strPLotNo & " " & strQty & " " & strEDt)
'		If strNo = "" then
'			Exit For
'		End if
'		if AddLotNo(objDb,rsLotNo,strNo,strModel,strPLotNo,strQty,strEDt,objBk.Name) <> "" Then
'			Exit For
'		End if
	Next
End Function

Function LoadSGRow(objSt,objDb,rsZaiko,byVal lngRow)
	Call Debug("LoadSGRow():SheetName=" & objSt.Name & ":" & lngRow)
	dim	rngTop
	dim	rngCur
	set rngTop = objSt.Range("D4")
	set rngCur = objSt.Range("D" & lngRow)
	dim	strJCode
	dim	strPn
	strJCode 	= "00021185"
	strPn		= objSt.Range("B" & lngRow)
	do while True
		Call Debug("LoadSGRow():" & rngTop & "," & rngCur)
		if rngTop = "" then
			exit do
		end if
		if SyushiCheck(rngTop,rngCur) <> "" Then
			Call AddZaiko(objDb,rsZaiko,strJCode,strPn,rngTop,rngCur)
		end if
		set rngTop = rngTop.Offset(0,1)
		set rngCur = rngCur.Offset(0,1)
	Loop
End Function

Function SyushiCheck(byVal strSyushi,byVal strQty)
	SyushiCheck = ""
	if strSyushi = "" then
		exit function
	end if
	if Len(strSyushi) <> 2 then
		exit function
	end if
	if strQty = "" then
		exit function
	end if
	SyushiCheck = strSyushi
End Function

Function AddZaiko(objDb,rsZaiko,byVal strJCode,byVal strPn,byVal strSyushi,byVal strQty)
	On Error Resume Next
	Call Debug("AddZaiko(" & strJCode & "," & strPn & "," & strSyushi & "," & strQty & ")")
	rsZaiko.AddNew
	rsZaiko.Fields("JCode") 	= strJCode
	rsZaiko.Fields("Pn") 		= strPn
	rsZaiko.Fields("Syushi")	= strSyushi
	rsZaiko.Fields("Qty")		= strQty
	rsZaiko.UpdateBatch
	if Err.Number <> 0 then
		Call DispErr(Err)
		rsZaiko.CancelBatch
	end if
	On Error Goto 0
End Function

Function AddLotNo(objDb,rsLotNo,byVal strNo,byVal strModel,byVal strPLotNo,byVal strQty,byVal strEDt,strBookName)
	dim	strEntID
	strEntID = "LotNo.vbs"
	dim	strEntDtm
	strEntDtm = GetDateTime(Now())
	Call Debug("AddLotNo():" & strNo & " " & strModel & " " & strPLotNo & " " & strQty & " " & strEDt & " " & strBookName)
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

