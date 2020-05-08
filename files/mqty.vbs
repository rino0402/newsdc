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
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "出荷実績データ(Excel)読込"
	Wscript.Echo "mqty.vbs [option] <filename.xls>"
	Wscript.Echo " -?"
End Sub

'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	i
	dim	strArg
	dim	strFilename
	dim	strCenterCD
	dim	strYYYY
	dim	strSheetName
	dim	objFSO
	dim	objLog

	strFilename = ""
	For Each strArg In WScript.Arguments.UnNamed
    	select case strArg
		case "-?"
			call usage()
			Main = 1
			exit Function
		case else
			if strFilename = "" then
				strFilename = strArg
			else
				usage()
				Main = 1
				exit Function
			end if
		end select
	Next
	if strFilename = "" then
		usage()
		Main = 1
		exit Function
	end if
	Call LoadExcel(strFilename)
	Main = 0
End Function
'-----------------------------------------------------------------------
'Excel読込
'-----------------------------------------------------------------------
Private Sub LoadExcel(byval strFilename)
'	On Error Resume Next

	Call DispMsg("LoadExcel(" & strFilename & ")")

	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	dim	objXL
	Call DispMsg("CreateObject(Excel.Application)")
	Set objXL = WScript.CreateObject("Excel.Application")
'	objXL.Application.Visible = True
	dim	objBk
	Call DispMsg("Workbooks.Open(" & strFilename & ")")
	Set objBk = objXL.Workbooks.Open(strFilename,False,True,,"")

	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	dim	strDbName
	Call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	strDbName = "newsdc-4"
	Call objDb.Open(strDbName)
	'-------------------------------------------------------------------
	' テーブルOpen
	'-------------------------------------------------------------------
	Call DispMsg("Open(MonthlyQty)")
	dim	rsMQty
	Set rsMQty = Wscript.CreateObject("ADODB.Recordset")
	rsMQty.MaxRecords = 1
	rsMQty.Open "MonthlyQty", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	'-------------------------------------------------------------------
	' 経営資料データ読込＆テーブル登録
	'-------------------------------------------------------------------
	dim	strStat
	dim	strShtName
	strStat = ""
	dim	objSt
	for each objSt in objBk.Worksheets
		strShtName = objSt.Name
		Call DispMsg("シート名：" & strShtName)
		Call LoadSheetBW(objSt,objDb,rsMQty)
	next
	'-------------------------------------------------------------------
	'データベースの後処理
	'-------------------------------------------------------------------
	Call rsMQty.Close
	Set rsMQty = Nothing
	Call objDb.Close
	Set objDb = Nothing
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	Call objBk.Close(False)
	Set objXL = Nothing
End Sub
'-----------------------------------------------------------------------
'出荷実績読込 from ～ to
'-----------------------------------------------------------------------
Function LoadSheetBW(objSt,objDb,rsMQty)
	dim	strFrom
	strFrom = GetOption("from","")
	dim	strTo
	strTo = GetOption("to",strFrom)
	dim strCol
	strCol = "AA"
	dim	objRg
	set objRg = objSt.Range("AA2")
	do while objRg <> ""
		Call DispMsg(objRg.Address & ":" & objRg)
		if isNumeric(objRg) and len(objRg) = 6 then
			if objRg > strTo then
				exit do
			end if
			if objRg >= strFrom then
				Call DispMsg(objRg & "=" & strFrom & " " & objRg.Address)
				strCol = Split(objRg.Address,"$")(1)
				Call LoadSheet(objSt,objDb,rsMQty,strCol)
			end if
		end if
		set objRg = objRg.Offset(0,1)
	loop

End Function
'-----------------------------------------------------------------------
'出荷実績読込
'-----------------------------------------------------------------------
Function LoadSheet(objSt,objDb,rsMQty,byVal strCol)
	Call DispMsg("LoadSheet(" & objSt.Name & "," & strCol & ")")
'	dim	strCol
'	strCol = "AE"
	dim strYM
	strYM = objSt.Range(strCol & "2")
	dim	strSql
	strSql = "delete from MonthlyQty"
	strSql = strSql & " where DT='" & strYM & "'"
	strSql = strSql & "   and JGYOBU='A'"
	strSql = strSql & "   and NAIGAI='0'"
	Call DispMsg(strSql)
	Call ExecuteAdodb(objDb,strSql)
	dim	lngRow
	lngRow = 3
	do while true
		dim	strPn
		strPn = RTrim(objSt.Range("C" & lngRow))
		if LoopCheck(lngRow,strPn) then
			exit do
		end if
		dim	lngQty
		lngQty = objSt.Range(strCol & lngRow)
		dim	strAdd
		strAdd = ""
		if lngQty <> 0 then
			' insert into MonthlyQty (DT,JGYOBU,NAIGAI,HIN_GAI,SyukaCnt,SyukaQty) values (	'200404'	,'A'	,'0'	,'2SC3169'	,'1'	,'-1'	) #
			strAdd = " AddNew"
			rsMQty.AddNew
			rsMQty.Fields("DT") = strYM
			rsMQty.Fields("JGYOBU") = "A"
			rsMQty.Fields("NAIGAI") = "0"
			rsMQty.Fields("HIN_GAI") = strPn
			rsMQty.Fields("SyukaCnt") = "1"
			rsMQty.Fields("SyukaQty") = lngQty
			rsMQty.UpdateBatch
		end if
		Call DispMsg(makeMsg(strYM,-7) & makeMsg(lngRow,6) & ":" & makeMsg(strPn,-20) & makeMsg(lngQty,6) & strAdd)
		lngRow = lngRow + 1
	loop
End Function

Function LoopCheck(byval lRow,byval sPn)
	LoopCheck = False
	if sPn = "" then
		LoopCheck = True
	end if
	dim	lLimit
	lLimit = CLng(GetOption("limit",0))
	if lLimit > 0 then
		if lRow > lLimit then
			LoopCheck = True
		end if
	end if
End Function
