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

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "実績表報告(Excel)データ変換"
	Wscript.Echo "YOrder.vbs [option] <ファイル名> [シート名]"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo "sc32 YOrder.vbs /debug F:\★★★報告\実績表報告.xlsm 201406"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	dim	strSheetname
	strFilename = ""
	strSheetname = ""
	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		else
			strSheetname = strArg
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
	call LoadYOrder(strFilename,strSheetname)
	Main = 0
End Function

Function LoadYOrder(byVal strFilename,byVal strSheetname)
	Call Debug("LoadYOrder(" & strFilename & "," & strSheetname & ")")
	'-------------------------------------------------------------------
	'Excelファイル名
	'-------------------------------------------------------------------
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
	Call LoadYOrderXls(objXL,objBk,strSheetname)
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadYOrder():End")
End Function

Function LoadYOrderXls(objXL,objBk,byVal strSheetname)
	Call Debug("LoadYOrderXls(" & strSheetname & ")")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsYOrder
	set rsYOrder = OpenRs(objDb,"YOrder")

	Call Debug("LoadYOrderXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadYOrderXls():SheetName=" & strShtName)
		if strSheetname = "" or strSheetname = strShtName then
			Call LoadYOrderXst(objXL,objBk,objSt,objDb,rsYOrder)
		end if
	Next
	
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsYOrder = CloseRs(rsYOrder)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadYOrderXst(objXL,objBk,objSt,objDb,rsYOrder)
	Call Debug("LoadYOrderXst():SheetName=" & objSt.Name)

	dim	strYM
	strYM = chkYM(objSt.Name)
	if strYM = "" then
		Exit Function
	end if

	Call Debug("delete strYM:" & strYM)
	Call ExecuteAdodb(objDb,"delete from YOrder where YM = '" & strYM & "'")

	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"G",lngMaxRow)
	lngMaxRow = excelGetMaxRow(objSt,"H",lngMaxRow)
	dim	lngRow
	For lngRow = 5 to lngMaxRow
		Call LoadYOrderRow(strYM,objXL,objBk,objSt,objDb,rsYOrder,lngRow)
	Next
End Function

Function LoadYOrderRow(byVal strYM,objXL,objBk,objSt,objDb,rsYOrder,byVal lngRow)
	Call Debug("LoadYOrderRow():" & objSt.Name & ":" & lngRow)
	rsYOrder.AddNew
	rsYOrder.Fields("SName")	= objSt.Name					'// シート名
	rsYOrder.Fields("Row")		= lngRow						'// 行番号
	rsYOrder.Fields("YM")		= strYM							'//  処理年月(シート名)
	Call SetField(rsYOrder,objSt,"JDt",			"G" , lngRow)	'//  G 受注日
	Call SetField(rsYOrder,objSt,"DenNo",		"H" , lngRow)	'//  H 売上No
	Call SetField(rsYOrder,objSt,"SDt",			"I" , lngRow)	'//  I:J 売上日
	Call SetField(rsYOrder,objSt,"DDt",			"K" , lngRow)	'//  K:L 納入日
	Call SetField(rsYOrder,objSt,"OdCd1",		"O" , lngRow)	'//  O 発注元
	Call SetField(rsYOrder,objSt,"OdCd2",		"P" , lngRow)	'//  P 発注元
	Call SetField(rsYOrder,objSt,"TkName",		"R" , lngRow)	'//  R 得意先名
	Call SetField(rsYOrder,objSt,"Article1",	"S" , lngRow)	'//  S 工事名・件名
	Call SetField(rsYOrder,objSt,"Article2",	"T" , lngRow)	'//  T ・分譲物件名・マンション名・環境システム
	Call SetField(rsYOrder,objSt,"Amount",		"V" , lngRow)	'//  V 受注金額
	Call SetField(rsYOrder,objSt,"AmountEHN",	"W" , lngRow)	'//  W 受注金額(EHN)
	Call SetField(rsYOrder,objSt,"Place",		"X" , lngRow)	'//  X 納品場所
	rsYOrder.UpdateBatch
	LoadYOrderRow = lngRow
End Function

Private Function chkYM(byVal strYM)
	dim	a
	if strYM = "当月" then
		strYM = Year(Now()) & Right("0" & Month(Now()),2)
	end if
	if Len(strYM) <> 6 then
		strYM = ""
	end if
	if isNumeric(strYM) <> True then
		strYM = ""
	end if
	chkYM = strYM
End Function

Function SetField(objRs,objSt,byVal strField,byVal strCol,byVal lngRow)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & objSt.Range(strCol & lngRow))
	dim	v
	v = RTrim(objSt.Range(strCol & lngRow))
	select case strField
	case "JDt"
		v = Replace(v,"/","")
	case "SDt","DDt"	'//  I:J 売上日 '//  K:L 納入日
		v = v & "/" & objSt.Range(strCol & lngRow).Offset(0,1)
		if isDate(v) then
			v = CDate(v)
			v = Replace(v,"/","")
		else
			v = ""
		end if
	case "Amount","AmountEHN"
		if isNumeric(v) <> True then
			v = 0
		end if
		v = CCur(v)
	end select
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & v)
	objRs.Fields(strField) = v
End Function
