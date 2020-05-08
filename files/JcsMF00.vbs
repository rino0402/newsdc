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
Call Include("file.vbs")
Call Include("get_b.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "JCS MF(Excel)データ変換"
	Wscript.Echo "JcsMF.vbs [option] <ファイル名> [シート名]"
	Wscript.Echo " /db:newsdc7"
	Wscript.Echo " /debug"
	Wscript.Echo "sc32 JcsMF.vbs jcs\ＭＦ.xls 部品ＭＦ /debug"
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
	call LoadMF(strFilename,strSheetname)
	Main = 0
End Function

Function LoadMF(byVal strFilename,byVal strSheetname)
	'-------------------------------------------------------------------
	'Excelファイル名
	'-------------------------------------------------------------------
	strFilename = GetAbsPath(strFilename)
	Call Debug("LoadMF(" & strFilename & "," & strSheetname & ")")
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
	Call LoadMFXls(objXL,objBk,strSheetname)
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadMF():End")
End Function

Function JcsTableName(byVal strSheetname)
	dim	strTableName
	strTableName = ""
	select case strSheetname
	case "部品ＭＦ"
		strTableName = "JcsItem"
	end select
	JcsTableName = strTableName
End Function
Function LoadMFXls(objXL,objBk,byVal strSheetname)
	Call Debug("LoadMFXls(" & strSheetname & ")")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsJcsItem
	set rsJcsItem = OpenRs(objDb,JcsTableName(strSheetname))

	Call Debug("LoadMFXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadMFXls():SheetName=" & strShtName)
		if strSheetname = "" or strSheetname = strShtName then
			Call LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem)
		end if
	Next
	
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsJcsItem = CloseRs(rsJcsItem)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem)
	Call Debug("LoadMFXst():SheetName=" & objSt.Name)

	Call Debug("delete from " & JcsTableName(objSt.Name))
	Call ExecuteAdodb(objDb,"delete from " & JcsTableName(objSt.Name))

	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"C",lngMaxRow)
	dim	lngRow
	For lngRow = 3 to lngMaxRow
		Call LoadMFRow(objXL,objBk,objSt,objDb,rsJcsItem,lngRow)
	Next
End Function

Function LoadMFRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
	Call Debug("LoadMFRow():" & objSt.Name & ":" & lngRow)
	Call rsJcsItem.AddNew
	Call SetField(rsJcsItem,objSt,"MazdaPn",		"A" , lngRow)	'// マツダ品番 
	Call SetField(rsJcsItem,objSt,"NameE",			"B" , lngRow)	'// 品名（英語） 
	Call SetField(rsJcsItem,objSt,"Pn",				"C" , lngRow)	'// ＪＣＳ品番 
	Call SetField(rsJcsItem,objSt,"NameJ",			"D" , lngRow)	'// 品名（日本語） 
	Call SetField(rsJcsItem,objSt,"Color",			"E" , lngRow)	'// 色 
	Call SetField(rsJcsItem,objSt,"LabelType",		"F" , lngRow)	'// ラベル MorF
	Call SetField(rsJcsItem,objSt,"LabelCut",		"G" , lngRow)	'//  切断
	Call SetField(rsJcsItem,objSt,"Location",		"H" , lngRow)	'// 在庫 9999-XX
	Call SetField(rsJcsItem,objSt,"SSpec",			"J" , lngRow)	'// 商品化 仕様書
	Call SetField(rsJcsItem,objSt,"SType",			"K" , lngRow)	'//  タイプ
	Call SetField(rsJcsItem,objSt,"GPn",			"L" , lngRow)	'// 外装 品番
	Call SetField(rsJcsItem,objSt,"GQty",			"M" , lngRow)	'//  入数
	Call SetField(rsJcsItem,objSt,"LastPn",			"N" , lngRow)	'// 最終荷姿 品番
	Call SetField(rsJcsItem,objSt,"LastQty",		"O" , lngRow)	'//  入数
	Call SetField(rsJcsItem,objSt,"CheckConf",		"P" , lngRow)	'// チェック確認 
	Call SetField(rsJcsItem,objSt,"Location1",		"Q" , lngRow)	'// ロケーション 棚番
	Call SetField(rsJcsItem,objSt,"Location2",		"R" , lngRow)	'//  別置
	Call SetField(rsJcsItem,objSt,"ShareNum",		"S" , lngRow)	'// 共用 件数
	Call SetField(rsJcsItem,objSt,"ShareNo",		"T" , lngRow)	'//  No.
	Call SetField(rsJcsItem,objSt,"AlterDate",		"Y" , lngRow)	'// 登録 月／日
	Call SetField(rsJcsItem,objSt,"AlterPerson",	"Z" , lngRow)	'//  登録者
	Call SetField(rsJcsItem,objSt,"LastShipDate",	"AA" , lngRow)	'// 前回出荷実績 月／日
	Call SetField(rsJcsItem,objSt,"LastShipQty",	"AB" , lngRow)	'//  数量
	Call SetField(rsJcsItem,objSt,"LastShipPn",		"AC" , lngRow)	'//  ＪＣＳ品番③
	Call SetField(rsJcsItem,objSt,"CoStockDate",	"AD" , lngRow)	'// 繰り越し在庫 月／日
	Call SetField(rsJcsItem,objSt,"CoStockQty",		"AE" , lngRow)	'//  数量
	On Error Resume Next
		Call rsJcsItem.UpdateBatch
		if Err <> 0 Then
			WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
			' ﾚｺｰﾄﾞのｷｰ ﾌｨｰﾙﾄﾞに重複するｷｰ値があります(Btrieve Error 5)
			if Err.Number <> -2147467259 then
				lngRow = 0
			end if
			Call rsJcsItem.CancelUpdate
		end if
	On Error GoTo 0
	LoadMFRow = lngRow
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
	v = Get_LeftB(v,objRs.Fields(strField).DefinedSize)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & v)
	objRs.Fields(strField) = v
End Function

Class JcsItem
	Public pTableName
    Private Sub Class_Initialize
        pTableName = "JcsItem"
    End Sub
End Class
