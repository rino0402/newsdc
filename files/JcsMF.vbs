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
	Wscript.Echo "cscript JcsMF.vbs jcs\ＭＦ.xls 部品ＭＦ /debug"
	Wscript.Echo "cscript JcsMF.vbs jcs\ＭＦ.xls タイプＭＦ /debug"
	Wscript.Echo "cscript JcsMF.vbs jcs\ＤＦ.xls 移動ＤＦ /debug"
	Wscript.Echo "cscript JcsMF.vbs jcs\ＤＦ.xls 受注ＤＦ /debug"
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

Function LoadMFXls(objXL,objBk,byVal strSheetname)
	Call Debug("LoadMFXls(" & strSheetname & ")")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'JCSクラス
	'-------------------------------------------------------------------
	dim	objJcs
	select case strSheetname
	case "部品ＭＦ"
		Set objJcs = New JcsItem
	case "タイプＭＦ"
		Set objJcs = New JcsType
 	case "移動ＤＦ"
		Set objJcs = New JcsIdo
 	case "受注ＤＦ"
		Set objJcs = New JcsOrder
	end select
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsJcsItem
	set rsJcsItem = OpenRs(objDb,objJcs.pTableNameTmp)

	Call Debug("LoadMFXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadMFXls():SheetName=" & strShtName)
		if strSheetname = "" or strSheetname = strShtName then
			Call LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem,objJcs)
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

Function LoadMFXst(objXL,objBk,objSt,objDb,rsJcsItem,objJcs)
	Call Debug("LoadMFXst():SheetName=" & objSt.Name)

	Call objJcs.InitRecord(objDb)

	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"C",lngMaxRow)
	dim	lngRow
	For lngRow = 3 to lngMaxRow
		Call DispMsg("登録中..." & objSt.Name & ":" & lngRow & "/" & lngMaxRow)
		Call objJcs.LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,lngRow)
	Next
	Call DispMsg("テーブルコピー..." & objJcs.pTableNameTmp & "→" & objJcs.pTableName)
	Call CopyTable(objDb,objJcs.pTableNameTmp,objJcs.pTableName)
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

'-------------------------------------------------------
' 部品ＭＦ
'-------------------------------------------------------
Class JcsItem
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsItem"
        pTableNameTmp	= "JcsItem_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
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
		Call SetField(rsJcsItem,objSt,"Sagyo",			"X" , lngRow)	'// 作業標準書
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
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' タイプＭＦ
'-------------------------------------------------------
Class JcsType
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsType"
        pTableNameTmp	= "JcsType_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 6 then
			LoadRow = lngRow
			exit function
		end if
		Call rsJcsItem.AddNew
		rsJcsItem("XlsRow") = lngRow							'// Excel行
		Call SetField(rsJcsItem,objSt,"SType",		"A" , lngRow)	'// タイプ
		Call SetField(rsJcsItem,objSt,"SPn1",		"B" , lngRow)	'// 資材品番１
		Call SetField(rsJcsItem,objSt,"SQty1",		"C" , lngRow)	'// 資材員数１
		Call SetField(rsJcsItem,objSt,"SPn2",		"D" , lngRow)	'// 資材品番２
		Call SetField(rsJcsItem,objSt,"SQty2",		"E" , lngRow)	'// 資材員数２
		Call SetField(rsJcsItem,objSt,"SPn3",		"F" , lngRow)	'// 資材品番３
		Call SetField(rsJcsItem,objSt,"SQty3",		"G" , lngRow)	'// 資材員数３
		Call SetField(rsJcsItem,objSt,"SPn4",		"H" , lngRow)	'// 資材品番４
		Call SetField(rsJcsItem,objSt,"SQty4",		"I" , lngRow)	'// 資材員数４
                                                                        
		Call SetField(rsJcsItem,objSt,"PUnit",		"K" , lngRow)	'// 梱包単位
		Call SetField(rsJcsItem,objSt,"MCP1",		"L" , lngRow)	'// 材料費１
		Call SetField(rsJcsItem,objSt,"MCP2",		"M" , lngRow)	'// 材料費２

		Call SetField(rsJcsItem,objSt,"PCP1",		"O" , lngRow)	'// 加工費１
		Call SetField(rsJcsItem,objSt,"PCP2",		"P" , lngRow)	'// 加工費２
                                                                        
		Call SetField(rsJcsItem,objSt,"OCP1",		"R" , lngRow)	'// その他１
		Call SetField(rsJcsItem,objSt,"OCP2",		"S" , lngRow)	'// その他２
                                                                        
                                                                        
		Call SetField(rsJcsItem,objSt,"AlterDate",	"V" , lngRow)	'// 登録日
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
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' 移動ＤＦ
'-------------------------------------------------------
Class JcsIdo
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsIdo"
        pTableNameTmp	= "JcsIdo_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		Call rsJcsItem.AddNew
		rsJcsItem.Fields("ID") = lngRow									'// 行番号
		Call SetField(rsJcsItem,objSt,"OrderDt",		"A" , lngRow)	'// 受注 月日
		Call SetField(rsJcsItem,objSt,"NohinNo",		"B" , lngRow)	'// 納品番号 
		Call SetField(rsJcsItem,objSt,"NohinNo2",		"C" , lngRow)	'//  ②
		Call SetField(rsJcsItem,objSt,"Location",		"D" , lngRow)	'// 部品 庫
		Call SetField(rsJcsItem,objSt,"DlvDt",			"E" , lngRow)	'// 納期 
		Call SetField(rsJcsItem,objSt,"MazdaPn",		"F" , lngRow)	'// マツダ品番 
		Call SetField(rsJcsItem,objSt,"Pn",				"H" , lngRow)	'// q 
		Call SetField(rsJcsItem,objSt,"Qty",			"I" , lngRow)	'// 指示数 
		Call SetField(rsJcsItem,objSt,"L46",			"K" , lngRow)	'// L46 
		Call SetField(rsJcsItem,objSt,"Dt",				"L" , lngRow)	'// 月日 
		Call SetField(rsJcsItem,objSt,"DestCode",		"M" , lngRow)	'// 受払先or処理 CODE
		Call SetField(rsJcsItem,objSt,"DestName",		"N" , lngRow)	'//  受払先名or項目
		Call SetField(rsJcsItem,objSt,"IQty",			"O" , lngRow)	'// 移動数量 入庫
		Call SetField(rsJcsItem,objSt,"OQty",			"P" , lngRow)	'//  出庫
		Call SetField(rsJcsItem,objSt,"SSpec",			"Q" , lngRow)	'// 商品化 Ｎ
		Call SetField(rsJcsItem,objSt,"SType",			"R" , lngRow)	'//  タイプ
		Call SetField(rsJcsItem,objSt,"SPrice",			"S" , lngRow)	'//  単価
		Call SetField(rsJcsItem,objSt,"SAmont",			"T" , lngRow)	'//  金額
		Call SetField(rsJcsItem,objSt,"GPn",			"U" , lngRow)	'// 外装 品番
		Call SetField(rsJcsItem,objSt,"GQty",			"V" , lngRow)	'//  数量
		Call SetField(rsJcsItem,objSt,"LastPn",			"W" , lngRow)	'// 最終荷姿 品番
		Call SetField(rsJcsItem,objSt,"LastQty",		"X" , lngRow)	'//  数量
		Call rsJcsItem.UpdateBatch
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' 受注ＤＦ
'-------------------------------------------------------
Class JcsOrder
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "JcsOrder"
        pTableNameTmp	= "JcsOrder_Tmp"
    End Sub
    Public Sub InitRecord(objDb)
		Call Debug("delete from " & pTableNameTmp)
		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsJcsItem,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 11 then
			LoadRow = lngRow
			exit function
		end if
		Call rsJcsItem.AddNew
		rsJcsItem.Fields("XlsRow")	= lngRow
		Call SetField(rsJcsItem,objSt,	"OrderDt",	"A", lngRow)	'//	受注日
		Call SetField(rsJcsItem,objSt,	"NohinNo",	"B", lngRow)	'//	納品番号①
		Call SetField(rsJcsItem,objSt,	"NohinNo2",	"C", lngRow)	'//	納品番号②
		Call SetField(rsJcsItem,objSt,	"PartWH",	"D", lngRow)	'//	部品庫
		Call SetField(rsJcsItem,objSt,	"DlvDt",	"E", lngRow)	'//	納期
		Call SetField(rsJcsItem,objSt,	"MazdaPn",	"F", lngRow)	'//	マツダ品番
							
		Call SetField(rsJcsItem,objSt,	"Pn",		"H", lngRow)	'//	JCS品番
		Call SetField(rsJcsItem,objSt,	"Qty",		"I", lngRow)	'//	指示数
							
		Call SetField(rsJcsItem,objSt,	"Location",	"K", lngRow)	'//	在庫場所
		Call SetField(rsJcsItem,objSt,	"MPn",		"L", lngRow)	'//	マスター登録JCS品番
		Call SetField(rsJcsItem,objSt,	"InfoDlvDt","M", lngRow)	'//	納品情報 納期
		Call SetField(rsJcsItem,objSt,	"InfoTotal","N", lngRow)	'//	類型
		Call SetField(rsJcsItem,objSt,	"InfoZan",	"O", lngRow)	'//	残
		Call SetField(rsJcsItem,objSt,	"NG",		"P", lngRow)	'//	NG
		Call SetField(rsJcsItem,objSt,	"Biko",		"Q", lngRow)	'//	備考
		Call SetField(rsJcsItem,objSt,	"No",		"R", lngRow)	'//	No
		Call rsJcsItem.UpdateBatch
		LoadRow = lngRow
	End Function
End Class
