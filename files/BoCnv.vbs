Option Explicit
Const xlUp = -4162
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "BoCnv.vbs [option]"
	Wscript.Echo " /db:newsdc1	データベース"
	Wscript.Echo " /j:4			事業部"
	Wscript.Echo " /s:10000		開始行"
	Wscript.Echo " /l:100		読み込む行数"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc1 /j:4 bo\炊飯カテゴリー.xls"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc1 /j:5 bo\品目カテゴリー.xls"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc3 bo\品薄リスト補足データ.xlsx"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc3 bo\サファイア.xlsx"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc4 00025800.xls"
End Sub

'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objBoCnv
	Set objBoCnv = New BoCnv
	if objBoCnv.Init() <> "" then
		call usage()
		exit function
	end if
	call objBoCnv.Run()
End Function

'-----------------------------------------------------------------------
'BoCnv
'-----------------------------------------------------------------------
Class BoCnv
	Private	strDBName
	Private	objDB
	Private	objRs
	Public	strJGYOBU
	Private	strAction
	Private	strFileName
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
	'-----------------------------------------------------------------------
	Function GetOption(byval strName _
					  ,byval strDefault _
					  )
		dim	strValue

		if strName = "" then
			strValue = ""
			if strDefault < WScript.Arguments.UnNamed.Count then
				strValue = WScript.Arguments.UnNamed(strDefault)
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		strAction = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "j"
			case "s"
			case "l"
			case "debug"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'CheckFunction
	'-----------------------------------------------------------------------
	Private Function CheckFunction(byval strA)
		Debug ".CheckFunction():" & strA
		CheckFunction = False
		if strAction = "" then
			exit function
		end if
		if WScript.Arguments.Named.Exists(strA) then
			exit function
		end if
		CheckFunction = True
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		strJGYOBU = GetOption("j"	,"4")
		set objDB = nothing
		set objRs = nothing
		set objBk = nothing
		set objXL = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		if not objBk is nothing then
			Call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
		set objRs = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			Debug ".Run():" & strArg
			strFileName = strArg
			Call Load()
		Next
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load():" & strFileName
		select case FileType()
		case "excel"
			Call CreateExcelApp()
			Call OpenExcel()
			Call LoadExcel()
		case "csv"
'			Call OpenCsv()
'			Call LoadCsv()
'			Call CloseCsv()
		end select
	End Function
	'-------------------------------------------------------------------
	'ファイルの種類
	'-------------------------------------------------------------------
	Private Function FileType()
		FileType = ""
		select case lcase(fileExt(strFileName))
		case "xls","xlsx"	FileType = "excel"
		case "csv"			FileType = "csv"
		end select
		Debug(".FileType():" & FileType)
	End Function
	'-------------------------------------------------------------------
	'拡張子
	'-------------------------------------------------------------------
	Private Function fileExt(byVal f)
		dim	fobj
		set fobj = CreateObject("Scripting.FileSystemObject")
		dim	strExt
		strExt = fobj.GetextensionName(f)
		set fobj = Nothing
		fileExt = strExt
	End Function
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private	objXL
	Private Function CreateExcelApp()
		Debug(".CreateExcelApp()")
		if objXL is nothing then
			Debug(".CreateExcelApp():CreateObject(Excel.Application)")
			Set objXL = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private	objBk
	Private Function OpenExcel()
		Debug(".OpenExcel()")
		if objBk is nothing then
			Debug(".OpenExcel().Open=" & GetAbsPath(strFileName))
			Set objBk = objXL.Workbooks.Open(GetAbsPath(strFileName),False,True,,"")
		end if
	end function
	'-------------------------------------------------------------------
	'絶対パス
	'-------------------------------------------------------------------
	Private Function GetAbsPath(byVal strPath)
		Dim objFileSys
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		strPath = objFileSys.GetAbsolutePathName(strPath)
		Set objFileSys = Nothing
		GetAbsPath = strPath
	End Function
	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Public	objSt
	Private Function LoadExcel()
		Debug ".LoadExcel()"
		if objBk is nothing then
			exit function
		end if
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
	end function
	'-------------------------------------------------------------------
	'読込処理(シート)
	'-------------------------------------------------------------------
	Private Function LoadXls()
		Debug ".LoadXls()"
		if objSt is nothing then
			exit function
		end if
		Call LoadData()
	end function
	'-------------------------------------------------------------------
	'シート読込
	'-------------------------------------------------------------------
	Private	clsData
	Private Function LoadData()
		Debug ".LoadData():" & objSt.Name
		' 品目カテゴリー/炊飯カテゴリー
		Set clsData = New BoHinmoku
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
		end if
		' 品薄リスト補足データ
		Set clsData = New BoHosoku
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
		end if
		' サファイア納入予定
		Set clsData = New SaDelv
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
		end if
		' Activeデータ
		Set clsData = New AcData
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
		end if
	End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
'		on error resume next
		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field名
	'-------------------------------------------------------------------
	Public Function GetFields(byVal strTable)
		Debug ".GetFields():" & strTable
		dim	strFields
		strFields = ""
		dim	objRs
		set objRS = objDB.Execute("select top 1 * from " & strTable)
		dim	objF
		for each objF in objRS.Fields
			if strFields <> "" then
				strFields = strFields & ","
			end if
			strFields = strFields & objF.Name
		next
		set objRs = nothing
		GetFields = strFields
	End Function
End Class

' タイトル文字列
Private Function getTitle(byVal strT)
	getTitle = Replace(strT,vbLf,"")
End Function

' タイトル比較
Private Function CompTitle(byVal strS,byVal strD)
	CompTitle = true
	if getTitle(strS) = getTitle(strD) then
		CompTitle = false
	end if
End Function

' Excel最終行
Private Function excelGetMaxRow(objSt,byVal strCol,byVal lngRow)
	dim lngRowMax
	lngRowMax = objSt.rows.count
	lngRowMax = objSt.Range(strCol & lngRowMax).End(xlUp).Row
	if lngRow > lngRowMax then
		lngRowMax = lngRow
	end if
	excelGetMaxRow = lngRowMax
End Function

' 品目カテゴリー/炊飯カテゴリー
Class BoHinmoku
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		'PN共通_品目番号	品目_品目コード	品目_品目カテゴリー名
		'PN共通_品目番号	品目_品目コード	品目_品目カテゴリー別名
		if	getTitle(objSt.Range("A4")) <> "" then
			exit function
		end if
		if	getTitle(objSt.Range("B4")) <> "PN共通_品目番号" then
			exit function
		end if
		if	getTitle(objSt.Range("C4")) <> "品目_品目コード" then
			exit function
		end if
		select case getTitle(objSt.Range("D4"))
		case "品目_品目カテゴリー名","品目_品目カテゴリー別名"
		case else
			exit function
		end select
		Call Load()
		Init = true
	End Function
	Public Function Load()
		lngRowTop = objParent.GetOption("s",5)
		lngRowEnd = excelGetMaxRow(objSt,"B",5)
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
		next
	End Function
	Private	strJGYOBU
	Private	strHIN_GAI
	Private	strHinmokuCode
	Private	strHinmokuName
	Public Function Disp()
		strJGYOBU		= objParent.strJGYOBU
		strHIN_GAI		= objSt.Range("B" & lngRow)
		strHinmokuCode	= objSt.Range("C" & lngRow)
		strHinmokuName	= objSt.Range("D" & lngRow)
		objParent.Disp	lngRow & "/" & lngRowEnd	_
				& " " & strJGYOBU	_
				& " " & strHIN_GAI	_
				& " " & strHinmokuCode	_
				& " " & strHinmokuName
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into PnHinmoku"
		strSql = strSql & " (JGYOBU"
		strSql = strSql & " ,HIN_GAI"
		strSql = strSql & " ,HinmokuCode"
		strSql = strSql & " ,EntID"
		strSql = strSql & " ) values ("
		strSql = strSql & "  '" & strJGYOBU & "'" 
		strSql = strSql & " ,'" & strHIN_GAI & "'"
		strSql = strSql & " ,'" & strHinmokuCode & "'"
		strSql = strSql & " ,'BoCnv'"
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
		strSql = strSql & "insert into Hinmoku"
		strSql = strSql & " (JGYOBU"
		strSql = strSql & " ,HinmokuCode"
		strSql = strSql & " ,HinmokuName"
		strSql = strSql & " ,EntID"
		strSql = strSql & " ) values ("
		strSql = strSql & "  '" & strJGYOBU & "'" 
		strSql = strSql & " ,'" & strHinmokuCode & "'"
		strSql = strSql & " ,'" & strHinmokuName & "'"
		strSql = strSql & " ,'BoCnv'"
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
	End Function
End Class

' 品薄リスト補足データ
Class BoHosoku
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		'PN共通_資産管理事業場コード	PN共通_品目番号	PN名称_品目名	PN共通（PN)_代表機種品目コード	機種品目_機種品目カテゴリー名	PN共通_品目コード	品目_品目カテゴリー名	ＰＮ供期_国内供給開始年月	ＰＮ供期_国内生産打切年月	PN共通_国内供給部品区分	ＰＮ供期_輸出供給開始年月	ＰＮ供期_輸出生産打切年月	PN共通_海外供給部品区分	PN共通_備考欄
		if	getTitle(objSt.Range("A4")) <> "" then
			exit function
		end if
		if	getTitle(objSt.Range("B4")) <> "PN共通_資産管理事業場コード" then
			exit function
		end if
		if	getTitle(objSt.Range("C4")) <> "PN共通_品目番号" then
			exit function
		end if
		if	getTitle(objSt.Range("D4")) <> "PN名称_品目名" then
			exit function
		end if
		if	getTitle(objSt.Range("E4")) <> "PN共通（PN)_代表機種品目コード" then
			exit function
		end if
		if	getTitle(objSt.Range("F4")) <> "機種品目_機種品目カテゴリー名" then
			exit function
		end if
		if	getTitle(objSt.Range("G4")) <> "PN共通_品目コード" then
			exit function
		end if
		if	getTitle(objSt.Range("H4")) <> "品目_品目カテゴリー名" then
			exit function
		end if
		if	getTitle(objSt.Range("I4")) <> "ＰＮ供期_国内供給開始年月" then
			exit function
		end if
		if	getTitle(objSt.Range("J4")) <> "ＰＮ供期_国内生産打切年月" then
			exit function
		end if
		if	getTitle(objSt.Range("K4")) <> "PN共通_国内供給部品区分" then
			exit function
		end if
		if	getTitle(objSt.Range("L4")) <> "ＰＮ供期_輸出供給開始年月" then
			exit function
		end if
		if	getTitle(objSt.Range("M4")) <> "ＰＮ供期_輸出生産打切年月" then
			exit function
		end if
		if	getTitle(objSt.Range("N4")) <> "PN共通_海外供給部品区分" then
			exit function
		end if
		if	getTitle(objSt.Range("O4")) <> "PN共通_備考欄" then
			exit function
		end if
		' データ読込
		Call Load()
		Init = true
	End Function
	Private	lngLimit
	Private	lngCount
	Public Function Limit()
		objParent.Debug ".Limit():" & lngCount & "/" & lngLimit
		Limit = false
		lngCount = lngCount + 1
		if lngLimit <> 0 then
			if lngCount >= lngLimit then
				Limit = true
			end if
		end if
	End Function
	Private	strFields
	Public Function Load()
		strFields = objParent.GetFields("PnHosoku")
		lngRowTop = objParent.GetOption("s",5)
		lngRowEnd = excelGetMaxRow(objSt,"B",5)
		lngCount  = 0
		lngLimit  = CLng(objParent.GetOption("l",0))
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
			if Limit() then
				exit for
			end if
		next
	End Function
	Private	aryCell(255)
	Public Function Disp()
		dim	i
		dim	strMsg
		strMsg = lngRow & "/" & lngRowEnd
		dim	strValues
		strValues = ""
		for i = 1 to 14
			aryCell(i) = objSt.Range("A" & lngRow).Offset(0,i)
			strMsg = strMsg & " " & aryCell(i)
			if strValues <> "" then
				strValues = strValues & ","
			end if
			strValues = strValues & "'" & aryCell(i) & "'"
		next
		objParent.Disp	strMsg
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into PnHosoku"
		strSql = strSql & " ("
		strSql = strSql & "  ShisanJCode" 	
		strSql = strSql & " ,Pn" 			
		strSql = strSql & " ,PName" 			
		strSql = strSql & " ,DModel" 		
		strSql = strSql & " ,DModelName"		
		strSql = strSql & " ,Hinmoku" 		
		strSql = strSql & " ,HinmokuName"	
		strSql = strSql & " ,NaiSupplyYm"	
		strSql = strSql & " ,NaiBldOutYm"	
		strSql = strSql & " ,NaiKbn" 		
		strSql = strSql & " ,GaiSupplyYm"	
		strSql = strSql & " ,GaiBldOutYm"	
		strSql = strSql & " ,GaiKbn" 		
		strSql = strSql & " ,Biko" 			
		strSql = strSql & " ) values ("
		strSql = strSql & strValues
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
'		Call objParent.CallSql(strSql)
	End Function
End Class

'対象年月	年月
'出庫先/払出元	91H
'予定/実績	予定
'指定日付	2016年05月01日 ～2016年12月31日
'正式/仮	正式のみ
'過剰	全て
'数量/金額	全て
'部品コード	部品名	調達年月度	調達№	予定日付	AMPM	時刻	実績日付	伝票№	取引コード	仕入/支給先	納入場所	予定数	実績数	出庫先/払出元	通貨コード	仕入/支給単価	金額	改定理由1	注文№	保管場所	在庫区分	対応伝票№	資材ライン	通貨コード	直送支給単価	改定理由2	検査日付	検査区分	日計区分	補修バイヤー	先方部品コード	環境区分	予約番号	予約時調達年月	予約時調達№	予約時変更№	Daily区分	納期ランク	国籍コード	都道府県コード	原産国コード	PDM機種コード	色区分	仕様区分	成形ユニット区分	ユーザー	ユーザー名	更新日付
' サファイア納入予定
Class SaDelv
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		dim	strTitle
		strTitle = "部品コード	部品名	調達年月度	調達№	予定日付	AMPM	時刻	実績日付	伝票№	取引コード	仕入/支給先	納入場所	予定数	実績数	出庫先/払出元	通貨コード	仕入/支給単価	金額	改定理由1	注文№	保管場所	在庫区分	対応伝票№	資材ライン	通貨コード	直送支給単価	改定理由2	検査日付	検査区分	日計区分	補修バイヤー	先方部品コード	環境区分	予約番号	予約時調達年月	予約時調達№	予約時変更№	Daily区分	納期ランク	国籍コード	都道府県コード	原産国コード	PDM機種コード	色区分	仕様区分	成形ユニット区分	ユーザー	ユーザー名	更新日付"
		dim	aryTitle
		aryTitle = Split(strTitle,vbTab)
		dim	objR
		for each objR in objSt.Range("A10:AW10")
			objParent.Debug ".Init():" & objR & ":" & objR.Column & ":" & aryTitle(objR.Column - 1)
			if CompTitle(objR,aryTitle(objR.Column - 1)) then
				exit function
			end if
		next
		' データ読込
		Call Load()
		Init = true
	End Function
	Private	lngLimit
	Private	lngCount
	Public Function Limit()
		objParent.Debug ".Limit():" & lngCount & "/" & lngLimit
		Limit = false
		lngCount = lngCount + 1
		if lngLimit <> 0 then
			if lngCount >= lngLimit then
				Limit = true
			end if
		end if
	End Function
	Private	strFields
	Public Function Load()
		strFields = objParent.GetFields("SaDelv")
		strFields = Replace(strFields,",EntID,EntTm,UpdID,UpdTm","")
		lngRowTop = objParent.GetOption("s",11)
		lngRowEnd = excelGetMaxRow(objSt,"A",11)
		lngCount  = 0
		lngLimit  = CLng(objParent.GetOption("l",0))
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
			if Limit() then
				exit for
			end if
		next
	End Function
	Private	aryCell(255)
	Public Function Disp()
		dim	i
		dim	strMsg
		strMsg = lngRow & "/" & lngRowEnd
		dim	strValues
		strValues = ""
		dim	objR
		for each objR in objSt.Range("A" & lngRow & ":AW" & lngRow)
			strMsg = strMsg & " " & objR
			if strValues <> "" then
				strValues = strValues & ","
			end if
			strValues = strValues & "'" & objR & "'"
		next
		objParent.Disp	strMsg
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into SaDelv"
		strSql = strSql & " ("
		strSql = strSql & strFields
		strSql = strSql & " ) values ("
		strSql = strSql & strValues
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
'		Call objParent.CallSql(strSql)
	End Function
End Class

' Activeデータ
Class AcData
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		dim	strTitle
		strTitle = "ID-NO	事業場コード	資産管理事業場コード	販売区分	品目番号	得意先コード	得意先略称	直送相手先コード	進捗　サービスデータ進捗区分	収略　在庫収支略式名	出荷予定数	前回納期回答年月日	購買担当納期回答年月日	購買担当自由使用欄	納期回答データ送信区分	納期回答データ送信回数	出荷予定年月日	出庫予定年月日	受注年月日	指定納期日　指定納期年月日	出荷指定年月日	納期回答日　納期回答年月日	帳端区分	帳端区分再付与	オーダーNO	ITEM-NO	伝票番号	サービス販売ルートコード	注文区分	仕入先ワークセンターコード	購買担当　購買担当者コード	納期回答データ訂正区分	納期回答自動付与区分	納期回答自由使用欄	登録ユーザｰID	登録日付	登録時刻	更新ユーザｰID	更新日付	更新時刻"
		dim	aryTitle
		aryTitle = Split(strTitle,vbTab)
		dim	objR
		for each objR in objSt.Range("A1:AN1")
			objParent.Debug ".Init():" & objR & ":" & objR.Column & ":" & aryTitle(objR.Column - 1)
			if CompTitle(objR,aryTitle(objR.Column - 1)) then
				exit function
			end if
		next
		' データ読込
		Call Load()
		Init = true
	End Function
	Private	lngLimit
	Private	lngCount
	Public Function Limit()
		objParent.Debug ".Limit():" & lngCount & "/" & lngLimit
		Limit = false
		lngCount = lngCount + 1
		if lngLimit <> 0 then
			if lngCount >= lngLimit then
				Limit = true
			end if
		end if
	End Function
	Private	strFields
	Private	aryFields
	Public Function Load()
		strFields = objParent.GetFields("A_Data")
		strFields = Replace(strFields,",EntID,EntTm,UpdID,UpdTm","")
		aryFields = Split(strFields,",")
		lngRowTop = objParent.GetOption("s",3)
		lngRowEnd = excelGetMaxRow(objSt,"A",3)
		lngCount  = 0
		lngLimit  = CLng(objParent.GetOption("l",0))
		dim	strSql
		strSql = ""
		strSql = strSql & "delete from A_Data"
		Call objParent.CallSql(strSql)
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
			if Limit() then
				exit for
			end if
		next
	End Function
	Private	aryCell(255)
	Public Function Disp()
		dim	i
		dim	strMsg
		strMsg = lngRow & "/" & lngRowEnd
		dim	strValues
		strValues = ""
		dim	objR
		i = 0
		for each objR in objSt.Range("A" & lngRow & ":AN" & lngRow)
			objParent.Debug ".Load():" & aryFields(i) & ":" & objR
			strMsg = strMsg & " " & objR
			if strValues <> "" then
				strValues = strValues & ","
			end if
			dim	strV
			strV = objR
			strV = Replace(strV,"'","")
			if Right(aryFields(i),2) = "Dt" then
				strV = Replace(strV,"/","")
			end if
			dim	strQ
			strQ = "'"
			if Right(aryFields(i),3) = "Qty" then
				strQ = ""
			end if
			strValues = strValues & strQ & strV & strQ
			i = i + 1
		next
		objParent.Disp	strMsg
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into A_Data"
		strSql = strSql & " ("
		strSql = strSql & strFields
		strSql = strSql & " ) values ("
		strSql = strSql & strValues
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
	End Function
End Class

