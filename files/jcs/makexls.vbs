Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "makexls.vbs [option]"
	Wscript.Echo " /db:newsdc	:データベース"
	Wscript.Echo " /a:10		:追加用dummy件数(default:0)"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript makexls.vbs /db:newsdc7 l164157.csv /a:10"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objJcsJoin
	Set objJcsJoin = New JcsJoin
	if objJcsJoin.Init() <> "" then
		call usage()
		exit function
	end if
	call objJcsJoin.Run()
End Function
'-----------------------------------------------------------------------
'JcsJoin
'-----------------------------------------------------------------------
Const xlEdgeTop		=	8
Const xlContinuous	=	1
Const xlThin		=	2
Class JcsJoin
	Private	strDBName
	Private	objDB
	Private	objRs
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
	Private strScriptPath
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		if WScript.Arguments.UnNamed.Count = 0 then
			Init = "ファイル未指定"
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "a"
			case "debug"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
		strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	End Function
	'-----------------------------------------------------------------------
	'CheckFunction
	'-----------------------------------------------------------------------
	Private Function CheckFunction(byval strA)
		Debug ".CheckFunction():" & strA
		CheckFunction = False
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
		set objDB = nothing
		set objRs = nothing
		set	objExcel = nothing
		set	objBook = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
		set	objBook = nothing
		set	objExcel = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			strFileName = strArg
			Call Load()
		Next
		Call CloseDb()
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
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
	Private	strBookName
    Public Function Load()
		Debug ".Load():" & strFileName
		strBookName = GetBaseName(strFileName) & ".xls"
		Debug ".Load():" & strBookName
		Call CreateExcel()
		Call OpenBook("引取リスト.xls")
		Call MakeBook()
		Call SaveBook("csv\" & strBookName)
		Call CloseBook()
	End Function
	'-----------------------------------------------------------------------
	'引取リスト作成
	'-----------------------------------------------------------------------
	Private	objSheet
    Public Function MakeBook()
		Debug ".MakeBook():" & objBook.Name
		set objSheet = objBook.ActiveSheet
		objSheet.Name = "引取リスト " & GetBaseName(strFilename)
		Debug ".MakeBook():" & objSheet.Name
		'行削除
		objSheet.Range("2:65536").Delete
		'問合せ
		Call OpenRs()
		'レコードセット
		Call SetData()
		'問合せClose
		Call CloseRs()
		'フィルター 数量<>0
'		Call objSheet.Range("K1").AutoFilter(1, "<>0")
		'フィルター後の書式セット
'		prvPn = ""
'		curPn = ""
'		lngRow = 2
'		do while true
'			curPn = objSheet.Range("G" & lngRow)
'			if curPn = "" then
'				exit do
'			end if
'			if objSheet.Rows(lngRow).hidden = false then
'				Call FormatRow()
'				prvPn = curPn
'			end if
'			lngRow = lngRow + 1
'		loop
	End Function
	'-----------------------------------------------------------------------
	'レコードセット
	'-----------------------------------------------------------------------
	Private	curPn
	Private	prvPn
	Private	lngRow
    Public Function SetData()
		Debug ".SetData()"
		if objRs is nothing then
			exit function
		end if
		prvPn = ""
		curPn = ""
		lngRow = 2
		do while objRs.Eof = false
			Call SetDataRow()
			curPn = GetField("Pn")
			Call FormatRow()
			Call YNyuka()
			prvPn = curPn
			lngRow = lngRow + 1
			objRs.Movenext
		loop
		curPn = ""
		'追加用Dummy
		dim	lngDummy
		lngDummy = GetOption("a",0) - 1
		dim	i
		for i=0 to lngDummy
			Call SetDummy(i)
			Call FormatRow()
			Call YNyukaDummy(i)
			lngRow = lngRow + 1
		next
		Call FormatRow()
	End Function
	'-----------------------------------------------------------------------
	'追加用ダミー行セット
	'-----------------------------------------------------------------------
    Public Function SetDummy(byVal i)
		Debug ".SetDummy():" & lngRow & ":" & i
		i = i + 1
		dim	strID
		strID = "J"
		strID = strID & strSyukaYmd
		strID = strID & Right(GetBaseName(strFilename),6)
		strID = strID & "A"
		strID = strID & Right("00" & i,2)
		objSheet.Range("Q" & lngRow) = "*" & strID & "*"			 '受入ID
	End Function
	'-----------------------------------------------------------------------
	'YNyukaInsert
	'-----------------------------------------------------------------------
	Private	strIdNo
    Public Function YNyukaDummy(byVal i)
		Debug ".YNyukaDummy()"
		i = i + 1
		strTextNo = Right(GetBaseName(strFilename),6)
		strTextNo = strTextNo & "A"
		strTextNo = strTextNo & Right("00" & i,2)
		strIdNo = UCase(GetBaseName(strFilename))
		strIdNo = strIdNo & "A"
		strIdNo = strIdNo & Right("0000" & i,4)
		strSql = "insert into Y_NYUKA"
		strSql = strSql & " ("
		strSql = strSql & " DT_SYU"
		strSql = strSql & ",JGYOBU"
		strSql = strSql & ",NAIGAI"
		strSql = strSql & ",TEXT_NO"
		strSql = strSql & ",ID_NO"
		strSql = strSql & ",ID_NO2"
		strSql = strSql & ",SYUKO_YMD"
		strSql = strSql & ",SYUKA_YMD"
		strSql = strSql & ",MAEGARI_SURYO"
		strSql = strSql & ",INS_TANTO"
		strSql = strSql & " ) values ( "
		strSql = strSql & " '1'"			'DT_SYU"
		strSql = strSql & ",'J'"       		'JGYOBU"
		strSql = strSql & ",'1'"        	'NAIGAI"
		strSql = strSql & ",'" & strTextNo & "'"
		strSql = strSql & ",'" & strIdNo & "'"
		strSql = strSql & ",'" & strIdNo & "'"			'ID_NO2
		strSql = strSql & ",'" & strSyukaYmd & "'"
		strSql = strSql & ",'" & strSyukaYmd & "'"
		strSql = strSql & ",'00000000'"		'MAEGARI_SURYO"
		strSql = strSql & ",'makex'"		'INS_TANTO"
		strSql = strSql & " )"
		Debug ".YNyukaDummy():" & strSql
		on error resume next
			Call objDb.Execute(strSql)
			Debug ".YNyukaDummy():0x" & Hex(Err.Number) & ":" & Err.Description
		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'YNyuka
	'-----------------------------------------------------------------------
	Private	lngQty
	Private	strSyukaYmd
	Private	strTextNo
    Public Function YNyuka()
		Debug ".YNyuka()"
		if objRs is nothing then
			exit function
		end if
		dim	lngRcptQty
		lngRcptQty = CLng(GetField("RcptQty"))
		if curPn = prvPn then
			lngQty = lngQty + lngRcptQty
			Call YNyukaUpdate()
		else
			strSyukaYmd	= GetField(".SYUKA_YMD")
			strTextNo	= GetField(".TEXT_NO")
			lngQty = lngRcptQty
			Call YNyukaInsert()
		end if
	End Function
	'-----------------------------------------------------------------------
	'YNyukaUpdate
	'-----------------------------------------------------------------------
    Public Function YNyukaUpdate()
		Debug ".YNyukaUpdate()"
		strSql = "update Y_NYUKA"
		strSql = strSql & " set SURYO='" & lngQty & "'"
		strSql = strSql & " ,UPD_TANTO='makex'"
		strSql = strSql & " where JGYOBU='J'"
		strSql = strSql & "   and SYUKA_YMD='" & strSyukaYmd &  "'"
		strSql = strSql & "   and TEXT_NO='" & strTextNo &  "'"
		Debug ".YNyukaUpdate():" & strSql
'		on error resume next
			Call objDb.Execute(strSql)
			Debug ".YNyukaUpdate():0x" & Hex(Err.Number) & ":" & Err.Description
'		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'YNyukaInsert
	'-----------------------------------------------------------------------
    Public Function YNyukaInsert()
		Debug ".YNyukaInsert()"
		strSql = "insert into Y_NYUKA"
		strSql = strSql & " ("
		strSql = strSql & " DT_SYU"
		strSql = strSql & ",JGYOBU"
		strSql = strSql & ",NAIGAI"
		strSql = strSql & ",TEXT_NO"
		strSql = strSql & ",ID_NO"
		strSql = strSql & ",ID_NO2"
		strSql = strSql & ",HIN_NO"
		strSql = strSql & ",HIN_NAI"
		strSql = strSql & ",DEN_NO"
		strSql = strSql & ",SURYO"
		strSql = strSql & ",MUKE_CODE"
		strSql = strSql & ",SYUKO_YMD"
		strSql = strSql & ",SYUKA_YMD"
		strSql = strSql & ",HIN_NAME"
		strSql = strSql & ",NOUKI_YMD"
		strSql = strSql & ",SHIIRE_WORK_CENTER"
		strSql = strSql & ",MAEGARI_SURYO"
		strSql = strSql & ",INS_TANTO"
		strSql = strSql & " ) values ( "
		strSql = strSql & " '1'"			'DT_SYU"
		strSql = strSql & ",'J'"       		'JGYOBU"
		strSql = strSql & ",'1'"        	'NAIGAI"
		strSql = strSql & ",'" & GetField(".TEXT_NO") & "'"
		strSql = strSql & ",'" & GetField(".ID_NO") & "'"
		strSql = strSql & ",'" & GetField(".ID_NO") & "'"			'ID_NO2
		strSql = strSql & ",'" & GetField(".HIN_NO") & "'"
		strSql = strSql & ",'" & GetField(".HIN_NAI") & "'"
		strSql = strSql & ",'" & GetField(".DEN_NO") & "'"
		strSql = strSql & ",'" & lngQty & "'"
		strSql = strSql & ",'" & GetField(".MUKE_CODE") & "'"
		strSql = strSql & ",'" & GetField(".SYUKO_YMD") & "'"
		strSql = strSql & ",'" & GetField(".SYUKA_YMD") & "'"
		strSql = strSql & ",'" & GetField(".HIN_NAME") & "'"
		strSql = strSql & ",'" & GetField(".NOUKI_YMD") & "'"
		strSql = strSql & ",'" & GetField(".SHIIRE_WORK_CENTER") & "'"
		strSql = strSql & ",'00000000'"		'MAEGARI_SURYO"
		strSql = strSql & ",'makex'"		'INS_TANTO"
		strSql = strSql & " )"
		Debug ".YNyukaInsert():" & strSql
		on error resume next
			Call objDb.Execute(strSql)
			Debug ".YNyukaInsert():0x" & Hex(Err.Number) & ":" & Err.Description
		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'行書式
	'-----------------------------------------------------------------------
    Public Function FormatRow()
		objSheet.Range("A" & lngRow).RowHeight = 60
		if curPn = prvPn then
			objSheet.Range("Q" & lngRow) = ""
			exit function
		end if
		if curPn <> "" then
			if objSheet.Range("Q" & lngRow) = "" then
				dim	i
				i = 1
				do while objSheet.Range("Q" & lngRow) = ""
					objSheet.Range("Q" & lngRow) = objSheet.Range("Q" & lngRow - i)
					i = i - 1
				loop
			end if
		end if
	    objSheet.Range("A" & lngRow & ":Q" & lngRow).Borders(xlEdgeTop).LineStyle = xlContinuous
	    objSheet.Range("A" & lngRow & ":Q" & lngRow).Borders(xlEdgeTop).Weight = xlThin
	End Function
	'-----------------------------------------------------------------------
	'レコードセット行
	'-----------------------------------------------------------------------
    Public Function SetDataRow()
		Debug ".SetData():" & lngRow
		if objRs is nothing then
			exit function
		end if
		objSheet.Range("A" & lngRow) = GetField("CstDlvNo")	'顧客納入指示番号
		objSheet.Range("B" & lngRow) = GetField("CstPn")	'顧客品番
		objSheet.Range("C" & lngRow) = GetField("CstDlvDt")	'顧客納期
		objSheet.Range("D" & lngRow) = GetField("CstQty")	'顧客納入数量
		objSheet.Range("E" & lngRow) = GetField("OrderNo")	'発注番号
		objSheet.Range("F" & lngRow) = GetField("TestKb")	'試区
		objSheet.Range("G" & lngRow) = GetField("Pn")		'品番
		objSheet.Range("H" & lngRow) = GetField("PName")	'部名
		objSheet.Range("I" & lngRow) = GetField("DlvDt")	'納期日
		objSheet.Range("J" & lngRow) = GetField("DlvTm")	'時刻
		objSheet.Range("K" & lngRow) = GetField("RcptQty")	'数量
		objSheet.Range("L" & lngRow) = GetField("CenterCd")	'拠点
		objSheet.Range("M" & lngRow) = GetField("Location")	'納入場所
		objSheet.Range("N" & lngRow) = GetField("ClientNo")	'取引先
		objSheet.Range("O" & lngRow) = GetField("DlvMdfDt")	'納期変更日
		objSheet.Range("P" & lngRow) = GetField("SdcStkQty")	'SDC在庫数量
		objSheet.Range("Q" & lngRow) = "*" & GetField(".ID") & "*"			 '受入ID
'		objSheet.Range("R" & lngRow) = GetField("SdcStkQty")	'バーコード
	End Function
	'-----------------------------------------------------------------------
	'Fields 値
	'-----------------------------------------------------------------------
    Public Function GetField(byVal strFldNm)
		Debug ".GetField():" & strFldNm
		if objRs is nothing then
			exit function
		end if
		if left(strFldNm,1) <> "." then
			GetField = RTrim(objRs.Fields(strFldNm))
		else
			GetField = "."
		end if
		if GetField <> "" then
			select case strFldNm
			case "CstDlvDt"	'顧客納期	03/16 D
				GetField = Right(GetField,4)
				GetField = Left(GetField,2) & "/" & Right(GetField,2)
				GetField = GetField & " " & RTrim(objRs.Fields("CstDlvSft"))
			case "DlvDt"	'納期日	03/16 D
				GetField = Right(GetField,4)
				GetField = Left(GetField,2) & "/" & Right(GetField,2)
				GetField = GetField & " " & RTrim(objRs.Fields("DlvSft"))
			case "DlvTm"	'時刻	17:30
				GetField = Left(GetField,4)
				GetField = Left(GetField,2) & ":" & Right(GetField,2)
			case "DlvMdfDt"	'納期変更日	03/16
					GetField = Right(GetField,4)
					GetField = Left(GetField,2) & "/" & Right(GetField,2)
							' 123456789							
			case ".TEXT_NO"	'L151121
				dim	strTextNo
				strTextNo = Right(GetBaseName(strFilename),6)
				strTextNo = strTextNo & Right("000" & GetField("Row"),3)
				GetField = strTextNo
			case ".ID_NO"
				dim	strIdNo
				strIdNo = UCase(GetBaseName(strFilename))
				strIdNo = strIdNo & Right("00000" & GetField("Row"),5)
				GetField = strIdNo 
			case ".ID"
				dim	strId
				strId = "J"
				strId = strId & GetField(".SYUKA_YMD")
				strId = strId & GetField(".TEXT_NO")
				GetField = strId 
			case ".HIN_NO"
				GetField = RTrim(objRs.Fields("Pn"))
			case ".HIN_NAI"
				GetField = RTrim(objRs.Fields("CstPn"))
			case ".DEN_NO"
				GetField = RTrim(objRs.Fields("OrderNo"))
			case ".SURYO"
				GetField = RTrim(objRs.Fields("RcptQty"))
			case ".MUKE_CODE"
				GetField = RTrim(objRs.Fields("Location"))
			case ".SYUKO_YMD",".SYUKA_YMD"
				GetField = RTrim(objRs.Fields("DlvMdfDt"))
				if GetField = "" then
					GetField = RTrim(objRs.Fields("DlvDt"))
				end if
			case ".HIN_NAME"
				GetField = RTrim(objRs.Fields("PName"))
			case ".NOUKI_YMD"
				GetField = RTrim(objRs.Fields("DlvDt"))
			case ".SHIIRE_WORK_CENTER"
				GetField = RTrim(objRs.Fields("ClientNo"))
			end select
		end if
	End Function
	'-----------------------------------------------------------------------
	'問合せ
	'-----------------------------------------------------------------------
	Private strSql
    Public Function OpenRs()
		Debug ".OpenRs()"
		if objDb is nothing then
			exit function
		end if
		strSql = "select"
		strSql = strSql & " *"
		strSql = strSql & " from JcsTakeOn"
		strSql = strSql & " where Filename = '" & strFilename & "'"
		strSql = strSql & "   and RcptQty <> 0"
		strSql = strSql & " order by"
		strSql = strSql & "  Pn"
		strSql = strSql & " ,Row"
		set objRs = objDb.Execute(strSql)
	End Function
	'-----------------------------------------------------------------------
	'問合せClose
	'-----------------------------------------------------------------------
    Public Function CloseRs()
		Debug ".CloseRs()"
		if objRs is nothing then
			exit function
		end if
		Call objRs.Close()
		set objRs = nothing
	End Function
	'-------------------------------------------------------------------
	'ファイル名(パス、拡張子 除く)
	'-------------------------------------------------------------------
	Private Function GetBaseName(byVal f)
		dim	fobj
		set fobj = CreateObject("Scripting.FileSystemObject")
		dim	strBaseName
		strBaseName = fobj.GetBaseName(f)
		set fobj = Nothing
		GetBaseName = strBaseName
	End Function
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private	objExcel
	Private Function CreateExcel()
		Debug(".CreateExcel()")
		if objExcel is nothing then
			Debug(".CreateExcel():CreateObject(Excel.Application)")
			Set objExcel = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private	objBook
	Private Function OpenBook(byVal strBkName)
		Debug(".OpenBook()")
		if objBook is nothing then
			strBkName = strScriptPath & strBkName
			Debug(".OpenBook().Open:" & strBkName)
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
		end if
	end function
	'-------------------------------------------------------------------
	'Excel ファイル名前を付けて保存
	'-------------------------------------------------------------------
	Private Function SaveBook(byVal strBkName)
		Debug(".SaveBook()")
		if not objBook is nothing then
			strBkName = strScriptPath & strBkName
			Debug(".SaveBook().Save:" & strBkName)
			objExcel.DisplayAlerts = False
			Call objBook.SaveAs(strBkName)
			objExcel.DisplayAlerts = True
		end if
	end function
	'-------------------------------------------------------------------
	'Excel ファイルクローズ
	'-------------------------------------------------------------------
	Private Function CloseBook()
		Debug(".CloseBook()")
		if not objBook is nothing then
			Debug(".CloseBook().Close:" & objBook.Name)
			Call objBook.Close(False)
			set objBook = nothing
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
End Class
