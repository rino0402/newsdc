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
	Wscript.Echo "BO出荷データ(Excel)変換"
	Wscript.Echo "BoSyukaX.vbs [option] <ファイル名> [シート名]"
	Wscript.Echo " /db:newsdc9"
	Wscript.Echo " /debug"
	Wscript.Echo "cscript BoSyukaX.vbs ""I:\0SDC_honsya\事業部別商品化出荷金額まとめ\ＡＣ　ＮＰＬからの出荷実績\【１２月度】出荷実績_.xls"" /debug"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	dim	strStName
	strStName = ""
	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		elseif strStName = "" then
			strStName = strArg
		end if
	next
	'名前付きオプションチェック
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "conv"
		case "?"
			strFilename = ""
		case else
			strFilename = ""
		end select
	next
	dim	strConv
	strConv = GetOption("conv","")
	if strConv <> "" then
		dim	objC
		select case strConv
		case "AcCZaiko"
			Set objC = New AcCZaiko
		case "AcYNyuka"
			Set objC = New AcYNyuka
		end select
		'-------------------------------------------------------------------
		'データベースの準備
		'-------------------------------------------------------------------
		dim	objDb
		Set objDb = OpenAdodb(GetOption("db","newsdc9"))

		Call DispMsg("登録中...InitRecord")
		Call objC.InitRecord(objDb)
		Call DispMsg("登録中...ContRecord")
		Call objC.ConvRecord(objDb)
		'-------------------------------------------------------------------
		'データベースのクローズ
		'-------------------------------------------------------------------
		set objDb = CloseAdodb(objDb)
		Set objC = Nothing

		Main = 0
		exit function
	end if
	if strFilename = "" then
		usage()
		Main = 1
		exit Function
	end if
	call Load(strFilename,strStName)
	Main = 0
End Function

Function Load(byVal strFilename,byval strStName)
	'-------------------------------------------------------------------
	'Excelファイル名
	'-------------------------------------------------------------------
	strFilename = GetAbsPath(strFilename)
	Call Debug("Load():" & strFilename & "," & strStName & "")
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
	Dim	objSt
	For each objSt in objBk.Worksheets
		if strStName = "" or strStName = objSt.Name then
			Call LoadXls(objXL,objBk,objSt)
		end if
	Next
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("Load():End")
End Function

Function LoadXls(objXL,objBk,objSt)
	Call Debug("LoadXls():" & objSt.Name & "")
	'-------------------------------------------------------------------
	'クラス
	'-------------------------------------------------------------------
	dim	objC
	dim	lngMaxRow
	lngMaxRow = -1
	dim i
	for i = 0 to 4
		select case i
		case 0
			Set objC = New AcSyuka
		case 1
			Set objC = New NrSyuka
		case 2
			Set objC = New NrSyFuri
		case 3
			Set objC = New AcCZaiko
		case 4
			Set objC = New AcYNyuka
		end select
		lngMaxRow = objC.CheckHead(objXL,objBk,objSt)
		if lngMaxRow > 0 then
			exit for
		end if
		Set objC = Nothing
	next
	if lngMaxRow <= 0 then
		Exit Function
	end if
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc9"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	Call objC.CreateTmp(objDb)
	dim	rsTmp
	set rsTmp = OpenRs(objDb,objC.pTableNameTmp)

	'-------------------------------------------------------------------
	'読込
	'-------------------------------------------------------------------
	Call LoadSt(objXL,objBk,objSt,objDb,rsTmp,objC)

	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsTmp = CloseRs(rsTmp)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
	Set objC = Nothing
End Function

Function LoadSt(objXL,objBk,objSt,objDb,rsTmp,objC)
	Call Debug("LoadSt():" & objSt.Name)


	dim	lngMaxRow
	lngMaxRow = 0
	lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
	dim	lngRow
	For lngRow = 1 to lngMaxRow
		Call DispMsg("登録中..." & objSt.Name & ":" & lngRow & "/" & lngMaxRow)
		Call objC.LoadRow(objXL,objBk,objSt,objDb,rsTmp,lngRow)
	Next

	Call DispMsg("登録中...InitRecord")
	Call objC.InitRecord(objDb)
	Call DispMsg("登録中...ContRecord")
	Call objC.ConvRecord(objDb)
End Function

'-------------------------------------------------------
' AC入荷予定20150625.xls
'-------------------------------------------------------
Class AcYNyuka
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "AcYNyuka"
        pTableNameTmp	= "AcYNyukaTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("AcYNyuka.CreateTmp()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcYNyukaTmp"
		Call Debug("AcYNyuka.CreateTmp():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcYNyukaTmp using 'AcYNyukaTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = 1 To 65
			c = getColumnName2(i)
			select case c
			case "A"
				strSql = strSql & " x" & c & " Char(60) default '' not null" & vbCrLf
			case else
				strSql = strSql & ",x" & c & " Char(60) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   xA" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcYNyuka.CreateTmp():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	End Sub
    Public Sub InitRecord(objDb)
		Call Debug("AcYNyuka.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcYNyuka"
		Call Debug("AcYNyuka.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcYNyuka using 'AcYNyuka.DAT' with replace (" & vbCrLf
		strSql = strSql & " ID		Char(12) default '' not null" & vbCrLf
		strSql = strSql & ",Stat	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",Wc		Char( 8) default '' not null" & vbCrLf
		strSql = strSql & ",WcName	Char(40) default '' not null" & vbCrLf
		strSql = strSql & ",Pn		Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",Qty		CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",NYDt	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",ORDt	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",SNDt	Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   ID" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcYNyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into AcYNyuka "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(xA)"
		strSql = strSql & ",Left(RTrim(xB),1)"
		strSql = strSql & ",RTrim(xG)"
		strSql = strSql & ",RTrim(xH)"
		strSql = strSql & ",RTrim(xJ)"
		strSql = strSql & ",convert(xK,sql_decimal)"
		strSql = strSql & ",RTrim(xV)"
		strSql = strSql & ",RTrim(xW)"
		strSql = strSql & ",RTrim(xX)"
		strSql = strSql & " from AcYNyukaTmp"
		Call Debug("AcYNyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("AcYNyuka.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"Sheet1") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"発注納入管理番号") then
			Exit Function
		end if
		if CompHead(objSt.Range("B1"),"進捗　サービスデータ進捗区分") then
			Exit Function
		end if
		if CompHead(objSt.Range("C1"),"事業場コード") then
			Exit Function
		end if
		if CompHead(objSt.Range("D1"),"会計事業　会計用事業場コード") then
			Exit Function
		end if
		if CompHead(objSt.Range("E1"),"資産事業　資産管理事業場コード") then
			Exit Function
		end if
		if CompHead(objSt.Range("F1"),"調達先区分") then
			Exit Function
		end if
		if CompHead(objSt.Range("G1"),"仕入先WCコード") then
			Exit Function
		end if
		if CompHead(objSt.Range("H1"),"仕入先WC名称") then
			Exit Function
		end if
		if CompHead(objSt.Range("I1"),"仕入先品目番号") then
			Exit Function
		end if
		if CompHead(objSt.Range("J1"),"品目番号") then
			Exit Function
		end if
		if CompHead(objSt.Range("K1"),"入出庫予定数") then
			Exit Function
		end if
		if CompHead(objSt.Range("BM1"),"更新回数") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("AcYNyuka.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("AcYNyuka.LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 3 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = 1 To 65
			c = GetColumnName2(i)
			Call SetField(rsTmp,objSt,"x" & c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ﾚｺｰﾄﾞのｷｰ ﾌｨｰﾙﾄﾞに重複するｷｰ値があります(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class
'-------------------------------------------------------
' ACセンター在庫150625.xlsx,
'-------------------------------------------------------
Class AcCZaiko
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "AcCZaiko"
        pTableNameTmp	= "AcCZaikoTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("AcCZaiko.CreateTmp()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcCZaikoTmp"
		Call Debug("AcCZaiko.CreateTmp():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcCZaikoTmp using 'AcCZaikoTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("J")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " " & c & " Char(20) default '' not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & "  ,B" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcCZaiko.CreateTmp():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	End Sub
    Public Sub InitRecord(objDb)
		Call Debug("AcCZaiko.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcCZaiko"
		Call Debug("AcCZaiko.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcCZaiko using 'AcCZaiko.DAT' with replace (" & vbCrLf
		strSql = strSql & " JCode	Char( 8) default '' not null" & vbCrLf
		strSql = strSql & ",Pn 		Char(20) default '' not null" & vbCrLf
		strSql = strSql & ",Tanto	Char(10) default '' not null" & vbCrLf
		strSql = strSql & ",C_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000440_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000441_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000443_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000444_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",S22000446_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",T_Qty	CURRENCY default  0 not null" & vbCrLf
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   JCode" & vbCrLf
		strSql = strSql & "  ,Pn" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcCZaiko.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into AcCZaiko "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(A)"
		strSql = strSql & ",RTrim(B)"
		strSql = strSql & ",RTrim(C)"
		strSql = strSql & ",convert(D,sql_decimal)"
		strSql = strSql & ",convert(E,sql_decimal)"
		strSql = strSql & ",convert(F,sql_decimal)"
		strSql = strSql & ",convert(G,sql_decimal)"
		strSql = strSql & ",convert(H,sql_decimal)"
		strSql = strSql & ",convert(I,sql_decimal)"
		strSql = strSql & ",convert(J,sql_decimal)"
		strSql = strSql & " from AcCZaikoTmp"
		Call Debug("AcCZaiko.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("AcCZaiko.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"在庫") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"事業場CD") then
			Exit Function
		end if
		if CompHead(objSt.Range("B1"),"品目番号") then
			Exit Function
		end if
		if CompHead(objSt.Range("C1"),"購買担当者CD") then
			Exit Function
		end if
		if CompHead(objSt.Range("D1"),"センター倉庫") then
			Exit Function
		end if
		if CompHead(objSt.Range("J1"),"総在庫") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("AcCZaiko.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("AcCZaiko.LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 2 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("J")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ﾚｺｰﾄﾞのｷｰ ﾌｨｰﾙﾄﾞに重複するｷｰ値があります(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

'-------------------------------------------------------
' Ac出荷実績
'-------------------------------------------------------
Class AcSyuka
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "BoSyuka"
        pTableNameTmp	= "AcSyukaTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("AcSyuka.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table AcSyukaTmp"
		Call Debug("AcSyuka.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table AcSyukaTmp using 'AcSyukaTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("L")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " A UBIGINT  default 0  not null" & vbCrLf
			case "G"
				strSql = strSql & "," & c & " Char(40) default '' not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "Create Unique Index AcSyukaTmp_Key01 On AcSyukaTmp (" & vbCrLf
		strSql = strSql & "   B" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("AcSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	End Sub
    Public Sub InitRecord(objDb)
'		Call Debug("delete from " & pTableNameTmp)
'		Call ExecuteAdodb(objDb,"delete from " & pTableNameTmp)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = "delete from BoSyuka"
		strSql = strSql & " where ShisanJCode in (select distinct RTrim(C) from AcSyukaTmp)"
		strSql = strSql & "   and Left(Dt,6) in (select distinct Left(J,6) from AcSyukaTmp)"
		Call Debug("AcSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "insert into BoSyuka "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(B)"					'// 収支振過実_収支振替管理番号
		strSql = strSql & ",RTrim(C)"					'// 収支振過実_資産管理事業場コード
		strSql = strSql & ",''"							'// 収支振過実_収支振替コード
		strSql = strSql & ",RTrim(H)"					'// 収支振過実_入出庫取引区分
		strSql = strSql & ",''"							'// 収支振過実_在庫収支略式名
		strSql = strSql & ",''"							'// 収支振過実_在庫収支コード
		strSql = strSql & ",''"							'// 収支振過実_伝票番号
		strSql = strSql & ",RTrim(E)"					'// 収支振過実_品目番号
		strSql = strSql & ",convert(L,sql_decimal)"		'// 収支振過実_入出庫実績数
		strSql = strSql & ",RTrim(J)"					'// 収支振過実_収支振実年月日
		strSql = strSql & ",''"							'// 収支振過実_振替先在庫収支略式名
		strSql = strSql & ",''"							'// 収支振過実_振替先在庫収支コード
		strSql = strSql & ",''"							'// 在庫収支_倉庫コード
		strSql = strSql & " from AcSyukaTmp"
		Call Debug("AcSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("AcSyuka.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"前月出荷明細") then
			if CompHead(objSt.Name,"当月出荷明細") then
				Exit Function
			end if
			if objSt.Range("A1") = "" then
				objSt.Range("A1") = "NO"
			end if
		end if
		if CompHead(objSt.Range("A1"),"NO") then
			Exit Function
		end if
		if CompHead(objSt.Range("L1"),"出荷実績数") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("AcSyuka.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("AcSyuka.LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 2 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("L")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ﾚｺｰﾄﾞのｷｰ ﾌｨｰﾙﾄﾞに重複するｷｰ値があります(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

'-------------------------------------------------------
' 奈良東倉庫 出荷実績
'-------------------------------------------------------
Class NrSyuka
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "BoSyuka"
        pTableNameTmp	= "NrSyukaTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("NrSyuka.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table NrSyukaTmp"
		Call Debug("NrSyuka.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table NrSyukaTmp using 'NrSyukaTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " A UBIGINT  default 0  not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("NrSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "Create Unique Index NrSyukaTmp_Key01 On NrSyukaTmp (" & vbCrLf
		strSql = strSql & "   C" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("NrSyuka.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub InitRecord(objDb)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = "delete from BoSyuka"
		strSql = strSql & " where ShisanJCode in (select distinct RTrim(D) from NrSyukaTmp)"
		strSql = strSql & "   and Left(Dt,6) in (select distinct Left(L,6) from NrSyukaTmp)"
		Call Debug("NrSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "insert into BoSyuka "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(C)"					'// 収支振過実_収支振替管理番号
		strSql = strSql & ",RTrim(D)"					'// 収支振過実_資産管理事業場コード
		strSql = strSql & ",''"							'// 収支振過実_収支振替コード
		strSql = strSql & ",RTrim(G)"					'// 収支振過実_入出庫取引区分
		strSql = strSql & ",''"							'// 収支振過実_在庫収支略式名
		strSql = strSql & ",''"							'// 収支振過実_在庫収支コード
		strSql = strSql & ",RTrim(F)"					'// 収支振過実_伝票番号
		strSql = strSql & ",RTrim(E)"					'// 収支振過実_品目番号
		strSql = strSql & ",convert(M,sql_decimal)"		'// 収支振過実_入出庫実績数
		strSql = strSql & ",RTrim(L)"					'// 収支振過実_収支振実年月日
		strSql = strSql & ",''"							'// 収支振過実_振替先在庫収支略式名
		strSql = strSql & ",''"							'// 収支振過実_振替先在庫収支コード
		strSql = strSql & ",RTrim(B)"					'// 在庫収支_倉庫コード
		strSql = strSql & " from NrSyukaTmp"
		Call Debug("NrSyuka.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("NrSyuka.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"東倉庫_出荷実績明細") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"◆東倉庫_出荷実績明細") then
			Exit Function
		end if
		if CompHead(objSt.Range("A2"),"NO") then
			Exit Function
		end if
		if CompHead(objSt.Range("M2"),"出荷実績数") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("NrSyuka.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("LoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 3 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ﾚｺｰﾄﾞのｷｰ ﾌｨｰﾙﾄﾞに重複するｷｰ値があります(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

'-------------------------------------------------------
' 奈良東倉庫 収支振替明細
'-------------------------------------------------------
Class NrSyFuri
	Public pTableName
	Public pTableNameTmp
    Private Sub Class_Initialize
        pTableName		= "BoSyFuri"
        pTableNameTmp	= "NrSyFuriTmp"
    End Sub
    Public Sub CreateTmp(objDb)
		Call Debug("NrSyFuri.InitRecord()")
		dim	strSql
		strSql = ""
		strSql = strSql & "Drop Table NrSyFuriTmp"
		Call Debug("NrSyFuri.InitRecord():" & strSql)
		On Error Resume Next
		Call ExecuteAdodb(objDb,strSql)
		On Error goto 0
		strSql = ""
		strSql = strSql & "Create Table NrSyFuriTmp using 'NrSyFuriTmp.DAT' with replace (" & vbCrLf
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			select case c
			case "A"
				strSql = strSql & " " & c & " Char(20) default '' not null" & vbCrLf
			case else
				strSql = strSql & "," & c & " Char(20) default '' not null" & vbCrLf
			end select
		Next
		strSql = strSql & ",PRIMARY KEY(" & vbCrLf
		strSql = strSql & "   A" & vbCrLf
		strSql = strSql & "  ,I" & vbCrLf
		strSql = strSql & "  ,K" & vbCrLf
		strSql = strSql & " )" & vbCrLf
		strSql = strSql & ")" & vbCrLf
		Call Debug("NrSyFuri.InitRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
'		strSql = ""
'		strSql = strSql & "Create Unique Index NrSyFuriTmp_Key01 On NrSyFuriTmp (" & vbCrLf
'		strSql = strSql & "   C" & vbCrLf
'		strSql = strSql & ")" & vbCrLf
'		Call Debug("NrSyFuri.InitRecord():" & strSql)
'		Call ExecuteAdodb(objDb,strSql)
    End Sub
    Public Sub InitRecord(objDb)
    End Sub
    Public Sub ConvRecord(objDb)
		dim	strSql
		strSql = "delete from BoSyFuri"
		strSql = strSql & " where ShisanJCode in (select distinct RTrim(C) from NrSyFuriTmp)"
		strSql = strSql & "   and Left(Dt,6) in (select distinct Left(Replace(L,'/',''),6) from NrSyFuriTmp)"
		Call Debug("NrSyFuri.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		strSql = ""
		strSql = strSql & "insert into BoSyFuri "
		strSql = strSql & " select"
		strSql = strSql & " RTrim(A)"					'// 注文管理番号	500380470-40
		strSql = strSql & ",RTrim(B)"					'// 倉庫名	奈良東
		strSql = strSql & ",RTrim(C)"					'// 事業場CD	00021184
		strSql = strSql & ",RTrim(D)"					'// 品目番号	A0601-1E60S
		strSql = strSql & ",RTrim(E)"					'// 業務名	入庫:商品化
		strSql = strSql & ",RTrim(F)"					'// 収支振替CD	2611B
		strSql = strSql & ",RTrim(G)"					'// 収支CD	11B
		strSql = strSql & ",RTrim(H)"					'// 収支略式	11N1
		strSql = strSql & ",RTrim(I)"					'// 取引区分	45
		strSql = strSql & ",RTrim(J)"					'// 進捗区分	7
		strSql = strSql & ",RTrim(K)"					'// 伝票番号	005628
		strSql = strSql & ",Replace(RTrim(L),'/','')"	'// 実績年月日	2015/01/10
		strSql = strSql & ",convert(M,sql_decimal)"		'// 個数	2,000
		strSql = strSql & " from NrSyFuriTmp"
		Call Debug("NrSyFuri.ConvRecord():" & strSql)
		Call ExecuteAdodb(objDb,strSql)
    End Sub
	Public Function CheckHead(objXL,objBk,objSt)
		Call Debug("NrSyFuri.CheckHead():" & objSt.Name)
		CheckHead = -1
		if CompHead(objSt.Name,"東倉庫_収支振替明細") then
			Exit Function
		end if
		if CompHead(objSt.Range("A1"),"◆東倉庫_収支振替明細") then
			Exit Function
		end if
		if CompHead(objSt.Range("A2"),"注文管理番号") then
			Exit Function
		end if
		if CompHead(objSt.Range("M2"),"個数") then
			Exit Function
		end if
		dim	lngMaxRow
		lngMaxRow = 0
		lngMaxRow = excelGetMaxRow(objSt,"A",lngMaxRow)
		Call Debug("NrSyFuri.CheckHead():" & lngMaxRow)
		CheckHead = lngMaxRow
	End Function
	Public Function LoadRow(objXL,objBk,objSt,objDb,rsTmp,byVal lngRow)
		Call Debug("NrSyFuriLoadRow():" & objSt.Name & ":" & lngRow)
		if lngRow < 3 then
			LoadRow = 0
			Exit Function
		end if

		Call rsTmp.AddNew
		Dim c
		Dim	i
		For i = Asc("A") To Asc("M")
			c = Chr(i)
			Call SetField(rsTmp,objSt,c,c,lngRow)
		Next
'		On Error Resume Next
			Call rsTmp.UpdateBatch
			if Err <> 0 Then
				WScript.Echo "line(" & lngRow & "):" & Err.Number & " : " & Err.Description
				' ﾚｺｰﾄﾞのｷｰ ﾌｨｰﾙﾄﾞに重複するｷｰ値があります(Btrieve Error 5)
				if Err.Number <> -2147467259 then
					lngRow = 0
				end if
				Call rsTmp.CancelUpdate
			end if
'		On Error GoTo 0
		LoadRow = lngRow
	End Function
End Class

Function CompHead(byval strV,strTitle)
	Call Debug("CompHead():" & strV & ":" & strTitle)
	if strV = strTitle then
		Call Debug("CompHead():====")
		CompHead = 0
		Exit Function
	end if
	strV = Replace(strV,vbCrLf,"")
	if strV = strTitle then
		CompHead = 0
		Call Debug("CompHead():====CrLf")
		Exit Function
	end if
	strV = Replace(strV,vbLf,"")
	if strV = strTitle then
		CompHead = 0
		Call Debug("CompHead():====Lf")
		Exit Function
	end if
	Call Debug("CompHead():<><>")
	CompHead = 1
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

Function getColumnName2(ByVal pColumn)

    Const cStartChr = 65 'Aの文字コード
    Const cAlphabet = 26 'アルファベットの種類

    Dim lColumnNum
    Dim sColumnName

    If pColumn < 1 Then

       getColumnName2 = "??"
       Exit Function

    End If

    Do

       lColumnNum = (pColumn - 1) Mod cAlphabet
       sColumnName = sColumnName & Chr(cStartChr + lColumnNum)

       pColumn = (pColumn - 1) \ cAlphabet

    Loop Until pColumn = 0

    getColumnName2 = StrReverse(sColumnName)

End Function
