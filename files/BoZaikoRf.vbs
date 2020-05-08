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
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "BO在庫データ 冷蔵庫用"
	Wscript.Echo "BoZaikoRf.vbs [option]"
	Wscript.Echo " /list"
	Wscript.Echo " /debug"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	For Each strArg In WScript.Arguments.UnNamed
    	select case strArg
		case else
			if strFilename <> "" then
				usage()
				Main = 1
				exit Function
			end if
			strFilename = strArg
		end select
	Next
	For Each strArg In WScript.Arguments.Named
    	select case lcase(strArg)
		case "db"
		case "debug"
		case "list"
		case "load"
		case "top"
		case else
			usage()
			Main = 1
			exit Function
		end select
	Next
	select case GetFunction()
	case "list"
		Call List()
	case "load"
		Call Load(strFilename)
	end select
	Main = 0
End Function

Private Function GetFunction()
	GetFunction = "list"
	if WScript.Arguments.Named.Exists("load") then
		GetFunction = "load"
	end if
End Function

Private Function Load(byval strFilename)
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	dim	objRs
	set objRs = OpenRs(objDb,"BoZaikoRf")
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	dim	objXL
	dim	objBk
	dim	objSt
	Call Debug("CreateObject(Excel.Application)")
	Set objXL = WScript.CreateObject("Excel.Application")
'	objXL.Application.Visible = True
	Call Debug("Workbooks.Open()" & strFilename)
	Set objBk = objXL.Workbooks.Open(strFilename,False,True)
	set objSt = objBk.ActiveSheet
	Call Debug("objSt.Name=" & objSt.Name)

	dim	cnt
	cnt = 0
	dim	cntAdd
	cntAdd = 0

	Const xlUp = -4162
	dim	lngRowMax
	lngRowMax = objSt.Range("B65536").End(xlUp).Row
	dim	strJCode
	dim	strShisanJCode
	strJCode 		= ""
	strShisanJCode	= ""
	dim	strStat
	strStat = "head"
	dim lngRow
	for lngRow = 1 to lngRowMax
		dim	strB
		strB = objSt.Range("B" & lngRow)
		Call Debug(lngRow & ":" & strStat & ":" & strB)
		select case strStat
		case "head"
			select case strB
			case ""
			case "草津センタ倉庫在庫"
				strStat = "title"
				strJCode		= "00021259"
				strShisanJCode	= "00021259"
			case "ＰＰＳＣ奈良在庫検索"
				strStat = "title"
				strJCode		= "00036003"
				strShisanJCode	= "00021259"
			case else
				Call DispMsg("ヘッダーエラー：" & strB)
			end select
		case "title"
			select case strB
			case ""
			case "品目番号"
				dim	strDeleteSql
				strDeleteSql = "delete from BoZaikoRf where jCode = '" & strJCode & "' and ShisanJCode = '" & strShisanJCode & "'"
				Call Debug(strDeleteSql)
				Call ExecuteAdodb(objDb,strDeleteSql)
				strStat = "value"
			case else
				Call DispMsg("項目名エラー：" & strB)
			end select
		case "value"
			'	 JCode			Char( 8) default '' not null	// 事業場コード
			'	,ShisanJCode	Char( 8) default '' not null	// 資産管理事業場コード
			'	,Pn				Char(20) default '' not null	// 品目番号
			'	,PName			Char(40) default '' not null	// 品目名
			'	,DModel			Char(20) default '' not null	// 代表機種
			'	,HikiQty		CURRENCY default 0  not null	// 正味引当可能在庫数
			'	,Hinmoku		Char(10) default '' not null	// ＰＮ共通_品目コード２
			'	,SyuShi			Char( 8) default '' not null	// 在庫収支コード
			cntAdd = cntAdd + 1
			objRs.Addnew
			objRs.Fields("JCode")		= strJCode
			objRs.Fields("ShisanJCode")	= strShisanJCode
			objRs.Fields("Pn")			= strB
			objRs.Fields("PName")		= RTrim(objSt.Range("C" & lngRow))
			objRs.Fields("DModel")		= RTrim(objSt.Range("D" & lngRow))
			objRs.Fields("HikiQty")		= RTrim(objSt.Range("E" & lngRow))
			objRs.Fields("Hinmoku")		= RTrim(objSt.Range("F" & lngRow))
			objRs.Fields("SyuShi")		= RTrim(objSt.Range("G" & lngRow))
			objRs.UpdateBatch
		end select
	next
	Call DispMsg("読込件数：" & lngRow)
	Call DispMsg("登録件数：" & cntAdd)
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	Call objBk.Close(False)
	set objBk = Nothing
	set objXL = Nothing
	'-------------------------------------------------------------------
	'データベースの後処理
	'-------------------------------------------------------------------
	set objRs = CloseRs(objRs)
	set objDb = nothing
End Function

Private Function chkJCode(byVal aryJCode(),byVal strJCode)
	dim	a
'	for each a in aryJCode
	dim	i
	Call Debug("chkJCode:" & LBound(aryJCode) & " to " & UBound(aryJCode))
	for i = LBound(aryJCode) to UBound(aryJCode)
		a = aryJCode(i)
		Call Debug("chkJCode:" & a & "=" & strJCode)
		if a = strJCode then
			strJCode = ""
			exit for
		end if
	next
	chkJCode = strJCode
End Function

Private Function GetFName(byval strTitle)
	dim	strFName
	strFName = ""
						'strFName = "BTKbn"		' 部品取引区分
	select case strTitle
	case "取引区分"
	case "サービスデータ進捗区分"
						strFName = "SrvDtSts"	' サービスデータ進捗区分
	case "資産管理事業場コード"
						strFName = "JCode"		' 資産事業　資産管理事業場コード
	case "品目番号"
						strFName = "Pn"			' 出荷品目番号
	case "グローバル品目番号"
	case "サービス品目番号"
	case "受付品目番号"
						strFName = "PnRcv"		' 受注品目番号
	case "相手先コード"
	case "相手先名"
	case "数量"
'						strFName = "QtyRcv"		' 受注実績数
						strFName = "QtySnd"		' 受注実績数
	case "単価"
						strFName = "Price"		' 単価　実際単価    9999999.0000
	case "実際金額"
						strFName = "Amount"		' 実際金額
	case "オーダーNo."
						strFName = "OrderNo"	' オーダーNO
	case "ITEM-No."
	case "伝票番号"
						strFName = "DenNo"		' 伝票番号
	case "ID-No."
						strFName = "IDNo"		' ID-NO
	case "在庫収支略式名"
						strFName = "ZSyushiRk"	' 在庫収支略式名
	case "在庫収支コード"
	case "資産管理在庫収支コード"
	case "補助在庫収支コード"
	case "帳端区分"
						strFName = "CHKbn"		' 帳端区分
	case "値差区分"
						strFName = "NSKbn"		' 値差区分
	case "返品区分"
	case "実績日(予定日)"
						strFName = "SalesDt"	' 売上予定年月日 yyyymmdd
	case "受発注年月日"
						strFName = "RcvDt"		' 受注年月日
	case "出庫年月日"
						strFName = "PckDt"		' 出庫予定年月日
	case "出荷年月日"
						strFName = "SndDt"		' 出荷予定年月日
	case "発送年月日"
	case "出荷指定年月日"
	case "指定納期年月日"
						strFName = "DlvDt"		' 指定納期日　指定納期年月日
	case "納期回答年月日"
						strFName = "AnsDt"		' 納期回答日　納期回答年月日
	case "受注出荷・販売区分"
	case "受注出荷・直送先コード"
	case "受注出荷・注文区分"
						strFName = "ChuKbn"		' 注文区分
	end select
	GetFName = strFName
End Function

Private Function List()
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	Call DispMsg("objDb.Execute(" & strSql & ")")
	set rsList = objDb.Execute(strSql)

	do while rsList.Eof = False
		Call DispMsg("■" _
			 & " " & rsList.Fields("IDNo") _
			 & " " & rsList.Fields("JCode") _
			 & " " & rsList.Fields("Pn") _
			 & " " & rsList.Fields("PnRcv") _
			 & " " & rsList.Fields("BTKbn") _
			 & " " & rsList.Fields("TKCode") _
			 & " " & rsList.Fields("ChokuCode") _
			 & " " & rsList.Fields("SrvDtSts") _
					)
		Call rsList.MoveNext
	loop

	Call Debug("CloseAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = Nothing
End Function

Private Function makeSql()
	dim	strSql
	dim	strTop
	strTop = GetOption("top","")
	if strTop <> "" then
		strTop = " top " & strTop
	end if
	strSql = "select" & strTop
	strSql = strSql & " *"
	strSql = strSql & " from BoZaikoRf"
	makeSql = strSql
End Function

