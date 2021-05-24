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
Call Include("get_b.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "出荷商品化費用請求データ"
	Wscript.Echo "Bill.vbs [option]"
	Wscript.Echo " /db:<database>"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
	Wscript.Echo " /debug"
	Wscript.Echo "Ex."
	Wscript.Echo "Bill.vbs /db:newsdc-ono /load ""\\hs1\sec\ppsc\PPSC提出請求書過去分\201204請求\小野\04PPSC請求書（IHCS分）.xlsx"""
	Wscript.Echo "----"
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
		case "list"
		case "jgyobu"
		case "load"
		case "top"
		case "debug"
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
	case "usage"
		Call usage()
	end select
	Main = 0
End Function

Private Function GetFunction()
	GetFunction = "usage"
	if WScript.Arguments.Named.Exists("load") then
		GetFunction = "load"
	elseif WScript.Arguments.Named.Exists("list") then
		GetFunction = "list"
	end if
End Function

'-------------------------------------------------------------------
'請求データ(Excel)→Bill
'-------------------------------------------------------------------
Private Function Load(byval strFilename)
	'-------------------------------------------------------------------
	'データベース準備
	'-------------------------------------------------------------------
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	'-------------------------------------------------------------------
	'登録用レコードセット準備
	'-------------------------------------------------------------------
	dim	objRs
	set objRs = OpenRs(objDb,"Bill")
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	dim	objXL
	dim	objBk
	Call Debug("CreateObject(Excel.Application)")
	Set objXL = WScript.CreateObject("Excel.Application")
'	objXL.Application.Visible = True
	Call Debug("Workbooks.Open()" & strFilename)
	Set objBk = objXL.Workbooks.Open(strFilename,False,True)
	dim	strJGyobu
	strJGyobu = GetJGyobu(objBk)
	if strJGyobu = "" then
		call DispMsg("事業部不明")
	else
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"⑤入庫")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"⑦部内出庫")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"①②③④出荷明細")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"②③④⑤ＡＣ商品化")
		Call LoadBill(objDb,objRs,objBk,strJGyobu,"②③商品化工料")
	end if
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

Function GetJGyobu(objBk)
	dim	strJGyobu
	strJGyobu = GetOption("jgyobu","")
	if strJGyobu <> "" then
		GetJGyobu = strJGyobu
		exit function
	end if
	dim	objSt
	for each objSt in objBk.Worksheets
		dim	strName
		select case objSt.Name
'		case "請求明細"
'			strName = Trim(objSt.Range("D9"))
'			select case strName
'			case "エアコン"
'				strJGyobu = "A"
'			case "冷蔵庫"
'				strJGyobu = "R"
'			end select
'			Call Debug("GetJGyobu():" & strName & ":" & strJGyobu)
'			exit for
		case "請求明細","請求明細フォーム (2)","請求明細フォーム(2)","請求明細（提出用）","請求明細 (2)"
			strName = Trim(objSt.Range("C8"))
			if strName = "" then
				strName = Trim(objSt.Range("D8"))
				if strName = "" or strName = "滋賀物流センター" then
					' 請求明細 エアコン/冷蔵庫
					strName = Trim(objSt.Range("D9"))
				end if
			end if
			select case strName
			case "小野パーツセンター　（IHｸｯｷﾝｸﾞﾋｰﾀｰ分）"
				strJGyobu = "D"
			case "小野パーツセンター　（ビューティ・リビングBU分）","小野パーツセンター　（ビューティ・リビング分）","小野パーツセンター　（調理小物分）"
				strJGyobu = "5"
			case "小野パーツセンター　（炊飯機器分）"
				strJGyobu = "4"
			case "滋賀パーツセンター"
				strJGyobu = "7"
			case "エアコン"
				strJGyobu = "A"
			case "冷蔵庫"
				strJGyobu = "R"
			end select
			Call Debug("GetJGyobu():" & strName & ":" & strJGyobu)
			exit for
		end select
	next
	GetJGyobu = strJGyobu
End Function

Function GetBillSheet(objBk,byVal strSheetName)
	set GetBillSheet = Nothing
	dim	objSt
	for each objSt in objBk.Worksheets
		if objSt.Name = strSheetName then
			set GetBillSheet = objSt
			exit for
		end if
		select case strSheetName
		case "①②③④出荷明細"
			if objSt.Name = "③④⑤⑥出荷明細" then
				set GetBillSheet = objSt
				exit for
			end if
		case "⑦部内出庫"
			if objSt.Name = "⑩部内出庫" then
				set GetBillSheet = objSt
				exit for
			end if
		case "⑤入庫"
			if objSt.Name = "⑦入庫" then
				set GetBillSheet = objSt
				exit for
			end if
		end select
	next
End Function


Function LoadBill(objDb,objRs,objBk,byVal strJGyobu,byVal strSheetName)

	Call Debug("LoadBill(" & strJGyobu & "," & strSheetName & ")")
	dim	objSt
	set objSt = GetBillSheet(objBk,strSheetName)
	if objSt is Nothing then
		Call DispMsg("LoadBill():" & strSheetName & "：" & "指定シートがありません.")
		exit function
	end if
	dim	strBillDt
	strBillDt = ""
	dim	strYM
	strYM = ""
	dim	strKBN
	select case strSheetName
	case "①②③④出荷明細","③④⑤⑥出荷明細"
		strKBN = "A"
		strBillDt = GetDt(RTrim(objSt.Range("C4")))
	case "⑦部内出庫","⑩部内出庫"
		strKBN = "B"
		strBillDt = GetDt(RTrim(objSt.Range("C3")))
		if strBillDt = "" then
			strBillDt = GetDt(RTrim(objSt.Range("C4")))
		end if
	case "⑤入庫","⑦入庫"
		strKBN = "C"
		strBillDt = GetDt(RTrim(objSt.Range("B4")))
	case "②③④⑤ＡＣ商品化"
		strKBN = "D"
		strBillDt = GetDt(RTrim(objSt.Range("M3")))
	case "②③商品化工料"
		strKBN = "E"
		strBillDt = GetDt(RTrim(objSt.Range("H2")))
	case else
		Exit Function
	end select
	strYM = GetYM(strBillDt)
	if strYM <> "" then
		Call DispMsg(strSheetName & "：" & strYM)
		'-------------------------------------------------------------------
		'年月データ削除
		'-------------------------------------------------------------------
		dim	strSql
		strSql = "delete from Bill" _
			   & " where JGyobu = '" & strJGyobu & "'" _
			   & "   and BillDt = '" & strBillDt & "'" _
			   & "   and YM = '" & strYM & "'" _
			   & "   and KBN = '" &	strKBN & "'"
		Call Debug("削除:" & strSql)
		Call ExecuteAdodb(objDb,strSql)
	end if
	'-------------------------------------------------------------------
	'Excel最終行
	'-------------------------------------------------------------------
	Const xlUp = -4162
	dim	lngRowTop
	dim	lngRowMax
	select case strKBN
	case "D"
		lngRowTop = 11
		lngRowMax = objSt.Range("B65536").End(xlUp).Row
	case "E"
		lngRowTop = 11
		lngRowMax = objSt.Range("A65536").End(xlUp).Row
	case else
		lngRowTop = 4
		lngRowMax = objSt.Range("A65536").End(xlUp).Row
	end select

	dim	cntAdd
	cntAdd = 0

	'-------------------------------------------------------------------
	'ループ：3～最終行
	'-------------------------------------------------------------------
	dim	lngRow
	for lngRow = lngRowTop to lngRowMax
		Call Debug(strJGyobu & " " & strYM & " " & strBillDt & " " & lngRow & _
					" " & RTrim(objSt.Range("A" & lngRow)) & _
					" " & RTrim(objSt.Range("B" & lngRow)) & _
					" " & RTrim(objSt.Range("C" & lngRow)) & _
					" " & RTrim(objSt.Range("D" & lngRow)) & _
					" " & RTrim(objSt.Range("E" & lngRow)) & _
					" " & RTrim(objSt.Range("F" & lngRow)) _
					)
		dim	lngNo
		if strKBN = "D" then
			lngNo = GetNumValue(RTrim(objSt.Range("B" & lngRow)))
		else
			lngNo = GetNumValue(RTrim(objSt.Range("A" & lngRow)))
		end if
		if lngNo = 0 then
			Call Debug("Exit for:lngNo = 0")
			exit for
		end if
		cntAdd = cntAdd + 1
		objRs.AddNew
		objRs.Fields("JGyobu") 		= strJGyobu			'// 事業部
		objRs.Fields("BillDt")		= strBillDt			'// 請求日
		objRs.Fields("YM") 			= strYM				'// 請求年月
		objRs.Fields("KBN") 		= strKBN			'// 請求区分
														'// 1:PPSC出荷
														'// 2:PPSC部内出庫
		objRs.Fields("No") 			= lngNo				'// 請求書明細No
		select case strKBN
		case "A"	'①②③④出荷明細
			objRs.Fields("IdNo") 		= RTrim(objSt.Range("B" & lngRow))		'// ID-No
			objRs.Fields("Dt") 			= Replace(RTrim(objSt.Range("C" & lngRow)),"/","")		'// 出荷日
			objRs.Fields("DenNo") 		= RTrim(objSt.Range("D" & lngRow))		'// 伝票番号
			objRs.Fields("SyukaCd")		= RTrim(objSt.Range("E" & lngRow))		'// 出荷先
			objRs.Fields("SyukaNm")		= RTrim(objSt.Range("F" & lngRow))		'// 出荷先名
			objRs.Fields("Pn") 			= RTrim(objSt.Range("G" & lngRow))		'// 品番
			objRs.Fields("PnName") 		= RTrim(objSt.Range("H" & lngRow))		'// 品名
			objRs.Fields("Qty") 		= RTrim(objSt.Range("I" & lngRow))		'// 出荷数
			objRs.Fields("Pick") 		= RTrim(objSt.Range("J" & lngRow))		'// 出庫工料
			objRs.Fields("Ship") 		= RTrim(objSt.Range("K" & lngRow))		'// 出荷工料
			objRs.Fields("AnyKbn") 		= RTrim(objSt.Range("M" & lngRow))		'// 区分
			objRs.Fields("KoryoPrc")	= RTrim(objSt.Range("N" & lngRow))		'// 個装工料単価
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("O" & lngRow))		'// 個装工料
			objRs.Fields("HakoPrc") 	= RTrim(objSt.Range("P" & lngRow))		'// 個装箱代単価
			objRs.Fields("Hako") 		= RTrim(objSt.Range("Q" & lngRow))		'// 個装箱代
		case "B"	'⑦部内出庫
			objRs.Fields("IdNo") 		= RTrim(objSt.Range("B" & lngRow))		'// ID-No
			objRs.Fields("Dt") 			= Replace(RTrim(objSt.Range("C" & lngRow)),"/","")		'// 出荷日
			objRs.Fields("DenNo") 		= RTrim(objSt.Range("D" & lngRow))		'// 伝票番号
			objRs.Fields("SyukaCd")		= RTrim(objSt.Range("E" & lngRow))		'// 出荷先
			objRs.Fields("SyukaNm")		= RTrim(objSt.Range("F" & lngRow))		'// 出荷先名
			objRs.Fields("Pn") 			= RTrim(objSt.Range("G" & lngRow))		'// 品番
			objRs.Fields("PnName") 		= ""									'// 品名
			objRs.Fields("Qty") 		= RTrim(objSt.Range("I" & lngRow))		'// 出荷数
			objRs.Fields("Pick") 		= RTrim(objSt.Range("K" & lngRow))		'// 出庫工料
			objRs.Fields("Ship") 		= 0										'// 出荷工料
			dim	strCol
			strCol = ""
'			if objSt.Range("P2") = "未or完" then
'				strCol = "P"
'			elseif objSt.Range("Q2") = "未or完" then
'				strCol = "Q"
'			else
'			end if
			if strCol <> "" then
				objRs.Fields("AnyKbn") 		= RTrim(objSt.Range(strCol & lngRow))	'// 区分
			end if
			objRs.Fields("KoryoPrc")	= RTrim(objSt.Range("L" & lngRow))		'// 個装工料単価
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("M" & lngRow))		'// 個装工料
			objRs.Fields("HakoPrc") 	= RTrim(objSt.Range("N" & lngRow))		'// 個装箱代単価
			objRs.Fields("Hako") 		= RTrim(objSt.Range("O" & lngRow))		'// 個装箱代
		case "C"	'⑤入庫
			objRs.Fields("Dt") 			= Replace(RTrim(objSt.Range("B" & lngRow)),"/","")		'入庫日
			objRs.Fields("DenNo") 		= RTrim(objSt.Range("C" & lngRow))						'伝№
			objRs.Fields("SyukaCd")		= RTrim(objSt.Range("D" & lngRow))						'相手先
			objRs.Fields("Pn") 			= RTrim(objSt.Range("E" & lngRow))		'品番
			objRs.Fields("PnName") 		= RTrim(objSt.Range("F" & lngRow))		'品名
			objRs.Fields("Qty") 		= RTrim(objSt.Range("G" & lngRow))		'数量
																				'棚番
			objRs.Fields("AnyKbn") 		= Get_LeftB(RTrim(objSt.Range("I" & lngRow)),10)		'入庫区分
																				'入庫工料 単価
			objRs.Fields("Pick") 		= RTrim(objSt.Range("K" & lngRow))		'入庫工料 金額
		case "D"	'②③④⑤ＡＣ商品化
			objRs.Fields("Dt") 			= GetDt(objSt.Range("C" & lngRow))		'受入日
			objRs.Fields("Pn") 			= Get_LeftB(RTrim(objSt.Range("D" & lngRow)),20)		'品番
			objRs.Fields("PnName") 		= RTrim(objSt.Range("E" & lngRow))		'品名
			objRs.Fields("Qty") 		= RTrim(objSt.Range("F" & lngRow))		'数量
			objRs.Fields("KoryoPrc") 	= RTrim(objSt.Range("G" & lngRow))		'工料＠
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("H" & lngRow))		'工料金額
			objRs.Fields("HakoPrc") 	= RTrim(objSt.Range("I" & lngRow))		'箱代＠
			objRs.Fields("Hako") 		= RTrim(objSt.Range("J" & lngRow))		'箱代金額
			objRs.Fields("GaisoPrc") 	= RTrim(objSt.Range("K" & lngRow))		'外装＠
			objRs.Fields("Gaiso") 		= RTrim(objSt.Range("L" & lngRow))		'外装金額
			objRs.Fields("FutaiPrc") 	= RTrim(objSt.Range("M" & lngRow))		'付帯＠
			objRs.Fields("Futai") 		= RTrim(objSt.Range("N" & lngRow))		'付帯金額
		case "E"	'②③商品化工料
			objRs.Fields("Dt") 			= GetDt(objSt.Range("B" & lngRow))		'受入日
			objRs.Fields("Pn") 			= RTrim(objSt.Range("C" & lngRow))		'品番
			objRs.Fields("PnName") 		= RTrim(objSt.Range("D" & lngRow))		'品名
			objRs.Fields("Qty") 		= RTrim(objSt.Range("E" & lngRow))		'数量
			objRs.Fields("KoryoPrc") 	= RTrim(objSt.Range("F" & lngRow))		'工料＠
			objRs.Fields("Koryo") 		= RTrim(objSt.Range("G" & lngRow))		'工料金額
			objRs.Fields("FutaiPrc") 	= RTrim(objSt.Range("H" & lngRow))		'付加作業＠
			objRs.Fields("Futai") 		= RTrim(objSt.Range("I" & lngRow))		'付加作業金額
		end select
		objRs.UpdateBatch
	next
	dim	strStat
	strStat = "head"

	Call DispMsg("  読込件数：" & lngRow)
	Call DispMsg("  登録件数：" & cntAdd)

End Function

Function GetDt(byVal strDt)
	strDt = RTrim(strDt)
	if inStr(strDt,"/") > 0 then
		strDt = Replace(strDt,"/","")
	end if
	GetDt = left(strDt,8)
End Function

Function GetNumValue(strV)
	dim	dblV
' for debug
'	Wscript.Echo "GetNumValue(" & len(rtrim(strV)) & " " & rtrim(strV) & ")"
' for debug
	dblV = 0
	if isnumeric(strV) = True then
		dblV = cdbl(strV)
	end if
	GetNumValue = dblV
End Function

Private Function GetYM(byVal strDt)
	dim	strYM
	dim	iY
	dim	iM
	dim	iD
	Call Debug("GetYM(" & strDt & ")")
	if inStr(strDt,"/") > 0then
		iY = CInt(Split(strDt,"/")(0))
		iM = CInt(Split(strDt,"/")(1))
		iD = CInt(Split(strDt,"/")(2))
	else
		iY = CInt(Left(strDt,4))
		iM = CInt(Mid(strDt,5,2))
		iD = CInt(Right(strDt,2))
	end if
	if iD > 20 Then
		iM = iM + 1
	end if
	if iM > 12 Then
		iY = iY + 1
		iM = 1
	end if
	strYM = iY & Right("0" & iM,2)
	GetYM = strYM
End Function

Private Function List()
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	Call DispMsg("objDb.Execute(" & strSql & ")")
	set rsList = objDb.Execute(strSql)

	do while rsList.Eof = False
		' DT 	JGYOBU 	NAIGAI 	HIN_GAI 	SyukaCnt 	SyukaQty
		Call DispMsg("■" _
			 & " " & rsList.Fields("DT") _
			 & " " & rsList.Fields("JGYOBU") _
			 & " " & rsList.Fields("NAIGAI") _
			 & " " & rsList.Fields("HIN_GAI") _
			 & " " & rsList.Fields("SyukaCnt") _
			 & " " & rsList.Fields("SyukaQty") _
					)
		Call rsList.MoveNext
	loop

	Call DispMsg("CloseAdodb(" & GetOption("db","newsdc") & ")")
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
	strSql = strSql & " from MonthlyQty"
	makeSql = strSql
End Function

