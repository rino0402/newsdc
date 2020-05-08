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
	Wscript.Echo "エアコン 出荷データ 201309-201004"
	Wscript.Echo "AcSyuka.vbs [option]"
	Wscript.Echo " /db:<database>"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
	Wscript.Echo " /debug"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript AcSyuka.vbs /db:newsdc4 /load ""I:\0SDC_honsya\事業部別商品化出荷金額まとめ\AC NPLからの出荷実績\201309AC出荷実績.xlsx"""
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
'冷蔵庫出荷データ(Excel)変換→MonthlyQty
'-------------------------------------------------------------------
Private Function Load(byval strFilename)
	'-------------------------------------------------------------------
	'データベース準備
	'-------------------------------------------------------------------
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc4") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc4"))
	'-------------------------------------------------------------------
	'登録用レコードセット準備
	'-------------------------------------------------------------------
	dim	objRs
	set objRs = OpenRs(objDb,"BoSyuka")
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
	'-------------------------------------------------------------------
	'Excel最終行
	'-------------------------------------------------------------------
	Const xlUp = -4162
	dim	lngRowMax
	lngRowMax = objSt.Range("B65536").End(xlUp).Row

	dim	cntAdd
	cntAdd = 0

	'-------------------------------------------------------------------
	'年月取得
	'-------------------------------------------------------------------
	dim	rngYM
	set rngYM = objSt.Range("BF2")
	dim	strCol
	strCol = ""
	do while strCol <> "K"
		'-------------------------------------------------------------------
		'Excel列名取得
		'-------------------------------------------------------------------
		strCol = Split(rngYM.Address,"$")(1)

		dim	strYM
		strYM = GetYM(rngYM)
		Call DispMsg(strCol & ":年月:" & strYM)
		if strYM <> "" then
			'-------------------------------------------------------------------
			'年月データ削除
			'-------------------------------------------------------------------
			dim	strSql
			strSql = "delete from BoSyuka"
			strSql = strSql & " where ShisanJCode='00025800'"
			strSql = strSql & "   and DT like '" & strYM & "%'"
			Call DispMsg("削除:" & strSql)
			Call ExecuteAdodb(objDb,strSql)
			'-------------------------------------------------------------------
			'ループ：3～最終行
			'-------------------------------------------------------------------
			dim lngRow
			for lngRow = 3 to lngRowMax
				'-------------------------------------------------------------------
				'A：資産管理事業場
				'-------------------------------------------------------------------
				dim	strJCode
				strJCode = RTrim(objSt.Range("A" & lngRow))
				'-------------------------------------------------------------------
				'C：品番
				'-------------------------------------------------------------------
				dim	strPn
				strPn = RTrim(objSt.Range("C" & lngRow))
				'-------------------------------------------------------------------
				'出荷数
				'-------------------------------------------------------------------
				dim	strQty
				strQty = RTrim(objSt.Range(strCol & lngRow))
				'-------------------------------------------------------------------
				'IdNo
				'-------------------------------------------------------------------
				dim	strIdNo
				strIdNo = strYM & Right("00000" & lngRow,5)
				'-------------------------------------------------------------------
				'レコード追加
				'-------------------------------------------------------------------
				Call DispMsg("年月:" & strYM & ":" & strCol & lngRow & ":" & strPn & " " & strQty)
				if strQty <> "" then
					if strPn <> RTrim(objSt.Range("C" & lngRow - 1)) then
						cntAdd = cntAdd + 1
						objRs.AddNew
						objRs.Fields("IdNo") = strIdNo
						objRs.Fields("ShisanJCode") = strJCode
						objRs.Fields("Dt") = strYM & "01"
						objRs.Fields("Pn") = strPn
						objRs.Fields("Qty") = strQty
						objRs.UpdateBatch
					end if
				end if
			next
		end if
		set rngYM = rngYM.Offset(0,-1)
	loop
	dim	strStat
	strStat = "head"

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

Private Function GetYM(rngYM)
	GetYM = ""
	dim	strYM
	strYM = rngYM
	if Len(strYM) <> 6 then
		exit function
	end if
	if isNumeric(strYM) = false then
		exit function
	end if
	GetYM = strYM
End Function

Private Function List()
End Function
