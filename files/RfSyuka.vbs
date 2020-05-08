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
	Wscript.Echo "冷蔵庫 出荷データ"
	Wscript.Echo "RfSyuka.vbs [option]"
	Wscript.Echo " /db:<database>"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
	Wscript.Echo " /debug"
	Wscript.Echo "Ex."
	Wscript.Echo "RfSyuka.vbs /db:newsdc-4 /load ""冷蔵庫サービス部品月別出荷実績2012.02から03. (1).xls"""
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
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	'-------------------------------------------------------------------
	'登録用レコードセット準備
	'-------------------------------------------------------------------
	dim	objRs
	set objRs = OpenRs(objDb,"MonthlyQty")
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
	set rngYM = objSt.Range("H2")
	do while GetYM(rngYM) <> ""
		dim	strYM
		strYM = GetYM(rngYM)
		Call Debug("年月:" & strYM)
		'-------------------------------------------------------------------
		'年月データ削除
		'-------------------------------------------------------------------
		dim	strSql
		strSql = "delete from MonthlyQty" _
			   & " where JGYOBU = 'R'" _
			   & "   and NAIGAI = '0'" _
			   & "   and DT = '" & strYM & "'"
		Call Debug("削除:" & strSql)
		Call ExecuteAdodb(objDb,strSql)
		'-------------------------------------------------------------------
		'Excel列名取得
		'-------------------------------------------------------------------
		dim	strCol
		strCol = Split(rngYM.Address,"$")(1)
		'-------------------------------------------------------------------
		'ループ：3～最終行
		'-------------------------------------------------------------------
		dim lngRow
		for lngRow = 3 to lngRowMax
			'-------------------------------------------------------------------
			'C列：品番
			'-------------------------------------------------------------------
			dim	strPn
			strPn = RTrim(objSt.Range("C" & lngRow))
			'-------------------------------------------------------------------
			'出荷数
			'-------------------------------------------------------------------
			dim	strQty
			strQty = RTrim(objSt.Range(strCol & lngRow))
			'-------------------------------------------------------------------
			'レコード追加
			'-------------------------------------------------------------------
			Call Debug("年月:" & strYM & ":" & strCol & lngRow & ":" & strPn & " " & strQty)
			if strQty <> "" then
				if strPn <> RTrim(objSt.Range("C" & lngRow - 1)) then
					cntAdd = cntAdd + 1
					objRs.AddNew
					objRs.Fields("DT") = strYM
					objRs.Fields("JGYOBU") = "R"
					objRs.Fields("NAIGAI") = "0"
					objRs.Fields("HIN_GAI") = strPn
					objRs.Fields("SyukaCnt") = "1"
					objRs.Fields("SyukaQty") = strQty
					objRs.UpdateBatch
				end if
			end if
		next
		set rngYM = rngYM.Offset(0,1)
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
	dim	strYM
	strYM = rngYM
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

