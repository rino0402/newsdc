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
	Wscript.Echo "Glicsユニットマスター"
	Wscript.Echo "Unit.vbs [option]"
	Wscript.Echo " /db:<database>"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
	Wscript.Echo " /debug"
	Wscript.Echo "Ex."
	Wscript.Echo "Unit.vbs /db:newsdc-ono /load ""I:\pos\商品化計画\ユニット\DOM13_15 【検討表】ユニットマスタ.xls"""
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
	set objRs = OpenRs(objDb,"Unit")
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
	strJGyobu = "D"
	if strJGyobu = "" then
		call DispMsg("事業部不明")
	else
		Call LoadUnit(objDb,objRs,objBk,strJGyobu)
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

Function GetUnitSheet(objBk)
	set GetUnitSheet = Nothing
	dim	objSt
	for each objSt in objBk.Worksheets
		set GetUnitSheet = objSt
		exit for
	next
End Function


Function LoadUnit(objDb,objRs,objBk,byVal strJGyobu)

	Call Debug("LoadUnit(" & strJGyobu & ")")
	dim	objSt
	set objSt = GetUnitSheet(objBk)
	if objSt is Nothing then
		Call DispMsg("LoadUnit():指定シートがありません.")
		exit function
	end if
	'-------------------------------------------------------------------
	'既存データ削除
	'-------------------------------------------------------------------
	dim	strSql
	strSql = "delete from Unit"
	Call Debug("削除:" & strSql)
	Call ExecuteAdodb(objDb,strSql)
	'-------------------------------------------------------------------
	'Excel最終行
	'-------------------------------------------------------------------
	Const xlUp = -4162
	dim	lngRowTop
	dim	lngRowMax
	lngRowTop = 3
	lngRowMax = objSt.Range("B65536").End(xlUp).Row

	dim	cntAdd
	cntAdd = 0

	'-------------------------------------------------------------------
	'ループ：3～最終行
	'-------------------------------------------------------------------
	dim	lngRow
	for lngRow = lngRowTop to lngRowMax
		Call Debug(strJGyobu & " " & lngRow & _
					" " & RTrim(objSt.Range("A" & lngRow)) & _
					" " & RTrim(objSt.Range("B" & lngRow)) & _
					" " & RTrim(objSt.Range("C" & lngRow)) & _
					" " & RTrim(objSt.Range("D" & lngRow)) _
					)
		dim	lngNo
		cntAdd = cntAdd + 1
		objRs.AddNew
		objRs.Fields("Pn") 		= RTrim(objSt.Range("B" & lngRow))		'// ユニット_親品目番号
		objRs.Fields("PnC") 	= RTrim(objSt.Range("C" & lngRow))		'// ユニット_子品目番号
		objRs.Fields("Qty") 	= RTrim(objSt.Range("D" & lngRow))		'// ユニット_構成品員数
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
	GetDt = strDt
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

