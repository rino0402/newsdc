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

dim	lngRet
lngRet = Main()
WScript.Quit lngRet

Function FormatDt(byval cnvMOTO,byval iLen)
	dim	strD
	strD = ""
	If IsDate(cnvMOTO) Then
		strD = strD & Right("0000"	&   Year(cnvMOTO), 4)  '年
		strD = strD & Right("00"	&  Month(cnvMOTO), 2)  '月
		strD = strD & Right("00"	&    Day(cnvMOTO), 2)  '日
		strD = strD & Right("00"	&   Hour(cnvMOTO), 2)  '時
		strD = strD & Right("00"	& Minute(cnvMOTO), 2)  '分
		strD = strD & Right("00"	& Second(cnvMOTO), 2)  '秒
		if iLen > 0 then
			strD = Left(strD,iLen)
		end if
	end if
	FormatDt = strD
End Function

'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "商品化比例データ作成"
	Wscript.Echo "s_hirei.vbs [option] [<Read filename>] [<Save filename>]"
	Wscript.Echo "Ex."
	Wscript.Echo "s_hirei.vbs s_hirei.xlsm s_hirei_" & FormatDt(Now(), 12) & ".xlsm"
'	dim	strDate
'	strDate = FormatDt(Now(), 0)
'	Wscript.Echo strDate
'	strDate = FormatDt(Now(), 12)
'	Wscript.Echo strDate
End Sub

'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilenameR
	dim	strFilenameS
	strFilenameR = ""
	strFilenameS = ""
	dim	strOption
	strOption = ""
	For Each strArg In WScript.Arguments
'		Call DispMsg(strArg)
		select case strOption
		case else
			strArg = lcase(strArg)
	    	select case Split(strArg,":")(0)
			case "/debug"
			case "/?"
				usage()
				Main = 1
				exit Function
			case else
				if strFilenameR = "" then
					strFilenameR = strArg
				elseif strFilenameS = "" then
					strFilenameS = strArg
				else
					Call DispMsg("ファイル名が多すぎます.")
					usage()
					Main = 1
					exit Function
				end if
			end select
		end select
	Next
	if strOption <> "" then
		Call DispMsg("OptionErorr:" & strOption )
		usage()
		Main = 1
		exit Function
	end if
	if strFilenameR = "" then
		strFilenameR = "s_hirei.xlsm"
	end if
	if strFilenameS = "" then
		strFilenameS = "s_hirei_" & FormatDt(Now(),12) & ".xlsm"
	end if
	dim	strMsg
	strMsg = ""
	strMsg = Load(strFilenameR,strFilenameS)
	Call DispMsg(strMsg)
	Main = 0
End Function

Function Load(byval strFilenameR,byval strFilenameS)
	Call Debug("Load():FilenameR:" & strFilenameR)
	Call Debug("Load():FilenameS:" & strFilenameS)
	dim	strMsg
	strMsg = ""
	'-------------------------------------------------------------------
	'Excelファイル名
	'-------------------------------------------------------------------
	strFilenameR = GetAbsPath(strFilenameR)
	Call Debug("Load():FilenameR:" & strFilenameR)
	if strFilenameR = "" then
		Call DispMsg("ファイル名を指定して下さい")
		Exit Function
	end if
	strFilenameS = GetAbsPath(strFilenameS)
	Call Debug("Load():FilenameS:" & strFilenameS)
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	dim	objXL
	Set objXL = WScript.CreateObject("Excel.Application")
	Call Debug("Load():CreateObject(Excel.Application)")
	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	dim	strPassword
	strPassword = ""
	dim	objBk
	Set objBk = objXL.Workbooks.Open(strFilenameR,False,True,,strPassword)
	Call Debug("Load():Workbooks.Open:" & objBk.Name)
	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	if LoadXls(objXL,objBk) = 0 then
		'-------------------------------------------------------------------
		'Excelの後処理 名前を付けて保存
		'-------------------------------------------------------------------
		'警告を非表示
		objXL.DisplayAlerts = False
		'同じ名前のファイルがあったときには強制的に上書
		Call objBk.SaveAs(strFilenameS)
		strMsg = "正常終了:" & strFilenameS
	else
		strMsg = "読込エラー"
	end if
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("Load():End")
	Load = strMsg
End Function

Function LoadXls(objXL,objBk)
	Call Debug("LoadXls():" & objBk.Name)
	'-------------------------------------------------------------------
	'データベース接続を削除
	'-------------------------------------------------------------------
	dim	objCnn
	for each objCnn in objBk.Connections
		Call Debug("LoadXls():Refresh():" & objCnn.Name)
		objCnn.ODBCConnection.BackgroundQuery = False
		Call objCnn.Refresh()
		Call Debug("LoadXls():Delete():" & objCnn.Name)
		Call objCnn.Delete()
	next
	dim objSt
	for each objSt in objBk.Worksheets
		Call Debug("LoadXls():Sheet:" & objSt.Name)
		select case objSt.Name
		case "商品化金額チェック"
			Call Debug("LoadXls():Activate:" & objSt.Name)
			dim p
		    For Each p In objSt.PivotTables
				Call Debug("LoadXls():PivotTable.PivotCache.Refresh():" & p.Name)
		        call p.PivotCache.Refresh
			next
			objSt.Activate
			objSt.Range("E1").Select
		case "Bo振替実績"
		case else
			Call DispMsg("シート名が不正です。" & objSt.Name)
			LoadXls = -1
			Exit Function
		end select
	next
	LoadXls = 0
End Function
