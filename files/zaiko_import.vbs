Option Explicit
Function Include(byval strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
	strFileName = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")) & strFileName
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
End Function
Call Include("const.vbs")

Wscript.Quit Main()

Function usage()
    Wscript.Echo "zaiko_import(2011.12.24)"
    Wscript.Echo "zaiko_import.vbs <在庫データファイル>"
    Wscript.Echo "<例>"
    Wscript.Echo "zaiko_import.vbs ＰＰＳＣ在庫111219.xls"
    Wscript.Echo "WScript.Path          =" & WScript.Path
    Wscript.Echo "WScript.ScriptName    =" & WScript.ScriptName
    Wscript.Echo "WScript.ScriptFullName=" & WScript.ScriptFullName
    Wscript.Echo Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
End Function

Function Main()
	dim	strFilename
	dim	strArg
	dim	objFSO
	dim	objLog

	set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	set objLog = OpenLogFile(objFSO)
	call WriteLogFile(objLog,"start")
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		else
			call usage()
			Main = -1
			exit function
		end if
	next
	if strFilename = "" then
		call usage()
		Main = -1
		exit function
	end if
	if lcase(Right(strFilename,4)) = ".zip" then
		strFilename = UnZip(strFilename)
	end if

	Call ConvertZaiko(strFilename,objLog)
	call CloseLogFile(objLog)
	set objFSO = Nothing
	Main = 0
End Function
'-----------------------------------------------------------------------
'zipファイル展開
'-----------------------------------------------------------------------
Const FOF_SILENT 			= &H04 	'進捗ダイアログを表示しない。
Const FOF_RENAMEONCOLLISION = &H08 	'ファイルやフォルダ名が重複するときは「コピー ～ 」のようなファイル名にリネームする。
Const FOF_NOCONFIRMATION 	= &H10 	'上書き確認ダイアログを表示しない（[すべて上書き]と同じ）。
Const FOF_ALLOWUNDO 		= &H40 	'操作の取り消し（[編集]-[元に戻す]や{ctrl}+{z}）を有効にする。
Const FOF_FILESONLY 		= &H80 	'ワイルドカードが指定された場合のみ実行する。
Const FOF_SIMPLEPROGRESS 	= &H100 '進捗ダイアログは表示するがファイル名は表示しない。
Const FOF_NOCONFIRMMKDIR 	= &H200 'フォルダ作成確認ダイアログを表示しない（自動で作成）。
Const FOF_NOERRORUI 		= &H400 'コピーや移動ができなかった場合の実行時エラーを発生させない。ただし、対象のファイルを飛ばして処理を続けるわけではないことに注意。
Const FOF_NORECURSION 		= &H1000 'サブフォルダ内のファイルはコピーしない（ただし、フォルダは作成される）。
Function UnZip(byVal strFilename)
	dim	objShell
	Set objShell = WScript.CreateObject("Shell.Application")
	dim	objNs
	Set objNs = objShell.NameSpace(strFilename)
	dim	objItm
	For Each objItm in objNs.Items
		strFilename = GetScriptPath() & objItm.Name
		Call XDeleteFile(strFilename)
		dim	objCurr
		Set objCurr = objShell.NameSpace(GetScriptPath())
		Call objCurr.CopyHere(objItm,FOF_NOCONFIRMATION)
		Exit For
	Next
	set objShell = Nothing
	UnZip = strFilename
End Function
Function GetScriptPath()
	GetScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
End Function
Function XDeleteFile(byVal strFilename)
	Dim objFileSys
	
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	
	if objFileSys.FileExists(strFilename) = True Then
		objFileSys.DeleteFile strFilename, True
	End if
	
	Set objFileSys = Nothing
	XDeleteFile = strFilename
End Function

'-----------------------------------------------------------------------
'ログファイル Open
'-----------------------------------------------------------------------
Private Function OpenLogFile(byval objFSO)
	dim	objFile
	dim	strFilename

	strFilename = Wscript.ScriptFullName
	strFilename = left(strFilename,len(strFilename)-3)
	strFilename = strFilename & "log"
	Set objFile = objFSO.OpenTextFile(strFilename, ForWriting, True)
	set OpenLogFile = objFile
End Function
'-----------------------------------------------------------------------
'ログファイル Close
'-----------------------------------------------------------------------
Private Function CloseLogFile(byval objFile)
	objFile.Close
	set CloseLogFile = Nothing
End Function
'-----------------------------------------------------------------------
'ログファイル Write
'-----------------------------------------------------------------------
Private Function WriteLogFile(byval objFile,byval strMsg)
	Wscript.Echo strMsg
	objFile.WriteLine strMsg
End Function
'-----------------------------------------------------------------------
'ログファイル Err表示
'-----------------------------------------------------------------------
Private Function ErrLogFile(byval objFile,byval objErr)
	dim	strMsg
	if objErr.Number <> 0 then
		strMsg = "Error.Number:" & objErr.Number
		Call WriteLogFile(objFile,strMsg)
		strMsg = "Error.Description:" & objErr.Description
		Call WriteLogFile(objFile,strMsg)
	end if
End Function

'-----------------------------------------------------------------------
'在庫データ変換
'-----------------------------------------------------------------------
Private Function ConvertZaiko(byval strFilename _
							 ,byval objLog)
	call WriteLogFile(objLog,"ConvertZaiko(" & strFilename & ")")
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	dim	objXL
	dim	objBk
	dim	objSt
	Set objXL = WScript.CreateObject("Excel.Application")
	Call ErrLogFile(objLog,Err)
'	objXL.Application.Visible = True
	Call WriteLogFile(objLog,"Workbooks.Open()")
	Set objBk = objXL.Workbooks.Open(strFilename,False,True)
	Call ErrLogFile(objLog,Err)
	Call WriteLogFile(objLog,"set ActiveSheet")
	set objSt = objBk.ActiveSheet
'	Call WriteLogFile(objLog,"set Worksheets(1)")
'	set objSt = objBk.Worksheets(1)
	Call ErrLogFile(objLog,Err)
	Call WriteLogFile(objLog,"objSt.Name=" & objSt.Name)
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	cnnDb
	dim	rsZaiko
	dim	strDbName
	Set cnnDb = Wscript.CreateObject("ADODB.Connection")
												Call ErrLogFile(objLog,Err)
	strDbName = "newsdc"
	Call cnnDb.Open(strDbName)
												Call ErrLogFile(objLog,Err)
	Call cnnDb.Execute("delete from ZaikoGlics where JKubun = 'A'")
	' テーブルOpen
	Set rsZaiko = Wscript.CreateObject("ADODB.Recordset")
												Call ErrLogFile(objLog,Err)
	rsZaiko.MaxRecords = 1
	rsZaiko.Open "ZaikoGlics", cnnDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
												Call ErrLogFile(objLog,Err)
	'-------------------------------------------------------------------

	'-------------------------------------------------------------------
	'Excelの読込
	'-------------------------------------------------------------------
	dim	lngRow
	dim	lngMaxRow
	dim	strJGyobu
	dim	strPn
	dim	strSyushi
	dim	strCol
	dim	aryCol
	dim	strSql
	aryCol = Array("D","E","F","G","H","I")
	lngMaxRow = 65535
	for lngRow = 1 to lngMaxRow
		if isempty(objSt.Range("A" & lngRow)) then
			exit for
		end if
		strJGyobu = objSt.Range("A" & lngRow)
		strPn = objSt.Range("B" & lngRow)
		if lngRow > 1 then
			if trim(objSt.Range("C" & lngRow)) <> "" then
				strSql = "update item"
				strSql = strSql & " set CS_TANTO_CD = '" & objSt.Range("C" & lngRow) & "'"
				strSql = strSql & " where jgyobu = 'A'"
				strSql = strSql & "   and naigai = '1'"
				strSql = strSql & "   and hin_gai = '" & strPn & "'"
				strSql = strSql & "   and CS_TANTO_CD <> '" & objSt.Range("C" & lngRow) & "'"
				Call cnnDb.Execute(strSql)
			end if
			for each strCol in aryCol
				if objSt.Range(strCol & lngRow) <> 0 then
					strSyushi = objSt.Range(strCol & "1")
					select case strSyushi
					case "センター倉庫"
							strSyushi = "11"
					end select
					rsZaiko.Addnew
																Call ErrLogFile(objLog,Err)
					rsZaiko.Fields("JKubun")		= "A"
					rsZaiko.Fields("JCode")			= objSt.Range("A" & lngRow)
					rsZaiko.Fields("Syushi")		= strSyushi
					rsZaiko.Fields("SSyushi")		= objSt.Range("A" & lngRow)
					rsZaiko.Fields("HojoSyushi")	= "00000000"
					rsZaiko.Fields("Pn")			= objSt.Range("B" & lngRow)
					rsZaiko.Fields("ShizaiPn")		= ""
					rsZaiko.Fields("PName")			= ""
					rsZaiko.Fields("Location1")		= objSt.Range("C" & lngRow)
					rsZaiko.Fields("Qty")			= objSt.Range(strCol & lngRow)
					rsZaiko.Fields("Tm")			= ""
					rsZaiko.UpdateBatch
																	Call ErrLogFile(objLog,Err)
				end if
			next
		end if
'		Call WriteLogFile(objLog,strJGyobu & " " & strPn)
'		if lngRow >= 100 then
'			exit for
'		end if
	next
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	'-------------------------------------------------------------------
	'データベースの後処理
	'-------------------------------------------------------------------
	Call rsZaiko.Close
	set rsZaiko = Nothing
	Call cnnDb.Close
	set cnnDb = Nothing
End Function
