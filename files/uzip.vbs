Option Explicit

Dim objSA
Dim objBefore
Dim objAfter

Const FOF_SILENT 			= &H04 	'進捗ダイアログを表示しない。
Const FOF_RENAMEONCOLLISION = &H08 	'ファイルやフォルダ名が重複するときは「コピー ～ 」のようなファイル名にリネームする。
Const FOF_NOCONFIRMATION 	= &H10 	'上書き確認ダイアログを表示しない（[すべて上書き]と同じ）。
Const FOF_ALLOWUNDO 		= &H40 	'操作の取り消し（[編集]-[元に戻す]や{ctrl}+{z}）を有効にする。
Const FOF_FILESONLY 		= &H80 	'ワイルドカードが指定された場合のみ実行する。
Const FOF_SIMPLEPROGRESS 	= &H100 '進捗ダイアログは表示するがファイル名は表示しない。
Const FOF_NOCONFIRMMKDIR 	= &H200 'フォルダ作成確認ダイアログを表示しない（自動で作成）。
Const FOF_NOERRORUI 		= &H400 'コピーや移動ができなかった場合の実行時エラーを発生させない。ただし、対象のファイルを飛ばして処理を続けるわけではないことに注意。
Const FOF_NORECURSION 		= &H1000 'サブフォルダ内のファイルはコピーしない（ただし、フォルダは作成される）。

dim	sa
Set sa = WScript.CreateObject("Shell.Application")

dim	arg
For Each arg In WScript.Arguments
	dim	src
	Set src = sa.NameSpace(arg)
'	src.ParentFolder.CopyHere(src.Items)
	Wscript.Echo arg
	dim	itm
	For Each itm in src.Items
		Wscript.Echo itm.Name
		Call XDeleteFile(itm.Name)
		Call src.ParentFolder.CopyHere(itm,FOF_NOCONFIRMATION)
	Next
'	Wscript.Echo "FOF=0x" & Hex(FOF_NOCONFIRMATION)
Next

Function XDeleteFile(byVal txtFilename)
	Dim objFileSys
	Dim strScriptPath
	Dim strDeleteFrom
	
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	
	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	
	strDeleteFrom = objFileSys.BuildPath(strScriptPath, "\" & txtFilename)
	
	WScript.echo "DeleteFile:" & strDeleteFrom

	if objFileSys.FileExists(strDeleteFrom) = True Then
		objFileSys.DeleteFile strDeleteFrom, True
	End if
	
	Set objFileSys = Nothing
End Function


