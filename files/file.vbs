'-----------------------------------------------------------------------
Function GetCD()
	Dim objWshShell
	'①WScript.Shellオブジェクトの作成
	Set objWshShell = CreateObject("WScript.Shell")
	'カレントディレクトリを表示
	dim	strCD
	strCD = objWshShell.CurrentDirectory
	Set objWshShell = Nothing
	GetCD = strCD
End Function

Function GetAbsPath(byVal strPath)
	Dim objFileSys
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	strPath = objFileSys.GetAbsolutePathName(strPath)
	Set objFileSys = Nothing
	GetAbsPath = strPath
End Function

Function GetScriptPath()
	GetScriptPath = Left(WScript.ScriptFullName,Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
End Function

Function GetFileName(byVal strFullName)
	dim	strFileName
	strFileName = strFullName
	dim	c
	for each c in split(strFileName,"\")
		Call Debug("GetFileName():" & c)
		if c <> "" then
			strFileName = c
		end if
	next
	GetFileName = strFileName
End Function
'-----------------------------------------------------------------------
'ファイル一覧
'-----------------------------------------------------------------------
Function FileList(byval strPath,byval strRcv)
	Dim objFileSys
	Dim strScriptPath
	Dim strTargetPath
	Dim objFolder
	Dim objItem

	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

	strTargetPath = objFileSys.BuildPath(strScriptPath, strPath)

	Set objFolder = objFileSys.GetFolder(strTargetPath)

	Call Debug("FileList():" & strTargetPath)
	dim	aryFile()
	dim	i
	i = 0
	dim	strList
	strList = ""
	For Each objItem In objFolder.Files
	    Call Debug("FileList():" & objItem.Name)
	    Call Debug("FileList():" & objItem.Path)
	    Call Debug("FileList():" & objItem.DateLastModified)
		strList = strList & TmFormat(objItem.DateLastModified)
		strList = strList & " " & NumFormat(objItem.Size)
		strList = strList & " " & objItem.Name
		strList = strList & vbCrlf

		Redim Preserve aryFile(i)
		aryFile(i) = objItem.Path
		i = i + 1
	Next

	Call Debug("FileList():ファイル数：" & objFolder.Files.Count)

	Set objFolder = Nothing
	Set objFileSys = Nothing
	if strRcv = "list" then
		FileList = strList
	else
		FileList = aryFile
	end if
End Function

Private Function TmFormat(v)
	dim	dt
	dim	tm
	dt = Split(v," ")(0)
	tm = Split(v," ")(1)
	TmFormat = dt & Right("  " & tm,9)
End Function

Private Function NumFormat(v)
	NumFormat = Right(Space(12) & FormatNumber(v,0,,-1),12)
End Function

Function DeleteFile(byVal path)
	Dim objFileSys
	Dim strScriptPath
	Dim strDeleteFrom
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
'	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
'	strDeleteFrom = objFileSys.BuildPath(strScriptPath, "\backup\TestData.csv")
	objFileSys.DeleteFile path, True
'	WScript.echo "BackUpからTestData.csvを削除しました。"
	Set objFileSys = Nothing
End Function
