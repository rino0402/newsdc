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
	Wscript.Echo "JCS MF(Excel)データ変換"
	Wscript.Echo "TblCopy.vbs [option] <コピー元> <コピー先>"
	Wscript.Echo " /db:newsdc7"
	Wscript.Echo " /debug"
	Wscript.Echo "sc32 TblCopy.vbs /db:newsdc7 JcsItem_Tmp JcsItem"
	Wscript.Echo "sc32 TblCopy.vbs /db:newsdc7 JcsIdo_Tmp JcsIdo"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strSrc
	dim	strDst
	strSrc = ""
	strDst = ""
	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		if strSrc = "" then
			strSrc = strArg
		elseif strDst = "" then
			strDst = strArg
		else
			usage()
			Main = 1
			exit Function
		end if
	next
	'名前付きオプションチェック
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			strSrc = ""
		case else
			strSrc = ""
		end select
	next
	if strSrc = "" then
		usage()
		Main = 1
		exit Function
	end if
	call TblCopy(strSrc,strDst)
	Main = 0
End Function

Function TblCopy(byVal strSrc,byVal strDst)
	Call Debug("TblCopy(" & strSrc & "," & strDst & ")")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'テーブルコピー
	'-------------------------------------------------------------------
	Call DispMsg("テーブルコピー:" & strSrc & "→" & strDst)
	Call CopyTable(objDb,strSrc,strDst)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

