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
Call Include("BoConv_sub.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "BOデータ変換(メール自動受信)"
	Wscript.Echo "newsdc9.vbs [option]"
	Wscript.Echo "/save	受信メールをサーバーに残す(default)"
	Wscript.Echo "/savd	受信メールを削除する"
	Wscript.Echo "/load	<filename>"
	Wscript.Echo "Ex."
	Wscript.Echo "newsdc9.vbs"
	Wscript.Echo "newsdc9.vbs /load temp\棚番ﾃﾞｰﾀ.csv"

	dim	c
	for each c in FileList("temp","path")
		Call DispMsg("FileList():" & c)
	next

End Sub

'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	dim	strOption
	strOption = ""
	dim	strDb
	strDb = "newsdc9"
'	For Each strArg In WScript.Arguments.UnNamed
'    	select case strArg
'		case else
'			usage()
'			Main = 1
'			exit Function
'		end select
'	Next
	For Each strArg In WScript.Arguments
'		Call DispMsg(strArg)
		select case strOption
		case "/load"
			strFilename = strArg
			strOption = ""
		case else
			strArg = lcase(strArg)
	    	select case Split(strArg,":")(0)
			case "/db"
				strDb = Split(strArg,":")(1)
			case "/debug"
			case "/save"
			case "/savd"
			case "/load"
				strOption = strArg
			case else
				Call DispMsg("unknown:" & strArg )
				usage()
				Main = 1
				exit Function
			end select
		end select
	Next
	if strOption <> "" then
		Call DispMsg("option:" & strOption )
		usage()
		Main = 1
		exit Function
	end if
	if strFilename <> "" then
		dim	strMsg
		strMsg = Load(strDb,strFilename)
		Call DispMsg(strMsg)
	else
		call RcvMail()
	end if

	Main = 0
End Function

'-----------------------------------------------------------------------
'RcvMail()のオプション
'-----------------------------------------------------------------------
Private Function RcvMailOpt()
	RcvMailOpt = "SAVE 1-10"
	if WScript.Arguments.Named.Exists("savd") then
		RcvMailOpt = "SAVD 1-10"
	end if
End Function

'-----------------------------------------------------------------------
'メール受信
'-----------------------------------------------------------------------
Private Function RcvMail()
	Call Debug("RcvMail()")
	' メール送受信APIの宣言
	dim	bobj
	Set bobj = CreateObject("Basp21")

	dim	svname
	dim	user
	dim	pass
	dim	dirname
	dim	strDb

	strDb = "newsdctest"

	svname	= "ns"						' POP3サーバマシン名
	user	= "newsdc9"					' メールボックス名
	pass	= "123daa@Z"				' パスワード
	dirname = "rcvtemp"					' 保存ディレクトリ名

'       SAVE n[-n2] .... n番目のメールを受信します
'                        n2を指定するとn2番目までのメールを受信します。
'       SAVD n[-n2] ... n番目のメールを受信し、サーバのメールボックスから
'                   削除します
'                   n2を指定するとn2番目までのメールを受信して削除します
	Call DispMsg("メール受信中:" & RcvMailOpt())
	dim	outarray
	outarray = bobj.RcvMail(svname,user,pass,RcvMailOpt(),dirname)
	if IsArray(outarray) then	' OK ?
		dim	file
		for each file in outarray
			dim	strSubject
			dim	strSubject1
			dim	strSubject2
			dim	strFrom
			dim	strBody
			strSubject	= ""
			strSubject1	= ""
			strSubject2	= ""
			strFrom		= ""
			strBody		= ""
			strBody = strBody & "BOデータ変換メールを受信しました。" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "■添付ファイル" & vbCrlf
			dim	array2
			array2 = bobj.ReadMail(file,"subject:from:date:",">" & dirname)
			if IsArray(array2) then	' OK ?
				dim	data
				for each data in array2
					'1行目を表示
					Call DispMsg(Split(data,vbCrLf)(0))
					dim	strHead
					strHead = Split(data,":")(0)
					select case strHead
					case "Subject"
									strSubject = ""
									if Len(data) > 9 then
										strSubject = Right(data,Len(data) - 9)
									end if
									strSubject1 = "開始:BOデータ変換 " & strSubject
									if instr(trim(lcase(strSubject)),"bodata") > 0 then
										strDb = "newsdc9"
									end if
					case "From"
									strFrom	= Right(data,Len(data) - 6)
									if instr(trim(lcase(strSubject)),"bodata") > 0 then
										if inStr(strFrom,"akagi.seisuke@kk.jp.panasonic.com") > 0 then
											strFrom = "kitamura.ryuichi@kk.jp.panasonic.com"
											strFrom = strFrom & vbTab & "cc"
											strFrom = strFrom & vbTab & "kitamura@kk-sdc.co.jp"
											strFrom = strFrom & vbTab & "system@kk-sdc.co.jp"
										elseif inStr(strFrom,"kubo@kk-sdc.co.jp") > 0 then
											strFrom = strFrom & vbTab & "cc"
											strFrom = strFrom & vbTab & "kitamura.ryuichi@kk.jp.panasonic.com"
											strFrom = strFrom & vbTab & "kitamura@kk-sdc.co.jp"
											strFrom = strFrom & vbTab & "kishimi@kk-sdc.co.jp"
										else
											strFrom = strFrom & vbTab & "cc"
											strFrom = strFrom & vbTab & "kubo@kk-sdc.co.jp"
										end if
									else
										strFrom = strFrom & vbTab & "cc"
										strFrom = strFrom & vbTab & "kubo@kk-sdc.co.jp"
									end if
					case "Body"
					case "File"
						Call SaveFileEx(bobj,data)
						strBody = strBody & "　" & Right(data,Len(data) - 6) & vbCrlf
					end select
				next
			end if
			strBody = strBody & vbCrlf
			strBody = strBody & "■展開ファイル" & vbCrlf
			strBody = strBody & FileList("temp","list")

			strBody = strBody & vbCrlf
			strBody = strBody & "BOデータ変換処理を開始します。" & vbCrlf
			strBody = strBody & "処理終了後に完了通知メールを送信します。" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "db=" & GetOption("db",strDb) & vbCrlf

			Call TanaMake()
			dim	strFile
			strFile = "\\mint\newsdc\tana.csv"
			strFile = strFile & vbTab & "\\mint\newsdc\tananew.csv"

			'返信
			Call DispMsg("SendMail:" & svname & ":" & strFrom & ":" & user & ":" & strSubject)
			dim strMsg
			strMsg = bobj.SendMail(svname,strFrom,"newsdc9@kk-sdc.co.jp", strSubject1,strBody,strFile)
			Call DispMsg(strMsg)
			'変換処理
			strSubject2 = "完了:BOデータ変換 " & strSubject
			strBody = ""
			strBody = strBody & "BOデータ変換処理が完了しました。" & vbCrlf
			strBody = strBody & vbCrlf
			strBody = strBody & "db=" & GetOption("db",strDb) & vbCrlf
			dim f
			for each f in FileList("temp","path")
				Call DispMsg("変換処理:" & f)
				strMsg = Load(strDb,f)
				strBody = strBody & strMsg & vbCrlf
				Call DeleteFile(f)
			next
			strMsg = bobj.SendMail(svname,strFrom,"newsdc9@kk-sdc.co.jp", strSubject2,strBody,"")
			Call DispMsg(strMsg)
		next
	end if
End Function

Function TanaMake()
	Dim objIE

	'IEオブジェクトを作成します
	Set objIE = CreateObject("InternetExplorer.Application")

	'ウィンドウの大きさを変更します
	objIE.Width = 800
	objIE.Height = 600

	'表示位置を変更します
	objIE.Left = 0
	objIE.Top = 0

	'ステータスバーとツールバーを非表示にします
	objIE.Statusbar = False
	objIE.ToolBar = False

	'インターネットエクスプローラ画面を表示します
	objIE.Visible = True

	'①指定したURLを表示します
	objIE.Navigate "http://mint/newsdc/tanamake.php"

	'②ページの読み込みが終わるまでココでグルグル回る
	Do Until objIE.Busy = False
	   '空ループだと無駄にCPUを使うので250ミリ秒のインターバルを置く
	   WScript.sleep(250)
	Loop

	'ステータスバーとツールバーを表示します
	objIE.Statusbar = True
	objIE.ToolBar = True

	WScript.Sleep 5000

	objIE.Quit
	Set objIE = Nothing
End Function

'-----------------------------------------------------------------------
'添付ファイル保存＆LZH展開
'-----------------------------------------------------------------------
Private Function SaveFileEx(bobj,byVal m)
	Call Debug("SaveFileEx(" & m & ")")
	dim	strFilename
	strFilename = Right(m,Len(m) - 6)
	Call Debug("strFilename=" & strFilename)

	'rc = bobj.Execute("cmd.exe /c c:\lha.exe l basp21.lzh",1,stdout)
	dim cmd
	cmd = "cmd.exe /c lha32 e " & strFilename & " temp\"
	Call DispMsg(cmd)
	dim	rc
	dim	stdout
	rc = bobj.Execute(cmd,1,stdout)
	Call DispMsg(stdout)
End Function
'-----------------------------------------------------------------------
'ファイル一覧
'-----------------------------------------------------------------------
Private Function FileList(byval strPath,byval strRcv)
	Dim objFileSys
	Dim strScriptPath
	Dim strTargetPath
	Dim objFolder
	Dim objItem

	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

	strTargetPath = objFileSys.BuildPath(strScriptPath, strPath)

	Set objFolder = objFileSys.GetFolder(strTargetPath)

	Call Debug("ファイル名一覧")
	dim	aryFile()
	dim	i
	i = 0
	dim	strList
	strList = ""
	For Each objItem In objFolder.Files
	    Call Debug(objItem.Name)
	    Call Debug(objItem.Path)
	    Call Debug(objItem.DateLastModified)
		strList = strList & TmFormat(objItem.DateLastModified)
		strList = strList & " " & NumFormat(objItem.Size)
		strList = strList & " " & objItem.Name
		strList = strList & vbCrlf

		Redim Preserve aryFile(i)
		aryFile(i) = objItem.Path
		i = i + 1
	Next

	Call Debug("ファイル数：" & objFolder.Files.Count)

	Set objFolder = Nothing
	Set objFileSys = Nothing
	if strRcv = "list" then
		FileList = strList
	else
		FileList = aryFile
	end if
End Function

Private Function NumFormat(v)
	NumFormat = Right(Space(12) & FormatNumber(v,0,,-1),12)
End Function

Private Function TmFormat(v)
	dim	dt
	dim	tm
	dt = Split(v," ")(0)
	tm = Split(v," ")(1)
	TmFormat = dt & Right("  " & tm,9)
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
