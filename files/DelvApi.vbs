Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "DelvApi.vbs <問合せNo> [option]"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo DelvApi.vbs 4343394041210"
End Sub
'-----------------------------------------------------------------------
'DelvApi
'2016.10.19 新規 WebAPI
'-----------------------------------------------------------------------
Const READYSTATE_COMPLETE	= 4

Class DelvApi
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		set	objIE = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set	objIE = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Disp GetTrackingBody
	End Function
	'-------------------------------------------------------------------
	' 福通の問合せNoから配達状況を取得
	'-------------------------------------------------------------------
	Private	objIE
	Private	strID
	Private	strUrl
	Private	strBody
	Private Function GetTrackingBody()
	    strUrl = "http://thira.plavox.info/transport/api/?t=fukutsu&no=" & strID
		strUrl = "http://www4.kurumeunsou.co.jp/kurume-trans/kamotsu.asp?w_no=" & strID & "&toikbn=2"
	    strBody = ""
		Debug "接続:" & strUrl
	    'IEの起動
		if objIE is nothing then
			Debug "InternetExplorer.Application"
			Set objIE = CreateObject("InternetExplorer.Application")
			objIE.Visible = False
		end if
        objIE.Navigate strUrl
        ' ページが取り込まれるまで待つ
        Do While objIE.Busy or objIE.readyState <> READYSTATE_COMPLETE
			WScript.StdOut.Write "."
            WScript.Sleep 3000
        Loop
        ' テキスト形式で出力
		strBody = objIE.Document.Body.InnerText
'		strBody = objIE.Document.Body.textContent
'		strBody = objIE.Document.Body.InnerHtml
	    GetTrackingBody = strBody
	End Function
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
	'-----------------------------------------------------------------------
	Function GetOption(byval strName _
					  ,byval strDefault _
					  )
		dim	strValue

		if strName = "" then
			strValue = ""
			if strDefault < WScript.Arguments.UnNamed.Count then
				strValue = WScript.Arguments.UnNamed(strDefault)
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		strId = ""
		For Each strArg In WScript.Arguments.UnNamed
			strId = strArg
		Next
		if strId = "" then
			Disp Init
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objDelvApi
	Set objDelvApi = New DelvApi
	if objDelvApi.Init() <> "" then
		call usage()
		exit function
	end if
	call objDelvApi.Run()
End Function
