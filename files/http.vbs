Option Explicit
'xoxb-205975029010-KP0QW7Odl7rnDAFoz72E72cj
'sdc-w1
'sdc-bot-w1
'https://slack.com/api/chat.postMessage?token=xoxb-205975029010-KP0QW7Odl7rnDAFoz72E72cj&channel=sdc-bot-w1&text=Hello

'HTTPDownload "http://google.co.jp/index.html" , "c:\text.txt"
HTTPDownload "https://slack.com/api/chat.postMessage?token=xoxb-205975029010-KP0QW7Odl7rnDAFoz72E72cj&channel=sdc-bot-w1&text=http.vbs"

Sub HTTPDownload(ByVal STR_URL)
	Dim OBJ_HTTP

	Set OBJ_HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    OBJ_HTTP.Open "GET", STR_URL, False
	OBJ_HTTP.Send
	Wscript.Echo OBJ_HTTP.ResponseText
End Sub
