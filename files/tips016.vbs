'[tips0161.vbs]
Option Explicit
On Error Resume Next

Dim strUrl      ' 表示するページ
Dim objIE       ' IE オブジェクト

strUrl = "http://www.whitire.com/"
strUrl = "http://thira.plavox.info/transport/api/?t=fukutsu&no=41103160350"

Set objIE = WScript.CreateObject("InternetExplorer.Application")
If Err.Number = 0 Then
    objIE.Navigate strUrl

    ' ページが取り込まれるまで待つ
    Do While objIE.busy
        WScript.Sleep(100)
    Loop
    Do While objIE.Document.readyState <> "complete"
        WScript.Sleep(100)
    Loop

    ' テキスト形式で出力
    WScript.Echo objIE.Document.Body.InnerText

    ' ＨＴＭＬ形式で出力
'    WScript.Echo objIE.Document.Body.InnerHtml
Else
    WScript.Echo "エラー：" & Err.Description
End If
Set objIE = Nothing
