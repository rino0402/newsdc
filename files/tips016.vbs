'[tips0161.vbs]
Option Explicit
On Error Resume Next

Dim strUrl      ' �\������y�[�W
Dim objIE       ' IE �I�u�W�F�N�g

strUrl = "http://www.whitire.com/"
strUrl = "http://thira.plavox.info/transport/api/?t=fukutsu&no=41103160350"

Set objIE = WScript.CreateObject("InternetExplorer.Application")
If Err.Number = 0 Then
    objIE.Navigate strUrl

    ' �y�[�W����荞�܂��܂ő҂�
    Do While objIE.busy
        WScript.Sleep(100)
    Loop
    Do While objIE.Document.readyState <> "complete"
        WScript.Sleep(100)
    Loop

    ' �e�L�X�g�`���ŏo��
    WScript.Echo objIE.Document.Body.InnerText

    ' �g�s�l�k�`���ŏo��
'    WScript.Echo objIE.Document.Body.InnerHtml
Else
    WScript.Echo "�G���[�F" & Err.Description
End If
Set objIE = Nothing
