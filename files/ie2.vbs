use_ie

Sub use_ie()

    ' IE�N��
'    Set ie = CreateObject("InternetExplorer.Application")
    Set ie = GetObject("InternetExplorer.Application")
    ie.Navigate "http://www.google.co.jp/"
    ie.Visible = True
    waitIE ie
    
    ' �����L�[���[�h�����
'    ie.Document.getElementById("q").Value = "�z�Q���b�`��"
    ie.Document.getElementById("lst-ib").Value = "�z�Q���b�`��"
    WScript.Sleep 100
    
    ' �����{�^���N���b�N
    ie.Document.all("btnG").Click
    waitIE ie
    
    ' 1���ڂ̃T�C�g�̃^�C�g����\��
    MsgBox ie.Document.getElementById("res") _
        .getElementsByTagName("li")(0) _
        .getElementsByTagName("h3")(0) _
        .innerText
    
    ' �����j��
    ie.Quit
    Set ie = Nothing

End Sub


' IE���r�W�[��Ԃ̊ԑ҂��܂�
Sub waitIE(ie)
    
    Do While ie.Busy = True Or ie.readystate <> 4
        WScript.Sleep 100
    Loop
    
    WScript.Sleep 1000
    
End Sub
