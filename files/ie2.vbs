use_ie

Sub use_ie()

    ' IE起動
'    Set ie = CreateObject("InternetExplorer.Application")
    Set ie = GetObject("InternetExplorer.Application")
    ie.Navigate "http://www.google.co.jp/"
    ie.Visible = True
    waitIE ie
    
    ' 検索キーワードを入力
'    ie.Document.getElementById("q").Value = "ホゲラッチョ"
    ie.Document.getElementById("lst-ib").Value = "ホゲラッチョ"
    WScript.Sleep 100
    
    ' 検索ボタンクリック
    ie.Document.all("btnG").Click
    waitIE ie
    
    ' 1件目のサイトのタイトルを表示
    MsgBox ie.Document.getElementById("res") _
        .getElementsByTagName("li")(0) _
        .getElementsByTagName("h3")(0) _
        .innerText
    
    ' 制御を破棄
    ie.Quit
    Set ie = Nothing

End Sub


' IEがビジー状態の間待ちます
Sub waitIE(ie)
    
    Do While ie.Busy = True Or ie.readystate <> 4
        WScript.Sleep 100
    Loop
    
    WScript.Sleep 1000
    
End Sub
