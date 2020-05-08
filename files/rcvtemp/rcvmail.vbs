' --------------------------------------------------------
' メールを受信するサンプル(VBS)
' Basp21.dllとBsmtp.dllをC:\Windowsにコピーしています
' [Regsvr32.exe Basp21.dll]を実行しています

' メール送受信APIの宣言
Set BASP21 = CreateObject("Basp21")
    
' メール受信およびメールボックスから削除
outary = BASP21.RcvMail( _
    "ns", "newsdc9", "123daa@Z", _
    "SAVD 1-1", ".")

' 受信メールチェック
If IsArray(outary) Then
    outary2 = BASP21.ReadMail( _
        outary(0), "subject:from:date:", ".")
    Wscript.Echo "メール有り:" & outary2(1)
Else
    Wscript.Echo "メール無し"
End If
