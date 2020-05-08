WScript.Echo "AAA"
dim	strRetVal
dim	strHTML
strHTML = "http://thira.plavox.info/transport/api/?t=fukutsu&no=41103160350"
strHTML = "http://yahoo.co.jp/"
WScript.Echo GetHtmlSource(strHTML,strRetVal,0,"","")
WScript.Echo strRetVal

'####################################################################################
'#
'# 関数名：GetHtmlSource
'#-----------------------------------------------------------------------------------
'# 機能  ：指定のURLからHTMLソースを取得する
'# 引数  ：strURL       I URL
'#         strRetVal    O 取得した文字列
'#         isSJIS       I ソースが Shift-JIS の場合 True
'#         strID        I ドメイン認証が必要な場合のユーザーID
'#         strPass      I ドメイン認証が必要な場合のパスワード
'# 戻り値：True 正常、False 失敗
'#
'####################################################################################
Private Function GetHtmlSource(ByVal strURL, _
                               ByRef strRetVal, _
                      		   ByVal isSJIS, _
                      		   ByVal strID, _
                               ByVal strPass)

    Dim oHttp

    'オブジェクト変数に参照をセットします
On Error Resume Next
    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    If (Err.Number <> 0) Then
        Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
		WScript.Echo "MSXML.XMLHTTPRequest"
    End If
On Error GoTo 0
    If oHttp Is Nothing Then
        MsgBox "XMLHTTP オブジェクトを作成できませんでした。", vbCritical
        Exit Function
    End If

    'ドメイン認証が必要な場合
    If strID <> "" Then
        oHttp.Open "GET", strURL, False, strID, strPass
    Else
        oHttp.Open "GET", strURL, False
    End If
    Call oHttp.Send(Null)

	do
		WScript.Echo "oHttp.readyState=" & oHttp.readyState
'		WScript.Echo "oHttp.Status=" & oHttp.Status
'		WScript.Echo "oHttp.statusText=" & oHttp.statusText

'		if oHttp.readyState = 1 then
			exit do
'		end if
	loop


    '失敗した場合は関数を終了します。
    If (oHttp.Status < 200 Or oHttp.Status >= 300) Then Exit Function

    'ソースを格納します
    If isSJIS Then
        'ソースが Shift-JIS の場合
        strRetVal = StrConv(oHttp.responseBody, vbUnicode)
    Else
        'ソースが Unicode の場合
        strRetVal = oHttp.responseText
    End If

    'オブジェクト変数の参照を解放します
    Set oHttp = Nothing

    '戻り値をセットします
    GetHtmlSource = True

End Function

