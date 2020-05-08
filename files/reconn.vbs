' ----------------------------------------------
' ネットワークインターフェイスの再起動スクリプト
' ----------------------------------------------
' Ver 1.0
' 2011/05/20
'

Option Explicit

Dim objArgs
Dim strTarget
Dim strNIC

strTarget = "192.168.4.31" '検査用IPアドレスを指定
strNIC = ""

If PingResult(strTarget) = True Then
    'Wscript.Echo "Ping応答あり"
Else
    'Wscript.Echo "Ping応答なし"
	Call NICReboot(strNIC)
End If

'Wscript.Echo "スクリプト終了"

Set objArgs = Nothing

'strTargetにpingを行って成功したらPingResultにTrueを返す
Function PingResult(strTarget) 

    Dim objWMIService
    Dim colItems
    Dim objItem
 
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery _
        ("Select * from Win32_PingStatus " & _
            "Where Address = '" & strTarget & "'")
    For Each objItem in colItems
        If objItem.StatusCode = 0 Then
            PingResult = True
        Else
            PingResult = False
        End If
    Next
 
    Set objWMIService = Nothing
    Set colItems = Nothing
End Function


Function NICReboot(strNIC)
 'Wscript.Echo "NIC再起動実行"
Const ssfCONTROLS = 3 
Const sConPaneName    = "ネットワーク接続"
Const sConnectionName = "ローカル エリア接続" '対象NICを指定
Const sDisableVerb    = "無効にする(&B)" 
Const sEnableVerb     = "有効にする(&A)" 

Dim shellApp
Dim oControlPanel
Dim oNetConnections
Dim folderitem
Dim oLanConnection
Dim verb

set shellApp = createobject("shell.application") 
set oControlPanel = shellApp.Namespace(ssfCONTROLS) 
set oNetConnections = nothing 
for each folderitem in oControlPanel.items 
  if folderitem.name = sConPaneName then 
    set oNetConnections = folderitem.getfolder: exit for 
  end if 
next 
if oNetConnections is nothing then 
  wscript.quit 
end if 
set oLanConnection = nothing 
for each folderitem in oNetConnections.items 
  if lcase(folderitem.name) = lcase(sConnectionName) then 
    set oLanConnection = folderitem: exit for 
  end if 
next 
if oLanConnection is nothing then 
  wscript.quit 
end if 

for each verb in oLanConnection.verbs 
  if verb.name = sDisableVerb then 
    verb.Doit
    WScript.Sleep 2000
  end if 
next 
for each verb in oLanConnection.verbs 
  if verb.name = sEnableVerb then 
    verb.Doit
    WScript.Sleep 2000
  end if 
next 


End Function

'EOF