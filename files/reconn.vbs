' ----------------------------------------------
' �l�b�g���[�N�C���^�[�t�F�C�X�̍ċN���X�N���v�g
' ----------------------------------------------
' Ver 1.0
' 2011/05/20
'

Option Explicit

Dim objArgs
Dim strTarget
Dim strNIC

strTarget = "192.168.4.31" '�����pIP�A�h���X���w��
strNIC = ""

If PingResult(strTarget) = True Then
    'Wscript.Echo "Ping��������"
Else
    'Wscript.Echo "Ping�����Ȃ�"
	Call NICReboot(strNIC)
End If

'Wscript.Echo "�X�N���v�g�I��"

Set objArgs = Nothing

'strTarget��ping���s���Đ���������PingResult��True��Ԃ�
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
 'Wscript.Echo "NIC�ċN�����s"
Const ssfCONTROLS = 3 
Const sConPaneName    = "�l�b�g���[�N�ڑ�"
Const sConnectionName = "���[�J�� �G���A�ڑ�" '�Ώ�NIC���w��
Const sDisableVerb    = "�����ɂ���(&B)" 
Const sEnableVerb     = "�L���ɂ���(&A)" 

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