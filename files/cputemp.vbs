Option Explicit

'WMI�ɂĎg�p����e��I�u�W�F�N�g���`�E��������B
Dim oClassSet
Dim oClass
Dim oLocator
Dim oService
Dim sMesStr

'���[�J���R���s���[�^�ɐڑ�����B
Set oLocator = Wscript.CreateObject("WbemScripting.SWbemLocator")
Set oService = oLocator.ConnectServer(, "Root\WMI")
'�N�G���[������ WQL �ɂĎw�肷��B
Set oClassSet = oService.ExecQuery("Select * From MSAcpi_ThermalZoneTemperature")

'�R���N�V��������͂���B
For Each oClass In oClassSet

sMesStr = "�C���X�^���X��: " & oClass.InstanceName & vbCrLf & _
"���x: " & CStr((oClass.CurrentTemperature - 2732) / 10 ) & vbCrLf & vbCrLf

Next

MsgBox("CPU ���x�Ɋւ�����ł��B" & vbCrLf & vbCrLf & sMesStr)

'�g�p�����e��I�u�W�F�N�g����Еt������B
Set oClassSet = Nothing
Set oClass = Nothing
Set oService = Nothing
Set oLocator = Nothing
