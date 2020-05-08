Option Explicit

'WMIにて使用する各種オブジェクトを定義・生成する。
Dim oClassSet
Dim oClass
Dim oLocator
Dim oService
Dim sMesStr

'ローカルコンピュータに接続する。
Set oLocator = Wscript.CreateObject("WbemScripting.SWbemLocator")
Set oService = oLocator.ConnectServer(, "Root\WMI")
'クエリー条件を WQL にて指定する。
Set oClassSet = oService.ExecQuery("Select * From MSAcpi_ThermalZoneTemperature")

'コレクションを解析する。
For Each oClass In oClassSet

sMesStr = "インスタンス名: " & oClass.InstanceName & vbCrLf & _
"温度: " & CStr((oClass.CurrentTemperature - 2732) / 10 ) & vbCrLf & vbCrLf

Next

MsgBox("CPU 温度に関する情報です。" & vbCrLf & vbCrLf & sMesStr)

'使用した各種オブジェクトを後片付けする。
Set oClassSet = Nothing
Set oClass = Nothing
Set oService = Nothing
Set oLocator = Nothing
