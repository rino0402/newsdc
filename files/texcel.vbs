'On Error Resume Next
dim	objXL
Set objXL = WScript.CreateObject("Excel.Application.15")
WScript.Echo Err
WScript.Echo objXL.Name
set objXL = Nothing

Set objXL = WScript.CreateObject("Excel.Application")
WScript.Echo Err
set objXL = Nothing
