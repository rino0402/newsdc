Option Explicit
'xoxb-205975029010-KP0QW7Odl7rnDAFoz72E72cj
'sdc-w1
'sdc-bot-w1
'https://slack.com/api/chat.postMessage?token=xoxb-205975029010-KP0QW7Odl7rnDAFoz72E72cj&channel=sdc-bot-w1&text=Hello

'HTTPDownload "http://google.co.jp/index.html" , "c:\text.txt"
Call Main()

Sub Main()
	dim	strArg
	dim	strText
	strText = ""
	dim	strFile
	strFile = ""
	dim	strToken
	strToken = "xoxb-205975029010-KP0QW7Odl7rnDAFoz72E72cj"
	dim	strUser
'	strUser	= "sdc-w1"
	strUser	= ""
	dim	strChannel
'	strChannel	= "sdc-bot-w1"
'	strChannel	= "w3"
	strChannel	= ""
	For Each strArg In WScript.Arguments.UnNamed
		if strUser = "" then
			strUser = strArg
		elseif strChannel = "" then
			strChannel = strArg
		elseif strText = "" then
			strText = strArg
		elseif strFile = "" then
			strFile = strArg
		end if
	Next
	dim	strUrl
	strUrl = "https://slack.com/api/chat.postMessage?token=" & strToken & "&username=" & strUser & "&channel=" & strChannel
'	strUrl = strUrl & "&text=```" & strText & vbLf & Att(strFile) & vbLf & "```"
'	strUrl = strUrl & "&text=" & strText
'org	strUrl = strUrl & "&text=```" & strText & "%0D%0A" & Att(strFile) & "```"
	dim	strAtt
	strAtt = Att(strFile)
	if strAtt <> "" then
		strAtt = "%0D%0A```" & strAtt & "```"
	end if
	strUrl = strUrl & "&text=" & strText & strAtt
'	strUrl = strUrl & "&text=```" & Att(strFile) & "```"
	'strUrl = strUrl & "&attachments=[{""text"":""- 111\n- 222""}]"
	'strUrl = strUrl & "&attachments=[{""text"":""111\n222""}]"
'	strUrl = strUrl & "&attachments=[{""pretext"":""" & Att(strFile) & """}]"
'	strUrl = strUrl & "&attachments=[{""text"":""```" & Att(strFile) & "```""}]"
'	strUrl = strUrl & "&attachments=[{""text"":""" & Att(strFile) & """}]"
'	strUrl = strUrl & "&attachments=[{""pretext"":""PreText"",""text"":""" & Att(strFile) & """}]"
'	strUrl = strUrl & "&attachments=[{""text"":""" & Att(strFile) & """}]"
'	strUrl = strUrl & "&attachments=[{""text"":""```" & Att(strFile) & "```""}]"
	HTTPDownload strUrl
End Sub

Const ForReading = 1

Function Att(byVal strFile)
	Att = strFile
	if strFile = "" then
		exit function
	end if
	dim	objFs
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Dim objFile
	on error resume next
    Set objFile = objFs.OpenTextFile(strFile, ForReading, False)
	if Err.Number <> 0 then
		Att = strFile & ":オープンエラー:" & Err.Number
		Wscript.Echo Att
		exit function
	end if
	on error goto 0
	dim	strText
	If objFile.AtEndOfStream Then
	    strText = ""
	else
	    strText = objFile.ReadAll
	end if
    objFile.Close
	Set objFile = nothing
	Set objFs = nothing
'	strText = Replace(strText,vbCrLf,"\n")
'	strText = Replace(strText,vbCrLf,"%0D%0A")
'	strText = Replace(strText,0," ")
	strText = Replace(strText,vbCrLf,"%0A")
'	strText = Replace(strText," ","%20")
	Att = strText
End Function

Sub HTTPDownload(ByVal STR_URL)
	Dim OBJ_HTTP

	Set OBJ_HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

	Wscript.Echo STR_URL
    OBJ_HTTP.Open "GET", STR_URL, False
	OBJ_HTTP.Send
	Wscript.Echo OBJ_HTTP.ResponseText
End Sub
