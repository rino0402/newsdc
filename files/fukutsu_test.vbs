Option Explicit
'-----------------------------------------------------------------------
'���C���ďo���C���N���[�h
'-----------------------------------------------------------------------
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
	dim	strScriptPath
	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strFileName = strScriptPath & strFileName
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
	Set fso = Nothing
End Function
Call Include("const.vbs")
Call Include("debug.vbs")
Call Include("get_b.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "�o�׏��i����p�����f�[�^"
	Wscript.Echo "fukutsu.vbs [option] No"
	Wscript.Echo "Ex."
	Wscript.Echo "fukutsu.vbs 41103160350"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	For Each strArg In WScript.Arguments.Named
    	select case lcase(strArg)
		case "db"
		case "list"
		case "load"
		case "top"
		case "debug"
		case else
			usage()
			Main = 1
			exit Function
		end select
	Next
	For Each strArg In WScript.Arguments.UnNamed
		Wscript.Echo strArg & ":" & GetFukutsu(strArg)
	Next
'	select case GetFunction()
'	case "list"
'		Call List()
'	case "load"
'		Call Load(strFilename)
'	case "usage"
'		Call usage()
'	end select
	Main = 0
End Function

Function GetFukutsu(byVal strID)
	dim	strFukutsu
	strFukutsu = ""
	strFukutsu = GetFukutsuBody(strID)
	GetFukutsu = strFukutsu
End Function

Function GetFukutsuBody(byVal strID)
On Error Resume Next

	Dim strUrl      ' �\������y�[�W
	Dim objIE       ' IE �I�u�W�F�N�g
	dim	strBody

	strBody = ""
'	strUrl = "http://www.whitire.com/"
	strUrl = "http://thira.plavox.info/transport/api/?t=fukutsu&no=" & strID
	strUrl = "http://thira.plavox.info/transport/api/?t=fukutsu&no=" & strID & "&nojump=1"

	Set objIE = WScript.CreateObject("InternetExplorer.Application")
	If Err.Number = 0 Then
	    objIE.Navigate strUrl

	    ' �y�[�W����荞�܂��܂ő҂�
	    Do While objIE.busy
	        WScript.Sleep(100)
	    Loop
	    Do While objIE.Document.readyState <> "complete"
	        WScript.Sleep(100)
	    Loop

	    ' �e�L�X�g�`���ŏo��
	    strBody = objIE.Document.Body.InnerText

	    ' �g�s�l�k�`���ŏo��
'	    strBody = objIE.Document.Body.InnerHtml
	Else
	    strBody "�G���[�F" & Err.Description
	End If
	objIE.Close
	Set objIE = Nothing
	GetFukutsuBody = strBody
End Function

