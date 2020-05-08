Option Explicit
Call Main()
WScript.Quit 0

Private Sub usage()
	Wscript.Echo "Excel読込テスト"
	Wscript.Echo "exceltest.vbs [option] <filename>"
	Wscript.Echo " -?"
End Sub

Private Sub Main()
	dim	i
	dim	strArg
	dim	strFilename

	strFilename = ""
	For Each strArg In WScript.Arguments
	    	select case strArg
		case "-?"
			call usage()
			exit sub
		case else
			if strFilename = "" then
				strFilename = strArg
			else
				usage()
				exit sub
			end if
		end select
	Next
	if strFilename = "" then
		usage()
		exit sub
	end if
	call LoadExcel(strFilename)
End Sub

Private Sub LoadExcel(byval strFilename)
	dim	objXL
	dim	objBk
	dim	objSt
	dim	lngRow

	Wscript.Echo "LoadExcel(" & strFilename & ")"
	Set objXL = WScript.CreateObject("Excel.Application")
	objXL.Application.Visible = True
	Set objBk = objXL.Workbooks.Open(strFilename,,True)
	set objSt = objBk.Worksheets("単票")
	lngRow = 3
	do while objSt.Range("B" & lngRow) <> ""
		objSt.Range("B" & lngRow).Select
		lngRow = lngRow + 1
	loop
End Sub
