Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Call Main()
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objExcel
	Set objExcel = New Excel
	objExcel.Run
	Set objExcel = Nothing
End Function
'-----------------------------------------------------------------------
'Excel
'2017.03.07 �V�K
'-----------------------------------------------------------------------
Const xlUp = -4162

Class Excel
	'-----------------------------------------------------------------------
	'�g�p���@
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "AcZaiko.vbs [option] <*.xlsx>"
		Echo "/db:newsdc4"
		Echo "/debug"
		Echo "Ex."
		Echo "cscript//nologo AcZaiko.vbs �Z���^�[�݌�170515.xls /db:newsdc4"
		Echo "cscript//nologo AcZaiko.vbs �T�e�݌�.xlsx /db:newsdc1"
	End Sub
	'-----------------------------------------------------------------------
	'Private �ϐ�
	'-----------------------------------------------------------------------
	Private	strDBName
	Private	objDB
	Private	strFileName
	Private	strBookName
	Private	objExcel
	Private	objBook
	Private	objSheet
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName		= GetOption("db","newsdc")
		set objDB		= nothing
		set	objExcel	= nothing
		set	objBook		= nothing
		set	objSheet	= nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB		= nothing
		set	objSheet	= nothing
		set	objBook		= nothing
		set	objExcel	= nothing
    End Sub
	'-----------------------------------------------------------------------
	'Quit() �����I��
	'-----------------------------------------------------------------------
	Private Function Quit()
		Debug ".Quit()"
		Wscript.Quit
	End Function
	'-----------------------------------------------------------------------
	'Echo()
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		Wscript.Echo s
	End Sub
	'-----------------------------------------------------------------------
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
	Private Function Init()
		Debug ".Init()"
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Echo "Error:�I�v�V����:" & strArg
				Disp Init
				Usage
				Quit
			end if
		Next
		if strFileName = "" then
			Echo "Error:�t�@�C�����w��."
			Usage
			Quit
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Echo "Error:�I�v�V����:" & strArg
				Usage
				Quit
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Init
		OpenDb
		Load
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Load() �Ǎ�
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load():" & strFileName
		CreateExcel
		OpenBook strFileName
		LoadBook
		CloseBook
	End Function
	'-------------------------------------------------------------------
	'LoadBook()
	'-------------------------------------------------------------------
	Private	lngTopRow
	Private	intCol
	Private	strMaxCol
    Private Function LoadBook()
		Debug ".LoadBook()"
		for each objSheet in objBook.Worksheets
			Write objSheet.Name & ":"
			select case SheetType()
			case "�Z���^�[�݌�","�T�e���C�g�݌�","���|�[�g 1"
				WriteLine "Load"
				LoadSheet
			case else
				WriteLine "skip"
			end select
		next
    End Function
	'-------------------------------------------------------------------
	'SheetType()
	'-------------------------------------------------------------------
	Private	Function SheetType()
		SheetType = ""
		if Trim(objSheet.Name) = "�Z���^�[�݌�" then
			SheetType = Trim(objSheet.Name)
			exit function
		end if
		if Trim(objSheet.Name) = "�T�e���C�g�݌�" then
			SheetType = Trim(objSheet.Name)
			exit function
		end if
		if Trim(objSheet.Name) = "���|�[�g 1" then
			SheetType = Trim(objSheet.Name)
			exit function
		end if
	End Function
	'-------------------------------------------------------------------
	'LoadSheet()
	'-------------------------------------------------------------------
	Private	lngMaxRow
	Private	lngRow
    Private Function LoadSheet()
		Debug ".LoadSheet():" & objBook.Name & ":" & objSheet.Name
		lngMaxRow = objSheet.Range("A65535").End(xlUp).Row
		Debug ".LoadSheet():MaxRow=" & lngMaxRow
		for lngRow = 2 to lngMaxRow
			Write objSheet.Name & ":" & lngRow & "/" & lngMaxRow
			if Trim(objSheet.Name) = "���|�[�g 1" then
				LoadReport
			else
				LoadLine
			end if
			WriteLine ""
		next
    End Function
	'-------------------------------------------------------------------
	'LoadReport()
	'-------------------------------------------------------------------
    Private Function LoadReport()
		Debug ".LoadReport()"
		dim	strSyushi
		strSyushi = GetValue(objSheet.Range("A" & lngRow))
		dim	strPn
		strPn = GetValue(objSheet.Range("B" & lngRow))
		dim	strQty
		strQty = GetValue(objSheet.Range("C" & lngRow))
		Write " " & strSyushi & " " & strPn & " " & strQty
		Insert strPn,strSyushi,strQty
    End Function
	'-------------------------------------------------------------------
	'LoadLine()
	'-------------------------------------------------------------------
    Private Function LoadLine()
		Debug ".LoadLine()"
		dim	strPn
		strPn = GetValue(objSheet.Range("B" & lngRow))
		Write " " & strPn
		dim	objTop
		set objTop = objSheet.Range("D1")
		dim	objQty
		set objQty = objSheet.Range("D" & lngRow)
		do while True
			dim	strSyushi
			strSyushi = GetValue(objTop)
			if strSyushi = "���v" then
				exit do
			end if
			if strSyushi = "" then
				exit do
			end if
			dim	strQty
			strQty = GetValue(objQty)
			Insert strPn,strSyushi,strQty
			set objTop = objTop.Offset(0,1)
			set objQty = objQty.Offset(0,1)
		loop
    End Function
	'-------------------------------------------------------------------
	'Insert()
	'-------------------------------------------------------------------
    Private Function Insert(byVal strPn,byVal strSyushi,byVal strQty)
		Debug ".Insert()"
		if isNumeric(strQty) = False then
			exit function
		end if
		if CLng(strQty) = 0 then
			exit function
		end if
		Write " " & strSyushi & ":" & strQty
		AddSql ""
		AddSql "insert into AcZaiko"
		AddSql "(Pn"
		AddSql ",Syushi"
		AddSql ",Qty"
		AddSql ") values ("
		AddSql " '" & strPn & "'"
		AddSql ",'" & strSyushi & "'"
		AddSql "," & strQty
		AddSql ")"
		CallSql strSql
	End Function
	'-------------------------------------------------------------------
	'GetValue()
	'-------------------------------------------------------------------
	Private	Function GetValue(objR)
		dim	strValue
		on error resume next
		strValue = Trim(objR)
		on error goto 0
		if Err.Number <> 0 then
'			Wscript.StdOut.WriteLine ".GetValue():0x" & Hex(Err.Number) & ":" & Err.Description
'			Wscript.StdOut.WriteLine
'			Wscript.StdOut.WriteLine objR.Address & ":(" & objR.Text & ")"
'			Wscript.Quit
			strValue = Trim(objR.Text)
		end if
		if strValue <> "" then
			if Asc(strValue) = 0 then
				strValue = ""
			end if
		end if
		strValue = Replace(strValue,vbCr,"")
		strValue = Replace(strValue,vbLf,"")
'		GetValue = Replace(GetValue,vbCrLf,"")
'		GetValue = Replace(GetValue,0,"")
		Debug "GetValue():" & objR.Address & ":" & strValue & ":"
		GetValue = strValue
	End Function
	'-------------------------------------------------------------------
	'Sql���s
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		select case Err.Number
		case -2147467259	'�d��
		case 0,500
		case else
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine Err.Number & "(0x" & Hex(Err.Number) & "):" & Err.Description
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end select
		on error goto 0
'		on error resume next
'		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'������ǉ� strSql
	'-------------------------------------------------------------------
	dim	strSql
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
	End Function
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	Private Function CreateExcel()
		Debug ".CreateExcel()"
		if objExcel is nothing then
			Debug ".CreateExcel():CreateObject(Excel.Application)"
			Set objExcel = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'AbsPath() ��΃p�X
	'-------------------------------------------------------------------
	Private	Function AbsPath(byVal strPath)
		dim	objFso
		Set objFso = CreateObject("Scripting.FileSystemObject")
		AbsPath = objFso.GetAbsolutePathName(strPath)
		Set objFso = Nothing
	End Function
	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	Private Function OpenBook(byVal strBkName)
		Debug ".OpenBook()"
		if objBook is nothing then
			strBkName = AbsPath(strBkName)
			Write strBkName & " :"
			on error resume next
'			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True)
			WriteLine Err.Number
			if Err.Number <> 0 then
				WriteLine
				WriteLine "0x" & Hex(Err.Number) & ":" & Err.Description
				Quit
			end if
			on error goto 0
		end if
	end function
	'-------------------------------------------------------------------
	'Excel �t�@�C���N���[�Y
	'-------------------------------------------------------------------
	Private Function CloseBook()
		Debug ".CloseBook()"
		if not objBook is nothing then
			Debug ".CloseBook().Close:" & objBook.Name
			Call objBook.Close(False)
			set objBook = nothing
		end if
	end function
	'-----------------------------------------------------------------------
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'���b�Z�[�W�\��
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-------------------------------------------------------------------
	'Write
	'-------------------------------------------------------------------
	Private	Sub Write(byVal s)
		Wscript.StdOut.Write s
	End Sub
	'-------------------------------------------------------------------
	'WriteLine
	'-------------------------------------------------------------------
	Private	Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine s
	End Sub
	'-----------------------------------------------------------------------
	'�I�v�V�����擾
	'-----------------------------------------------------------------------
	Function GetOption(byval strName _
					  ,byval strDefault _
					  )
		dim	strValue

		if strName = "" then
			strValue = ""
			if strDefault < WScript.Arguments.UnNamed.Count then
				strValue = WScript.Arguments.UnNamed(strDefault)
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
End Class
