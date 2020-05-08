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
		Wscript.Echo "AcSyuka.vbs [option] <*.xlsx>"
		Wscript.Echo "/db:newsdc4"
		Wscript.Echo "/debug"
		Wscript.Echo "Ex."
		Wscript.Echo "cscript AcSyuka.vbs �y�S���x�z�o�׎���.xlsx"
		Wscript.Echo "cscript AcSyuka.vbs 2017.12.AC.���ߋ��o�׎���2017-12.xlsx"
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
		strDBName		= GetOption("db"	,"newsdc4")
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
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "?"
				Usage
				Quit
			case else
				Echo "Error:�I�v�V����:" & strArg
				Usage
				Quit
			end select
		Next
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
			case "�ߋ��o�׎���"
				WriteLine "Load"
				lngRowTop = 3
				LoadSheet
			case "�C�ߋ��o�׎��сi���ʁj","�ߋ��o�׎��сi���ʁj"
				WriteLine "Load"
				lngRowTop = 2
				strYmF = ""
				strYmT = ""
				LoadSheet
			case "RF"
				WriteLine "�①�ɔ̔�����"
				lngRowTop = 2
				strYmF = ""
				strYmT = ""
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
		if Trim(objSheet.Name) = "�ߋ��o�׎���" then
			SheetType = "�ߋ��o�׎���"
			exit function
		end if
		if Trim(objSheet.Name) = "�C�ߋ��o�׎��сi���ʁj" then
			SheetType = "�C�ߋ��o�׎��сi���ʁj"
			exit function
		end if
		if Trim(objSheet.Name) = "�ߋ��o�׎��сi���ʁj" then
			SheetType = "�C�ߋ��o�׎��сi���ʁj"
			exit function
		end if
		if Trim(objSheet.Name) = "���|�[�g 1" then
			SheetType = "RF"
			exit function
		end if
	End Function
	'-------------------------------------------------------------------
	'LoadSheet()
	'-------------------------------------------------------------------
	Private	lngMaxRow
	Private	lngRowTop
	Private	lngRow
    Private Function LoadSheet()
		Debug ".LoadSheet():" & objBook.Name & ":" & objSheet.Name
		lngMaxRow = objSheet.Range("A65535").End(xlUp).Row
		Debug ".LoadSheet():MaxRow=" & lngMaxRow
		for lngRow = lngRowTop to lngMaxRow
			Write objSheet.Name & ": " & lngRow & "/" & lngMaxRow
			LoadLine
			WriteLine ""
		next
    End Function
	'-------------------------------------------------------------------
	'LoadLine()
	'-------------------------------------------------------------------
    Private Function LoadLine()
		Debug ".LoadLine():" & SheetType
		select case SheetType()
		case "�ߋ��o�׎���"
			LoadLine1
		case "�C�ߋ��o�׎��сi���ʁj"
			LoadLine0
		case "RF"
			LoadLineRf
		case else
		end select
	End Function
	'-------------------------------------------------------------------
	'LoadLine0() �C�ߋ��o�׎��сi���ʁj
	'-------------------------------------------------------------------
	Private	strYmF	'from
	Private	strYmT	'to
    Private Function LoadLine0()
		Debug ".LoadLine0()"
		dim	strPn
		dim	strYmRange
		if objSheet.Range("A1") = "�i�ڔԍ�" then
			strPn = GetValue(objSheet.Range("A" & lngRow))
			strYmRange = "K1:ZZ1"
		else
			strPn = GetValue(objSheet.Range("D" & lngRow))
			strYmRange = "P1:ZZ1"
		end if
		Write " " & strPn
		dim	objYm
		if strYmF = "" then
			for each objYm in objSheet.Range(strYmRange)
				if strYmF = "" then
					strYmF = objYm
				end if
				if objYm = "" then
					exit for
				end if
				if objYm = "�����v" then
					exit for
				end if
				strYmT = objYm
			next
			Debug "�͈�:" & strYmF & "�`" & strYmT
		end if
		Delete0 strPn,strYmF,strYmT
		set	objYm = objSheet.Range(left(strYmRange,2))
		dim	objQty
		set	objQty = objSheet.Range(left(strYmRange,1) & lngRow)
		do while True
			dim	strYm
			dim	strQty
			strYm = GetValue(objYm)
			if strYm = "" then
				exit do
			end if
			strQty = GetValue(objQty)
			Insert strPn,strYm,strQty
			set objYm = objYm.offset(0,1)
			set objQty = objQty.offset(0,1)
		loop
    End Function
	'-------------------------------------------------------------------
	'Insert0()
	'-------------------------------------------------------------------
    Private Function Delete0(byVal strPn,byVal strYmF,byVal strYmT)
		Debug ".Delete0():" & strPn & "," & strYmF & "," & strYmT
		AddSql ""
		AddSql "delete from AcSyuka"
		AddSql "where Pn = '" & strPn & "'"
		AddSql "and Ym between '" & strYmF & "' and '" & strYmT & "'"
		Write " del(" & strYmF & "�`" & strYmT & ")"
		CallSql strSql
	End Function
	'-------------------------------------------------------------------
	'LoadLineRf() ���|�[�g 1
	'-------------------------------------------------------------------
    Private Function LoadLineRf()
		Debug ".LoadLineRf()"
		dim	strPn
		strPn = GetValue(objSheet.Range("B" & lngRow))
		Write " " & strPn
		dim	objYm
		dim	objQty
		set	objYm = objSheet.Range("G1")
		set	objQty = objSheet.Range("G" & lngRow)
		dim	strYm
		strYm = GetValue(objYm)
		dim	strQty
		strQty = GetValue(objQty)
		Insert strPn,strYm,strQty
    End Function
	'-------------------------------------------------------------------
	'LoadLine1() �ߋ��o�׎���
	'-------------------------------------------------------------------
    Private Function LoadLine1()
		Debug ".LoadLine1()"
		dim	strPn
		strPn = GetValue(objSheet.Range("B" & lngRow))
		Write " " & strPn
		dim	objYm
		dim	objQty
		set	objYm = objSheet.Range("K2")
		set	objQty = objSheet.Range("K" & lngRow)
		do while True
			dim	strYm
			dim	strQty
			strYm = GetValue(objYm)
			if strYm = "" then
				exit do
			end if
			strQty = GetValue(objQty)
			Insert strPn,strYm,strQty
			set objYm = objYm.offset(0,1)
			set objQty = objQty.offset(0,1)
		loop
    End Function
	'-------------------------------------------------------------------
	'Insert()
	'-------------------------------------------------------------------
    Private Function Insert(byVal strPn,byVal strYm,byVal strQty)
		Debug ".Insert()"
		if isNumeric(strYm) = False then
			exit function
		end if
		Write " " & strYm
		if isNumeric(strQty) = False then
			exit function
		end if
		if CLng(strQty) = 0 then
			exit function
		end if
		Write " " & strQty
		AddSql ""
		AddSql "insert into AcSyuka"
		AddSql "(Pn"
		AddSql ",Ym"
		AddSql ",Qty"
		AddSql ") values ("
		AddSql " '" & strPn & "'"
		AddSql ",'" & strYm & "'"
		AddSql "," & strQty
		AddSql ")"
		CallSql strSql
		Write "."
	End Function
	'-------------------------------------------------------------------
	'DeleteSheet
	'-------------------------------------------------------------------
    Private Function DeleteSheet()
		Debug ".DeleteSheet()"
		Wscript.StdOut.Write objBook.Name & " " & objSheet.Name & ":�폜��..."

		AddSql ""
		AddSql "delete from CsvTemp"
		AddSql " where Filename = '" & objBook.Name & "'"
		AddSql " and Sheetname = '" & objSheet.Name & "'"
		Wscript.StdOut.Write ":" & strSql
		CallSql strSql
		Wscript.StdOut.WriteLine
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
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
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
