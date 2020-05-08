Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objExcel
	Set objExcel = New Excel
	Set objExcel = Nothing
End Function
'-----------------------------------------------------------------------
'Excel
'-----------------------------------------------------------------------
Const xlUp = -4162
Const xlLastCell = 11
Class Excel
	Private	strDBName
	Private	objDB
	Private	strFileName
	Private	objExcel
	Private	objBook
	'-----------------------------------------------------------------------
	'Echo()
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		 WScript.Echo s
	End Sub
	'-----------------------------------------------------------------------
	'Usage() 使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "xls2temp.vbs [option] <*.xlsx>"
		Echo "Ex."
		Echo "cscript//nologo xls2temp.vbs "
	End Sub
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Private Function Init()
		Debug ".Init()"
		Init = True
		dim	strArg
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "?"
				Usage
				Exit Function
			case else
				Echo "オプションエラー:" & strArg
				Exit Function
			end select
		Next
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Echo "オプションエラー:" & strArg
				Exit Function
			end if
		Next
		if strFileName = "" then
			Echo "ファイルを指定して下さい."
			Exit Function
		end if
		Init = False
	End Function
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Private Function Run()
		Debug ".Run()"
		if Init() then
			Exit Function
		end if
		OpenDb
		Load
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
	Private	objSheet
	Private	strBookName
    Public Function Load()
		Debug ".Load():" & strFileName
		Call CreateExcel()
		Call OpenBook(strFileName)
		Call LoadBook()
		Call CloseBook()
	End Function
	'-------------------------------------------------------------------
	'LoadBook
	'-------------------------------------------------------------------
	Private	lngTopRow
	Private	intCol
	Private	strMaxCol
    Private Function LoadBook()
		Debug ".LoadBook():" & objBook.Name
		for each objSheet in objBook.Worksheets
			LoadSheet
		next
	End Function
	'-------------------------------------------------------------------
	'LoadSheet
	'-------------------------------------------------------------------
	Private	lngMaxRow
	Private	lngRow
	Private	intMaxCol
    Private Function LoadSheet()
		Debug ".LoadSheet():" & objSheet.Name
		Debug "  xlLastCell:" & objSheet.Range("A1").SpecialCells(xlLastCell).Address
		lngMaxRow = objSheet.Range("A1").SpecialCells(xlLastCell).Row
		intMaxCol = objSheet.Range("A1").SpecialCells(xlLastCell).Column
		for lngRow = 1 to lngMaxRow
			if lngRow = 1 then
				Write "Delete:"
				WriteLine Delete()
			end if
			if lngRow > 100 then
'				exit for
			end if
			Write lngRow & "/" & lngMaxRow & ":" & objSheet.Name
			Write ":" & objSheet.Range("A" & lngRow)
			Write " " & objSheet.Range("B" & lngRow)
			Write " " & objSheet.Range("C" & lngRow)
			Write " " & objSheet.Range("D" & lngRow)
			Write ":"
			WriteLine Insert()
		next
	End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
    Private Function Delete()
		Debug ".Delete()"
		AddSql ""
		AddSql "delete from CsvTemp"
		AddSql " where Filename = '" & objBook.Name & "'"
		AddSql " and SheetName = '" & objSheet.Name & "'"
		Delete = CallSql(strSql)
    End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
    Private Function Insert()
		Debug ".Insert()"

		AddSql ""
		AddSql "insert into CsvTemp"
		AddSql "(Filename"
		AddSql ",SheetName"
		AddSql ",Row"
		dim	i
		for	i = 1 to intMaxCol
			if i < 100 then
				AddSql ",Col" & right("00" & i,2)
			else
				AddSql ",Col" & right("000" & i,3)
			end if
		next
		AddSql ",Col"
		AddSql ") values ("
		AddSql " '" & objBook.Name & "'"
		AddSql ",'" & objSheet.Name & "'"
		AddSql "," & lngRow
		dim	objRange
		set objRange = objSheet.Range("A" & lngRow)
		for	i = 1 to intMaxCol
			AddSql ",'" & GetValue(objRange) & "'"
			set objRange = objRange.Offset(0,1)
		next
		AddSql "," & i
		AddSql ")"
		Insert = CallSql(strSql)
    End Function
	'-------------------------------------------------------------------
	'GetValue()
	'-------------------------------------------------------------------
	Private	Function GetValue(objR)
		dim	strValue
		on error resume next
		strValue = Trim(objR)
		if Err.Number <> 0 then
'			Wscript.StdOut.WriteLine ".GetValue():0x" & Hex(Err.Number) & ":" & Err.Description
'			Wscript.StdOut.WriteLine
'			Wscript.StdOut.WriteLine objR.Address & ":(" & objR.Text & ")"
'			Wscript.Quit
			strValue = Trim(objR.Text)
		end if
		on error goto 0
		if strValue <> "" then
			if Asc(strValue) = 0 then
				strValue = ""
			end if
		end if
		strValue = Replace(strValue,"'","''")
		strValue = Replace(strValue,vbCr,"")
		strValue = Replace(strValue,vbLf,"")
'		Debug "GetValue():" & objR.Address & ":" & strValue & ":"
		GetValue = strValue
	End Function
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private Function CreateExcel()
		Debug ".CreateExcel()"
		if objExcel is nothing then
			Debug ".CreateExcel():CreateObject(Excel.Application)"
			Set objExcel = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private Function OpenBook(byVal strBkName)
		Debug ".OpenBook()"
		if objBook is nothing then
			strBkName = AbsPath(strBkName)
			Debug ".OpenBook().Open:" & strBkName
			Write strBkName & " :"
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
			WriteLine "Ok"
		end if
	end function
	'-------------------------------------------------------------------
	'AbsPath
	'-------------------------------------------------------------------
	Private	Function AbsPath(byVal strPath)
		dim	objFso
		Set objFso = CreateObject("Scripting.FileSystemObject")
		AbsPath = objFso.GetAbsolutePathName(strPath)
		Set objFso = Nothing
	End Function
	'-------------------------------------------------------------------
	'Excel ファイルクローズ
	'-------------------------------------------------------------------
	Private Function CloseBook()
		Debug ".CloseBook()"
		if not objBook is nothing then
			Debug ".CloseBook().Close:" & objBook.Name
			Call objBook.Close(False)
			set objBook = nothing
		end if
	end function
	'-------------------------------------------------------------------
	'文字列追加 strSql
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
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql()"
		Debug strSql
		on error resume next
		Call objDb.Execute(strSql)
		if Err.Number <> 0 then
			WriteLine ".CallSql():0x" & Hex(Err.Number) & ":" & Err.Description
			WriteLine ""
			WriteLine strSql
			Wscript.Quit
		end if
		on error goto 0
		CallSql = RowCount()
'		on error resume next
'		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-----------------------------------------------------------------------
	'RowCount()
	'-----------------------------------------------------------------------
    Public Function RowCount()
		Debug ".RowCount()"
		dim	objRow
		set	objRow = objDb.Execute("select @@rowcount")
		RowCount = objRow.Fields(0)
		set	objRow = Nothing
	End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
'		objDB.commandTimeout = 0
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
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		set	objExcel = nothing
		set	objBook = nothing
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		Run
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set	objBook = nothing
		if not objExcel is nothing then
			objExcel.Quit
		end if
		set	objExcel = nothing
		set objDB = nothing
		strDBName = GetOption("db"	,"newsdc")
    End Sub
	'-------------------------------------------------------------------
	'Write
	'-------------------------------------------------------------------
	Private Sub Write(byVal s)
		Wscript.StdOut.Write s
	End Sub
	'-------------------------------------------------------------------
	'WriteLine
	'-------------------------------------------------------------------
	Private Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine s
	End Sub
	'-----------------------------------------------------------------------
	'Debug
	'-----------------------------------------------------------------------
	Private Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
	'-----------------------------------------------------------------------
	Private Function GetOption(byval strName ,byval strDefault)
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
