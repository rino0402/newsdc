Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "b2excel.vbs [option] <b2file.xlsx>"
	Wscript.Echo " /db:newsdc1 データベース"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo b2excel.vbs /db:newsdc1 b2out.xls"
End Sub
'-----------------------------------------------------------------------
'HtDelvNo
'2016.10.26 新規 b2list.xls→ B2List
'-----------------------------------------------------------------------
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CommandTypeEnum Values ----
Const adCmdUnspecified	= -1	' Unspecified type of command 
Const adCmdText			= 1		' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2		' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4		' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8		' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

Const xlUp = -4162

Class B2Excel
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strFileName
	Private	objExcel
	Private	objBook
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		set	objExcel = nothing
		set	objBook = nothing
		optNew = "new"
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
		set	objBook = nothing
		set	objExcel = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		Call Load()
		Call CloseDb()
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
		for each objSheet in objBook.Worksheets
			Call LoadSheet()
		next
		Call CloseBook()
	End Function
	'-------------------------------------------------------------------
	'LoadSheet() シート読込
	'-------------------------------------------------------------------
	Private	lngRow
	Private	lngRowMax
	Private Function LoadSheet()
		Debug ".LoadSheet():" & objSheet.Name
		Call objDb.Execute("delete from b2excel")
		lngRow = 1
		lngRowMax = GetRowMax()
		for lngRow = 1 to lngRowMax
			Call LoadRow()
		next
	End Function
	'-------------------------------------------------------------------
	'GetRowMax() 最終行
	'-------------------------------------------------------------------
	Private Function GetRowMax()
		Debug ".GetRowMax():" & objSheet.Name
		lngRowMax = objSheet.Rows.Count
		lngRowMax = objSheet.Range("B" & lngRowMax).End(xlUp).Row
		GetRowMax = lngRowMax
	End Function
	'-------------------------------------------------------------------
	'1行表示
	'-------------------------------------------------------------------
	Private	intCol
	Private Function LoadRow()
		Debug ".LoadRow()"
		WScript.StdOut.Write lngRow & "/" & lngRowMax
'		WScript.StdOut.Write " " & RTrim(objSheet.Range("A" & lngRow))
		WScript.StdOut.Write " " & RTrim(objSheet.Range("D" & lngRow))
		WScript.StdOut.Write " " & RTrim(objSheet.Range("H" & lngRow))
		WScript.StdOut.Write " " & RTrim(objSheet.Range("N" & lngRow))
		WScript.StdOut.Write "" & RTrim(objSheet.Range("O" & lngRow))
		WScript.StdOut.Write "" & RTrim(objSheet.Range("P" & lngRow))
		dim	strSql1
		dim	strSql2
		strSql1 = "insert into b2excel (idRow"
		strSql2 = "(" & lngRow
		for intCol = 1 to 50
'			Debug objSheet.Cells(lngRow,intCol)
			strSql1 = strSql1 & ",c" & Right("0" & intCol,2)
			strSql2 = strSql2 & ",'" & Trim(objSheet.Cells(lngRow,intCol)) & "'"
		next
		strSql1 = strSql1 & ")"
		strSql2 = strSql2 & ")"
		strSql = strSql1 & " values " & strSql2
		Debug strSql
		Call objDb.Execute(strSql)
		WScript.StdOut.WriteLine
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
	'-----------------------------------------------------------------------
	'SQL文字列追加
	'-----------------------------------------------------------------------
	Private	strSql
	Public Function SetSql(byVal s)
		if s = "" then
			strSql = ""
		else
			if strSql <> "" then
				strSql = strSql & " "
			end if
			strSql = strSql & s
		end if
	End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
'		on error resume next
		Call objDB.Execute(strSql)
'		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		strField = RTrim("" & objRs.Fields(strName))
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
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
'			strBkName = strScriptPath & strBkName
			Debug ".OpenBook().Open:" & strBkName
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
		end if
	end function
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
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
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
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
	Private	optNew
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		if strFileName = "" then
			Init = "ファイルを指定して下さい."
			Disp Init
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "new"
				optNew = "new"
			case "all"
				optNew = "all"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objB2Excel
	Set objB2Excel = New B2Excel
	if objB2Excel.Init() <> "" then
		call usage()
		exit function
	end if
	call objB2Excel.Run()
End Function
