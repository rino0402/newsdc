Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "csvtemp.vbs [option]"
	Wscript.Echo " /db:newsdc9	データベース"
	Wscript.Echo " /i:0			表示のみ"
	Wscript.Echo " /i:1			Insert (default)"
	Wscript.Echo " /i:2			AddNew"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript csvtemp.vbs /db:newsdc9 pop3w9\tmp\在庫_棚番1.csv"
End Sub
'-----------------------------------------------------------------------
'BoCnv
'-----------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
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

Const	cInsertNone		= 0
Const	cInsertSql		= 1
Const	cInsertAddNew	= 2
Class Csv
	Private	intInsert
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strPathName
	Private	strFileName
	Private	objFileSys
	Private	objFile
	Private	strDT
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "i"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
		select case intInsert
		case 0,1,2
		case else
			Init = "オプションエラー /i:" & intInsert
			Disp Init
			Exit Function
		end select
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		intInsert = GetOption("i"	,1)
		strDBName = GetOption("db"	,"newsdc")
		set objRs = nothing
		set objDB = nothing
		set	objFile	= nothing
		Set objFileSys	= WScript.CreateObject("Scripting.FileSystemObject")
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
		set	objFile	= nothing
		set objFileSys	= nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			Debug ".Run():" & strArg
			strPathName = strArg
			strFileName = GetFileName(strPathName)
			Call Load()
		Next
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load():" & strPathName
		select case FileType()
'		case "excel"
'			Call CreateExcelApp()
'			Call OpenExcel()
'			Call LoadExcel()
		case "csv"
			Call OpenCsv()
			Call LoadCsv()
			Call CloseCsv()
		end select
	End Function
	'-------------------------------------------------------------------
	'ファイルの種類
	'-------------------------------------------------------------------
	Private Function FileType()
		FileType = ""
		select case lcase(fileExt(strPathName))
		case "xls","xlsx"	FileType = "excel"
		case "csv"			FileType = "csv"
		end select
		Debug(".FileType():" & FileType)
	End Function
	'-------------------------------------------------------------------
	'ファイル名(パスを除く)
	'-------------------------------------------------------------------
	Private Function GetFileName(byVal f)
		dim	strFName
		strFName = objFileSys.GetBaseName(f)
		strFName = strFName & "."
		strFName = strFName & objFileSys.GetextensionName(f)
		GetFileName	= strFName
	End Function
	'-------------------------------------------------------------------
	'拡張子
	'-------------------------------------------------------------------
	Private Function fileExt(byVal f)
		dim	strExt
		strExt = objFileSys.GetextensionName(f)
		fileExt = strExt
	End Function
	'-------------------------------------------------------------------
	'絶対パス
	'-------------------------------------------------------------------
	Private Function GetAbsPath(byVal strPath)
		strPath		= objFileSys.GetAbsolutePathName(strPath)
		GetAbsPath	= strPath
	End Function
	'-------------------------------------------------------------------
	'CSV ファイルクローズ
	'-------------------------------------------------------------------
	Private Function CloseCsv()
		Debug ".CloseCsv()"
		objFile.Close
		set objFile		= nothing
	end function
	'-------------------------------------------------------------------
	'CSV ファイルオープン
	'-------------------------------------------------------------------
	Private Function OpenCsv()
		Debug ".OpenCsv():" & GetAbsPath(strPathName)
		Set objFile	= objFileSys.OpenTextFile(GetAbsPath(strPathName), ForReading, False)
	end function
	'-------------------------------------------------------------------
	'Delete
	'-------------------------------------------------------------------
	Private Function DeleteCsv()
		Debug ".DeleteCsv()"
		if intInsert = 0 then
			Exit Function
		end if
		Call AddSql("")
		Call AddSql("delete from CsvTemp")
		Call AddSql("where Filename = '" & strFileName & "'")
		Call Disp(strSql)
		Call CallSql(strSql)
	End Function
	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Private	lngRow
	Private	strBuff
	Private Function LoadCsv()
		Debug ".LoadCsv()"
		Call DeleteCsv()
		lngRow = 0
		do while ( objFile.AtEndOfStream = False )
			strBuff = objFile.ReadLine()
			lngRow = lngRow + 1
			Call LoadLine()
		loop
	end function
	'-------------------------------------------------------------------
	'読込処理(CSV 行)
	'-------------------------------------------------------------------
	Private	aryBuff
	Private	intCol
	Private Function LoadLine()
		Debug ".LoadLine():" & lngRow & ":" & strBuff
		aryBuff = GetTab(strBuff)
		intCol = 0
		Call InsertRow()
	end function
	'-------------------------------------------------------------------
	'行追加
	'-------------------------------------------------------------------
	Private Function InsertRow()
		Debug ".InsertRow():" & intInsert
		Disp strFilename & ":" & lngRow & ":" & Left(Replace(strBuff,vbTab,""),50)
		select case intInsert
		case 1:
			Call InsertSql()
		case 2:
			Call InsertAddNew()
		end select
	End Function
	'-------------------------------------------------------------------
	'AddNew
	'-------------------------------------------------------------------
	Private	Function InsertAddNew()
		Debug ".InsertAddNew()"
		if objRs is nothing then
			Call OpenRs()
		end if
		objRs.AddNew
		objRs.Fields("Filename")	= strFileName
		objRs.Fields("Row")			= lngRow
		intCol = 0
		dim	c
		for each c in (aryBuff)
			c = GetTrim(c)
			intCol = intCol + 1
			Debug ".InsertAddNew():" & intCol & ":(" & c & ")"
			objRs.Fields("Col" & right("00" & intCol,2)) = c
		next
		objRs.Fields("Col")			= intCol
		objRs.Update
	End Function
	'-------------------------------------------------------------------
	'Sql:insert values
	'-------------------------------------------------------------------
	Private	strValues
	Private	Function AddValue(byVal strV)
		if strV = "" then
			strValues = strV
		end if
		if strValues <> "" then
			strValues = strValues & ","
		end if
		strValues = strValues & strV
	End Function
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
	'1行追加
	'-------------------------------------------------------------------
	Private	Function InsertSql()
		Debug ".InsertSql()"
		dim	c
		Call AddValue("")
		Call AddValue("'" & strFileName & "'")
		Call AddValue(lngRow)
		for each c in (aryBuff)
			c = GetTrim(c)
			intCol = intCol + 1
			Debug ".InsertSql():" & lngRow & ":" & intCol & ":" & c
			Call AddValue("'" & c & "'")
		next
		Call AddValue(intCol)

		Call AddSql("")
		Call AddSql("insert into CsvTemp")
		Call AddSql("(Filename")
		Call AddSql(",Row")
		dim	i
		for	i = 1 to intCol
			Call AddSql(",Col" & right("00" & i,2))
		next
		Call AddSql(",Col")
		Call AddSql(") values (")
		Call AddSql(strValues)
		Call AddSql(")")
		Debug ".InsertSql():" & strSql
		Call CallSql(strSql)
	End Function
	'-------------------------------------------------------------------
	'CSV Trim
	'-------------------------------------------------------------------
	Private Function GetTrim(byval c)
		if left(c,1) = """" then
			if right(c,1) = """" then
				c = Right(c,Len(c) -1 )
				if Len(c) > 0 then
					c = Left(c,Len(c) -1 )
				end if
			end if
		end if
		GetTrim = c
	End Function
	'-------------------------------------------------------------------
	'CSV配列(Tab)
	'-------------------------------------------------------------------
	Private Function GetTab(ByVal s)
	    Dim r
		if inStr(s,vbTab) > 0 then
			r = Split(s,vbTab)
		else
			'カンマ区切り
			r = Split(GetTrim(trim(s)),""",""")
		end if
		GetTab = r
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
	'OpenRs
	'-------------------------------------------------------------------
    Private Function OpenRs()
		Debug ".OpenRs()"
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
'		Call objRs.Open("CsvTemp", objDb, adOpenKeyset, adLockOptimistic,adCmdTableDirect)
		Call objRs.Open("CsvTemp", objDb, adOpenForwardOnly, adLockOptimistic,adCmdTableDirect)
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
	'Field名
	'-------------------------------------------------------------------
	Public Function GetFields(byVal strTable)
		Debug ".GetFields():" & strTable
		dim	strFields
		strFields = ""
		dim	objRs
		set objRS = objDB.Execute("select top 1 * from " & strTable)
		dim	objF
		for each objF in objRS.Fields
			if strFields <> "" then
				strFields = strFields & ","
			end if
			strFields = strFields & objF.Name
		next
		set objRs = nothing
		GetFields = strFields
	End Function
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
'			WScript.StdErr.WriteLine
			Wscript.Echo strMsg
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
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objCsv
	Set objCsv = New Csv
	if objCsv.Init() <> "" then
		call usage()
		exit function
	end if
	call objCsv.Run()
End Function
