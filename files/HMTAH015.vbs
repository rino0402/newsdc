Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "HMTAH015.vbs [option]"
	Wscript.Echo " /db:newsdc1	データベース"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript HMTAH015.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'HMTAH015
'-----------------------------------------------------------------------
Class HMTAH015
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strPathName
	Private	strFileName
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		strFileName = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName	= GetOption("db"	,"newsdc")
		set objDB	= nothing
		set objRs	= nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		Call GetFilename()
		Call Insert()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Insert() 登録
	'-----------------------------------------------------------------------
    Private Function GetFilename()
		Debug ".GetFilename()"

		if strFilename <> "" then
			Wscript.StdOut.WriteLine "Filename:" & strFileName
			exit function
		end if

		AddSql ""
		AddSql "select max(Filename) from HMTAH015_t"

		Wscript.StdOut.Write "HMTAH015_t:"
		CallSql
		do while objRs.Eof = False
			strFileName = RTrim(objRs.Fields(0))
			exit do
		loop
		Wscript.StdOut.WriteLine strFileName
	End Function
	'-----------------------------------------------------------------------
	'Insert() 登録
	'-----------------------------------------------------------------------
    Public Function Insert()
		Debug ".Insert()"

		'追加 SQL
		AddSql ""
		AddSql "insert into HMTAH015 "
		'フィールド名 HMTAH015
		Set objRs = objDB.Execute("select top 1 * from HMTAH015")
		dim	strC
		strC = "("
		dim	objF
		for each objF in objRs.Fields
			Debug "HMTAH015:" & objF.Name
			AddSql strC & objF.Name
			strC = ","
		next
		AddSql ") select "
		'フィールド名 HMTAH015_t
		Set objRs = objDB.Execute("select top 1 * from HMTAH015_t")
		strC = ""
		for each objF in objRs.Fields
			Debug "HMTAH015_t:" & objF.Name
			dim	strFeild
			strFeild = objF.Name
			select case strFeild
			case "Filename","Row"
				strFeild = ""
			case "Qty"
				strFeild = "convert(Qty,sql_decimal)"
			case else
			end select
			if strFeild <> "" then
				AddSql strC & strFeild
				strC = ","
			end if
		next
		AddSql " from HMTAH015_t"
		AddSql " where Filename = '" & strFileName & "'"

		Wscript.StdOut.Write "削除:"
		Call objDB.Execute("delete from HMTAH015")
		Wscript.StdOut.WriteLine RowCount()

		Wscript.StdOut.Write "追加:"
		CallSql
		Wscript.StdOut.WriteLine RowCount()
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
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql()
		Debug ".CallSql():" & strSql
'		on error resume next
		Set objRs = objDB.Execute(strSql)
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
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.StdErr.WriteLine strMsg
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
	dim	objClass
	Set objClass = New HMTAH015
	if objClass.Init() <> "" then
		call usage()
		exit function
	end if
	call objClass.Run()
End Function
