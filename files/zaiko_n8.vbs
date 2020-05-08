Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objZaikoN8
	Set objZaikoN8 = New ZaikoN8
	objZaikoN8.Run
	Set objZaikoN8 = nothing
End Function
'-----------------------------------------------------------------------
'ZaikoN8
'-----------------------------------------------------------------------
Class ZaikoN8
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "Zaiko_N8.vbs [option]"
		Echo "Ex."
		Echo "cscript//nologo Zaiko_N8.vbs /db:newsdc4"
	End Sub
	Private	strDBName
	Private	objDB
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		if Init() = True then
			OpenDb
			Load
			CloseDb
		end if
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
	Private Function Load()
		Debug ".Load()"
		AddSql ""
		AddSql "select"
		AddSql "distinct"
		AddSql " z.JGYOBU"
		AddSql ",z.HIN_GAI"
		AddSql ",z.Soko_No + z.Retu + z.Ren + z.Dan	Loc"
		AddSql ",ifnull(i.ST_Soko + i.ST_Retu + i.ST_Ren + i.ST_Dan,'')	ST_Loc"
'		AddSql ",convert(z.YUKO_Z_QTY,sql_decimal) Qty"
		AddSql "from Zaiko z"
		AddSql "left outer join Item i on (z.JGYOBU = i.JGYOBU and z.NAIGAI = i.NAIGAI and z.HIN_GAI = i.HIN_GAI)"
		AddSql "where z.JGYOBU = 'R'"
		AddSql "and z.NAIGAI = '1'"
		AddSql "and convert(z.YUKO_Z_QTY,sql_decimal) <> 0"
		AddSql "and Loc not in ('90010101','N8000000')"
		AddSql "and Loc <> ST_Loc"
'		AddSql "and (z.Soko_No + z.Retu + z.Ren + z.Dan) <> (i.ST_Soko + i.ST_Retu + i.ST_Ren + i.ST_Dan)"
		AddSql "order by"
		AddSql " z.JGYOBU"
		AddSql ",z.HIN_GAI"
		AddSql ",Loc"
		Write "検索中..."
		set objRs = objDb.Execute(strSql)
		WriteLine "ok"
		do while objRs.Eof = False
			Load1
			objRs.MoveNext
		loop
	End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
	Private	intNo
    Private Function Load1()
		Debug ".Load1()"

		Write objRs.Fields("JGYOBU") & " "
		Write objRs.Fields("HIN_GAI")
		Write objRs.Fields("Loc") & " "
		Write objRs.Fields("ST_Loc")
		dim	strJGYOBU
		strJGYOBU = RTrim(objRs.Fields("JGYOBU"))
		dim	strHIN_GAI
		strHIN_GAI = RTrim(objRs.Fields("HIN_GAI"))
		dim	strLoc
		strLoc = RTrim(objRs.Fields("Loc"))
		'標準棚番セット
		AddSql ""
		AddSql "update item"
		AddSql "set ST_Soko = '" & Mid(strLoc,1,2) & "'"
		AddSql "  , ST_Retu = '" & Mid(strLoc,3,2) & "'"
		AddSql "  , ST_Ren  = '" & Mid(strLoc,5,2) & "'"
		AddSql "  , ST_Dan  = '" & Mid(strLoc,7,2) & "'"
		AddSql "where JGYOBU = '" & strJGYOBU & "'"
		AddSql "and NAIGAI = '1'"
		AddSql "and HIN_GAI = '" & strHIN_GAI & "'"
'		Write strSql
		Execute strSql
		WriteLine ":" & RowCount()
    End Function
	'-----------------------------------------------------------------------
	'RowCount()
	'-----------------------------------------------------------------------
    Private Function RowCount()
		Debug ".RowCount()"
		dim	objRow
		set	objRow = objDb.Execute("select @@rowcount")
		RowCount = objRow.Fields(0)
		set	objRow = Nothing
	End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private Function Execute(byVal strSql)
		Debug ".Execute():" & strSql
		on error resume next
		Call objDb.Execute(strSql)
		Execute = Err.Number & "(0x" & Hex(Err.Number) & ")" & Err.Description
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Private	objRs
	Private Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		Call objDb.Execute(strSql)
		if Err.Number <> 0 then
			Wscript.StdOut.WriteLine "0x" & Hex(Err.Number) & ":" & Err.Description
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end if
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
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
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
	'WriteLine
	'-----------------------------------------------------------------------
	Private Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine s
	End Sub
	'-----------------------------------------------------------------------
	'Write
	'-----------------------------------------------------------------------
	Private Sub Write(byVal s)
		Wscript.StdOut.Write s
	End Sub
	'-----------------------------------------------------------------------
	'Echo
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		Wscript.Echo s
	End Sub
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = False
		For Each strArg In WScript.Arguments.UnNamed
				Echo "オプションエラー:" & strArg
				Usage
				Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "?"
				Usage
				Exit Function
			case else
				Echo "オプションエラー:" & strArg
				Usage
				Exit Function
			end select
		Next
		Init = True
	End Function
End Class
