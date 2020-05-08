Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Bill
'2017.06.20 新規
'-----------------------------------------------------------------------
Class Bill
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "BillPnYm.vbs [option]"
		Echo "Ex."
		Echo "cscript//nologo BillPnYm.vbs /db:newsdc1 201504 201705"
		Echo "cscript//nologo BillPnYm.vbs /db:newsdc3 201504 201705"
		Echo "cscript//nologo BillPnYm.vbs /db:newsdc4 201504 201705"
	End Sub
	Private	strDBName
	Private	objDB
	Private	strYm1
	Private	strYm2
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		strYm1		= ""
		strYm2		= ""
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
		if Init() <> "" then
			usage
			exit function
		end if
		OpenDb
		Load
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Private Function Load()
		Debug ".Load():" & strYm1 & "," & strYm2
		if strYm2 = "" then
			strYm2 = strYm1
		end if
		do while true
			Debug ".Load():" & strYm1 & "," & strYm2
			if strYm1 > strYm2 then
				exit do
			end if
			Delete strYm1
			Insert strYm1
			strYm1 = NextYm(strYm1)
		loop
'		Call InsertCsv()
	End Function
	'-------------------------------------------------------------------
	'NextYm
	'-------------------------------------------------------------------
	Private	Function NextYm(byVal strYm)
		dim	iYear
		dim	iMonth
		iYear = CInt(Left(strYm,4))
		iMonth = CInt(Right(strYm,2))
		iMonth = iMonth + 1
		if iMonth > 12 then
			iYear = iYear + 1
			iMonth = 1
		end if
		NextYm = iYear & Right("0" & iMonth,2)
	End Function
	'-------------------------------------------------------------------
	'Delete
	'-------------------------------------------------------------------
    Private Function Delete(byVal strYm)
		Debug ".Delete():" & strYm
		Write "削除:" & strYm & "..."
		AddSql ""
		AddSql "delete from BillPnYm"
		AddSql "where YM = '" & strYm & "'"
		CallSql strSql
		WriteLine RowCount()
    End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
    Private Function Insert(byVal strYm)
		Debug ".Insert():" & strYm
		Write "登録:" & strYm & "..."
		AddSql ""
		AddSql "insert into BillPnYm"
		AddSql "("
		AddSql "	 JGYOBU		"
		AddSql "	,Pn			"
		AddSql "	,Ym			"
		AddSql "	,BillKbn	"
		AddSql "	,DelvKbn	"
		AddSql "	,Cnt		"
		AddSql "	,Qty		"
		AddSql "	,PickFee	"
		AddSql "	,DelvFee	"
		AddSql "	,WorkFee	"
		AddSql "	,PackFee	"
		AddSql "	,WrapFee	"
		AddSql "	,AttdFee	"
		AddSql ")"
		AddSql "select"
		AddSql " JGYOBU"
		AddSql ",Pn"
		AddSql ",YM"
		AddSql ",KBN BillKbn"
		AddSql ",case"
		AddSql " when KBN = 'A' and Left(SyukaCd,1) = 'A'	then '1'"
		AddSql " when KBN = 'A' and SyukaCd = '39040'		then '2'"
		AddSql " when KBN = 'A' and Left(SyukaCd,2) = '22'	then '3'"
		AddSql " when KBN = 'A'								then '4'"
		AddSql " when KBN = 'B'								then '5'"
		AddSql " when KBN = 'C'								then '6'"
		AddSql " else ''"
		AddSql " end DelvKbn"
		AddSql ",count(*)"
		AddSql ",Sum(Qty)"
		AddSql ",Sum(Pick)"
		AddSql ",Sum(Ship)"
		AddSql ",Sum(Koryo)"
		AddSql ",Sum(Hako)"
		AddSql ",Sum(Gaiso)"
		AddSql ",Sum(Futai)"
		AddSql "from Bill"
		AddSql "where YM = '" & strYm & "'"
		AddSql "group by"
		AddSql " JGYOBU"
		AddSql ",Pn"
		AddSql ",YM"
		AddSql ",BillKbn"
		AddSql ",DelvKbn"
		CallSql strSql
		WriteLine RowCount()
    End Function
	'-------------------------------------------------------------------
	'RowCount
	'-------------------------------------------------------------------
	Public Function RowCount()
		Debug ".RowCount()"
		dim	objRow
		set objRow = objDb.Execute("select @@rowcount")
		RowCount = objRow.Fields(0)
		set objRow = nothing
    End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		call objDb.Execute(strSql)
		if Err.Number <> 0 then
			WriteLine ""
			WriteLine "CallSql():0x" & Hex(Err.Number)
			WriteLine Err.Description
			WriteLine ""
			Debug strSql
			Wscript.Quit
		end if
		on error goto 0
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
	'標準出力:Write
	'-----------------------------------------------------------------------
	Public Sub Write(byVal strMsg)
		Wscript.StdOut.Write strMsg
	End Sub
	'-----------------------------------------------------------------------
	'標準出力:WriteLine
	'-----------------------------------------------------------------------
	Public Sub WriteLine(byVal strMsg)
		Wscript.StdOut.WriteLine strMsg
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Echo(byVal strMsg)
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
			if strYm1 = "" then
				strYm1 = strArg
			elseif strYm2 = "" then
				strYm2 = strArg
			else
				Init = "オプションエラー:" & strArg
				Echo Init
				Exit Function
			end if
		Next
		if strYm1 = "" then
			Init = "年月を指定して下さい."
			Echo Init
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Init = "オプションエラー:" & strArg
				Echo Init
				Exit Function
			end select
		Next
	End Function
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objBill
	Set objBill = New Bill
	objBill.Run
	Set objBill = Nothing
End Function
