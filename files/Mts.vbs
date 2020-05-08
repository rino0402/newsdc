Option Explicit
'-----------------------------------------------------------------------
'Mts.vbs
'出荷先マスター更新
'2016.11.28 新規
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "Mts.vbs [option]"
	Wscript.Echo " /db:newsdc	データベース"
	Wscript.Echo "Ex."
	Wscript.Echo "sc32//nologo Mts.vbs /db:newsdc5"
End Sub

Class Mts
	Private	strDBName
	Private	objDB
	Private	objRs
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
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
		Call Make()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Make() 才数データ更新
	'-----------------------------------------------------------------------
	Private	strPrev
	Private	strCurr
    Private Function Make()
		Debug ".Make()"
		SetSql	""
		SetSql	"select"
		SetSql	"distinct"
		SetSql	"COL_OKURISAKI_CD"
		SetSql	",OKURISAKI_CD"
		SetSql	",OKURISAKI"
		SetSql	",Max(SYUKA_YMD) SYUKA_YMD"
		SetSql	"from ("
		SetSql	"select"
		SetSql	"distinct"
		SetSql	"COL_OKURISAKI_CD"
		SetSql	",OKURISAKI_CD"
		SetSql	",OKURISAKI"
		SetSql	",SYUKA_YMD"
		SetSql	"from Y_SYUKA_H"
		SetSql	"where OKURISAKI_CD <> ''"
		SetSql	"union"
		SetSql	"select"
		SetSql	"distinct"
		SetSql	"COL_OKURISAKI_CD"
		SetSql	",OKURISAKI_CD"
		SetSql	",OKURISAKI"
		SetSql	",SYUKA_YMD"
		SetSql	"from DEL_SYUKA_H"
		SetSql	"where OKURISAKI_CD <> ''"
		SetSql	"and SYUKA_YMD >= '20150401'"
		SetSql	") y"
		SetSql	"where COL_OKURISAKI_CD <> ''"
		SetSql	"group by"
		SetSql	"COL_OKURISAKI_CD"
		SetSql	",OKURISAKI_CD"
		SetSql	",OKURISAKI"
		SetSql	"order by"
		SetSql	"COL_OKURISAKI_CD"
		SetSql	",OKURISAKI_CD"
		SetSql	",SYUKA_YMD desc"
		Debug ".Make():" & strSql
		WScript.StdErr.Write "検索中..."
		set objRs = objDB.Execute(strSql)
		WScript.StdErr.WriteLine "Eof:" & objRs.Eof
		strPrev = ""
		strCurr = ""
		do while objRs.Eof = False
			if DispData() <> "" then
				Call MakeData()
			end if
			WScript.StdOut.WriteLine
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function DispData()
		Debug ".DispData()"
		dim	strCode
		strCurr = GetField("OKURISAKI_CD")
		strCode = strCurr
		if strPrev = strCurr then
			strCode = ""
		end if
		strPrev = strCurr
		WScript.StdOut.Write GetStr("COL_OKURISAKI_CD"	,21)
		WScript.StdOut.Write Format(strCode		,10)
		WScript.StdOut.Write GetStr("SYUKA_YMD"	, 9)
		WScript.StdOut.Write GetStr("OKURISAKI"	,40)
		DispData = strCode
	End Function
	'-------------------------------------------------------------------
	'フィールド値 書式
	'-------------------------------------------------------------------
	Private Function GetStr(byVal strName,byVal intLen)
		GetStr = Format(GetField(strName),intLen)
	End Function
	'-------------------------------------------------------------------
	'書式
	'-------------------------------------------------------------------
	Private Function Format(byVal strV,byVal intLen)
		Format = strV
		if intLen > 0 then
			Format = Left(Format & space(intLen),intLen)
		else
			intLen = Abs(intLen)
			Format = Right(space(intLen) & Format,intLen)
		end if
	End Function
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData()"
		if Update() = 0 then
			Call Insert()
		end if
	End Function
	'-------------------------------------------------------------------
	'Update
	'-------------------------------------------------------------------
	Private	Function Update()
		Debug ".Update()"
		SetSql ""
		SetSql "update Mts"
		SetSql "set"
		SetSql " MUKE_NAME = '" & GetField("OKURISAKI") & "'"
		SetSql "where NAIGAI = '1'"
		SetSql "and MUKE_CODE = '" & GetField("OKURISAKI_CD") & "'"
'		SetSql "and MUKE_NAME <> '" & GetField("OKURISAKI") & "'"
		WScript.StdOut.Write " Upd:"
		Update = CallSql(strSql)
		WScript.StdOut.Write Update
	End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
	Private Function Insert()
		Debug ".Insert()"
		SetSql ""
		SetSql "insert into Mts ("
		SetSql " NAIGAI"	
		SetSql ",MUKE_CODE"
		SetSql ",MUKE_NAME"
		SetSql ") values ("
		SetSql "'1'"
		SetSql ",'" & GetField("OKURISAKI_CD") & "'"
		SetSql ",'" & GetField("OKURISAKI") & "'"
		SetSql ")"
		Debug strSql
		WScript.StdOut.Write " Ins:"
		Insert = CallSql(strSql)
		WScript.StdOut.Write Insert
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
		on error resume next
		objDb.Execute strSql
		on error goto 0
		dim	intNumver
		dim	strDescription
		intNumver = Err.Number
		strDescription	= Err.Description
		if intNumver = 0 then
			dim	objRc
			set objRc = objDb.Execute("select @@rowcount")
			CallSql = objRc.Fields(0)
		else
			CallSql = -1
			WScript.StdOut.Write RTrim("0x" & Hex(intNumver) & " " & strDescription)
		end if
    End Function
	'-------------------------------------------------------------------
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		on error resume next
		strField = RTrim("" & objRs.Fields(strName))
		if Err.Number <> 0 then
			WScript.Echo "GetField(" & strName & "):Error:0x" & Hex(Err.Number)
			WScript.Quit
		end if
		on error goto 0
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
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
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			Init = "オプションエラー:" & strArg
			Disp Init
			Exit Function
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
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objMts
	Set objMts = New Mts
	if objMts.Init() <> "" then
		call usage()
		exit function
	end if
	call objMts.Run()
End Function
