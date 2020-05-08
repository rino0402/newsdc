Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objZaikoNo
	Set objZaikoNo = New ZaikoNo
	objZaikoNo.Run
	Set objZaikoNo = nothing
End Function
'-----------------------------------------------------------------------
'ZaikoNo
'-----------------------------------------------------------------------
Class ZaikoNo
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "ZaikoNo.vbs [option] <*.xlsx>"
		Echo "Ex."
		Echo "cscript//nologo ZaikoNo.vbs /db:newsdc1"
		Echo "cscript//nologo ZaikoNo.vbs /order /db:newsdc1   ※入荷日順セット"
		Echo "cscript//nologo ZaikoNo.vbs /loc0 /db:newsdc1    ※標準棚番０セット"
		Echo "cscript//nologo ZaikoNo.vbs /qty0 /db:newsdc1    ※在庫０削除"
	End Sub
	Private	strDBName
	Private	strDt
	Private	objDB
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		strDt = GetOption("dt"	,"")
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
			if WScript.Arguments.Named.Exists("order") then
				Order
			elseif WScript.Arguments.Named.Exists("loc0") then
				Loc0
			elseif WScript.Arguments.Named.Exists("qty0") then
				Qty0
			else
				Load
			end if
			CloseDb
		end if
	End Function
	'-----------------------------------------------------------------------
	'Qty0() 在庫０削除
	'-----------------------------------------------------------------------
    Public Function Qty0()
		Debug ".Qty0()"
		AddSql ""
		AddSql "select"
		AddSql " *"
		AddSql "from ZaikoNo z"
		AddSql "where (JGYOBU + HIN_GAI) not in (select distinct (JGYOBU + HIN_GAI) from zaiko)"
		AddSql "and Qty <> 0"
		AddSql "order by"
		AddSql " JGYOBU"
		AddSql ",Loc"
		Write "検索中...Qty0:"
		set objRs = objDb.Execute(strSql)
		WriteLine "ok"
		do while objRs.Eof = False
			DeleteQty0
			objRs.MoveNext
		loop
	End Function
	'-------------------------------------------------------------------
	'DeleteQty0
	'-------------------------------------------------------------------
    Private Function DeleteQty0()
		Debug ".DeleteQty0()"

		Write objRs.Fields("JGYOBU") & " "
		Write objRs.Fields("HIN_GAI")
		Write objRs.Fields("Loc") & " "
		Write objRs.Fields("No") & " "
		Write Right(space(8) & objRs.Fields("Qty"),8)
		Write Right(space(8) & objRs.Fields("QtyM"),8)
		Write Right(space(8) & objRs.Fields("QtyS"),8) & " "

		AddSql ""
		AddSql "update ZaikoNo"
		AddSql "set Qty = 0"
		AddSql "  , QtyM = 0"
		AddSql "  , QtyS = 0"
		AddSql "where JGYOBU = '" & objRs.Fields("JGYOBU") & "'"
		AddSql "  and HIN_GAI = '" & objRs.Fields("HIN_GAI") & "'"
		AddSql "  and Loc = '" & objRs.Fields("Loc") & "'"
		WriteLine ":" & Execute(strSql)
    End Function
	'-----------------------------------------------------------------------
	'Loc0() 標準棚番０セット
	'-----------------------------------------------------------------------
    Public Function Loc0()
		Debug ".Loc0()"
		AddSql ""
		AddSql "select"
		AddSql " z.JGYOBU"
		AddSql ",z.HIN_GAI"
		AddSql ",RTrim(ifnull(i.ST_Soko + i.ST_Retu + i.ST_Ren + i.ST_Dan,''))	StdLoc"
		AddSql "from ZaikoNo z"
		AddSql "left outer join Item i on (z.JGYOBU = i.JGYOBU and i.NAIGAI = '1' and z.HIN_GAI = i.HIN_GAI)"
		AddSql "where z.No = 1"
		AddSql "and (z.JGYOBU + z.HIN_GAI) not in (select (JGYOBU + HIN_GAI) from ZaikoNo where No = 0)"
		Write "検索中...Loc0:"
		set objRs = objDb.Execute(strSql)
		WriteLine "ok"
		do while objRs.Eof = False
			InsertLoc0
			objRs.MoveNext
		loop
	End Function
	'-------------------------------------------------------------------
	'InsertLoc0
	'-------------------------------------------------------------------
    Private Function InsertLoc0()
		Debug ".InsertLoc0()"

		Write objRs.Fields("JGYOBU") & " "
		Write objRs.Fields("HIN_GAI")
		Write objRs.Fields("StdLoc") & ""

		AddSql ""
		AddSql "Insert into ZaikoNo"
		AddSql "("
		AddSql "	JGYOBU	"
		AddSql ",	HIN_GAI	"
		AddSql ",	Loc		"
		AddSql ") values ("
		AddSql " '" & objRs.Fields("JGYOBU") & "'"
		AddSql ",'" & objRs.Fields("HIN_GAI") & "'"
		AddSql ",'" & objRs.Fields("StdLoc") & "'"
		AddSql ")"
		WriteLine ":" & Execute(strSql)
    End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Public Function Order()
		Debug ".Order()"
		AddSql ""
		AddSql "select"
		AddSql " *"
		AddSql "from ZaikoNo"
		AddSql "where No > 0"
		AddSql "order by"
		AddSql " JGYOBU"
		AddSql ",HIN_GAI"
		AddSql ",NyukaMin"
		AddSql ",Loc"
		Write "検索中...Order:"
		set objRs = objDb.Execute(strSql)
		WriteLine "ok"
		prvJGYOBU = ""
		prvHIN_GAI = ""
		do while objRs.Eof = False
			Update
			objRs.MoveNext
		loop
	End Function
	'-------------------------------------------------------------------
	'Update
	'-------------------------------------------------------------------
    Private Function Update()
		Debug ".Update()"
		Write objRs.Fields("JGYOBU") & " "
		Write objRs.Fields("HIN_GAI")
		Write objRs.Fields("Loc") & " "
		Write objRs.Fields("NyukaMin") & " "
		Write objRs.Fields("No")

		curJGYOBU = objRs.Fields("JGYOBU")
		curHIN_GAI = objRs.Fields("HIN_GAI")

		if Compare() = True then
			prvJGYOBU = curJGYOBU
			prvHIN_GAI = curHIN_GAI
			intNo = 0
		end if
		intNo = intNo + 1
		Write intNo
		if objRs.Fields("No") <> intNo then
			AddSql ""
			AddSql "update ZaikoNo"
			AddSql "set No = " & intNo
			AddSql "where JGYOBU = '" & objRs.Fields("JGYOBU") & "'"
			AddSql "  and HIN_GAI = '" & objRs.Fields("HIN_GAI") & "'"
			AddSql "  and Loc = '" & objRs.Fields("Loc") & "'"
			CallSql strSql
			Write " 更新"
		end if
		WriteLine ""

	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load()"
		AddSql ""
		AddSql "select"
		AddSql " z.JGYOBU"
		AddSql ",z.HIN_GAI"
		AddSql ",z.Soko_No + z.Retu + z.Ren + z.Dan	Loc"
		AddSql ",Sum(convert(z.YUKO_Z_QTY,sql_decimal)) Qty"
		AddSql ",Sum(if(z.GOODS_ON <>'0',convert(z.YUKO_Z_QTY,sql_decimal),0)) QtyM"
		AddSql ",Sum(if(z.GOODS_ON = '0',convert(z.YUKO_Z_QTY,sql_decimal),0)) QtyS"
		AddSql ",Min(z.NYUKA_DT)	NyukaMin"
		AddSql ",Max(z.NYUKA_DT)	NyukaMax"
		AddSql ",Min(z.NYUKO_DT)	NyukoMin"
		AddSql ",Max(z.NYUKO_DT)	NyukoMax"
		AddSql ",i.ST_Soko + i.ST_Retu + i.ST_Ren + i.ST_Dan	StdLoc"
'		AddSql ",Min(if((z.Soko_No + z.Retu + z.Ren + z.Dan) = (i.ST_Soko + i.ST_Retu + i.ST_Ren + i.ST_Dan),'',z.NYUKA_DT))	Order0"
'		AddSql ",Min(if((z.Soko_No + z.Retu + z.Ren + z.Dan) = (i.ST_Soko + i.ST_Retu + i.ST_Ren + i.ST_Dan),'',z.Soko_No + z.Retu + z.Ren + z.Dan))	Order0"
		AddSql "from Zaiko z"
		AddSql "left outer join Item i on (z.JGYOBU = i.JGYOBU and i.NAIGAI = '1' and z.HIN_GAI = i.HIN_GAI)"
		if strDt <> "" then
			AddSql "where z.HIN_GAI in (select distinct HIN_GAI from p_sagyo_log where HIN_GAI <> '' and (FROM_SOKO <> '' or TO_SOKO <> '') and JITU_DT = '" & strDt & "')"
		end if
'		AddSql "where z.JGYOBU = 'D'"
'		AddSql "  and z.NAIGAI = '1'"
		AddSql "group by"
		AddSql " z.JGYOBU"
		AddSql ",z.HIN_GAI"
		AddSql ",Loc"
		AddSql ",StdLoc"
		AddSql "order by"
		AddSql " z.JGYOBU"
		AddSql ",z.HIN_GAI"
		AddSql ",Loc"
		Write "検索中..." & strDt
		set objRs = objDb.Execute(strSql)
		WriteLine "ok"
		prvJGYOBU = ""
		prvHIN_GAI = ""
		do while objRs.Eof = False
			Insert
			objRs.MoveNext
		loop
	End Function
	'-------------------------------------------------------------------
	'Compare
	'-------------------------------------------------------------------
	Private Function Compare()
		Compare = True
		if curJGYOBU <> prvJGYOBU then
			exit function
		end if
		if curHIN_GAI <> prvHIN_GAI then
			exit function
		end if
		Compare = False
	End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
	Private	intNo
	Private	prvJGYOBU
	Private	prvHIN_GAI
	Private	curJGYOBU
	Private	curHIN_GAI
    Private Function Insert()
		Debug ".Insert()"

		Write objRs.Fields("JGYOBU") & " "
		Write objRs.Fields("HIN_GAI")
		Write objRs.Fields("Loc")
		Write Right(space(8) & objRs.Fields("Qty"),8)
		Write Right(space(8) & objRs.Fields("QtyM"),8)
		Write Right(space(8) & objRs.Fields("QtyS"),8) & " "
		Write objRs.Fields("NyukaMin") & " "
'		Write objRs.Fields("NyukaMax") & " "
'		Write objRs.Fields("NyukoMin") & " "
'		Write objRs.Fields("NyukoMax") & " "
		Write objRs.Fields("StdLoc") & " "
'		Write objRs.Fields("Order0")

		curJGYOBU = objRs.Fields("JGYOBU")
		curHIN_GAI = objRs.Fields("HIN_GAI")

		if Compare() = True then
			AddSql ""
			AddSql "delete From ZaikoNo"
			AddSql "where JGYOBU = '" & objRs.Fields("JGYOBU") & "'"
			AddSql "  and HIN_GAI = '" & objRs.Fields("HIN_GAI") & "'"
			Write "削除 "
			CallSql strSql

			prvJGYOBU = curJGYOBU
			prvHIN_GAI = curHIN_GAI
			intNo = 0
		end if

		AddSql ""
		AddSql "Insert into ZaikoNo"
		AddSql "("
		AddSql "	JGYOBU	"
		AddSql ",	HIN_GAI	"
		AddSql ",	No		"
		AddSql ",	NoLoc	"
		AddSql ",	Loc		"
		AddSql ",	Qty		"
		AddSql ",	QtyM	"
		AddSql ",	QtyS	"
		AddSql ",	NyukaMin"
		AddSql ",	NyukaMax"
		AddSql ",	NyukoMin"
		AddSql ",	NyukoMax"
		AddSql ") values ("
		AddSql " '" & objRs.Fields("JGYOBU") & "'"
		AddSql ",'" & objRs.Fields("HIN_GAI") & "'"
		if objRs.Fields("Loc") = objRs.Fields("StdLoc") then
			AddSql ",0"
			AddSql ",0"
		else
			intNo = intNo + 1
			AddSql "," & intNo
			AddSql "," & intNo
		end if
		AddSql ",'" & objRs.Fields("Loc") & "'"
		AddSql "," & objRs.Fields("Qty")
		AddSql "," & objRs.Fields("QtyM")
		AddSql "," & objRs.Fields("QtyS")
		AddSql ",'" & objRs.Fields("NyukaMin") & "'"
		AddSql ",'" & objRs.Fields("NyukaMax") & "'"
		AddSql ",'" & objRs.Fields("NyukoMin") & "'"
		AddSql ",'" & objRs.Fields("NyukoMax") & "'"
		AddSql ")"
		CallSql strSql
		WriteLine ""
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
			case "dt"
			case "debug"
			case "order"
			case "loc0"
			case "qty0"
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
