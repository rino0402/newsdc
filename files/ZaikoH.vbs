Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Call Main()
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objExcel
	Set objExcel = New Excel
	objExcel.Run
	Set objExcel = Nothing
End Function
'-----------------------------------------------------------------------
'Excel
'2017.03.07 新規
'-----------------------------------------------------------------------
Const xlUp = -4162

Class Excel
	'-----------------------------------------------------------------------
	'使用方法
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "ZaikoH.vbs [option] <*.xls>"
		Echo "/db:newsdc1"
		Echo "/debug"
		Echo "Ex.Satl[Satellite]"
		Echo "cscript//nologo ZaikoH.vbs 6月12日調理サテ在庫.xls /kubun:Satl /jcode:00021397 /dt:20170612 /db:newsdc1"
		Echo "cscript//nologo ZaikoH.vbs 6月12日炊飯サテ在庫.xls /kubun:Satl /jcode:00023410 /dt:20170612 /db:newsdc1"
		Echo "cscript//nologo ZaikoH.vbs 6月12日IHサテ在庫.xls   /kubun:Satl /jcode:00023510 /dt:20170612 /db:newsdc1"
		Echo "cscript//nologo ZaikoH.vbs ""調理　前月サテ在庫.xls"" /kubun:Satl /jcode:00021397 /dt:20170531 /db:newsdc1"
		Echo "cscript//nologo ZaikoH.vbs 炊飯前月サテ在庫.xls /kubun:Satl /jcode:00023410 /dt:20170531 /db:newsdc1"
		Echo "cscript//nologo ZaikoH.vbs IH前月サテ在庫.xls /kubun:Satl /jcode:00023510 /dt:20170531 /db:newsdc1"
		Echo ""
		Echo "指定オプション"
		Echo "Filename:" & strFileName
		Echo "     /db:" & strDBName
		Echo "  /Kubun:" & strKubun
		Echo "  /JCode:" & strJCode
		Echo "     /Dt:" & strDt
	End Sub
	'-----------------------------------------------------------------------
	'Private 変数
	'-----------------------------------------------------------------------
	Private	strKubun
	Private	strJCode
	Private	strDt
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
		strKubun		= GetOption("Kubun","")
		strJCode		= GetOption("JCode","")
		strDt			= GetOption("Dt","")
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
	'Quit() 強制終了
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
	'Init() オプションチェック
	'-----------------------------------------------------------------------
	Private Function Init()
		Debug ".Init()"
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Echo "Error:オプション:" & strArg
				Disp Init
				Usage
				Quit
			end if
		Next
'		if strFileName = "" then
'			Echo "Error:ファイル未指定."
'			Usage
'			Quit
'		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "dt"
			case "kubun"
			case "jcode"
			case else
				Echo "Error:オプション:" & strArg
				Usage
				Quit
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Init
		OpenDb
		if strFileName <> "" then
			Load
		else
			List
		end if
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'List() 一覧
	'-----------------------------------------------------------------------
    Private Function List()
		Debug ".List()"
		AddSql ""
		AddSql "select"
		AddSql " distinct"
		AddSql " DT"
		AddSql ",Kubun"
		AddSql ",JCode"
		AddSql "from ZaikoH"
		AddWhere "DT",strDt
		AddWhere "Kubun",strKubun
		AddWhere "JCode",strJCode
		set objRs = objDb.Execute(strSql)
		do while objRs.Eof = False
			Line
			DeleteLine
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	Private	objRs
	'-----------------------------------------------------------------------
	'Delete() 1行表示
	'-----------------------------------------------------------------------
    Private Function DeleteLine()
		Debug ".DeleteLine()"
		AddSql ""
		AddSql "delete"
		AddSql "from ZaikoH"
		AddWhere "DT",objRs.Fields("DT")
		AddWhere "Kubun",objRs.Fields("Kubun")
		AddWhere "JCode",objRs.Fields("JCode")
		Write ":Delete:"
		objDb.Execute strSql
		Write RowCount()
	End Function
	'-----------------------------------------------------------------------
	'Line() 1行表示
	'-----------------------------------------------------------------------
    Private Function Line()
		Debug ".Line()"
		Write objRs.Fields("DT")
		Write " " & objRs.Fields("kubun")
		Write " " & objRs.Fields("JCode")
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Private Function Load()
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
			case "レポート 1"
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
		if Trim(objSheet.Name) = "レポート 1" then
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
		lngMaxRow = objSheet.Range("B65535").End(xlUp).Row
		WriteLine objBook.Name & ":" & objSheet.Name & ":" & lngMaxRow
		Delete
		for lngRow = 3 to lngMaxRow
			Write lngRow & "/" & lngMaxRow
			LoadLine
			WriteLine ""
		next
    End Function
	'-------------------------------------------------------------------
	'LoadLine()
	'-------------------------------------------------------------------
    Private Function LoadLine()
		Debug ".LoadLine()"
		dim	strPn
		strPn = GetValue(objSheet.Range("B" & lngRow))
		Write " " & strDt
		Write " " & strKubun
		Write " " & strPn
		dim	objTop
		set objTop = objSheet.Range("C2")
		dim	objQty
		set objQty = objSheet.Range("C" & lngRow)
		do while True
			dim	strSyushi
			strSyushi = GetValue(objTop)
			if strSyushi = "合計" then
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
		AddSql "insert into ZaikoH"
		AddSql "(DT"
		AddSql ",Kubun"
		AddSql ",JCode"
		AddSql ",Pn"
		AddSql ",Syushi"
		AddSql ",Qty"
		AddSql ") values ("
		AddSql " '" & strDt & "'"
		AddSql ",'" & strKubun & "'"
		AddSql ",'" & strJCode & "'"
		AddSql ",'" & strPn & "'"
		AddSql ",'" & strSyushi & "'"
		AddSql "," & strQty
		AddSql ")"
		CallSql strSql
	End Function
	'-------------------------------------------------------------------
	'Delete()
	'-------------------------------------------------------------------
    Private Function Delete()
		Debug ".Delete()"
		Write "delete:" & strDt & " " & strKubun & " " & strJCode
		AddSql ""
		AddSql "delete from ZaikoH"
		AddSql "where DT = '" & strDt & "'"
		AddSql "  and Kubun = '" & strKubun & "'"
		AddSql "  and JCode = '" & strJCode & "'"
		CallSql strSql
		WriteLine ":" & RowCount()
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
	'Sql実行
	'-------------------------------------------------------------------
	Private Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		select case Err.Number
		case -2147467259	'重複
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
	'Where strSql
	'-------------------------------------------------------------------
	Private	Function AddWhere(byVal strF,byval strV)
		if strV = "" then
			exit function
		end if
		if inStr(strSql,"where") > 0 then
			AddSql " and "
		else
			AddSql " where "
		end if
		AddSql strF
		AddSql " = "
		AddSql " '" & strV & "'"
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
	'AbsPath() 絶対パス
	'-------------------------------------------------------------------
	Private	Function AbsPath(byVal strPath)
		dim	objFso
		Set objFso = CreateObject("Scripting.FileSystemObject")
		AbsPath = objFso.GetAbsolutePathName(strPath)
		Set objFso = Nothing
	End Function
	'-------------------------------------------------------------------
	'Excel ファイルオープン
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
