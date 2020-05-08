Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "NonConf.vbs [option] <*.xlsx>"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo NonConf.vbs I:\pos\不適合\不適合w3.xlsm"
	Wscript.Echo "cscript//nologo NonConf.vbs \\w3\y\不適合管理_物流\不適合.xls /db:newsdc3"
	Wscript.Echo
	Wscript.Echo "cscript//nologo NonConf.vbs \\w1\doc\◆男性\松井\不適合\IH2016年度.xlsx /db:newsdc1"
	Wscript.Echo "cscript//nologo NonConf.vbs \\w1\doc\◆男性\松井\不適合\BL2016年度.xlsx /db:newsdc1"
	Wscript.Echo "cscript//nologo NonConf.vbs \\w1\doc\◆男性\松井\不適合\炊回2016年度.xlsx /db:newsdc1"
	Wscript.Echo
	Wscript.Echo "cscript//nologo NonConf.vbs \\w4\manager\2011.01月ｓｄｃpubから移動したファイル\ISOデータ\入荷時不適合データ集計ｴｱｺﾝ (回復済み).xls /db:newsdc4"
	Wscript.Echo "cscript//nologo NonConf.vbs \\w4\manager\2011.01月ｓｄｃpubから移動したファイル\ISOデータ\入荷時不適合ﾃﾞ-ﾀ集計（冷蔵庫）.xlsx /db:newsdc4"
End Sub
'-----------------------------------------------------------------------
'Excel
'2017.03.07 新規
'-----------------------------------------------------------------------
Const xlUp = -4162

Class Excel
	Private	strDBName
	Private	objDB
	Private	strFileName
	Private	objExcel
	Private	objBook
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		set	objExcel = nothing
		set	objBook = nothing
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set	objBook = nothing
		set	objExcel = nothing
		set objDB = nothing
		strDBName = GetOption("db"	,"newsdc")
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
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
		Call InsertCsv()
		Call CloseBook()
	End Function
	'-------------------------------------------------------------------
	'Update
	'-------------------------------------------------------------------
	Private	lngTopRow
	Private	intCol
	Private	strMaxCol
    Private Function InsertCsv()
		Debug ".InsertCsv()"
		for each objSheet in objBook.Worksheets
			Debug ".InsertCsv():" & objBook.Name & "_" & objSheet.Name
			select case objSheet.Name
			case "保存データ"
				lngTopRow = 11
				intCol = 34
				strMaxCol	= "C"
				LoadSheet
				exit for
			case "不適合内容"
				lngTopRow = 1
				intCol = 19
				strMaxCol	= "A"
				LoadSheet
				exit for
			case else
				dim	strSheetName
				strSheetName = Replace(objSheet.Name,".","")
				if Len(strSheetName) = 6 and isNumeric(strSheetName) = True then
					Debug ".InsertCsv() 冷蔵庫:" & strSheetName
					for lngTopRow = 10 to 100
						if objSheet.Range("A" & lngTopRow) = "発生月" then
							exit for
						end if
					next
					if lngTopRow <> 100 then
						intCol = 8
						strMaxCol	= "A"
						LoadSheet
					end if
				end if
			end select
		next
    End Function
	'-------------------------------------------------------------------
	'LoadSheet
	'-------------------------------------------------------------------
	Private	lngMaxRow
	Private	lngRow
    Private Function LoadSheet()
		Debug ".LoadSheet():" & objBook.Name & "_" & objSheet.Name
		lngMaxRow = objSheet.Range(strMaxCol & "65535").End(xlUp).Row
		Debug ".LoadSheet():MaxRow=" & lngMaxRow
		DeleteSheet
		for lngRow = lngTopRow to lngMaxRow
			Debug ".LoadSheet():" _
					& lngRow & "/" & lngMaxRow _
					& " " & GetValue(objSheet.Range("A" & lngRow)) _
					& " " & GetValue(objSheet.Range("B" & lngRow)) _
					& " " & GetValue(objSheet.Range("C" & lngRow)) _
					& " " & GetValue(objSheet.Range("D" & lngRow))
			InsertSql
		next
    End Function
	'-------------------------------------------------------------------
	'DeleteSheet
	'-------------------------------------------------------------------
    Private Function DeleteSheet()
		Debug ".DeleteSheet()"
		Wscript.StdOut.Write objBook.Name & "_" & objSheet.Name & ":削除中..."

		AddSql ""
		AddSql "delete from CsvTemp"
		AddSql " where Filename = '" & objBook.Name & "_" & objSheet.Name & "'"
		Wscript.StdOut.Write ":" & strSql
		CallSql strSql
		Wscript.StdOut.WriteLine
    End Function
	'-------------------------------------------------------------------
	'InsertSql
	'-------------------------------------------------------------------
    Private Function InsertSql()
		Debug ".InsertSql()"
		Wscript.StdOut.Write objBook.Name & " " & objSheet.Name & ":" & lngRow & "/" & lngMaxRow

		AddSql ""
		AddSql "insert into CsvTemp"
		AddSql "(Filename"
		AddSql ",Row"
		dim	i
		for	i = 1 to intCol
			AddSql ",Col" & right("00" & i,2)
		next
		AddSql ",Col"
		AddSql ") values ("
		AddSql "'" & objBook.Name & "_" & objSheet.Name & "'"
		AddSql "," & lngRow
		dim	objRange
		set objRange = objSheet.Range("A" & lngRow)
		for	i = 1 to intCol
			AddSql ",'" & GetValue(objSheet.Range(objRange.Address)) & "'"
			set objRange = objRange.Offset(0,1)
		next
		AddSql "," & intCol
		AddSql ")"
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
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		on error goto 0
		if Err.Number <> 0 then
			Wscript.StdOut.WriteLine ".CallSql():0x" & Hex(Err.Number) & ":" & Err.Description
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end if
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
	'Excelの準備
	'-------------------------------------------------------------------
	Private Function CreateExcel()
		Debug ".CreateExcel()"
		if objExcel is nothing then
			Debug ".CreateExcel():CreateObject(Excel.Application)"
			Set objExcel = WScript.CreateObject("Excel.Application")
		end if
	end function
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
			Debug ".OpenBook().Open:" & strBkName
			Wscript.StdOut.Write strBkName & " :"
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
			Wscript.StdOut.WriteLine objBook.Name
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
	dim	objExcel
	Set objExcel = New Excel
	if objExcel.Init() <> "" then
		call usage()
		exit function
	end if
	call objExcel.Run()
End Function
