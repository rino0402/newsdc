Option Explicit
'-----------------------------------------------------------------------
'メイン呼出＆インクルード
'-----------------------------------------------------------------------
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
	dim	strScriptPath
	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strFileName = strScriptPath & strFileName
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
	Set fso = Nothing
End Function
Call Include("const.vbs")
Call Include("get_b.vbs")
Call Include("file.vbs")
Call Include("excel.vbs")
Call Include("debug.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "Bo出荷データ変換"
	Wscript.Echo "LsSyukaBo.vbs [option] <filename>"
	Wscript.Echo " /db:newsdc9"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript LsSyukaBo.vbs /db:newsdc9 ガス石出荷実績数.xls"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objLsSyukaBo
	Set objLsSyukaBo = New LsSyukaBo
	if objLsSyukaBo.Init() <> "" then
		call usage()
		exit function
	end if
	call objLsSyukaBo.Run()
End Function
'-----------------------------------------------------------------------
'Bo出荷データ変換
'-----------------------------------------------------------------------
Class LsSyukaBo
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	objXL
	Private	objBk
	Private	objSt
	Private	strFilename
	Private	strSql
	Private obj103
	Private strFunction
	Private lngRow
	Private lngMaxRow
	Private	strTable
	Private	strYM
	Private	strJCd
	Private	strSoko
	Private	strType
	Private	strDeleteSql
	Private	strColumn
	Private	strMsg

    Private Sub Class_Initialize
		Call Debug("LsSyukaBo.Class_Initialize()")
		strDBName = GetOption("db","newsdc9")
		set objDB = nothing
		set objRs = nothing
		set objXL = nothing
		set objBk = nothing
		set objSt = nothing
		strFilename = ""
        strFunction = "check"
		strTable = "LsSyuka"
		strYM = ""
		strJCd = ""
		strSoko = ""
		strDeleteSql = ""
    End Sub

    Private Sub Class_Terminate
		Call Debug("LsSyukaBo.Class_Terminate()")
'		Call Close()
		if not objBk is nothing then
			call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
    End Sub

    Public Function Init()
		Call Debug("LsSyukaBo.Init()")
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
	    	select case strArg
			case else
				if strFilename <> "" then
					Init = "filename error"
					exit Function
				end if
				strFilename = strArg
				Call Debug("strFilename=" & strFilename)
			end select
		Next
		if strFilename = "" then
			Init = "filename error"
			exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "c","check"
                strFunction = "check"
			case "i","info"
                strFunction = "info"
			case else
				Init = "unknown option:" & strArg
				Exit Function
			end select
		Next
	End Function

    Public Function Run()
		Call Debug("LsSyukaBo.Run()")
		Call CreateExcelApp()
		Call OpenExcel()
		Call LoadExcel()
	End Function

	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private Function CreateExcelApp()
		Call Debug("LsSyukaBo.CreateExcelApp()")
		if objXL is nothing then
			Call Debug("	CreateObject(Excel.Application)")
			Set objXL = WScript.CreateObject("Excel.Application")
		end if
	end function

	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private Function OpenExcel()
		Call Debug("LsSyukaBo.OpenExcel()")
		if objBk is nothing then
			Call Debug("	Workbooks.Open=" & GetAbsPath(strFilename))
			Set objBk = objXL.Workbooks.Open(GetAbsPath(strFilename),False,True,,"")
			Call Debug("	    objBk.Path=" & objBk.Path)
			Call Debug("	    objBk.Name=" & objBk.Name)
		end if
	end function
	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Private Function LoadExcel()
		Call Debug("LsSyukaBo.LoadExcel()")
		if objBk is nothing then
			exit function
		end if
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
	end function

	Private Function DataType()
		Call Debug("LsSyukaBo.DataType()")
		Call DispMsg(objSt.Name)
		Call DispMsg( ""_
					& " " & objSt.Range("A1") _
					& " " & objSt.Range("B1") _
					& " " & objSt.Range("C1") _
					& " " & objSt.Range("D1") _
					& " " & objSt.Range("E1") _
					)

		DataType = ""
		strType = ""
		if objSt is nothing then
			exit function
		end if
		'NO	倉庫CD	事業場CD	品目番号
		if	objSt.Range("A1") <> "NO" then
			exit function
		end if
		if	objSt.Range("B1") <> "倉庫CD" then
			exit function
		end if
		strSoko = objSt.Range("B2")
		if	objSt.Range("C1") <> "事業場CD" then
			exit function
		end if
		strJCd = objSt.Range("C2")
		if	objSt.Range("D1") <> "品目番号" then
			exit function
		end if
		strType = "BoSyuka"
		DataType = strType
	end Function

	Private Function LoadXls()
		Call Debug("LsSyukaBo.LoadXls()")
		if objSt is nothing then
			exit function
		end if
		Call Debug("	objSt.Name=" & objSt.Name)
		if DataType() = "" then
			exit function
		end if
		Call OpenDB()
		Call OpenRs()
		Call DeleteRs()
		Call DeleteYM()
		lngMaxRow = excelGetMaxRow(objSt,"A",2)
		for lngRow = 2 to lngMaxRow
			Call AddRecord()
		next
		Call CloseRs()
		Call CloseDB()
	end function

    Private Function OpenDB()
		Call Debug("LsSyukaBo.OpenDB():" & strDBName)
		set objDb = OpenAdodb(strDBName)
    End Function

    Private Function CloseDB()
		Call Debug("LsSyukaBo.CloseDB():" & strDBName)
		Call objDb.Close()
		set objDb = Nothing
    End Function

	Private Function OpenRs()
		Call Debug("LsSyukaBo.OpenRs()")
	End Function

	Private Function DeleteYM()
		Call Debug("LsSyukaBo.DeleteYM()")
		strColumn = "E"
		do while CheckYM()
			dim	strSql
			strSql = "delete from " & strTable
			strSql = strSql & " where YM = '" & strYM & "'"
			strSql = strSql & "   and JCd = '" & strJCd & "'"
			Call DispMsg("strSql=" & strSql)
			Call objDb.Execute(strSql)
			strColumn = excelNextColumn(objSt,strColumn)
		loop
	End Function

	Private Function DeleteRs()
		Call Debug("LsSyukaBo.DeleteRs()")
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strTable, objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function CloseRs()
		Call Debug("LsSyukaBo.CloseRs()")
		Call Debug("Table=" & strTable)
		Call objRs.Close()
		set objRs = Nothing
	End Function

	Private Function AddRecord()
		Call Debug("LsSyukaBo.AddRecord()")
		if objRs is Nothing then
			exit function
		end if
		strColumn = "E"
		do while CheckYM()
			if SetFields() then
				on error resume next
					objRs.UpdateBatch
					select case Err.Number
					case &h80004005
						strMsg = strMsg & "■二重登録■"
						Call objRs.CancelUpdate
					case 0
					case else
						strMsg = strMsg & "0x" & Hex(Err.Number) & " " & Err.Description
						Call objRs.CancelUpdate
					end select
				on error goto 0
				Call DispMsg(strMsg)
			end if
			strColumn = excelNextColumn(objSt,strColumn)
		loop
	End Function

	Private Function SetFields()
		strMsg = ""
		dim	strQty
		strQty = objSt.Range(strColumn & lngRow)
		if isNumeric(strQty) = false then
			SetFields = false
			exit function
		end if
		if CLng(strQty) = 0 then
			SetFields = false
			exit function
		end if
		objRs.AddNew
		objRs.Fields("YM")		= strYM
		objRs.Fields("No")		= objSt.Range("A" & lngRow)
		objRs.Fields("Soko")	= objSt.Range("B" & lngRow)
		objRs.Fields("JCd")		= objSt.Range("C" & lngRow)
		objRs.Fields("Pn")		= objSt.Range("D" & lngRow)
		objRs.Fields("Qty")		= CLng(strQty)
		strMsg = lngRow & "/" & lngMaxRow _
						& " " & objSt.Name _
						& " " & strYM _
						& " " & objSt.Range("A" & lngRow) _
						& " " & objSt.Range("B" & lngRow) _
						& " " & objSt.Range("C" & lngRow) _
						& " " & objSt.Range("D" & lngRow) _
						& " " & CLng(strQty)

		SetFields = true
	End Function

	Private Function CheckYM()
		strYM = objSt.Range(strColumn & "1")
		Call Debug("LsSyukaBo.CheckYM():" & strColumn & ":" & strYM)
		if IsNumeric(strYM) = false then
			CheckYM = false
			exit function
		end if
		CheckYM = True
	End Function
End Class

Function excelNextColumn(objSt,byVal strColumn)
	dim	strNextColumn
	strNextColumn = objSt.Range(strColumn & "1").Offset(0,1).Address
	strNextColumn = Split(strNextColumn,"$")(1)
	excelNextColumn = strNextColumn
End Function
