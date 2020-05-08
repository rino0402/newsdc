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
	Wscript.Echo "LsSyuka.vbs [option] <filename>"
	Wscript.Echo " /db:dns"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript LsSyuka.vbs /db:newsdc2 ""I:\0SDC_honsya\事業部別商品化出荷金額まとめ\奈良\ランドリー出荷実績(201507).xls"""
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objLsSyuka
	Set objLsSyuka = New LsSyuka
	if objLsSyuka.Init() <> "" then
		call usage()
		exit function
	end if
	call objLsSyuka.Run()
End Function
'-----------------------------------------------------------------------
'Bo出荷データ変換
'-----------------------------------------------------------------------
Class LsSyuka
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

    Private Sub Class_Initialize
		Call Debug("LsSyuka.Class_Initialize()")
		strDBName = GetOption("db","newsdc")
		set objDB = nothing
		set objRs = nothing
		set objXL = nothing
		set objBk = nothing
		set objSt = nothing
		strFilename = ""
        strFunction = "check"
		strTable = "LsSyuka"
    End Sub

    Private Sub Class_Terminate
		Call Debug("LsSyuka.Class_Terminate()")
'		Call Close()
		if not objBk is nothing then
			call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
    End Sub

    Public Function Init()
		Call Debug("LsSyuka.Init()")
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
		Call Debug("LsSyuka.Run()")
		Call CreateExcelApp()
		Call OpenExcel()
		Call LoadExcel()
	End Function

	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private Function CreateExcelApp()
		Call Debug("LsSyuka.CreateExcelApp()")
		if objXL is nothing then
			Call Debug("	CreateObject(Excel.Application)")
			Set objXL = WScript.CreateObject("Excel.Application")
		end if
	end function

	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private Function OpenExcel()
		Call Debug("LsSyuka.OpenExcel()")
		if objBk is nothing then
			Call Debug("	Workbooks.Open=" & strFilename)
			Set objBk = objXL.Workbooks.Open(strFilename,False,True,,"")
			Call Debug("	    objBk.Path=" & objBk.Path)
			Call Debug("	    objBk.Name=" & objBk.Name)
		end if
	end function
	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Private Function LoadExcel()
		Call Debug("LsSyuka.LoadExcel()")
		if objBk is nothing then
			exit function
		end if
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
	end function

	Private Function LoadXls()
		Call Debug("LsSyuka.LoadXls()")
		if objSt is nothing then
			exit function
		end if
		Call Debug("	objSt.Name=" & objSt.Name)
		if objSt.Range("A1") <> "NO" then
			exit function
		end if

		Call OpenDB()
		Call OpenRs()
		Call DeleteRs()
		lngMaxRow = excelGetMaxRow(objSt,"A",2)
		for lngRow = 2 to lngMaxRow
			dim	strMsg
			strMsg = AddRecord()
			Call DispMsg(lngRow & "/" & lngMaxRow _
						& " " & objSt.Range("E1") _
						& " " & objSt.Range("A" & lngRow) _
						& " " & objSt.Range("B" & lngRow) _
						& " " & objSt.Range("C" & lngRow) _
						& " " & objSt.Range("D" & lngRow) _
						& " " & objSt.Range("E" & lngRow) _
						& " " & objSt.Range("F" & lngRow) _
						& " " & strMsg _
						)
		next
		Call CloseRs()
		Call CloseDB()
	end function

    Private Function OpenDB()
		Call Debug("LsSyuka.OpenDB():" & strDBName)
		set objDb = OpenAdodb(strDBName)
    End Function

    Private Function CloseDB()
		Call Debug("LsSyuka.CloseDB():" & strDBName)
		Call objDb.Close()
		set objDb = Nothing
    End Function

	Private Function OpenRs()
		Call Debug("LsSyuka.OpenRs()")
		dim	strSql
		strSql = "delete from " & strTable
		strSql = strSql & " where YM = '" & objSt.Range("E1") & "'"
		strSql = strSql & "   and JCd = '" & objSt.Range("C2") & "'"
		Call DispMsg("strSql=" & strSql)
		Call objDb.Execute(strSql)
	End Function

	Private Function DeleteRs()
		Call Debug("LsSyuka.DeleteRs()")
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strTable, objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function CloseRs()
		Call Debug("LsSyuka.CloseRs()")
		Call Debug("Table=" & strTable)
		Call objRs.Close()
		set objRs = Nothing
	End Function

	Private Function AddRecord()
		Call Debug("LsSyuka.AddRecord()")
		if objRs is Nothing then
			exit function
		end if
		objRs.AddNew
		objRs.Fields("YM")		= objSt.Range("E1")
		objRs.Fields("No")		= objSt.Range("A" & lngRow)
		objRs.Fields("Soko")	= objSt.Range("B" & lngRow)
		objRs.Fields("JCd")		= objSt.Range("C" & lngRow)
		objRs.Fields("Pn")		= objSt.Range("D" & lngRow)
		objRs.Fields("Qty")		= CLng(objSt.Range("E" & lngRow))
		on error resume next
			objRs.UpdateBatch
			select case Err.Number
			case &h80004005
				AddRecord = "■二重登録■"
				Call objRs.CancelUpdate
			case 0
			case else
				AddRecord = "0x" & Hex(Err.Number) & " " & Err.Description
				Call objRs.CancelUpdate
			end select
		on error goto 0
	End Function

End Class
