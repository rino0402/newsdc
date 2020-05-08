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
	Wscript.Echo "BoZaikoX.vbs [option] <filename>"
	Wscript.Echo " /db:newsdc9"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript BoZaikoX.vbs /db:newsdc9 在庫ﾃﾞｰﾀ.xls"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objBoZaikoX
	Set objBoZaikoX = New BoZaikoX
	if objBoZaikoX.Init() <> "" then
		call usage()
		exit function
	end if
	call objBoZaikoX.Run()
End Function
'-----------------------------------------------------------------------
'Bo出荷データ変換
'-----------------------------------------------------------------------
Class BoZaikoX
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
		Call Debug("BoZaikoX.Class_Initialize()")
		strDBName = GetOption("db","newsdc9")
		set objDB = nothing
		set objRs = nothing
		set objXL = nothing
		set objBk = nothing
		set objSt = nothing
		strFilename = ""
        strFunction = "check"
		strTable = "BoZaiko"
		strYM = ""
		strJCd = ""
		strSoko = ""
		strDeleteSql = ""
    End Sub

    Private Sub Class_Terminate
		Call Debug("BoZaikoX.Class_Terminate()")
'		Call Close()
		if not objBk is nothing then
			call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
    End Sub

    Public Function Init()
		Call Debug("BoZaikoX.Init()")
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
		Call Debug("BoZaikoX.Run()")
		Call CreateExcelApp()
		Call OpenExcel()
		Call LoadExcel()
	End Function

	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private Function CreateExcelApp()
		Call Debug("BoZaikoX.CreateExcelApp()")
		if objXL is nothing then
			Call Debug("	CreateObject(Excel.Application)")
			Set objXL = WScript.CreateObject("Excel.Application")
		end if
	end function

	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private Function OpenExcel()
		Call Debug("BoZaikoX.OpenExcel()")
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
		Call Debug("BoZaikoX.LoadExcel()")
		if objBk is nothing then
			exit function
		end if
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
	end function

	Private Function DataType()
		Call Debug("BoZaikoX.DataType()")
		Call DispMsg(objSt.Name)
		Call DispMsg( ""_
					& " " & getTitle(objSt.Range("A1")) _
					& " " & getTitle(objSt.Range("B1")) _
					& " " & getTitle(objSt.Range("C1")) _
					& " " & getTitle(objSt.Range("D1")) _
					& " " & getTitle(objSt.Range("E1")) _
					)

		DataType = ""
		strType = ""
		if objSt is nothing then
			exit function
		end if
		'NO	倉庫CD	事業場CD	"資産管理事業場CD"	品目番号	"在庫収支CD"	棚在庫数	"正味引当可能在庫数"	"在庫収支略式名"	在庫収支名

		if	getTitle(objSt.Range("A1")) <> "NO" then
			exit function
		end if
		if	getTitle(objSt.Range("B1")) <> "倉庫CD" then
			exit function
		end if
		if	getTitle(objSt.Range("C1")) <> "事業場CD" then
			exit function
		end if
		if	getTitle(objSt.Range("D1")) <> "資産管理事業場CD" then
			exit function
		end if
		if	getTitle(objSt.Range("E1")) <> "品目番号" then
			exit function
		end if
		strType = "BoZaiko"
		DataType = strType
	end Function

	Private Function LoadXls()
		Call Debug("BoZaikoX.LoadXls()")
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
		Call Debug("BoZaikoX.OpenDB():" & strDBName)
		set objDb = OpenAdodb(strDBName)
    End Function

    Private Function CloseDB()
		Call Debug("BoZaikoX.CloseDB():" & strDBName)
		Call objDb.Close()
		set objDb = Nothing
    End Function

	Private Function OpenRs()
		Call Debug("BoZaikoX.OpenRs()")
	End Function

	Private Function DeleteYM()
		Call Debug("BoZaikoX.DeleteYM()")
		dim	strSql
		strSql = "delete from " & strTable
		Call DispMsg("strSql=" & strSql)
		Call objDb.Execute(strSql)
	End Function

	Private Function DeleteRs()
		Call Debug("BoZaikoX.DeleteRs()")
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strTable, objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function CloseRs()
		Call Debug("BoZaikoX.CloseRs()")
		Call Debug("Table=" & strTable)
		Call objRs.Close()
		set objRs = Nothing
	End Function

	Private Function AddRecord()
		Call Debug("BoZaikoX.AddRecord()")
		if objRs is Nothing then
			exit function
		end if
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
	End Function

	Private Function SetFields()
		strMsg = ""
		strMsg = lngRow & "/" & lngMaxRow _
						& " " & objSt.Name _
						& " " & objSt.Range("A" & lngRow) _
						& " " & objSt.Range("B" & lngRow) _
						& " " & objSt.Range("C" & lngRow) _
						& " " & objSt.Range("D" & lngRow) _
						& " " & objSt.Range("E" & lngRow) _
						& " " & objSt.Range("F" & lngRow) _
						& " " & objSt.Range("G" & lngRow) _
						& " " & objSt.Range("H" & lngRow) _
						& " " & objSt.Range("I" & lngRow) _
						& " " & objSt.Range("K" & lngRow)
		objRs.AddNew
		'Soko	JCode	ShisanJCode	Pn	SyuShi	TanaQty	HikiQty	SyuShiR	SyuShiName
		objRs.Fields("Soko")		= objSt.Range("B" & lngRow)
		objRs.Fields("JCode")		= objSt.Range("C" & lngRow)
		objRs.Fields("ShisanJCode")	= objSt.Range("D" & lngRow)
		objRs.Fields("Pn")			= objSt.Range("E" & lngRow)
		objRs.Fields("SyuShi")		= objSt.Range("F" & lngRow)
		objRs.Fields("TanaQty")		= objSt.Range("G" & lngRow)
		objRs.Fields("HikiQty")		= objSt.Range("H" & lngRow)
		objRs.Fields("SyuShiR")		= objSt.Range("I" & lngRow)
		objRs.Fields("SyuShiName")	= objSt.Range("J" & lngRow)
		objRs.Fields("Loc1")		= RTrim("" & objSt.Range("K" & lngRow))
		SetFields = true
	End Function

End Class

Function getTitle(byVal strT)
	getTitle = Replace(strT,vbLf,"")
End Function
