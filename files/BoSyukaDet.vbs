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
	Wscript.Echo "Bo詳細出荷データ変換"
	Wscript.Echo "BoSyukaDet.vbs [option] <filename> [sheetname]"
	Wscript.Echo " /db:newsdc9"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript BoSyukaDet.vbs /db:newsdc9 bo\◆【実績明細】過去・当月_2015.xlsx 当月以前実績"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objBoSyukaDet
	Set objBoSyukaDet = New BoSyukaDet
	if objBoSyukaDet.Init() <> "" then
		call usage()
		exit function
	end if
	call objBoSyukaDet.Run()
End Function
'-----------------------------------------------------------------------
'Bo出荷データ変換
'-----------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Class BoSyukaDet
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	objXL
	Private	objBk
	Private	objSt
	Private	strFilename
	Private	strSheetname
	Private	strSql
	Private strFunction
	Private lngRow
	Private lngMaxRow
	Private	strTable
	Private	strNewYm
	Private	strCurYm
	Private	strType
	Private	strDeleteSql
	Private	strColumn
	Private	strMsg

    Private Sub Class_Initialize
		Call Debug("BoSyukaDet.Class_Initialize()")
		strDBName = GetOption("db","newsdc9")
		set objDB = nothing
		set objRs = nothing
		set objXL = nothing
		set objBk = nothing
		set objSt = nothing
		strFilename = ""
		strSheetname = ""
        strFunction = "check"
		strTable = "BoSyukaDet"
		strNewYm = ""
		strCurYm = ""
		strDeleteSql = ""
    End Sub

    Private Sub Class_Terminate
		Call Debug("BoSyukaDet.Class_Terminate()")
'		Call Close()
		if not objBk is nothing then
			call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
    End Sub

    Public Function Init()
		Call Debug("BoSyukaDet.Init()")
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
	    	select case strArg
			case else
				if strFilename = "" then
					strFilename = strArg
					Call Debug("strFilename=" & strFilename)
				elseif strSheetname = "" then
					strSheetname = strArg
					Call Debug("strSheetname=" & strFilename)
				else
					Init = "option error"
					exit Function
				end if
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

	'-------------------------------------------------------------------
	'メイン処理
	'-------------------------------------------------------------------
    Public Function Run()
		Call Debug("BoSyukaDet.Run()")
		select case FileType()
		case "excel"
			Call CreateExcelApp()
			Call OpenExcel()
			Call LoadExcel()
		case "csv"
			Call OpenCsv()
			Call LoadCsv()
			Call CloseCsv()
		end select
	End Function

	'-------------------------------------------------------------------
	' csv
	'-------------------------------------------------------------------
	Private	objFSO
	Private	objFile
	'-------------------------------------------------------------------
	' csv Open
	'-------------------------------------------------------------------
	Private Function OpenCsv()
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	End Function
	'-------------------------------------------------------------------
	' csv Load
	'-------------------------------------------------------------------
	Private	strBuff
	Private	aryBuff
	Private bStop
	Private Function LoadCsv()
		Call OpenDB()
		Call OpenRs()
		lngRow = 0
		bStop = False
		do while ( objFile.AtEndOfStream = False )
			strBuff = objFile.ReadLine()
			aryBuff = GetTab(strBuff)
			lngRow = lngRow + 1
			if lngRow <> 1 then
				Call AddRecord()
			end if
			if bStop then
				exit do
			end if
		loop
		Call CloseRs()
		Call CloseDB()
	End Function
	Function GetTab(ByVal s)
	    Dim r
		r = Split(s,vbTab)
		GetTab = r
	End Function
	'-------------------------------------------------------------------
	' csv Close
	'-------------------------------------------------------------------
	Private Function CloseCsv()
		objFile.Close
		set objFile = nothing
		set objFSO = nothing
	End Function
	'-------------------------------------------------------------------
	'ファイルの種類
	'-------------------------------------------------------------------
	Private Function FileType()
		FileType = ""
		select case lcase(fileExt(strFilename))
		case "xls","xlsx"	FileType = "excel"
		case "csv"			FileType = "csv"
		end select
		Call Debug("BoSyukaDet.FileType():" & FileType)
	End Function
	'-------------------------------------------------------------------
	'拡張子
	'-------------------------------------------------------------------
	Private Function fileExt(byVal f)
		dim	fobj
		set fobj = CreateObject("Scripting.FileSystemObject")
		dim	strExt
		strExt = fobj.GetextensionName(f)
		set fobj = Nothing
		fileExt = strExt
	End Function

	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private Function CreateExcelApp()
		Call Debug("BoSyukaDet.CreateExcelApp()")
		if objXL is nothing then
			Call Debug("	CreateObject(Excel.Application)")
			Set objXL = WScript.CreateObject("Excel.Application")
		end if
	end function

	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private Function OpenExcel()
		Call Debug("BoSyukaDet.OpenExcel()")
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
		Call Debug("BoSyukaDet.LoadExcel()")
		if objBk is nothing then
			exit function
		end if
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
	end function

	Private Function DataType()
		Call Debug("BoSyukaDet.DataType()")
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
		'NO	"受注出荷管理番号"	事業場CD	"在庫収支"	伝票番号	出荷先CD	出荷先名	相手先CD	相手先名	注文区分	"売上実績年月日"	品目番号	"出荷実績数"

		if	getTitle(objSt.Range("A1")) <> "NO" then
			exit function
		end if
		if	getTitle(objSt.Range("B1")) <> "受注出荷管理番号" then
			exit function
		end if
		if	getTitle(objSt.Range("C1")) <> "事業場CD" then
			exit function
		end if
		if	getTitle(objSt.Range("D1")) <> "在庫収支" then
			exit function
		end if
		if	getTitle(objSt.Range("E1")) <> "伝票番号" then
			exit function
		end if
		strType = "BoSyukaDet"
		DataType = strType
	end Function

	Private Function LoadXls()
		Call Debug("BoSyukaDet.LoadXls()")
		if objSt is nothing then
			exit function
		end if
		Call Debug("	objSt.Name=" & objSt.Name)
		if strSheetname <> "" then
			if strSheetname <> objSt.Name then
				exit function
			end if
		end if
		if DataType() = "" then
			exit function
		end if
		Call OpenDB()
		Call OpenRs()
		lngMaxRow = excelGetMaxRow(objSt,"A",2)
		for lngRow = 2 to lngMaxRow
			Call AddRecord()
		next
		Call CloseRs()
		Call CloseDB()
	end function

    Private Function OpenDB()
		Call Debug("BoSyukaDet.OpenDB():" & strDBName)
		set objDb = OpenAdodb(strDBName)
    End Function

    Private Function CloseDB()
		Call Debug("BoSyukaDet.CloseDB():" & strDBName)
		Call objDb.Close()
		set objDb = Nothing
    End Function

	Private Function OpenRs()
		Call Debug("BoSyukaDet.OpenRs()")
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strTable, objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function DeleteYM(byVal strDt)
		Call Debug("BoSyukaDet.DeleteYM():" & strDt)
		strNewYm = left(strDt,6)
		if strCurYm = strNewYm then
			exit function
		end if
		strSql = "delete from " & strTable & " where JisekiDt like '" & strNewYm & "%'"
		Call DispMsg("Execute:" & strSql)
		objDb.CommandTimeout = 0
		Call objDb.Execute(strSql)
		strCurYm = strNewYm
	End Function

	Private Function DeleteRs()
		Call Debug("BoSyukaDet.DeleteRs()")
	End Function

	Private Function CloseRs()
		Call Debug("BoSyukaDet.CloseRs()")
		Call Debug("Table=" & strTable)
		Call objRs.Close()
		set objRs = Nothing
	End Function

	Private Function AddRecord()
		Call Debug("BoSyukaDet.AddRecord()")
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
					bStop = True
				end select
			on error goto 0
			Call DispMsg(strMsg)
		end if
	End Function

	Private Function CsvTrim(byval c)
		if left(c,1) = """" then
			if right(c,1) = """" then
				c = Right(c,Len(c) -1 )
				c = Left(c,Len(c) -1 )
			end if
		end if
		CsvTrim = c
	End Function

	Private Function SetFields()
		SetFields = false
		if FileType() = "csv" then 
			strMsg = ""
			strMsg = strMsg & " " & CsvTrim(aryBuff( 0))
			strMsg = strMsg & " " & CsvTrim(aryBuff( 1))
			strMsg = strMsg & " " & CsvTrim(aryBuff( 2))
			strMsg = strMsg & " " & CsvTrim(aryBuff( 3))
			strMsg = strMsg & " " & CsvTrim(aryBuff( 4))
			strMsg = strMsg & " " & CsvTrim(aryBuff( 7))
			if CsvTrim(aryBuff( 0)) = "受注出実_受注出荷管理番号" then
				bStop = True
				SetFields = false
				Exit Function
			end if
			objRs.AddNew
'			Call SetField("No"		,aryBuff( 0))
			Call SetField("IDNo"	,CsvTrim(aryBuff( 0)))	'"受注出荷過実_受注出荷管理番号"
			Call SetField("JCode"	,CsvTrim(aryBuff( 1)))	'"受注出荷過実_資産管理事業場コード"
			Call SetField("Syushi"	,CsvTrim(aryBuff( 2)))	'"受注出荷過実_在庫収支コード"
			Call SetField("DenNo"	,CsvTrim(aryBuff( 3)))	'"受注出荷過実_伝票番号"
			Call SetField("SyukaCd"	,CsvTrim(aryBuff( 4)))	'"受注出荷過実_得意先コード(相手先CD)"
			Call SetField("SyukaNm"	,CsvTrim(aryBuff( 5)))   '"受注出荷過実_得意先略称(相手先名)"
			Call SetField("ChuKb"	,CsvTrim(aryBuff( 6)))   '"受注出荷過実_注文区分"
			Call SetField("JisekiDt",CsvTrim(aryBuff( 7)))   '"受注出荷過実_売上実績年月日"
			Call SetField("Pn"		,CsvTrim(aryBuff( 8)))   '"受注出荷過実_品目番号"
			Call SetField("Qty"		,CsvTrim(aryBuff( 9)))   '"受注出荷過実_出荷実績数"
			Call SetField("AiteCd"	,CsvTrim(aryBuff(10)))   '"受注出荷過実_直送相手先コード"
'			Call SetField("AiteNm"	,aryBuff(12))
			Call DeleteYm(CsvTrim(aryBuff(7)))
			SetFields = true
			Exit Function
		end if

		Call DeleteYm(objSt.Range("K" & lngRow))
		strMsg = ""
		strMsg = lngRow & "/" & lngMaxRow _
						& " " & objSt.Name _
						& " " & objSt.Range("A" & lngRow) _
						& " " & objSt.Range("B" & lngRow) _
						& " " & objSt.Range("C" & lngRow) _
						& " " & objSt.Range("K" & lngRow) _
						& " " & objSt.Range("L" & lngRow) _
						& " " & objSt.Range("M" & lngRow)
		objRs.AddNew
		Call SetField("No"		,objSt.Range("A" & lngRow))
		Call SetField("IDNo"	,objSt.Range("B" & lngRow))
		Call SetField("JCode"	,objSt.Range("C" & lngRow))
		Call SetField("Syushi"	,objSt.Range("D" & lngRow))
		Call SetField("DenNo"	,objSt.Range("E" & lngRow))
		Call SetField("SyukaCd"	,objSt.Range("F" & lngRow))
		Call SetField("SyukaNm"	,objSt.Range("G" & lngRow))
		Call SetField("AiteCd"	,objSt.Range("H" & lngRow))
		Call SetField("AiteNm"	,objSt.Range("I" & lngRow))
		Call SetField("ChuKb"	,objSt.Range("J" & lngRow))
		Call SetField("JisekiDt",objSt.Range("K" & lngRow))
		Call SetField("Pn"		,objSt.Range("L" & lngRow))
		Call SetField("Qty"		,objSt.Range("M" & lngRow))
		SetFields = true
	End Function
	Private Function SetField(byVal strFName,byVal strV)
		strV = RTrim(strV)
		Call Debug("BoSyukaDet.SetField()" & strFName & "=" & strV)
		on error resume next
			objRs.Fields(strFName) = strV
			select case Err.Number
			case 0
			case else
				Call DispMsg(strFName & ":" & strV & ":" & "0x" & Hex(Err.Number) & " " & Err.Description)
				strMsg = strMsg & "0x" & Hex(Err.Number) & " " & Err.Description
				objRs.Fields(strFName) = ""
				bStop = True
			end select
		on error goto 0
	End Function

End Class

Function getTitle(byVal strT)
	getTitle = Replace(strT,vbLf,"")
End Function
