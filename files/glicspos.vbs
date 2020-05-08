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
Call Include("debug.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'使用方法
'HMTAH011SCS.dat.20150703-133754
'12345678901234567890123456789012
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "連携データ"
	Wscript.Echo "glicspos.vbs [option] <filename>"
	Wscript.Echo " /db:newsdc"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript glicspos.vbs /db:newsdcnr \\nr\glics\glics_pos\zaiko\hmtah007shu3.dat.20151202-221633"
	Wscript.Echo "cscript glicspos.vbs /db:newsdcnr \\nr\glics\glics_pos\zaiko\00021184zaiko.dat.20151202-2009"
	Wscript.Echo "cscript glicspos.vbs /db:newsdc1 \\w1\glics\glics_pos\outy\HMTAH015SZZ.dat.20160722-151145"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	For Each strArg In WScript.Arguments.UnNamed
    	select case strArg
		case else
			if strFilename <> "" then
				usage()
				Main = 1
				exit Function
			end if
			strFilename = strArg
		end select
	Next
	For Each strArg In WScript.Arguments.Named
    	select case lcase(strArg)
		case "db"
		case "load"
		case "top"
		case "debug"
		case else
			usage()
			Main = 1
			exit Function
		end select
	Next
	select case GetFunction()
	case "load"
		Call Load(strFilename)
	case else
		usage()
		Main = 1
		exit Function
	end select
	Main = 0
End Function

Private Function GetFunction()
	GetFunction = "?"
	if WScript.Arguments.UnNamed.Count > 0 then
		GetFunction = "load"
	end if
End Function

Function Load(byval strFilename)
	dim	objDL
	Set objDL = New DLink

	Call objDL.OpenDB()
	Call objDL.OpenTextFile(strFilename)
	Call objDL.ReadFile()

	Set objDL = nothing

End Function

Class DLink
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	objFSO
	Private	objFile
	Private	strPath
	Private	strFilename
	Private	strBuff
	Private	strTable
	Private	lngCnt
    Private Sub Class_Initialize
		Call Debug("DLink.Class_Initialize()")
		strDBName = GetOption("db","newsdc")
		set objDB = nothing
		set objRs = nothing
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		set objFile = nothing
		strPath = ""
		strFilename = ""
		strBuff = ""
		lngCnt = 0
		strTable = ""
    End Sub

    Private Sub Class_Terminate
		Call Debug("DLink.Class_Terminate()")
		Call Close()
    End Sub

    Public Function OpenDB()
		Call Debug("DLink.OpenDB()")
		set objDb = OpenAdodb(strDBName)
    End Function

    Public Function OpenTextFile(byval strFName)
		Call Debug("DLink.OpenTextFile()")
		strPath = strFName
		Call Debug("Path=" & strPath)
		strFilename = GetFilename(strPath)
		Call Debug("Filename=" & strFilename)
		Set objFile = objFSO.OpenTextFile(strPath, ForReading, False)
		Call SetTableName()
    End Function

	Private Function SetTableName()
		Call Debug("DLink.SetTableName():" & strFilename)

		'Glics 振替実績
		if InStr(lcase(strFilename),"hmem50") = 1 then
			strTable = "hmem500R"
		end if
		'Glics 出荷指示
		if InStr(lcase(strFilename),"hmem70") = 1 then
			strTable = "hmem700R"
		end if
		'Active 出荷実績確認
		if InStr(lcase(strFilename),"hmtac770") = 1 then
			strTable = ""
		end if
		'Active 在庫
		if InStr(lcase(strFilename),"hmtah007") = 1 then
			strTable = "hmtah007R"
		end if
		'Active 出荷指示
		if InStr(ucase(strFilename),"HMTAH011") = 1 then
			strTable = "HMTAH011R"
		end if
		'Active PN
		if InStr(lcase(strFilename),"hmtah012") = 1 then
			strTable = "hmtah012R"
		end if
		'Active 出荷予定(伝発済)
		if InStr(ucase(strFilename),"HMTAH015") = 1 then
			strTable = "HMTAH015_tR"
		end if
		'Active 振替実績
		if InStr(ucase(strFilename),"HMTAH500") = 1 then
			strTable = "HMTAH500R"
		end if
		'Active 在庫
		if InStr(lcase(strFilename),"hmtah007") = 1 then
			strTable = "hmtah007R"
		end if
		'Glics 在庫
		if InStr(lcase(strFilename),"zaiko.dat") = 9 then
			strTable = "GlicsZaikoR"
		end if
		if strTable = "" then
			Call DispMsg("未対応ファイル：" & strFilename)
			SetTableName = -1
		end if
		SetTableName = 0
    End Function

	Private Function OpenRs()
		Call Debug("DLink.OpenRs()")
		Call Debug("Table=" & strTable)
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strTable, objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function DeleteRecord()
		Call Debug("DLink.DeleteRecord()")
		dim	strSql
		strSql = "delete from " & strTable & " where Filename = '" & strFilename & "'"
		Call DispMsg("SQL:" & strSql)
		Call objDb.Execute(strSql)
		Call DispMsg("SQL:Finish")
	End Function

	Private Function AddRecord()
		Call Debug("DLink.AddRecord()")
		if objRs is Nothing then
			Call OpenRs()
		end if
		if strTable = "HMTAH500R" then
			if len(strBuff) > 0 then
				strBuff = left(strBuff,len(strBuff) - 1)
			end if
		end if

		objRs.AddNew
		objRs.Fields("Filename")	= strFilename
		objRs.Fields("Row")			= lngCnt
		objRs.Fields("RecBuff")		= strBuff
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

	Public Function ReadFile()
		Call Debug("DLink.ReadFile()")
		if strTable = "" then
			exit function
		end if
		Call DeleteRecord()
		do while ( objFile.AtEndOfStream = False )
			lngCnt = lngCnt + 1
			strBuff = objFile.ReadLine()
			dim	strMsg
			strMsg = AddRecord()
			Call DispMsg(strFilename & " " & lngCnt & ":" & strTable & "(" & Get_Lenb(strBuff) & ")" & strMsg)
'			Call Debug(strBuff)
		loop
		if lngCnt = 0 then
			strBuff = ""
			strMsg = AddRecord()
			Call DispMsg(strFilename & " " & lngCnt & "(" & Get_Lenb(strBuff) & ")" & strMsg)
		end if
    End Function

    Public Function Close()
		Call Debug("DLink.Close()")
		objFile.Close
		set objFile = nothing
		set objFSO = nothing
		set	objRs = nothing
		set objDB = nothing
	End Function
End Class

