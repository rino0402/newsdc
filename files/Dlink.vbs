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
	Wscript.Echo "dlink.vbs [option] <filename>"
	Wscript.Echo " /db:newsdc"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript dlink.vbs /db:newsdc5 \\192.168.5.31\newsdc\files\fukutsu\fukutsu-20151210.txt"
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
		if InStr(lcase(strFilename),"fukutsu") = 1 then
			strTable = "FukutsuR"
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

