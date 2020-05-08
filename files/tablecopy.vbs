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
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "テーブルコピー"
	Wscript.Echo "tablecopy.vbs <テーブル名> <dns 元> <dns 先> [option]"
	Wscript.Echo " /?"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript tablecopy.vbs item newsdc4 newsdc6 ""JGYOBU='A' and NAIGAI='1'"""
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objTableCopy
	Set objTableCopy = New TableCopy
	if objTableCopy.Init() <> "" then
		call usage()
		exit function
	end if
	call objTableCopy.Run()
End Function
'-----------------------------------------------------------------------
'Bo出荷データ変換
'-----------------------------------------------------------------------
Class TableCopy
	Private	strTableName
	Private	strSrcDBName
	Private	strDstDBName
	Private	objSrcDB
	Private	objDstDB
	Private	objSrcRs
	Private	objDstRs
	Private	strWhere
	Private	strSql
	Private strFunction
	Private	strMsg

    Private Sub Class_Initialize
		Call Debug("TableCopy.Class_Initialize()")
		strDstDBName = ""
		strSrcDBName = ""
		set objSrcDB = Nothing
		set objDstDB = Nothing
        strFunction = "check"
    End Sub

    Private Sub Class_Terminate
		Call Debug("TableCopy.Class_Terminate()")
'		Call Close()
    End Sub

    Public Function Init()
		Call Debug("TableCopy.Init()")
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			if		strTableName = "" then
				 strTableName	= strArg
			elseif	strSrcDBName = "" then
				 strSrcDBName	= strArg
			elseif	strDstDBName = "" then
				 strDstDBName	= strArg
			elseif	strWhere	 = "" then
				 strWhere		= strArg
			else
				Init = "option error:" & strArg
				Exit Function
			end if
		Next
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
		Call Debug("TableCopy.Run()")
		Call Load()
	End Function

	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Private Function Load()
		Call Debug("TableCopy.Load()")
		Call OpenDB()
		Call OpenRs()
		do while objSrcRs.EOF = false
			Call AddRecord()
			objSrcRs.MoveNext
		loop
		Call CloseRs()
		Call CloseDB()
	end function

    Private Function OpenDB()
		Call Debug("TableCopy.OpenDB():" & strSrcDBName & ":" & strDstDBName)
		set objSrcDb = OpenAdodb(strSrcDBName)
		set objDstDb = OpenAdodb(strDstDBName)
    End Function

    Private Function CloseDB()
		Call Debug("TableCopy.CloseDB()")
		Call objSrcDb.Close()
		set objSrcDb = Nothing
		Call objDstDb.Close()
		set objDstDb = Nothing
    End Function

	Private Function OpenRs()
		Call Debug("TableCopy.OpenRs()")
		if strWhere <> "" then
			strWhere = " where " & strWhere
		end if
		strSql = "select * from " & strTableName & strWhere
		Call DispMsg(strSql)
		objSrcDb.CommandTimeout = 0
		Set objSrcRs = objSrcDb.Execute(strSql)
		Call DispMsg("...finish")
		Set objDstRs = Wscript.CreateObject("ADODB.Recordset")
		objDstRs.Open strTableName, objDstDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function DeleteRs()
		Call Debug("TableCopy.DeleteRs()")
	End Function

	Private Function CloseRs()
		Call Debug("TableCopy.CloseRs()")
		Call objSrcRs.Close()
		set objSrcRs = Nothing
		Call objDstRs.Close()
		set objDstRs = Nothing
	End Function

	Private Function AddRecord()
		Call Debug("TableCopy.AddRecord()")
		if SetFields() then
			on error resume next
				objDstRs.UpdateBatch
				select case Err.Number
				case &h80004005
					strMsg = strMsg & "■二重登録■"
					Call objDstRs.CancelUpdate
				case 0
				case else
					strMsg = strMsg & "0x" & Hex(Err.Number) & " " & Err.Description
					Call objDstRs.CancelUpdate
				end select
			on error goto 0
			Call DispMsg(strMsg)
		end if
	End Function

	Private Function SetFields()
		SetFields = false
		Call SetMsg()
		objDstRs.AddNew
		dim	objF
		for each objF in objSrcRs.Fields
			on error resume next
				objDstRs.Fields(objF.Name) = objF
				select case Err.Number
				case 0
				case else
					call DispMsg(objF.Name & ":" & objF)
					call DispMsg("0x" & Hex(Err.Number) & " " & Err.Description)
				end select
			on error goto 0
		next
		SetFields = true
	End Function

	Private Function SetField()
		objDstRs.Fields(objF.Name) = RTrim(objF)
	End Function

	Private Function SetMsg()
		strMsg = ""
		select case lcase(strTableName)
		case "item"
			strMsg = strMsg & " " & objSrcRs.Fields("JGYOBU")
			strMsg = strMsg & " " & objSrcRs.Fields("NAIGAI")
			strMsg = strMsg & " " & objSrcRs.Fields("HIN_GAI")
		case else
			strMsg = strMsg & " " & objSrcRs.Fields(0)
			strMsg = strMsg & " " & objSrcRs.Fields(1)
			strMsg = strMsg & " " & objSrcRs.Fields(2)
			strMsg = strMsg & " " & objSrcRs.Fields(3)
			strMsg = strMsg & " " & objSrcRs.Fields(4)
		end select
	End Function

End Class
