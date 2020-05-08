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
	Wscript.Echo "作業ログ初回処理"
	Wscript.Echo "p_sagyo_sum.vbs [option]"
	Wscript.Echo " /db:newsdc"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript p_sagyo_sum.vbs"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objPSagyoSum
	Set objPSagyoSum = New PSagyoSum
	if objPSagyoSum.Init() <> "" then
		call usage()
		exit function
	end if
	call objPSagyoSum.Run()
End Function
'-----------------------------------------------------------------------
'Bo出荷データ変換
'-----------------------------------------------------------------------
Class PSagyoSum
	Private	strDBName
	Private	objDB
	Private	objSrcRs
	Private	objDstRs
	Private	strSql
	Private strFunction
	Private	strMsg

    Private Sub Class_Initialize
		Call Debug("PSagyoSum.Class_Initialize()")
		strDBName = GetOption("db","newsdc9")
		set objDB = nothing
        strFunction = "check"
    End Sub

    Private Sub Class_Terminate
		Call Debug("PSagyoSum.Class_Terminate()")
'		Call Close()
    End Sub

    Public Function Init()
		Call Debug("PSagyoSum.Init()")
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
	    	select case strArg
			case else
				Init = "option error"
			end select
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
		Call Debug("PSagyoSum.Run()")
		Call Load()
	End Function

	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Private Function Load()
		Call Debug("PSagyoSum.Load()")
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
		Call Debug("PSagyoSum.OpenDB():" & strDBName)
		set objDb = OpenAdodb(strDBName)
    End Function

    Private Function CloseDB()
		Call Debug("PSagyoSum.CloseDB():" & strDBName)
		Call objDb.Close()
		set objDb = Nothing
    End Function

	Private Function OpenRs()
		Call Debug("PSagyoSum.OpenRs()")
		strSql = "select * from P_SAGYO_LOG where JITU_DT >= (select max(JITU_DT) from P_SAGYO_SUM) order by JITU_DT,JITU_TM"
		Call DispMsg(strSql)
		objDb.CommandTimeout = 0
		Set objSrcRs = objDb.Execute(strSql)
		Call DispMsg("...finish")
		Set objDstRs = Wscript.CreateObject("ADODB.Recordset")
		objDstRs.Open "P_SAGYO_SUM", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function DeleteRs()
		Call Debug("PSagyoSum.DeleteRs()")
	End Function

	Private Function CloseRs()
		Call Debug("PSagyoSum.CloseRs()")
		Call objSrcRs.Close()
		set objSrcRs = Nothing
		Call objDstRs.Close()
		set objDstRs = Nothing
	End Function

	Private Function AddRecord()
		Call Debug("PSagyoSum.AddRecord()")
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
		strMsg = ""
		strMsg = strMsg & objSrcRs.Fields("JITU_DT")
		strMsg = strMsg & " " & objSrcRs.Fields("JITU_TM")
		strMsg = strMsg & " " & objSrcRs.Fields("JGYOBU")
		strMsg = strMsg & " " & objSrcRs.Fields("NAIGAI")
		strMsg = strMsg & " " & objSrcRs.Fields("HIN_GAI")
		strMsg = strMsg & " " & objSrcRs.Fields("RIRK_ID")
		if RTrim(objSrcRs.Fields("HIN_GAI")) = "" then
			exit function
		end if
		objDstRs.AddNew
		dim	objF
		for each objF in objSrcRs.Fields
			objDstRs.Fields(objF.Name) = objF
		next
		SetFields = true
	End Function

End Class
