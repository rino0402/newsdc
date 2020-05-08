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
	Wscript.Echo "LsSyukaAc.vbs [option] <filename>"
	Wscript.Echo " /db:newsdc9"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript LsSyukaAc.vbs /db:newsdc9"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objLsSyukaAc
	Set objLsSyukaAc = New LsSyukaAc
	if objLsSyukaAc.Init() <> "" then
		call usage()
		exit function
	end if
	call objLsSyukaAc.Run()
End Function
'-----------------------------------------------------------------------
'Bo出荷データ変換
'-----------------------------------------------------------------------
Class LsSyukaAc
	Private	strDBName
	Private	objDB
	Private	objDstRs
	Private	objSrcRs
	Private	strSql
	Private strFunction
	Private lngRow
	Private	strTable
	Private	strYM
	Private	strDeleteSql
	Private	strMsg

    Private Sub Class_Initialize
		Call Debug("LsSyukaAc.Class_Initialize()")
		strDBName = GetOption("db","newsdc9")
		set objDB = nothing
		set objSrcRs = nothing
		set objDstRs = nothing
        strFunction = "check"
		strTable = "LsSyuka"
		strYM = ""
		strDeleteSql = ""
    End Sub

    Private Sub Class_Terminate
		Call Debug("LsSyukaAc.Class_Terminate()")
'		Call Close()
    End Sub

    Public Function Init()
		Call Debug("LsSyukaAc.Init()")
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
	    	select case strArg
			case else
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
		Call Debug("LsSyukaAc.Run()")
		Call OpenDB()
		Call OpenSrcRs()
		Call OpenDstRs()
		Call AddRecord()
		Call CloseRs()
		Call CloseDB()
	end function

    Private Function OpenDB()
		Call Debug("LsSyukaAc.OpenDB():" & strDBName)
		set objDb = OpenAdodb(strDBName)
    End Function

    Private Function CloseDB()
		Call Debug("LsSyukaAc.CloseDB():" & strDBName)
		Call objDb.Close()
		set objDb = Nothing
    End Function

	Private Function OpenSrcRs()
		Call Debug("LsSyukaAc.OpenSrcRs()")
		strSql = ""
		strSql = strSql & "select"
		strSql = strSql & " Left(Dt,6)     YM"
		strSql = strSql & ",ShisanJCode    JCd"
		strSql = strSql & ",Pn"
		strSql = strSql & ",Sum(Qty)       Qty"
		strSql = strSql & " from BoSyuka"
		strSql = strSql & " where shisanjcode='00025800'"
		strSql = strSql & " and YM='201511'"
		strSql = strSql & " group by"
		strSql = strSql & " YM"
		strSql = strSql & ",JCd"
		strSql = strSql & ",Pn"
		Call DispMsg(strSql)
		set objSrcRs = objDb.Execute(strSql)
		Call DispMsg("End Execute.")
	End Function

	Private Function DeleteYM()
		Call Debug("LsSyukaAc.DeleteYM()")
	End Function

	Private Function OpenDstRs()
		Call Debug("LsSyukaAc.OpenDstRs()")
		Set objDstRs = Wscript.CreateObject("ADODB.Recordset")
		objDstRs.Open strTable, objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function CloseRs()
		Call Debug("LsSyukaAc.CloseRs()")
		if not objSrcRs is nothing then
			Call objSrcRs.Close()
			set objSrcRs = Nothing
		end if
		if not objDstRs is nothing then
			Call objDstRs.Close()
			set objDstRs = Nothing
		end if
	End Function

	Private Function AddRecord()
		Call Debug("LsSyukaAc.AddRecord()")
		if objSrcRs is Nothing then
			exit function
		end if
		if objDstRs is Nothing then
			exit function
		end if
		do while objSrcRs.Eof = false
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
			objSrcRs.MoveNext
		loop
	End Function

	Private Function SetFields()
		lngRow = lngRow + 1
		strMsg = ""
		objDstRs.AddNew
		objDstRs.Fields("YM")	= objSrcRs.Fields("YM")
		objDstRs.Fields("No")	= lngRow
		objDstRs.Fields("Soko")	= ""
		objDstRs.Fields("JCd")	= objSrcRs.Fields("JCd")
		objDstRs.Fields("Pn")	= objSrcRs.Fields("Pn")
		objDstRs.Fields("Qty")	= objSrcRs.Fields("Qty")
		strMsg = lngRow & "" _
						& " " & objDstRs.Fields("YM") _
						& " " & objDstRs.Fields("No") _
						& " " & objDstRs.Fields("Soko") _
						& " " & objDstRs.Fields("JCd") _
						& " " & objDstRs.Fields("Pn") _
						& " " & objDstRs.Fields("Qty")

		SetFields = true
	End Function
End Class
