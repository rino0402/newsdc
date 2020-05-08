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
Call Include("debug.vbs")
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
dim	objItem2
Set objItem2 = New Item2
dim	lngRet
lngRet = objItem2.Run()
Set objItem2 = Nothing
WScript.Quit lngRet
'-----------------------------------------------------------------------
'食洗移管クラス
'-----------------------------------------------------------------------
Class Item2
	'-----------------------------------
	' 使用方法
	'-----------------------------------
    Private Function Usage(byval lErr,byval sErr)
		Call Debug("Item2.Usage()")
		lngErr	= lErr
		strErr	= sErr
		Wscript.Echo "食洗移管"
		Wscript.Echo "item2.vbs [option]"
		Wscript.Echo " /db:dns"
		Wscript.Echo " /limit:10"
		Wscript.Echo " /table:item"
		Wscript.Echo " /pn:<pn>"
		Wscript.Echo " /copy"
		Wscript.Echo "Ex."
		Wscript.Echo "cscript item2.vbs /db:newsdc6"
		Call DispMsg(strErr)
		Usage = lngErr
    End Function

	Private	strDBName
	Private	objDB
	Private	objRs
	Private	objCopyRs
	Private	strSql
	Private	lngErr
	Private	strErr
	Private	lngLimit
	Private	strTable
	Private	strPn

    Private Sub Class_Initialize
		Call Debug("Item2.Class_Initialize()")
		strDBName = GetOption("db","newsdc")
		set objDB = nothing
		set objRs = nothing
		set objCopyRs = nothing
		lngErr	= 0
		strErr	= ""
		lngLimit = CLng(GetOption("limit",0))
		strTable = lcase(GetOption("table","item"))
		strPn = ucase(GetOption("pn",""))
    End Sub

    Private Sub Class_Terminate
		Call Debug("Item2.Class_Terminate()")
		Call CloseDB()
    End Sub

	'-----------------------------------
	' 初期処理
	'-----------------------------------
	Private Function	Init()
		Call Debug("Item2.Init()")
		lngErr = 0
		'-----------------------------------
		' オプションチェック
		'-----------------------------------
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
	    	select case strArg
			case else
				Init = Usage(-1,"unknown " & strArg)
				Exit Function
			end select
		Next
		'-----------------------------------
		' オプションチェック
		'-----------------------------------
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "limit"
			case "copy"
			case "delete"
			case "table"
			case "pn"
			case else
				Init = Usage(-1,"unknown option:/" & strArg)
				Exit Function
			end select
		Next
		select case strTable
		case "item"
		case "p_compo"
		case "p_compo_k"
		case else
			Init = Usage(-1,"bad table:" & strTable)
			Exit Function
		end select

		Init = lngErr
    End Function

	'-----------------------------------
	' オープンDB
	'-----------------------------------
	Private Function	OpenDB()
		Call Debug("Item2.OpenDb():" & strDBName)
		set objDB = OpenAdodb(strDBName)
	End Function
	'-----------------------------------
	' クローズDB
	'-----------------------------------
    Private Function CloseDB()
		Call Debug("Item2.CloseDB():" & strDBName)
		if not objRs is nothing then
			Call objRs.Close()
			set objRs = Nothing
		end if
		if not objCopyRs is nothing then
			Call objCopyRs.Close()
			set objCopyRs = Nothing
		end if
		if not objDB is nothing then
			Call objDB.Close()
			set objDB = Nothing
		end if
    End Function
	'-----------------------------------
	' SQL
	'-----------------------------------
	Private Function	GetSql()
		dim	strSql
		strSql = "select"
		if lngLimit > 0 then
			strSql = strSql & " top " & lngLimit
		end if
		select case strTable
		case "item"
			strSql = strSql & " *"
			strSql = strSql & " from Item"
			strSql = strSql & " where JGYOBU = '1'"
			strSql = strSql & " and NAIGAI = '1'"
	'		strSql = strSql & " and HIN_GAI in (select distinct top 10 HIN_GAI from p_compo where SHIMUKE_CODE = '04')"
			strSql = strSql & " and HIN_GAI in (select distinct Pn from PnNew where JCode = '00036003' and ShisanJCode = '00021529' and Pn not in (select distinct Pn from PnNew where JCode = '00036003' and ShisanJCode = '00023100' and (UnitKbn <> '0' or NaiKbn <> '0' or GaiKbn <> '0')))"
		case "p_compo"
			strSql = strSql & " *"
			strSql = strSql & " from p_compo"
			strSql = strSql & " where SHIMUKE_CODE = '04'"
			strSql = strSql & " and JGYOBU = '2'"
			strSql = strSql & " and NAIGAI = '1'"
			strSql = strSql & " and DATA_KBN = '0'"
			strSql = strSql & " and HIN_GAI in (select distinct HIN_GAI from item where JGYOBU='2' and NAIGAI='1')"
		case "p_compo_k"
			strSql = strSql & " *"
			strSql = strSql & " from p_compo_k"
			strSql = strSql & " where SHIMUKE_CODE = '04'"
			strSql = strSql & " and JGYOBU = '1'"
			strSql = strSql & " and NAIGAI = '1'"
			strSql = strSql & " and DATA_KBN <> '0'"
			strSql = strSql & " and HIN_GAI in (select distinct HIN_GAI from item where JGYOBU='2' and NAIGAI='1')"
		end select
		if strPn <> "" then
			strSql = strSql & " and HIN_GAI = '" & strPn & "'"
		end if
		GetSql = strSql
	end function
	'-----------------------------------
	' オープンRs
	'-----------------------------------
	Private Function	OpenRs()
		Call Debug("Item2.OpenRs()")
		dim	strSql
		strSql = GetSql()
		Call DispMsg("strSql=" & strSql)
		set	objRs = objDB.Execute(strSql)
'		Call Debug("	ADODB.Recordset:" & strTable)
'		Set objRs = Wscript.CreateObject("ADODB.Recordset")
'		objRs.Open strSql, objDB, adOpenKeyset, adLockOptimistic
	End Function
	'-----------------------------------
	' レコード表示
	'-----------------------------------
	Private Function	DispRs()
		dim	strBuff
		strBuff = ""
		select case strTable
		case "item"
					strBuff = strBuff & " " & objRs.Fields("JGYOBU")
					strBuff = strBuff & " " & objRs.Fields("NAIGAI")
					strBuff = strBuff & " " & objRs.Fields("HIN_GAI")
		case "p_compo"
					strBuff = strBuff & " " & objRs.Fields("SHIMUKE_CODE")
					strBuff = strBuff & " " & objRs.Fields("JGYOBU")
					strBuff = strBuff & " " & objRs.Fields("NAIGAI")
					strBuff = strBuff & " " & objRs.Fields("HIN_GAI")
					strBuff = strBuff & " " & objRs.Fields("DATA_KBN")
		case "p_compo_k"
					strBuff = strBuff & " " & objRs.Fields("SHIMUKE_CODE")
					strBuff = strBuff & " " & objRs.Fields("JGYOBU")
					strBuff = strBuff & " " & objRs.Fields("NAIGAI")
					strBuff = strBuff & " " & objRs.Fields("HIN_GAI")
					strBuff = strBuff & " " & objRs.Fields("DATA_KBN")
					strBuff = strBuff & " " & objRs.Fields("KO_SYUBETSU")
					strBuff = strBuff & " " & objRs.Fields("KO_JGYOBU")
					strBuff = strBuff & " " & objRs.Fields("KO_NAIGAI")
					strBuff = strBuff & " " & objRs.Fields("KO_HIN_GAI")
		end select
		DispRs = strBuff
	End Function
	'-----------------------------------
	' リードRs
	'-----------------------------------
	Private Function	ReadRs()
		Call Debug("Item2.ReadRs()")
		dim	lngCnt
		lngCnt = 0
		do while objRs.Eof = false
			lngCnt = lngCnt + 1
			if lngLimit > 0 then
'				Call Debug(lngCnt & ":" & lngLimit)
				if lngCnt > lngLimit then
					' select top が効かない
					Call DispMsg("limit over:" & lngCnt & ">" & lngLimit)
					exit function
				end if
			end if
			' レコード内容表示
			Call WScript.StdOut.Write(lngCnt & DispRs())
			' レコードCopy
			Call CopyRs()
			' レコード削除
			Call DeleteRs()
			' レコード内容表示(改行)
			Call WScript.StdOut.Write(vbCrLf)
			' 次レコード
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------
	' 削除Rs
	'-----------------------------------
	Private Function	DeleteRs()
		Call Debug("Item2.DeleteRs()")
		if WScript.Arguments.Named.Exists("delete") = false then
			exit function
		end if
		Call WScript.StdOut.Write(" Delete")
		dim	strSql
		select case strTable
		case "item"
			strSql = strSql & " delete"
			strSql = strSql & " from Item"
			strSql = strSql & " where JGYOBU = '" & objRs.Fields("JGYOBU") & "'"
			strSql = strSql & " and NAIGAI = '" & objRs.Fields("NAIGAI") & "'"
			strSql = strSql & " and HIN_GAI = '" & objRs.Fields("HIN_GAI") & "'"
		case "p_compo"
			strSql = strSql & " delete"
			strSql = strSql & " from p_compo"
			strSql = strSql & " where SHIMUKE_CODE = '" & objRs.Fields("SHIMUKE_CODE") & "'"
			strSql = strSql & " and JGYOBU = '" & objRs.Fields("JGYOBU") & "'"
			strSql = strSql & " and NAIGAI = '" & objRs.Fields("NAIGAI") & "'"
			strSql = strSql & " and HIN_GAI = '" & objRs.Fields("HIN_GAI") & "'"
			strSql = strSql & " and DATA_KBN = '" & objRs.Fields("DATA_KBN") & "'"
			strSql = strSql & " and SEQNO = '" & objRs.Fields("SEQNO") & "'"
		case "p_compo_k"
			strSql = strSql & " delete"
			strSql = strSql & " from p_compo_k"
			strSql = strSql & " where SHIMUKE_CODE = '" & objRs.Fields("SHIMUKE_CODE") & "'"
			strSql = strSql & " and JGYOBU = '" & objRs.Fields("JGYOBU") & "'"
			strSql = strSql & " and NAIGAI = '" & objRs.Fields("NAIGAI") & "'"
			strSql = strSql & " and HIN_GAI = '" & objRs.Fields("HIN_GAI") & "'"
			strSql = strSql & " and DATA_KBN = '" & objRs.Fields("DATA_KBN") & "'"
			strSql = strSql & " and SEQNO = '" & objRs.Fields("SEQNO") & "'"
		end select
		WScript.StdOut.Write(" " & strSql)
		Call objDB.Execute(strSql)
	End Function
	'-----------------------------------
	' コピーRs
	'-----------------------------------
	Private Function	CopyRs()
		if WScript.Arguments.Named.Exists("copy") = false then
			exit function
		end if
		Call WScript.StdOut.Write(" Copy")
		Call Debug("Item2.CopyRs()")
		if objCopyRs is nothing then
			Call Debug("	ADODB.Recordset:" & strTable)
			Set objCopyRs = Wscript.CreateObject("ADODB.Recordset")
'			objCopyRs.Open strTable, objDB, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
			objCopyRs.Open strTable, objDB, adOpenStatic, adLockOptimistic, adCmdTableDirect
		end if
		objCopyRs.AddNew
		dim	objF
		for each objF in objRs.Fields
'			Call Debug(objF.Name & ":" & objF.Value)
			select case UCase(objF.Name)
			case "SHIMUKE_CODE"
				objCopyRs.Fields(objF.Name) = "02"
			case "JGYOBU"
				objCopyRs.Fields(objF.Name) = "2"
			case "CLASS_CODE"
			case "KO_JGYOBU"
				if objF.Value = "1" then
					objCopyRs.Fields(objF.Name) = "2"
				end if
			case else
				objCopyRs.Fields(objF.Name) = objF.Value
			end select
		next
'		on error resume next
			objCopyRs.Update
			select case Err.Number
			case &h80004005
				Call WScript.StdOut.Write("■二重登録■ " & objCopyRs.Fields("JGYOBU"))
				Call objCopyRs.CancelUpdate
			case 0
			case else
				Call WScript.StdOut.Write("0x" & Hex(Err.Number) & " " & Err.Description)
				Call objCopyRs.CancelUpdate
			end select
'		on error goto 0

	End Function

	Public Function	Run()
		Call Debug("Item2.Run()")
		if Init() <> 0 then
			Run = lngErr
			Exit Function
		end if
		if OpenDb() <> 0 then
			Run = lngErr
			Exit Function
		end if
		if OpenRs() <> 0 then
			Run = lngErr
			Exit Function
		end if
		if ReadRs() <> 0 then
			Run = lngErr
			Exit Function
		end if
		Run = 0
    End Function

End Class
