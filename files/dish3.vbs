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
dim	objDish3
Set objDish3 = New Dish3
dim	lngRet
lngRet = objDish3.Run()
Set objDish3 = Nothing
WScript.Quit lngRet
'-----------------------------------------------------------------------
'食洗移管クラス
'-----------------------------------------------------------------------
Class Dish3
	'-----------------------------------
	' 使用方法
	'-----------------------------------
    Private Function Usage(byval lErr,byval sErr)
		Call Debug("Dish3.Usage()")
		lngErr	= lErr
		strErr	= sErr
		Wscript.Echo "食洗移管"
		Wscript.Echo "Dish3.vbs [option]"
		Wscript.Echo " /db:dns"
		Wscript.Echo " /limit:10"
		Wscript.Echo " /table:item"
		Wscript.Echo " /pn:<pn>"
		Wscript.Echo " /copy"
		Wscript.Echo "Ex."
		Wscript.Echo "cscript Dish3.vbs /db:newsdc6"
		Call DispMsg(strErr)
		Usage = lngErr
    End Function

	Private	strDBName
	Private	objDB
	Private	objRs
	Private	objRsM
	Private	strSql
	Private	lngErr
	Private	strErr
	Private	lngLimit
	Private	strTable
	Private	strPn
	Private	strPrev
	Private	strCurr
	Private	lngUpdate

    Private Sub Class_Initialize
		Call Debug("Dish3.Class_Initialize()")
		strDBName = GetOption("db","newsdc")
		set objDB = nothing
		set objRs = nothing
		set objRsM = nothing
		lngErr	= 0
		strErr	= ""
		lngLimit = CLng(GetOption("limit",0))
		strTable = lcase(GetOption("table","p_shiji"))
		strPn = ucase(GetOption("pn",""))
		strPrev = ""
		strCurr = ""
		lngUpdate = 0
    End Sub

    Private Sub Class_Terminate
		Call Debug("Dish3.Class_Terminate()")
		Call CloseDB()
    End Sub

	'-----------------------------------
	' 初期処理
	'-----------------------------------
	Private Function	Init()
		Call Debug("Dish3.Init()")
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
		case "p_shiji"
		case "p_shiji_k"
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
		Call Debug("Dish3.OpenDb():" & strDBName)
		set objDB = OpenAdodb(strDBName)
	End Function
	'-----------------------------------
	' クローズDB
	'-----------------------------------
    Private Function CloseDB()
		Call Debug("Dish3.CloseDB():" & strDBName)
		if not objRs is nothing then
			Call objRs.Close()
			set objRs = Nothing
		end if
		if not objRsM is nothing then
			Call objRsM.Close()
			set objRsM = Nothing
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
		strSql = strSql & " *"
		strSql = strSql & " from P_SSHIJI_O"
		strSql = strSql & " where SHIMUKE_CODE = '04'"
		strSql = strSql & " and HIN_GAI in (select distinct HIN_GAI from item where JGYOBU='2' and NAIGAI='1')"
		if strPn <> "" then
			strSql = strSql & " and HIN_GAI = '" & strPn & "'"
		end if
		strSql = strSql & " order by"
		strSql = strSql & " HIN_GAI"
		strSql = strSql & ",SHIJI_NO desc"
		GetSql = strSql
	end function
	'-----------------------------------
	' オープンRs
	'-----------------------------------
	Private Function	OpenRs()
		Call Debug("Dish3.OpenRs()")
		dim	strSql
		strSql = GetSql()
		Call DispMsg(strSql)
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
		if strCurr <> strPrev then
			strBuff = strBuff & " " & objRs.Fields("SHIMUKE_CODE")
			strBuff = strBuff & " " & objRs.Fields("JGYOBU")
			strBuff = strBuff & " " & objRs.Fields("NAIGAI")
			strBuff = strBuff & " " & objRs.Fields("HIN_GAI")
		else
			strBuff = strBuff & " " & String(2, " ")
			strBuff = strBuff & " " & String(1, " ")
			strBuff = strBuff & " " & String(1, " ")
			strBuff = strBuff & " " & String(20, " ")
		end if
		strBuff = strBuff & " " & objRs.Fields("SHIJI_NO")
		strBuff = strBuff & " " & objRs.Fields("HAKKO_DT")
		DispRs = strBuff
	End Function
	'-----------------------------------
	' リードRs
	'-----------------------------------
	Private Function	ReadRs()
		Call Debug("Dish3.ReadRs()")
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
			' 品番
			strPrev = strCurr
			strCurr = RTrim(objRs.Fields("HIN_GAI"))
			' レコード内容表示
			Call WScript.StdOut.Write(lngCnt & DispRs())
			' レコード内容表示(改行)
			Call WScript.StdOut.Write(vbCrLf)
			' 構成マスター更新
			Call UpdateMaster()
			' 次レコード
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------
	' リードRs
	'-----------------------------------
	Private Function	UpdateMaster()
		Call Debug("Dish3.UpdateMaster()")
		if strPrev = strCurr then
			exit function
		end if
		dim	strSql
		strSql = ""
		strSql = strSql & " select *"
		strSql = strSql & " from P_SSHIJI_K"
		strSql = strSql & " where SHIJI_NO = '" & objRs.Fields("SHIJI_NO") & "'"
		strSql = strSql & " order by"
		strSql = strSql & " DATA_KBN"
		strSql = strSql & ",SEQNO"
		Call Debug(strSql)
		dim	objRsK
		set	objRsK = objDB.Execute(strSql)
		do while objRsK.Eof = false
			WScript.StdOut.Write " " & objRsK.Fields("DATA_KBN")
			WScript.StdOut.Write " " & objRsK.Fields("SEQNO")
			WScript.StdOut.Write " " & objRsK.Fields("KO_SYUBETSU")
			WScript.StdOut.Write " " & objRsK.Fields("KO_JGYOBU")
			WScript.StdOut.Write " " & objRsK.Fields("KO_NAIGAI")
			WScript.StdOut.Write " " & objRsK.Fields("KO_HIN_GAI")
			WScript.StdOut.Write vbCrLf
			strSql = ""
			strSql = strSql & " select *"
			strSql = strSql & " from p_compo_k"
			strSql = strSql & " where SHIMUKE_CODE = '02'"
			strSql = strSql & " and JGYOBU = '2'"
			strSql = strSql & " and NAIGAI = '1'"
			strSql = strSql & " and HIN_GAI = '" & RTrim(objRs.Fields("HIN_GAI")) & "'"
			strSql = strSql & " and DATA_KBN = '" & RTrim(objRsK.Fields("DATA_KBN")) & "'"
			strSql = strSql & " and SEQNO = '" & RTrim(objRsK.Fields("SEQNO")) & "'"
			Call Debug(strSql)
			if objRsM is nothing then
				Set objRsM = Wscript.CreateObject("ADODB.Recordset")
			else
				objRsM.Close
			end if
			objRsM.Open strSql, objDB, adOpenKeyset, adLockOptimistic
			WScript.StdOut.Write " " & objRsM.Fields("DATA_KBN")
			WScript.StdOut.Write " " & objRsM.Fields("SEQNO")
			WScript.StdOut.Write " " & objRsM.Fields("KO_SYUBETSU")
			WScript.StdOut.Write " " & objRsM.Fields("KO_JGYOBU")
			WScript.StdOut.Write " " & objRsM.Fields("KO_NAIGAI")
			WScript.StdOut.Write " " & objRsM.Fields("KO_HIN_GAI")

			dim	iUpdate
			iUpdate = 1
			if objRsK.Fields("KO_SYUBETSU") <> objRsM.Fields("KO_SYUBETSU") then
				lngUpdate = lngUpdate + iUpdate
				iUpdate = 0
				WScript.StdOut.Write " ×"
			end if
			if objRsK.Fields("KO_JGYOBU") <> objRsM.Fields("KO_JGYOBU") then
				if objRsK.Fields("KO_JGYOBU") = "1" and objRsM.Fields("KO_JGYOBU") = "2" then
				else
					lngUpdate = lngUpdate + iUpdate
					iUpdate = 0
					WScript.StdOut.Write " ×"
					objRsM.Fields("KO_JGYOBU") = objRsK.Fields("KO_JGYOBU") 
				end if
			end if
			if objRsK.Fields("KO_NAIGAI") <> objRsM.Fields("KO_NAIGAI") then
				lngUpdate = lngUpdate + iUpdate
				iUpdate = 0
				WScript.StdOut.Write " ×"
				objRsM.Fields("KO_NAIGAI") = objRsK.Fields("KO_NAIGAI") 
			end if
			if objRsK.Fields("KO_HIN_GAI") <> objRsM.Fields("KO_HIN_GAI") then
				lngUpdate = lngUpdate + iUpdate
				iUpdate = 0
				WScript.StdOut.Write " ×"
				objRsM.Fields("KO_HIN_GAI") = objRsK.Fields("KO_HIN_GAI") 
			end if
			if iUpdate = 0 then
				WScript.StdOut.Write " " & lngUpdate
				objRsM.Update
			end if
			WScript.StdOut.Write vbCrLf

			objRsK.MoveNext
		loop
	End Function

	Public Function	Run()
		Call Debug("Dish3.Run()")
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
