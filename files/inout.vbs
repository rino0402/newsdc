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
Call Include("debug.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "InOutデータ"
	Wscript.Echo "inout.vbs [option]"
	Wscript.Echo " /Make"
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc"))
	Wscript.Echo "ConnectionString=" & objDB.ConnectionString
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	'名前無しオプションチェック
	select case WScript.Arguments.UnNamed.Count
	case 0
	case else
		usage()
		Main = 1
		exit Function
	end select
	'名前付きオプションチェック
	dim	strArg
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "make"
		case "center"
		case "debug"
		case "?"
			call usage()
			exit function
		case else
			call usage()
			exit function
		end select
	next
	call MakeInOut()
	Main = 0
End Function
Function MakeInOut()
	Call Debug("MakeInOut()")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc"))
	'-------------------------------------------------------------------
	'入出荷データ登録
	'-------------------------------------------------------------------
	dim	i
	dim	iCnt
	iCnt = GetOption("make","0")
	if iCnt > 0 then
		for i = 1 to iCnt
			Call DispMsg("MakeInOut():" & i & "/" & iCnt)
			Call InOutDelete(objDb,i)
			Call InOutInsert(objDb,i)
		next
	else
		Call DispMsg("MakeInOut():" & iCnt)
		Call InOutDelete(objDb,iCnt)
		Call InOutInsert(objDb,iCnt)
	end if
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function InOutDelete(objDb,byVal iDay)
	dim	strSql
	strSql = "delete from InOut"
	if iDay > 0 then
		strSql = strSql & " where JITU_DT = GetCharDate(DATEADD(day  ," & -iDay & ",CURDATE()))"
	else
		strSql = strSql & " where JITU_DT = 'LAST" & iDay & "'"
	end if
	Call Debug(strSql)
	Call objDb.Execute(strSql)
End Function
Function InOutInsert(objDb,byVal iDay)
	dim	strSql
	strSql = "insert into InOut("
	strSql = strSql & " JITU_DT"
	strSql = strSql & ",HIN_GAI"
	strSql = strSql & ",InCnt"
	strSql = strSql & ",InQty"
	strSql = strSql & ",OutCnt"
	strSql = strSql & ",OutQty"
	strSql = strSql & ")"
	if iDay > 0 then
		strSql = strSql & " select"
		strSql = strSql & " JITU_DT"
		strSql = strSql & ",HIN_GAI"
		strSql = strSql & ",sum(if(left(RIRK_ID,1) in ('1'),1,0)) InCnt"
		strSql = strSql & ",sum(if(left(RIRK_ID,1) in ('1'),convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC),0)) InQty"
		strSql = strSql & ",sum(if(" & GetSumIf("out") & ",1,0)) OutCnt"
		strSql = strSql & ",sum(if(" & GetSumIf("out") & ",convert(SUMI_JITU_QTY,SQL_NUMERIC) + convert(MI_JITU_QTY,SQL_NUMERIC),0)) OutQty"
		strSql = strSql & " From P_SAGYO_LOG"
		strSql = strSql & " where " & GetSumIf("where")
		strSql = strSql & "   and JITU_DT = GetCharDate(DATEADD(day  ," & -iDay & ",CURDATE()))"
		strSql = strSql & " group by"
		strSql = strSql & " JITU_DT"
		strSql = strSql & ",HIN_GAI"
		strSql = strSql & " having InCnt > 0 or OutCnt > 0"
	else
		strSql = strSql & " select"
		strSql = strSql & " 'LAST" & iDay & "'"
		strSql = strSql & ",HIN_GAI"
		strSql = strSql & ",sum(InCnt)"
		strSql = strSql & ",sum(InQty)"
		strSql = strSql & ",sum(OutCnt)"
		strSql = strSql & ",sum(OutQty)"
		strSql = strSql & " from InOut"
		strSql = strSql & " where JITU_DT between convert(GetCharDate(DATEADD(month,-3,CURDATE())),sql_char)"
		strSql = strSql & "                   and convert(GetCharDate(DATEADD(day,  -1,CURDATE())),sql_char)"
		strSql = strSql & " group by"
		strSql = strSql & " HIN_GAI"
	end if
	Call Debug(strSql)
	Call objDb.Execute(strSql)
End Function

Function GetSumIf(byVal strIO)
	select case strIO
	case "out"
		select case GetCenter()
		case "大阪"
			GetSumIf = "left(RIRK_ID,1) in ('2','4') or RIRK_ID in ('87','88')"
		case else
			GetSumIf = "left(RIRK_ID,1) in ('2','4')"
		end select
	case "where"
		select case GetCenter()
		case "大阪"
			GetSumIf = "(left(RIRK_ID,1) in ('1','2','4') or RIRK_ID in ('87','88','I0'))"
		case else
			GetSumIf = "(left(RIRK_ID,1) in ('1','2','4'))"
		end select
	end select
End Function

Function GetCenter()
	GetCenter = ""
	select case lcase(GetOption("center",""))
	case "osk"
		GetCenter = "大阪"
	end select
End Function
