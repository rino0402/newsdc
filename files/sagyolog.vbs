Option Explicit
' 作業ログ時間集計プログラム
' 2008.08.21 -StClear/-StNoClear対応
' 2008.09.16
' 2009.04.15 同一レコードで0になる不具合修正
'			 カーソルでUpdateするように変更

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

function CalcSecond(byVal strNow,byVal strPrv)
    dim tmNow
    dim tmPrv
    
    strNow = trim(strNow)
    strPrv = trim(strPrv)
    tmNow   = TimeSerial(Left(strNow,2),Mid(strNow,3,2),Right(strNow,2))
    tmPrv   = TimeSerial(Left(strPrv,2),Mid(strPrv,3,2),Right(strPrv,2))
    
    CalcSecond = DateDiff("s",tmPrv,tmNow)
	if CalcSecond < 0 then
		CalcSecond = 0
	end if
end function
'-----------------------------------------------------------------------
'オプションチェック
'-----------------------------------------------------------------------
Function GetOption(byval strName _
				  ,byval strDefault _
		  		  )
	dim	strValue

	if strName = "" then
		strValue = ""
		if strDefault < WScript.Arguments.UnNamed.Count then
			strValue = WScript.Arguments.UnNamed(strDefault)
		end if
	else
		strValue = strDefault
		if WScript.Arguments.Named.Exists(strName) then
			strValue = WScript.Arguments.Named(strName)
		end if
	end if
	GetOption = strValue
End Function

dim	db
dim	dbName
dim	sqlStr
dim	rsList
dim	strBuff
dim	strTanto
dim	strDt
dim	strTm
dim	strTantoPrev
dim	strDtPrev
dim	strTmPrev
dim	strDtPrevQ	' 処理日：問合せ用
dim	strTmPrevQ	' 処理時：問合せ用
dim	strTime
dim	strRId
dim	strRIdPrev
dim	strMenu
dim	strMenuPrev
dim	strJgyobu
dim	strJgyobuPrev
dim	i
dim	flgST   	' TRUE:ST でクリア/FALSE:STでクリアしない
dim	flgBefore	' 前の作業に時間をセット
dim	flgClean	' 作業をクリア 日付指定の時クリア

' strSoko = WScript.Arguments(0)
strDt		= ""
flgST		= True
flgBefore	= False
flgClean	= False
strTanto	= ""
dim	cLocation
cLocation = adUseServer
for i = 0 to WScript.Arguments.UnNamed.count - 1
    select case ucase(WScript.Arguments.UnNamed(i))
    case "-STCLEAR"
        flgST		= True
    case "-STNOCLEAR"
        flgST		= False
    case "-BEFORE"
        flgBefore	= True
    case "-CLIENT"
		cLocation = adUseClient
    case "-?"
        Wscript.Echo "作業ログ集計(2008.09.16)"
        Wscript.Echo "sagyolog.vbs [option] [処理日 [担当者]]"
        Wscript.Echo " -StClear   : STでクリア(default)"
        Wscript.Echo " -StNoClear : STでクリアしない"
        Wscript.Echo " -Before    : 前の作業に時間をセット"
        Wscript.Echo " -?"
        Wscript.Echo " 処理日     : Ex.20080916 省略:システム日付"
        Wscript.Echo " 担当者     : Ex.01212"
'	goto EndPrg
    case else
	if strDt = "" then
	        strDt = WScript.Arguments.UnNamed(i)
	else
	        strTanto = WScript.Arguments.UnNamed(i)
	end if
    end select
next

if strDt = "" then
    strDt = year(now()) & right("0" & month(now()),2) & right("0" & day(now()),2)
    flgClean	= False
else
    flgClean	= True
end if

Wscript.Echo "sagyolog.vbs " & strDt
Wscript.Echo "作業ログ集計"
Wscript.Echo "strDt = " & strDt
Wscript.Echo "flgSt = " & flgSt
Wscript.Echo "flgBefore = " & flgBefore
Wscript.Echo "strTanto = " & strTanto

dbName = GetOption("db","newsdc")

Set db = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbName
db.open dbName

	if flgClean = True then
		sqlStr = "update p_sagyo_log set WORK_TM = ''"
		sqlStr = sqlStr & " where JITU_DT = '" & strDt & "'"
		if strTanto <> "" then
			sqlStr = sqlStr & " and TANTO_CODE = '" & strTanto & "'"
		end if
	
		Wscript.Echo "sql : " & sqlStr

		call db.Execute(sqlStr)
		Wscript.Echo ""
	end if
	sqlStr = "select"
	sqlStr = sqlStr & " *"
	sqlStr = sqlStr & " from p_sagyo_log"
	sqlStr = sqlStr & " where JITU_DT like '" & strDt & "'"
	sqlStr = sqlStr & " and MENU_NO <> '99'"
	if strTanto <> "" then
		sqlStr = sqlStr & " and TANTO_CODE = '" & strTanto & "'"
	end if
'	sqlStr = sqlStr & " and WORK_TM = ''"
	sqlStr = sqlStr & " order by TANTO_CODE,JITU_DT,JITU_TM"

	Wscript.Echo "sql : " & sqlStr

'	set rsList = db.Execute(sqlStr)
	Set rsList = Wscript.CreateObject("ADODB.Recordset")
'	rsList.Open sqlStr, db, adOpenStatic, adLockOptimistic
'	rsList.Open sqlStr, db, adOpenKeyset, adLockBatchOptimistic
	rsList.CursorLocation = cLocation	'adUseClient	'adUseServer
	rsList.Open sqlStr, db, adOpenKeyset, adLockOptimistic
''	rsList.Open sqlStr, db, adOpenKeyset
	Wscript.Echo ""

	strTantoPrev    = ""
	strDtPrev       = ""
	strTmPrev       = ""
	strDtPrevQ	= ""
	strTmPrevQ	= ""
	strMenuPrev     = ""
	strJgyobuPrev   = ""

	Do While Not rsList.EOF
		strTanto    = rsList.Fields("TANTO_CODE")
		strDt       = rsList.Fields("JITU_DT")
		strTm       = rsList.Fields("JITU_TM")
		strRId      = rsList.Fields("RIRK_ID")
		strMenu     = rsList.Fields("MENU_NO")
		strJgyobu   = rsList.Fields("JGYOBU")

        if strJgyobu = "S" then
            if strJgyobuPrev <> "" then
                strJgyobu = strJgyobuPrev
            end if
        end if

		if strTanto <> strTantoPrev then
			strDtPrev       = ""
		    strTmPrev       = ""
       	end if

        if strMenu <> strMenuPrev then
            if flgSt = true then
	            strDtPrev       = ""
	            strTmPrev       = ""
       	    end if
        end if

        select case strRId
        case "ST"
            if flgSt = true then
    	        strDtPrev       = ""
		        strTmPrev       = ""
            end if
        end select

        strTime = "000000"
        select case strMenu
        case "99"
        case else
            if strDtPrev = strDt and strTmPrev <> "" then
				strTime = right("000000" & CalcSecond(strTm,strTmPrev),6)
            end if
        end select
	
		strBuff = rsList.Fields("TANTO_CODE")
		strBuff = strBuff & " " & rsList.Fields("JITU_DT")
		strBuff = strBuff & " " & rsList.Fields("JITU_TM")
		strBuff = strBuff & " " & rsList.Fields("MENU_NO")
		strBuff = strBuff & " " & rsList.Fields("RIRK_ID")
		strBuff = strBuff & " " & rsList.Fields("WORK_TM")
		strBuff = strBuff & " " & strTime

		if flgBefore = True then
            if strTmPrev <> "" then
				rsList.MovePrevious
				if rsList.Eof = False then
					rsList.Fields("WORK_TM") = strTime
					rsList.Fields("JGYOBU") = strJgyobu
					rsList.Update
					rsList.MoveNext
				end if
			end if
		else
			rsList.Fields("WORK_TM") = strTime
			rsList.Fields("JGYOBU") = strJgyobu
			on error resume next
			rsList.Update
			if Err <> 0 then
				Wscript.Echo strBuff & ":" & Err
				rsList.CancelUpdate
			end if
			on error goto 0
		end if
        

'		strBuff = strBuff & ":" & db.Errors.Error

		Wscript.Echo strBuff

		strTantoPrev	= strTanto
		strDtPrev	= strDt
		strTmPrev	= strTm
		strDtPrevQ	= strDt
		strTmPrevQ	= strTm
		strMenuPrev	= strMenu
		strJgyobuPrev	= strJgyobu
		strRIdPrev	= strRId
        select case strRId
        case "EN"
	        strDtPrev       = ""
	        strTmPrev       = ""
        end select

		rsList.movenext
	Loop
	Wscript.Echo ""

Wscript.Echo "close db : " & dbName
db.Close
set db = nothing
Wscript.Echo "end"
