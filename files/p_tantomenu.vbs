Option Explicit

function usage()
	Wscript.Echo "p_tantomenu.vbs - P_TANTOMENUàÍóóï\é¶ (2009.12.21)"
	Wscript.Echo "p_tantomenu.vbs [option] [íSìñé“CD]"
	Wscript.Echo "option : -set00000 : íSìñé“(00000)ÇÃÉÅÉjÉÖÅ[ì‡óeÇëSíSìñé“Ç…ìoò^"
end function

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly	= 0
Const adOpenKeyset		= 1
Const adOpenDynamic		= 2
Const adOpenStatic		= 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

dim db
dim	dbName
dim	sqlStr
dim	rsList
dim	rsList00000
dim	rsTanto
dim	rsPMenu
dim	strBuff
dim	strTanto
dim	strDt
dim	strTm
dim	strTantoPrev
dim	strDtPrev
dim	strTmPrev
dim	strTime
dim strRId
dim strMenu
dim strMenuPrev
dim strJgyobu
dim strJgyobuPrev
dim i
dim strNum
dim strNumPrev
dim iKonpo
dim iShiwake
dim	bSet00000

	bSet00000	= False
	strTanto	= ""
	for i = 0 to WScript.Arguments.count - 1
	    select case ucase(WScript.Arguments(i))
	    case "-SET00000"
			bSet00000	= True
	    case "-?"
			usage()
			Wscript.Quit
	    case else
			strTanto	= WScript.Arguments(i)
	    end select
	next

	dbName		= "newsdc"

	Set db = Wscript.CreateObject("ADODB.Connection")
	Wscript.Echo "open db : " & dbName
	db.open dbName

	sqlStr = "select"
	sqlStr = sqlStr & " *"
	sqlStr = sqlStr & " from p_tantomenu"
	if strTanto <> "" then
		sqlStr = sqlStr & " where tanto_code = '" & strTanto & "'"
	else
		if bSet00000 = true then
			sqlStr = sqlStr & " where tanto_code = '00000'"
		end if
	end if
	sqlStr = sqlStr & " order by tanto_code"

'    Wscript.Echo "sql : " & sqlStr

	if bSet00000 = true then
		Set rsList = Wscript.CreateObject("ADODB.Recordset")
		rsList.CursorLocation = adUseClient
		rsList.Open sqlStr, db, adOpenForwardOnly, adLockBatchOptimistic

		sqlStr = "select"
		sqlStr = sqlStr & " *"
		sqlStr = sqlStr & " from p_tantomenu"
		sqlStr = sqlStr & " where tanto_code = '00000'"
		set rsList00000 = db.Execute(sqlStr)
	else
		set rsList = db.Execute(sqlStr)
	end if


	Do While Not rsList.EOF
		strTanto   = rtrim(rsList.Fields("TANTO_CODE"))
        ' íSìñé“ñº TANTO_CODE TANTO_NAME 
       	sqlStr = "select"
    	sqlStr = sqlStr & " *"
    	sqlStr = sqlStr & " from tanto"
    	sqlStr = sqlStr & " where TANTO_CODE = '" & strTanto & "'"
      	set rsTanto = db.Execute(sqlStr)

        ' íSìñé“ ï\é¶
        strBuff = strTanto
        if rsTanto.eof = false then
            strBuff = strBuff & " " & rtrim(rsTanto.Fields("TANTO_NAME"))
        end if

		Wscript.Echo strBuff

        iKonpo      = 0
        iShiwake    = 0
        for i = 1 to 180
                ' JGYOBU_001 NAIGAI_001 MENU_NO_001 
  			strNum = right("000" & i,3)

			if bSet00000 = true then
				rsList.Fields("JGYOBU_" & strNum) = rsList00000.Fields("JGYOBU_" & strNum)
				rsList.Fields("NAIGAI_" & strNum) = rsList00000.Fields("NAIGAI_" & strNum)
				rsList.Fields("MENU_NO_" & strNum) = rsList00000.Fields("MENU_NO_" & strNum)
			else
		        if trim(rsList.Fields("MENU_NO_" & strNum)) = "" then
                		exit for
		        end if
			end if

	        ' ÉÅÉjÉÖÅ[ñº
       		sqlStr = "select"
       		sqlStr = sqlStr & " *"
       		sqlStr = sqlStr & " from p_menu"
       		sqlStr = sqlStr & " where JGYOBU = '" & trim(rsList.Fields("JGYOBU_" & strNum)) & "'"
       		sqlStr = sqlStr & " and NAIGAI = '" & trim(rsList.Fields("NAIGAI_" & strNum)) & "'"
       		sqlStr = sqlStr & " and MENU_NO = '" & trim(rsList.Fields("MENU_NO_" & strNum)) & "'"
       		set rsPMenu = db.Execute(sqlStr)
       		strBuff = " " & strNum
  			strBuff = strBuff & " " & rsList.Fields("JGYOBU_" & strNum)
   			strBuff = strBuff & " " & rsList.Fields("NAIGAI_" & strNum)
   			strBuff = strBuff & " " & rsList.Fields("MENU_NO_" & strNum)
   			if rsPMenu.eof = false then
       			strBuff = strBuff & " " & rtrim(rsPMenu.Fields("MENU_DSP"))
      		end if
			if bSet00000 = true then
	       		sqlStr = "update p_tantomenu"
	       		sqlStr = sqlStr & " set"
	       		sqlStr = sqlStr & "   JGYOBU_" & strNum & " = '" & trim(rsList.Fields("JGYOBU_" & strNum)) & "'"
	       		sqlStr = sqlStr & " , NAIGAI_" & strNum & " = '" & trim(rsList.Fields("NAIGAI_" & strNum)) & "'"
	       		sqlStr = sqlStr & " , MENU_NO_" & strNum & " = '" & trim(rsList.Fields("MENU_NO_" & strNum)) & "'"
				sqlStr = sqlStr & " where tanto_code <> '00000'"
	       		set rsPMenu = db.Execute(sqlStr)
       			strBuff = strBuff & " çXêV " & sqlStr
			end if
    		Wscript.Echo strBuff

        	next
    	rsList.movenext
	Loop
Wscript.Echo ""
rsList.Close
set rsList = nothing

Wscript.Echo "close db : " & dbName
db.Close

set db = nothing
Wscript.Echo "end"
