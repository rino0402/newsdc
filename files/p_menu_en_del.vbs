Option Explicit

dim db
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
dim	strTime
dim strRId
dim strMenu
dim strMenuPrev
dim strJgyobu
dim strJgyobuPrev
dim i
dim strNum

Wscript.Echo "p_menu_en_del.vbs - P_MENUàÍóóï\é¶ Åï çÏã∆èIóπ çÌèú (2008.08.25)"

dbName = "newsdc"

Set db = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbName
db.open dbName

	sqlStr = "select"
	sqlStr = sqlStr & " *"
	sqlStr = sqlStr & " from p_menu"

Wscript.Echo "sql : " & sqlStr

	set rsList = db.Execute(sqlStr)
	Wscript.Echo ""

	Do While Not rsList.EOF
		strJgyobu   = rsList.Fields("JGYOBU")

        ' JGYOBU NAIGAI MENU_NO MENU_DSP 
		strBuff = rsList.Fields("JGYOBU")
		strBuff = strBuff & " " & rsList.Fields("NAIGAI")
		strBuff = strBuff & " " & rsList.Fields("MENU_NO")
		strBuff = strBuff & " " & rsList.Fields("MENU_DSP")
		Wscript.Echo strBuff

        dim strYoinEn
        dim strYoin
        dim strDisp

        strYoinEn = ""

        for i = 1 to 20
            strNum = right("0" & i,2)

            strBuff = ""
    		strBuff = strBuff & " " & strNum & "(" & rsList.Fields("YOIN_" & strNum)
    		strBuff = strBuff & " " & rsList.Fields("PARAM_" & strNum)
    		strBuff = strBuff & " " & rsList.Fields("Disp_" & strNum)
    		strBuff = strBuff & " " & rsList.Fields("LOG_OUT_" & strNum)
    		strBuff = strBuff & ")"
    		Wscript.Echo strBuff

            strYoin = rtrim(rsList.Fields("YOIN_" & strNum))
            strDisp = rtrim(rsList.Fields("Disp_" & strNum))
            if strYoin = "EN" then
                if strDisp = "çÏã∆èIóπ" then
                    sqlStr = "update p_menu"
                    sqlStr = sqlStr & " set"
                    sqlStr = sqlStr & " YOIN_" & strNum & " = ''"
                    sqlStr = sqlStr & ",Disp_" & strNum & " = ''"
                    sqlStr = sqlStr & ",LOG_OUT_" & strNum & " = '1'"
                    sqlStr = sqlStr & " where JGYOBU = '" & rsList.Fields("JGYOBU") & "'"
                    sqlStr = sqlStr & " and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
                    sqlStr = sqlStr & " and MENU_NO = '" & rsList.Fields("MENU_NO") & "'"
            		Wscript.Echo sqlStr
            	    call db.Execute(sqlStr)
                    exit for
                end if
            end if
            if strYoin = "EN" then
                strYoinEn = "EN"
            end if
            ' YOIN_01 PARAM_01 Disp_01 LOG_OUT_01 
        next

    	rsList.movenext
	Loop
	Wscript.Echo ""

Wscript.Echo "close db : " & dbName
db.Close
set db = nothing
Wscript.Echo "end"
