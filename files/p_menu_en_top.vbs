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
dim strNumPrev

	Wscript.Echo "p_menu_en_top.vbs - P_MENUàÍóóï\é¶ Åï çÏã∆èIóπ êÊì™í«â¡ (2008.08.21)"

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
        
       	strYoinEn = ""

       	if rtrim(rsList.Fields("YOIN_01")) <> "EN" then
            for i = 20 to 2 step -1
                strNum = right("0" & i,2)
                strNumPrev = right("0" & i -1,2)
                ' YOIN_19 PARAM_19 Disp_19 LOG_OUT_19 
                sqlStr = "update p_menu"
                sqlStr = sqlStr & " set"
                sqlStr = sqlStr & " YOIN_" & strNum & " = YOIN_" & strNumPrev
                sqlStr = sqlStr & ",PARAM_" & strNum & " = PARAM_" & strNumPrev
                sqlStr = sqlStr & ",Disp_" & strNum & " = Disp_" & strNumPrev
                sqlStr = sqlStr & ",LOG_OUT_" & strNum & " = LOG_OUT_" & strNumPrev
                sqlStr = sqlStr & " where JGYOBU = '" & rsList.Fields("JGYOBU") & "'"
                sqlStr = sqlStr & " and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
                sqlStr = sqlStr & " and MENU_NO = '" & rsList.Fields("MENU_NO") & "'"
        		Wscript.Echo sqlStr
        	    call db.Execute(sqlStr)               
            next
            sqlStr = "update p_menu"
            sqlStr = sqlStr & " set"
            sqlStr = sqlStr & " YOIN_01 = 'EN'"
            sqlStr = sqlStr & ",PARAM_01 = ''"
            sqlStr = sqlStr & ",Disp_01 = 'ÅyçÏã∆èIóπÅz'"
            sqlStr = sqlStr & ",LOG_OUT_01 = '1'"
            sqlStr = sqlStr & " where JGYOBU = '" & rsList.Fields("JGYOBU") & "'"
            sqlStr = sqlStr & " and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
            sqlStr = sqlStr & " and MENU_NO = '" & rsList.Fields("MENU_NO") & "'"
      		Wscript.Echo sqlStr
      	    call db.Execute(sqlStr)               
        end if
		rsList.movenext
	Loop
	Wscript.Echo ""

	Wscript.Echo "close db : " & dbName
	db.Close
	set db = nothing
	Wscript.Echo "end"
