Option Explicit

dim db
dim	dbName
dim	sqlStr
dim	rsList
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

Wscript.Echo "p_tantomenu.vbs - P_TANTOMENU一覧表示 (2008.08.27)"

dbName = "newsdc"

Set db = Wscript.CreateObject("ADODB.Connection")
Wscript.Echo "open db : " & dbName
db.open dbName

	sqlStr = "select"
	sqlStr = sqlStr & " *"
	sqlStr = sqlStr & " from p_tantomenu"

'    Wscript.Echo "sql : " & sqlStr

	set rsList = db.Execute(sqlStr)
	Wscript.Echo ""

	Do While Not rsList.EOF
		strTanto   = rtrim(rsList.Fields("TANTO_CODE"))
        ' 担当者名 TANTO_CODE TANTO_NAME 
       	sqlStr = "select"
    	sqlStr = sqlStr & " *"
    	sqlStr = sqlStr & " from tanto"
    	sqlStr = sqlStr & " where TANTO_CODE = '" & strTanto & "'"
      	set rsTanto = db.Execute(sqlStr)

        ' 担当者 表示
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

            if trim(rsList.Fields("MENU_NO_" & strNum)) = "" then
                exit for
            end if
            if trim(rsList.Fields("MENU_NO_" & strNum)) = "33" then
                iShiwake = i
            end if

            ' メニュー名
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
    		if rsList.eof = false then
        		strBuff = strBuff & " " & rtrim(rsPMenu.Fields("MENU_DSP"))
                if iKonpo = 0 then
                    if left(rsPMenu.Fields("MENU_DSP"),4) = "出荷梱包" then
                        iKonpo = i
                    end if
                end if
      		end if
    		Wscript.Echo strBuff

        next

        if iKonpo > 0 then
      		Wscript.Echo "★★出荷梱包★★:" & iKonpo & " - " & i
'      		if iShiwake = 0 and strTanto = "99999" then
      		if iShiwake = 0 then
          		Wscript.Echo "★★仕分追加" & strTanto
                ' メニューを1つずらす
                for i = i to iKonpo step -1
                    strNum = right("000" & i,3)
                    strNumPrev = right("000" & i - 1,3)
        	        sqlStr = "update p_tantomenu set"
        	        if i = iKonpo then
        	            sqlStr = sqlStr & " JGYOBU_" & strNum & " = '7'"
        	            sqlStr = sqlStr & ",NAIGAI_" & strNum & " = '1'"
        	            sqlStr = sqlStr & ",MENU_NO_" & strNum & " = '33'"
        	        else
        	            sqlStr = sqlStr & " JGYOBU_" & strNum & " = JGYOBU_" & strNumPrev
        	            sqlStr = sqlStr & ",NAIGAI_" & strNum & " = NAIGAI_" & strNumPrev
        	            sqlStr = sqlStr & ",MENU_NO_" & strNum & " = MENU_NO_" & strNumPrev
        	        end if
        	        sqlStr = sqlStr & " where TANTO_CODE = '" & strTanto & "'"
          		    Wscript.Echo "★★" & sqlStr
        	        
            	    call db.Execute(sqlStr)
                next
          		
      		end if
   		end if
    	rsList.movenext
	Loop
	Wscript.Echo ""

Wscript.Echo "close db : " & dbName
db.Close
set db = nothing
Wscript.Echo "end"
