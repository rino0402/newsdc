Option Explicit
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
End Function
Call Include("const.vbs")

Call Main()

Function usage()
    Wscript.Echo "箱No修正(2012.01.27)"
    Wscript.Echo "hako_no.vbs "
    Wscript.Echo "<例>"
    Wscript.Echo "id_no.vbs 0073805 0495148"
End Function

Function Main()
	dim	strIdNoSrc
	dim	strIdNoDst
	dim	strArg

	strIdNoSrc	= ""
	strIdNoDst	= ""

	for each strArg in WScript.Arguments.UnNamed
	    select case strArg
		case "-?"
			call usage()
			exit Function
		case else
			if strIdNoSrc = "" then
				strIdNoSrc = strArg
			elseif strIdNoDst = "" then
				strIdNoDst = strArg
			else
				usage()
				exit Function
			end if
		end select
	next
	for each strArg in WScript.Arguments.Named
'		Wscript.Echo strArg,WScript.Arguments.Named(strArg)
		select case lcase(strArg)
		case "db"
		case "update"
		case else
			call usage()
			exit function
		end select
	next
	call HakoNo()
End Function

Function HakoNo()
	dim	objDb
	dim	strDbName
	dim	rsSrc
	dim	rsDst
	dim	strSql
	dim	strIdNoMax
	dim	intLUchiNo
	dim	strKonpoIdCurr
	dim	strKonpoIdPrev
	dim	strMsg
	dim	strLPageCurr
	dim	strLPagePrev

	strDbName = GetOption("db","newsdc-osk")
	call DispMsg("HakoNo()")

	call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	call DispMsg("Open:" & strDbName)
	objDb.Open strDbName

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsSrc = Wscript.CreateObject("ADODB.Recordset")
	call DispMsg("rsSrc.CursorLocation=" & rsSrc.CursorLocation)
	rsSrc.CursorLocation = adUseClient
	call DispMsg("rsSrc.CursorLocation=" & rsSrc.CursorLocation)

	strSql = GetSql()
	call DispMsg("Sql:" & strSql)
	rsSrc.Open strSql, objDb, adOpenKeyset, adLockBatchOptimistic
	strKonpoIdPrev = ""
	strLPagePrev = ""
	do while rsSrc.EOF = False
		strMsg = rsSrc.Fields("SND_YMD") & " " & _
				 rsSrc.Fields("SND_HMS") & " " & _
				 rsSrc.Fields("TOK_CD") & " " & _
				 rsSrc.Fields("CHU_CD") & " " & _
				 rsSrc.Fields("SYU_JUN") & " " & _
				 rsSrc.Fields("TEI_NM") & " " & _
				 rsSrc.Fields("L_PAGE") & " " & _
				 rsSrc.Fields("KONPO_ID") & " " & _
				 rsSrc.Fields("L_UCHI_NO")
		strLPageCurr = rtrim(rsSrc.Fields("L_PAGE"))
		if strLPagePrev = strLPageCurr then
			rsSrc.Fields("KONPO_ID") = strKonpoIdPrev 
		end if
		strLPagePrev = strLPageCurr

		strKonpoIdCurr = rtrim(rsSrc.Fields("KONPO_ID"))
		if strKonpoIdPrev <> strKonpoIdCurr then
			strKonpoIdPrev = strKonpoIdCurr
			intLUchiNo = 0
		end if
		intLUchiNo = intLUchiNo + 1
		strMsg = strMsg & right(space(5) & intLUchiNo,3)

		if cint(rsSrc.Fields("L_UCHI_NO")) <> intLUchiNo then
			strMsg = strMsg & " update"
			if WScript.Arguments.Named.Exists("update") then
				rsSrc.Fields("L_UCHI_NO") = right(space(10) & intLUchiNo,10)
				rsSrc.UpdateBatch
				strMsg = strMsg & " done"
			end if
		end if
		call DispMsg(strMsg)
		rsSrc.MoveNext
	loop
	rsSrc.Close
	set rsSrc = Nothing

	call DispMsg("Close:" & strDbName)
	objDb.Close
	set objDb = Nothing
End Function

Function UpdateIdNo(byval rsSrc _
				   ,byval strIdNoMax _
				   )
	dim	strUpdate
	if WScript.Arguments.Named.Exists("y_syuka") then
		call DispMsg("  " & _
					 rsSrc.Fields("KEY_ID_NO") & " " & _
					 rsSrc.Fields("KEY_SYUKA_YMD") & " " & _
					 rsSrc.Fields("KEY_MUKE_CODE") & " " & _
					 rsSrc.Fields("MUKE_NAME") _
					)
		strUpdate = ""
		if WScript.Arguments.Named.Exists("update") then
			rsSrc.Fields("KEY_ID_NO") = strIdNoMax
			rsSrc.Fields("ID_NO") = strIdNoMax
			rsSrc.Fields("UPD_NOW") = GetDateTime(Now())
			rsSrc.UpdateBatch
			strUpdate = " (更新)"
		end if
		call DispMsg("  " _
					& strIdNoMax _
					& strUpdate _
					)
	else
		call DispMsg("  " & _
					 rsSrc.Fields("ID_NO") & " " & _
					 rsSrc.Fields("SYUKA_YMD") & " " & _
					 rsSrc.Fields("MUKE_CODE") & " " & _
					 rsSrc.Fields("MUKE_NAME") _
					)
		strUpdate = ""
		if WScript.Arguments.Named.Exists("update") then
			rsSrc.Fields("ID_NO") = strIdNoMax
			rsSrc.UpdateBatch
			strUpdate = " (更新)"
		end if
		call DispMsg("  " _
					& strIdNoMax _
					& strUpdate _
					)
	end if
End Function

Function NextId(byval strIdNo)
	dim	lngIdNo
	lngIdNo = clng(strIdNo) + 1
	NextId = "0" & lngIdNo
End Function

Function GetSql()
	dim	strSql

	strSql = "select *"
	strSql = strSql & " from y_syuka_tei"
	strSql = strSql & " where SND_YMD = '20120126'"
	strSql = strSql & " and SND_HMS = '105033'"
	strSql = strSql & " and TOK_CD = '7401UH'"
	strSql = strSql & " and KONPO_ID <> ''"
	strSql = strSql & " order by SEQ_NO,L_PAGE,KONPO_ID,convert(L_UCHI_NO,sql_numeric)"

	GetSql = strSql
End Function

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub

Function GetOption(byval strName _
				  ,byval strDefault _
				  )
	dim	strValue

	strValue = strDefault
	if WScript.Arguments.Named.Exists(strName) then
		strValue = WScript.Arguments.Named(strName)
	end if
	GetOption = strValue
End Function
