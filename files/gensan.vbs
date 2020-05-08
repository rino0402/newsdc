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
dim	lngRet
lngRet = Main()
WScript.Quit lngRet

'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "原産国"
	Wscript.Echo "gensan.vbs [option]"
	Wscript.Echo " /list"
	Wscript.Echo " /nyuka[:update]"
	Wscript.Echo " /top:<num>"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg

	For Each strArg In WScript.Arguments.UnNamed
    	select case strArg
		case else
			usage()
			Main = 1
			exit Function
		end select
	Next
	For Each strArg In WScript.Arguments.Named
    	select case lcase(strArg)
		case "db"
		case "list"
		case "nyuka"
		case "top"
		case else
			usage()
			Main = 1
			exit Function
		end select
	Next
	select case GetFunction()
	case "list"
		Call List()
	case "nyuka"
		Call Nyuka()
	end select
	Main = 0
End Function

Private Function GetFunction()
	GetFunction = "list"
	if WScript.Arguments.Named.Exists("nyuka") then
		GetFunction = "nyuka"
	end if
End Function

Private Function Nyuka()
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	if GetOption("nyuka","") = "update" then
		Call DispMsg("CreateObject(" & strSql & ")")
		Set rsList = Wscript.CreateObject("ADODB.Recordset")
		rsList.Open strSql, objDb, adOpenKeyset, adLockBatchOptimistic
	else
		Call DispMsg("objDb.Execute(" & strSql & ")")
		set rsList = objDb.Execute(strSql)
	end if

	do while rsList.Eof = False
		dim	strMsg
		strMsg = ""
		strMsg ="■" _
			 & " " & rsList.Fields("JGYOBU") _
			 & " " & rsList.Fields("NAIGAI") _
			 & " " & rsList.Fields("HIN_GAI") _
			 & " " & rsList.Fields("GENSANKOKU") _
			 & " " & rsList.Fields("INS_TANTO") _
			 & " " & rsList.Fields("Ins_DateTime") _
			 & " " & rsList.Fields("UPD_TANTO") _
			 & " " & rsList.Fields("UPD_DATETIME") _
			 & ""
		Call DispMsg(strMsg)
		dim	rsNyuka
		set rsNyuka = GetNyuka(objDb,rsList)
		if rsNyuka.Eof = False then
			strMsg ="入" _
				 & " " & rsNyuka.Fields("JGYOBU") _
				 & " " & rsNyuka.Fields("NAIGAI") _
				 & " " & rsNyuka.Fields("HIN_NO") _
				 & " " & rsNyuka.Fields("GENSANKOKU") _
				 & " " & rsNyuka.Fields("UPD_TANTO") _
				 & " " & rsNyuka.Fields("UPD_DATETIME") _
				 & " " & rsNyuka.Fields("INS_TANTO") _
				 & " " & rsNyuka.Fields("Ins_DateTime") _
				 & ""
			dim	strChk
			strChk = NyukaCheck(rsList,rsNyuka)
			strMsg = strMsg & " " & strChk
			if GetOption("nyuka","") = "update" then
				if strChk = "<" then
					strMsg = strMsg & " " & "GENSA " & rsNyuka.Fields("Ins_DateTime")
					rsList.Fields("UPD_TANTO")		= "GENSA"
					rsList.Fields("UPD_DATETIME")	= rsNyuka.Fields("Ins_DateTime")
					rsList.UpdateBatch
				end if
			end if
			Call DispMsg(strMsg)
		end if
		set rsNyuka = Nothing
		Call rsList.MoveNext
	loop

	Call DispMsg("CloseAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = Nothing
End Function

Private Function NyukaCheck(rsList,rsNyuka)
	NyukaCheck = "="
	if Left(rsList.Fields("Ins_DateTime"),8) = Left(rsNyuka.Fields("Ins_DateTime"),8) then
		exit function
	end if
	NyukaCheck = "*"
	if Left(rsList.Fields("UPD_DATETIME"),8) < Left(rsNyuka.Fields("Ins_DateTime"),8) then
		NyukaCheck = "<"
	end if
End Function

Private Function List()
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	Call DispMsg("objDb.Execute(" & strSql & ")")
	set rsList = objDb.Execute(strSql)

	do while rsList.Eof = False
		Call DispMsg("■" _
			 & " " & rsList.Fields("JGYOBU") _
			 & " " & rsList.Fields("NAIGAI") _
			 & " " & rsList.Fields("HIN_GAI") _
			 & " " & rsList.Fields("GENSANKOKU") _
			 & " " & rsList.Fields("INS_TANTO") _
			 & " " & rsList.Fields("Ins_DateTime") _
			 & " " & rsList.Fields("UPD_TANTO") _
			 & " " & rsList.Fields("UPD_DATETIME") _
					)
		dim	strMsg
		strMsg = ""
		strMsg = CheckExist(strMsg,objDb,rsList,"入荷")
		strMsg = CheckExist(strMsg,objDb,rsList,"PN")
		strMsg = CheckExist(strMsg,objDb,rsList,"品目")
		strMsg = CheckExist(strMsg,objDb,rsList,"国歴")
		strMsg = CheckExist(strMsg,objDb,rsList,"PN歴")
		if strMsg = "" then
			Call DispMsg("--")
		end if
		Call rsList.MoveNext
	loop

	Call DispMsg("CloseAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = Nothing
End Function

Private Function CheckExist(byVal strMsg,objDb,rsList,byVal strKubun)
	if strMsg = "" then
		select case strKubun
		case "入荷"
			dim	rsNyuka
			set rsNyuka = GetNyuka(objDb,rsList)
			if rsNyuka.Eof = False then
				Call DispMsg("入" _
					 & " " & rsNyuka.Fields("JGYOBU") _
					 & " " & rsNyuka.Fields("NAIGAI") _
					 & " " & rsNyuka.Fields("HIN_NO") _
					 & " " & rsNyuka.Fields("GENSANKOKU") _
					 & " " & rsNyuka.Fields("INS_TANTO") _
					 & " " & rsNyuka.Fields("Ins_DateTime") _
					 & " " & rsNyuka.Fields("UPD_TANTO") _
					 & " " & rsNyuka.Fields("UPD_DATETIME") _
					)
				strMsg = strKubun
			end if
			set rsNyuka = Nothing
		case "品目"
			dim	rsItem
			set rsItem = GetItem(objDb,rsList)
			if rsItem.Eof = False then
				Call DispMsg("品" _
					 & " " & rsItem.Fields("JGYOBU") _
					 & " " & rsItem.Fields("NAIGAI") _
					 & " " & rsItem.Fields("HIN_GAI") _
					 & " " & makeMsg(rsItem.Fields("TORI_GENSANKOKU"),-20) _
					 & " " & rsItem.Fields("INS_TANTO") _
					 & " " & rsItem.Fields("Ins_DateTime") _
					 & " " & rsItem.Fields("UPD_TANTO") _
					 & " " & rsItem.Fields("UPD_DATETIME") _
							)
				strMsg = strKubun
			end if
			set rsItem = Nothing
		case "PN"
			dim	rsPn
			set rsPn = GetrsPn(objDb,rsList)
			if rsPn.Eof = False then
				Call DispMsg("PN" _
					 & " " & " " _
					 & " " & " " _
					 & " " & rsPn.Fields("Pn") _
					 & " " & rsPn.Fields("CountryName2") _
					 & "   " & rsPn.Fields("MadeInCode") _
					 & " " & rsPn.Fields("UPD_TM") _
					 & " " & rsPn.Fields("JCode") _
					 & " " & rsPn.Fields("ShisanJCode") _
							)
				strMsg = strKubun
			end if
			set rsPn = Nothing
		case "品歴"
			dim	rsPnH
			set rsPnH = GetrsPnH(objDb,rsList)
			if rsPnH.Eof = False then
				Call DispMsg("品歴  " _
					 & " " & rsPnH.Fields("Pn") _
					 & " " & rsPnH.Fields("CountryName2") _
					 & "   " & rsPnH.Fields("MadeInCode") _
					 & " " & rsPnH.Fields("AlterDate") _
					 & " " & rsPnH.Fields("JCode") _
					 & " " & rsPnH.Fields("ShisanJCode") _
							)
				strMsg = strKubun
			end if
			set rsPnH = Nothing
		case "PN歴"
			dim	rsPnHst
			set rsPnHst = GetrsPnHst(objDb,rsList)
			if rsPnHst.Eof = False then
				Call DispMsg("PN歴  " _
					 & " " & rsPnHst.Fields("Pn") _
					 & "   " & RTrim(rsPnHst.Fields("BefValue")) _
					 & " " & rsPnHst.Fields("TimeStamp") _
					 & " " & rsPnHst.Fields("JCode") _
					 & " " & rsPnHst.Fields("ShisanJCode") _
							)
				strMsg = strKubun
			end if
			set rsPnHst = Nothing
		end select
	end if
	CheckExist = strMsg
End Function

Private Function GetrsPnHst(objDb,rsList)
	set GetrsPnHst = objDb.Execute(makePnHstSql(rsList,"="))
	if GetrsPnHst.Eof then
'		Call DispMsg("+++")
		set GetrsPnHst = objDb.Execute(makePnHstSql(rsList,"in"))
	end if
End Function

Private Function makePnHstSql(rsList,byval strJgyobu)
	dim	strShisanJCode
	strShisanJCode = ""
	select case strJgyobu
	case "="
		select case rsList.Fields("JGYOBU")
		case "4"
			strShisanJCode = " = '00023410'"
		case "D"
			strShisanJCode = " = '00023510'"
		case "5"
			strShisanJCode = " = '00021397'"
		end select
	case "in"
		strShisanJCode = " in ('00023410','00023510','00021397')"
	case else
	end select
	dim	strSql
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from PnHistory m"
	strSql = strSql & " where m.JCode = '00036003'"
	strSql = strSql & "   and m.ShisanJCode " & strShisanJCode
	strSql = strSql & "   and m.Pn = '" & rsList.Fields("HIN_GAI") & "'"
	strSql = strSql & "   and FldName = 'MadeInCode' "
	strSql = strSql & "   and BefValue = (select CountryName2 from Country where CountryCode = '" & RTrim(rsList.Fields("GENSANKOKU")) & "')"
	makePnHstSql = strSql
End Function


Private Function GetrsPn(objDb,rsList)
	set GetrsPn = objDb.Execute(makePnSql(rsList,"="))
	if GetrsPn.Eof then
'		Call DispMsg("+++")
		set GetrsPn = objDb.Execute(makePnSql(rsList,"in"))
	end if
End Function

Private Function makePnSql(rsList,byval strJgyobu)
	dim	strShisanJCode
	strShisanJCode = ""
	select case strJgyobu
	case "="
		select case rsList.Fields("JGYOBU")
		case "4"
			strShisanJCode = " = '00023410'"
		case "D"
			strShisanJCode = " = '00023510'"
		case "5"
			strShisanJCode = " = '00021397'"
		end select
	case "in"
		strShisanJCode = " in ('00023410','00023510','00021397')"
	case else
	end select
	dim	strSql
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from pn5 m"
	strSql = strSql & " left outer join Country c on (m.MadeInCode=c.CountryCode)"
	strSql = strSql & " where m.JCode = '00036003'"
	strSql = strSql & "   and m.ShisanJCode " & strShisanJCode
	strSql = strSql & "   and m.Pn = '" & rsList.Fields("HIN_GAI") & "'"
	strSql = strSql & "   and c.CountryName2 = '" & rsList.Fields("GENSANKOKU") & "'"
	makePnSql = strSql
End Function

Private Function GetrsPnH(objDb,rsList)
	set GetrsPnH = objDb.Execute(makePnHSql(rsList,"="))
	if GetrsPnH.Eof then
'		Call DispMsg("+++")
		set GetrsPnH = objDb.Execute(makePnHSql(rsList,"in"))
	end if
End Function

Private Function makePnHSql(rsList,byval strJgyobu)
	'select top 100 m.JCode as "事業場",m.ShisanJCode as "資産管理事業場"
	',m.Pn as "品番",m.MadeInCode + rtrim(' ' + c.CountryName2) as "原産国コード"
	',m.AlterDate as "登録日時"
	' From PnMadeInCode m
	' left outer join Country c on (m.MadeInCode=c.CountryCode) where m.ShisanJCode = '00023410' and m.Pn = 'AQS10-134-0U' order by "登録日時" desc,"資産管理事業場","品番"
	dim	strShisanJCode
	strShisanJCode = ""
	select case strJgyobu
	case "="
		select case rsList.Fields("JGYOBU")
		case "4"
			strShisanJCode = " = '00023410'"
		case "D"
			strShisanJCode = " = '00023510'"
		case "5"
			strShisanJCode = " = '00021397'"
		end select
	case "in"
		strShisanJCode = " in ('00023410','00023510','00021397')"
	case else
	end select
	dim	strSql
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from PnMadeInCode m"
	strSql = strSql & " left outer join Country c on (m.MadeInCode=c.CountryCode)"
	strSql = strSql & " where m.JCode = '00036003'"
	strSql = strSql & "   and m.ShisanJCode " & strShisanJCode
	strSql = strSql & "   and m.Pn = '" & rsList.Fields("HIN_GAI") & "'"
	strSql = strSql & "   and c.CountryName2 = '" & rsList.Fields("GENSANKOKU") & "'"
	makePnHSql = strSql
End Function

Private Function GetItem(objDb,rsList)
	set GetItem = objDb.Execute(makeItemSql(rsList,"="))
	if GetItem.Eof then
		set GetItem = objDb.Execute(makeItemSql(rsList,"in"))
	end if
End Function

Private Function makeItemSql(rsList,byval strJgyobu)
	select case strJgyobu
	case "="
		strJgyobu = " = '" & rsList.Fields("JGYOBU") & "'"
	case "in"
		strJgyobu = " in ('4','5','D')"
	case else
	end select
	dim	strSql
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from item"
	strSql = strSql & " where JGYOBU " & strJgyobu & ""
	strSql = strSql & "   and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
	strSql = strSql & "   and HIN_GAI = '" & rsList.Fields("HIN_GAI") & "'"
	strSql = strSql & "   and TORI_GENSANKOKU = '" & rsList.Fields("GENSANKOKU") & "'"
	strSql = strSql & " order by Ins_DateTime desc"
	makeItemSql = strSql
End Function

Private Function GetNyuka(objDb,rsList)
	set GetNyuka = objDb.Execute(makeNyukaSql(rsList,"="))
	if GetNyuka.Eof then
'		Call DispMsg("+++")
		select case rsList.Fields("JGYOBU")
		case "4","5","D"
			set GetNyuka = objDb.Execute(makeNyukaSql(rsList,"in"))
		end select
	end if
End Function

Private Function makeNyukaSql(rsList,byval strJgyobu)
	select case strJgyobu
	case "="
		strJgyobu = " = '" & rsList.Fields("JGYOBU") & "'"
	case "in"
		strJgyobu = " in ('4','5','D')"
	case else
	end select
	dim	strSql
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from y_nyuka"
	strSql = strSql & " where JGYOBU " & strJgyobu & ""
	strSql = strSql & "   and NAIGAI = '" & rsList.Fields("NAIGAI") & "'"
	strSql = strSql & "   and HIN_NO = '" & rsList.Fields("HIN_GAI") & "'"
	strSql = strSql & "   and GENSANKOKU = '" & rsList.Fields("GENSANKOKU") & "'"
	strSql = strSql & " order by Ins_DateTime desc"
	makeNyukaSql = strSql
End Function


Private Function makeSql()
	dim	strSql
	dim	strTop
	strTop = GetOption("top","")
	if strTop <> "" then
		strTop = " top " & strTop
	end if
	strSql = "select" & strTop
	strSql = strSql & " *"
	strSql = strSql & " from gensan"
	strSql = strSql & " where JGYOBU+NAIGAI+HIN_GAI in ("
	strSql = strSql & " select"
	strSql = strSql & " JGYOBU+NAIGAI+HIN_GAI"
	strSql = strSql & " from gensan"
	strSql = strSql & " where JGYOBU<>''"
	strSql = strSql & "   and rtrim(GENSANKOKU) <> ''"
	strSql = strSql & " group by JGYOBU,NAIGAI,HIN_GAI"
	strSql = strSql & " having count(*) > 1"
	strSql = strSql & " )"
	strSql = strSql & " order by JGYOBU,NAIGAI,HIN_GAI,Ins_DateTime"
	makeSql = strSql
End Function
