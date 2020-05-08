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
    Wscript.Echo "品目MST代表機種更新(2011.12.27)"
    Wscript.Echo "item_dmodel.vbs [/db:DbName]"
    Wscript.Echo "<例>"
    Wscript.Echo "item_dmodel.vbs /db:newsdc-ono"
End Function

Function Main()
	dim	strArg
	dim	strDbName
	dim	strJgyobu
	dim	strNaigai

	strDbName	= "newsdc"
	strJgyobu	= ""
	strNaigai	= "1"

	for each strArg in WScript.Arguments.UnNamed
		call usage()
		exit function
	next
	for each strArg in WScript.Arguments.Named
'		Wscript.Echo strArg,WScript.Arguments.Named(strArg)
		select case lcase(strArg)
		case "db"
		case "jgyobu"
		case "naigai"
		case "limit"
		case "update"
		case "pn"
		case else
			call usage()
			exit function
		end select
	next
	call ItemDModel()
End Function

Function ItemDModel()
	dim	strDbName
	dim	strSql
	dim	cnnDb
	dim	rsItem
	dim	lngCnt
	dim	rsPn

	Wscript.Echo "ItemDModel()"
	call DispMsg("CreateObject(ADODB.Connection)")
	Set cnnDb = Wscript.CreateObject("ADODB.Connection")

	strDbName = "newsdc"
	if WScript.Arguments.Named.Exists("db") then
		strDbName = WScript.Arguments.Named("db")
	end if
	call DispMsg("Open:" & strDbName)
	cnnDb.Open strDbName
	'----------------------------------------------------------------
	Set rsItem = Wscript.CreateObject("ADODB.Recordset")
	strSql = GetSql()
	call DispMsg("SQL:" & strSql)
	rsItem.Open strSql, cnnDb, adOpenKeyset, adLockBatchOptimistic
	lngCnt = 0
	do while rsItem.EOF = False
		lngCnt = lngCnt + 1
		call DispItem(lngCnt,rsItem)
		strSql = GetSqlPn(rsItem.Fields("JGYOBU") _
				         ,rsItem.Fields("HIN_GAI") _
						 )
		set rsPn = cnnDb.Execute(strSql)
		if rsPn.EOF = False then
			call UpdateItem(rsItem,rsPn.Fields("DMODEL"))
		else
			call DispMsg("Not Found:" & strSql)
		end if

		if CheckLimit(lngCnt) then
			exit do
		end if
		call rsItem.MoveNext
	loop
	rsItem.Close
	set rsItem = Nothing
	'----------------------------------------------------------------
	call DispMsg("Close:" & cnnDb.DefaultDatabase)
	cnnDb.Close
	set cnnDb = Nothing
End Function
Function UpdateItem(byval rsItem,byval strDModel)
	dim	strMsg
	strMsg = ""
	strDModel = rtrim(strDModel)
	if rtrim(rsItem.Fields("D_MODEL")) <> strDModel then
		strMsg = "代表機種コード変更"
		if WScript.Arguments.Named.Exists("update") then
			strMsg = "代表機種コード変更(Update)"
			rsItem.Fields("D_MODEL") = strDModel
			rsItem.Fields("UPD_TANTO") = "D_MDL"
			rsItem.Fields("UPD_DATETIME") = GetDateTime(Now())
			rsItem.UpdateBatch
		end if
	end if
	call DispDModel(strDModel,strMsg)
End Function

Function CheckLimit(byval lngCnt)
	dim	retBool

	retBool = False
	if WScript.Arguments.Named.Exists("limit") then
'		Wscript.Echo "limit:" & lngCnt & ":" &  WScript.Arguments.Named("limit")
		if lngCnt >= clng(WScript.Arguments.Named("limit")) then
'			Wscript.Echo "True"
			retBool = True
		end if
	end if
	CheckLimit = retBool
End Function

Function GetSql()
	dim	strJgyobu
	dim	strNaigai
	dim	strSql
	dim	strPn
	dim	strJCode

	strJgyobu	= ""
	strNaigai	= ""
	strPn	= ""
	if WScript.Arguments.Named.Exists("pn") then
		strPn = WScript.Arguments.Named("pn")
	end if
	if WScript.Arguments.Named.Exists("Jgyobu") then
		strJgyobu = WScript.Arguments.Named("Jgyobu")
	end if
	if WScript.Arguments.Named.Exists("Naigai") then
		strNaigai = WScript.Arguments.Named("Naigai")
	end if
	strJcode = "00023410"
	strSql = "select *"
	strSql = strSql & " from item"
	strSql = strSql & " where JGYOBU = '" & strJgyobu & "'"
	strSql = strSql & "   and NAIGAI = '" & strNaigai & "'"
	strSql = strSql & "   and HIN_GAI in (select distinct Pn from pn3 where ShisanJCode = '" & strJCode & "')"
	if strPn <> "" then
		strSql = strSql & "   and HIN_GAI = '" & strPn & "'"
	end if
	GetSql = strSql
End Function

Function DispItem(byval lngCnt _
				 ,byval rsItem _
				 )
	Call DispMsg(right(space(6) & lngCnt,6) _
				     & " " & rsItem.Fields("JGYOBU") _
 					 & " " & rsItem.Fields("NAIGAI") _
 					 & " " & rsItem.Fields("HIN_GAI") _
 					 & " " & rsItem.Fields("L_KAISHA_CODE") _
 					 & " " & rsItem.Fields("L_JGYOBU_CODE") _
 					 & " " & rsItem.Fields("D_MODEL") _
					)
End Function
Function DispDModel(byval strDModel _
				   ,byval strMsg _
				   )
	Call DispMsg(space(6)_
				 & " " & space(1) _
				 & " " & space(1) _
				 & " " & space(20) _
				 & " " & space(2) _
				 & " " & space(2) _
				 & " " & strDModel _
				 & " " & strMsg _
				)
End Function
Function GetSqlPn(byval strJgyobu _
			   	 ,byval strPn _
				 )
	dim	strSql
	dim	strJCode

	select case strJgyobu
	case "D"
		strJCode = "00023510"
	case "4"
		strJCode = "00023410"
	case "7"
		strJCode = "00023210"
	end select

	strSql = "select *"
	strSql = strSql & " from pn3"
	strSql = strSql & " where ShisanJCode = '" & strJCode & "'"
	strSql = strSql & "   and pn = '" & rtrim(strPn) & "'"
	GetSqlPn = strSql
End Function
