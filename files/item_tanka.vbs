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
    Wscript.Echo "品目MST単価更新(2011.12.24)"
    Wscript.Echo "item_tanka.vbs [/db:DbName]"
    Wscript.Echo "<例>"
    Wscript.Echo "item_tanka.vbs /db:newsdc-ono"
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
			strDbName = WScript.Arguments.Named(strArg)
		case "jgyobu"
			strJgyobu = WScript.Arguments.Named(strArg)
		case "naigai"
			strNaigai = WScript.Arguments.Named(strArg)
		case "limit"
		case "update"
		case "pn"
		case "none"
		case else
			call usage()
			exit function
		end select
	next
	call ItemTanka(strDbName _
				  ,strJgyobu _
				  ,strNaigai _
				  )
End Function
Function ItemTanka(byval strDbName _
				  ,byval strJgyobu _
				  ,byval strNaigai _
				  )
	dim	strSql
	dim	cnnDb
	dim	rsItem
	dim	lngCnt
	dim	rsPCompo
	dim	dblGenka
	dim	dblBaika

	Wscript.Echo "ItemTanka(" & strDbName & _
				 		  "," &  strJgyobu & _
				 		  "," &  strNaigai & _
						  ")"
	call DispMsg("CreateObject(ADODB.Connection)")
	Set cnnDb = Wscript.CreateObject("ADODB.Connection")
	call DispMsg("Open:" & strDbName)
	cnnDb.Open strDbName
	call DispMsg("CursorLocation:" & cnnDb.CursorLocation)
	cnnDb.CursorLocation = adUseClient
	call DispMsg("CursorLocation:" & cnnDb.CursorLocation)
	'----------------------------------------------------------------
	Set rsItem = Wscript.CreateObject("ADODB.Recordset")
	strSql = GetSql(strJgyobu,strNaigai)
	call DispMsg("SQL:" & strSql)
	rsItem.Open strSql, cnnDb, adOpenKeyset, adLockBatchOptimistic
	lngCnt = 0
	do while rsItem.EOF = False
		lngCnt = lngCnt + 1
		call DispItem(lngCnt,rsItem)
		strSql = GetSqlPCompo(rsItem.Fields("JGYOBU") _
 					         ,rsItem.Fields("NAIGAI") _
 					         ,rsItem.Fields("SHIMUKE_CODE") _
 					         ,rsItem.Fields("HIN_GAI") _
							 )
		dblGenka = 0
		dblBaika = 0
		set rsPCompo = cnnDb.Execute(strSql)
		do while rsPCompo.EOF = False
'			call DispPCompo(rsPCompo)
			dblGenka = dblGenka + CalcGenka(rsPCompo)
			dblBaika = dblBaika + CalcBaika(rsPCompo)
			call rsPCompo.MoveNext
		loop
		call UpdateTanka(rsItem,dblGenka,dblBaika)

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
Function UpdateTanka(byval rsItem,byval dblGenka,byval dblBaika)
	dim	strMsg
	strMsg = ""
	if GetValue(rsItem.Fields("S_SHIZAI_GENKA")) <> dblGenka then
		strMsg = "資材原価変更"
		if WScript.Arguments.Named.Exists("update") then
			strMsg = "資材原価変更(Update)"
			rsItem.Fields("S_SHIZAI_GENKA") = dblGenka
			rsItem.Fields("UPD_TANTO") = "SETGN"
			rsItem.Fields("UPD_DATETIME") = GetDateTime(Now())
			rsItem.UpdateBatch
		end if
	end if
	call DispTanka(dblGenka,dblBaika,strMsg)
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

Function CalcGenka(byval rsPCompo)
	CalcGenka =	GetValue(rsPCompo.Fields("G_ST_SHITAN")) * GetValue(rsPCompo.Fields("KO_QTY"))
End Function
Function CalcBaika(byval rsPCompo)
	CalcBaika =	GetValue(rsPCompo.Fields("G_ST_URITAN")) * GetValue(rsPCompo.Fields("KO_QTY"))
End Function
Function GetValue(byval strValue)
	dim	strNumber
	dim	dblValue

	dblValue = 0
	strNumber = trim(strValue)
	if isnumeric(strNumber) = True then
		dblValue = ccur(strNumber)
	end if
	GetValue = dblValue
End Function

Function DispPCompo(byval rsPCompo)
	call DispMsg(space(6) _
			     & " " & rsPCompo.Fields("JGYOBU") _
			     & " " & rsPCompo.Fields("NAIGAI") _
			     & " " & rsPCompo.Fields("HIN_GAI") _
			     & " " & FrmNumber(rsPCompo.Fields("G_ST_SHITAN"),8,2) _
			     & " " & FrmNumber(rsPCompo.Fields("G_ST_URITAN"),8,2) _
				 & " " & FrmNumber(rsPCompo.Fields("KO_QTY"),8,2) _
				)
End Function

Function FrmNumber(byval strValue _
				  ,byval intCol _
				  ,byval intBeam _
				  )
	dim	strNumber
	strNumber = trim(strValue)
	if isnumeric(strNumber) = True then
		strNumber = ccur(strNumber)
		strNumber = FormatNumber(strNumber,intBeam,0,0)
		strNumber = right(space(intCol) & strNumber,intCol)
	else
		strNumber = left(strNumber & space(intCol),intCol)
	end if
	FrmNumber = strNumber
End Function

Function DispItem(byval lngCnt _
				 ,byval rsItem _
				 )
	Call DispMsg(right(space(6) & lngCnt,6) _
				     & " " & rsItem.Fields("JGYOBU") _
 					 & " " & rsItem.Fields("NAIGAI") _
 					 & " " & rsItem.Fields("HIN_GAI") _
 					 & " " & FrmNumber(rsItem.Fields("S_SHIZAI_GENKA"),8,2) _
 					 & " " & FrmNumber(rsItem.Fields("S_SHIZAI_BAIKA"),8,2) _
 					 & " " & rsItem.Fields("SHIMUKE_CODE") _
 					 & " " & rsItem.Fields("S_KOUSU_SET_DATE") _
					)
End Function

Function DispTanka(byval dblGenka _
				  ,byval dblBaika _
				  ,byval strMsg _
				  )
	Call DispMsg(space(6)_
				 & " " & space(1) _
				 & " " & space(1) _
				 & " " & space(20) _
				 & " " & FrmNumber(dblGenka,8,2) _
				 & " " & FrmNumber(dblBaika,8,2) _
				 & " " & strMsg _
				)
End Function


Function GetSql(byval strJgyobu _
			   ,byval strNaigai)
	dim	strSql
	dim	strPn
	dim	strNone

	strPn	= ""
	strNone	= ""
	if WScript.Arguments.Named.Exists("pn") then
		strPn = WScript.Arguments.Named("pn")
	end if
	if WScript.Arguments.Named.Exists("none") then
		strNone = "none"
	end if


	strSql = "select *"
	strSql = strSql & " from item"
	strSql = strSql & " where JGYOBU = '" & strJgyobu & "'"
	strSql = strSql & "   and NAIGAI = '" & strNaigai & "'"
	if strPn <> "" then
		strSql = strSql & "   and HIN_GAI = '" & strPn & "'"
	end if
	if strNone <> "" then
		strSql = strSql & "   and convert(S_SHIZAI_GENKA,sql_numeric) = 0"
		strSql = strSql & "   and convert(S_SHIZAI_BAIKA,sql_numeric) <> 0"
	end if
	strSql = strSql & "   and rtrim(S_KOUSU_SET_DATE) <> ''"
	strSql = strSql & "   and rtrim(S_SHIZAI_BAIKA) <> ''"
	GetSql = strSql
End Function

Function GetSqlPCompo(byval strJgyobu _
			   		 ,byval strNaigai _
			   		 ,byval strShimukeCode _
			   		 ,byval strPn _
					 )
	dim	strSql

	select case trim(strShimukeCode)
	case ""
		select case strJgyobu
		case "D"
			strShimukeCode = "01"
		case "4"
			strShimukeCode = "02"
		case "7"
			strShimukeCode = "01"
		end select
	end select

	strSql = "select "
	strSql = strSql & " k.ko_jgyobu jgyobu"
	strSql = strSql & ",k.ko_naigai naigai"
	strSql = strSql & ",k.ko_hin_gai hin_gai"
	strSql = strSql & ",s.G_ST_SHITAN G_ST_SHITAN"
	strSql = strSql & ",s.G_ST_URITAN G_ST_URITAN"
	strSql = strSql & ",k.KO_QTY KO_QTY"
	strSql = strSql & " from p_compo_k as k"
	strSql = strSql & " inner join item as s"
	strSql = strSql & " on (s.jgyobu = k.ko_jgyobu"
	strSql = strSql & " and s.naigai = k.ko_naigai"
	strSql = strSql & " and s.hin_gai = k.ko_hin_gai"
	strSql = strSql & " and s.sei_kbn <> '1'"
	strSql = strSql & " )"
	strSql = strSql & " where (k.data_kbn = '1' or (k.data_kbn = '3' and k.KO_SYUBETSU in ('03','90')))"
	strSql = strSql & "   and k.JGYOBU = '" & strJgyobu & "'"
	strSql = strSql & "   and k.NAIGAI = '" & strNaigai & "'"
	strSql = strSql & "   and k.SHIMUKE_CODE = '" & strShimukeCode & "'"
	strSql = strSql & "   and k.HIN_GAI = '" & strPn & "'"
	GetSqlPCompo = strSql
End Function
