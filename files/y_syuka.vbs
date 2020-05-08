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
    Wscript.Echo "y_syuka(2011.12.24)"
    Wscript.Echo "y_syuka.vbs [/db:DbName]"
    Wscript.Echo "<—á>"
    Wscript.Echo "y_syuka.vbs /db:newsdc-osk"
End Function

Function Main()
	dim	strArg
	dim	strDbName

	strDbName	= "newsdc"

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
		case "recover"
		case "reverse"
		case "id"
		case "dt"
		case "to"
		case else
			call usage()
			exit function
		end select
	next
	Call ySyukaRecober()
End Function
Function ySyukaRecober()
	dim	strSql
	dim	cnnDb
	dim	strDbName
	dim	rsSrc
	dim	rsDst
	dim	strTableSrc
	dim	strTableDst

	strDbName = "newsdc"
	if WScript.Arguments.Named.Exists("db") then
		strDbName = WScript.Arguments.Named("db")
	end if
	if WScript.Arguments.Named.Exists("reverse") then
		strTableSrc = "del_syuka"
		strTableDst = "y_syuka"
	else
		strTableSrc = "y_syuka"
		strTableDst = "del_syuka"
	end if

	call DispMsg("CreateObject(ADODB.Connection)")
	Set cnnDb = Wscript.CreateObject("ADODB.Connection")
	call DispMsg("Open:" & strDbName)
	cnnDb.Open strDbName

	Set rsSrc = Wscript.CreateObject("ADODB.Recordset")
	strSql = GetOpenSql(strTableSrc)
	call DispMsg("Open Src:" & strSql)
'	rsSrc.Open strSql, cnnDb, adOpenKeyset, adLockBatchOptimistic
	rsSrc.Open strSql, cnnDb, adOpenKeyset

	Set rsDst = Wscript.CreateObject("ADODB.Recordset")
	call DispMsg("Open Dst:" & strTableDst)
	rsDst.Open strTableDst, cnnDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

	'----------------------------------------------------------------
	Do While Not rsSrc.EOF
		Call DispYSyuka(rsSrc)
		Call MoveRecord(cnnDb,rsSrc,rsDst)
		dim	strDelSql
		strDelSql = ""
		if WScript.Arguments.Named.Exists("update") then
			strDelSql = "delete from " & strTableSrc
			strDelSql = strDelSql & " where key_id_no = '" & RTrim(rsSrc.Fields("key_id_no")) & "'"
			strDelSql = strDelSql & " and JGYOBU = '" & RTrim(rsSrc.Fields("JGYOBU")) & "'"
			strDelSql = strDelSql & " and KEY_CYU_KBN = '" & RTrim(rsSrc.Fields("KEY_CYU_KBN")) & "'"
			strDelSql = strDelSql & " and KEY_MUKE_CODE = '" & RTrim(rsSrc.Fields("KEY_MUKE_CODE")) & "'"
			strDelSql = strDelSql & " and KEY_SS_CODE = '" & RTrim(rsSrc.Fields("KEY_SS_CODE")) & "'"
			strDelSql = strDelSql & " and KEY_HIN_NO = '" & RTrim(rsSrc.Fields("KEY_HIN_NO")) & "'"
			strDelSql = strDelSql & " and KEY_SYUKA_YMD = '" & RTrim(rsSrc.Fields("KEY_SYUKA_YMD")) & "'"
		end if
		rsSrc.MoveNext
		if strDelSql <> "" then
			Call DispMsg("Delete:strIdNo:" & strDelSql)
			cnnDb.Execute(strDelSql)
		end if
	Loop

	'----------------------------------------------------------------
	Call rsSrc.Close
	set rsSrc = Nothing
	Call rsDst.Close
	set rsDst = Nothing
	Call DispMsg("Close:" & cnnDb.DefaultDatabase)
	Call cnnDb.Close
	set cnnDb = Nothing
End Function

Function MoveRecord(cnnDb,rsSrc,rsDst)
	dim	fldSrc
	dim	strMsg

	if WScript.Arguments.Named.Exists("update") then
		call rsDst.Addnew
	end if
	for each fldSrc in rsSrc.Fields
		if WScript.Arguments.Named.Exists("update") then
			rsDst.Fields(fldSrc.Name) = fldSrc
		end if
	next
	On Error Resume Next
		strMsg = ""
		if WScript.Arguments.Named.Exists("update") then
			rsDst.UpdateBatch
			if Err.Number = 0 then
				strMsg = "Ok"
'				cnnDb.Execute("delete from y_syuka where key_id_no = '" & RTrim(rsSrc.Fields("key_id_no")) & "'")
'				if Err.Number = 0 then
'					strMsg = "Ok"
'				else
'					strMsg = "Delete Err:" & Err.Number & " " & Err.Description
'				end if
			else
				strMsg = "Err:" & Err.Number & " " & Err.Description
			end if
			Call DispMsg(strMsg)
		end if
		Err.Clear
	On Error Goto 0
End Function

Function DispYSyuka(byval rsSrc)
	call DispMsg("" _
			     & " " & rsSrc.Fields("KEY_ID_NO") _
			     & " " & rsSrc.Fields("KAN_KBN") _
			     & " " & rsSrc.Fields("KEY_SYUKA_YMD") _
			     & " " & rsSrc.Fields("KEY_MUKE_CODE") _
			     & " " & rsSrc.Fields("MUKE_NAME") _
				)
End Function

Function GetOpenSql(byval strTable)
	dim	strSql
	dim	strId
	dim	strDt

	strId = ""
	if WScript.Arguments.Named.Exists("id") then
		strId = WScript.Arguments.Named("id")
	end if
	strDt = ""
	if WScript.Arguments.Named.Exists("dt") then
		strDt = WScript.Arguments.Named("dt")
	end if
	dim	strTo
	strTo = ""
	if WScript.Arguments.Named.Exists("to") then
		strTo = WScript.Arguments.Named("to")
	end if

	strSql = "select *"
	strSql = strSql & " from " & strTable
	if strTo = "" then
		strSql = strSql & " where key_syuka_ymd <= '" & strDt & "'"
	else
		strSql = strSql & " where key_syuka_ymd between '" & strDt & "' and '" & strTo & "'"
		strSql = strSql & " and KAN_KBN = '9'"
'		strSql = strSql & " and KEY_MUKE_CODE = 'G11'"
	end if
'	strSql = strSql & "   and JGYOBA = '00036003'"
'	strSql = strSql & "   and DATA_KBN = '1'"
'	strSql = strSql & "   and HAN_KBN = '2'"
'	strSql = strSql & "   and key_id_no = '" & strId & "'"
	GetOpenSql = strSql
End Function
