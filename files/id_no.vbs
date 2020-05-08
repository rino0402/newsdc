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
    Wscript.Echo "ID_NO変更(2011.12.22)"
    Wscript.Echo "id_no.vbs <ID_NO> [変更ID_NO]"
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
		case "y_syuka"
		case else
			call usage()
			exit function
		end select
	next
	if strIdNoSrc = "" then
		usage()
		exit Function
	end if

	call ChangeIdNo(strIdNoSrc,strIdNoDst)
End Function

Function ChangeIdNo(byval strIdNoSrc, _
					byval strIdNoDst)
	dim	objDb
	dim	strDbName
	dim	rsSrc
	dim	rsDst
	dim	strSql
	dim	strIdNoMax

	strDbName = "newsdc"
	call DispMsg("ChangeIdNo(" & strIdNoSrc & "," & strIdNoDst & ")")

	call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	call DispMsg("Open:" & strDbName)
	objDb.Open strDbName

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsSrc = Wscript.CreateObject("ADODB.Recordset")

	call DispMsg("CreateObject(ADODB.Recordset)")
	strSql = GetSql(strIdNoDst)
	call DispMsg("Dst:" & strSql)
	Set rsDst = objDb.Execute(strSql)
	strIdNoMax = strIdNoDst
	do while rsDst.EOF = False
		if WScript.Arguments.Named.Exists("y_syuka") then
			call DispMsg("  " & _
						 rsDst.Fields("KEY_ID_NO") & " " & _
						 rsDst.Fields("KEY_SYUKA_YMD") & " " & _
						 rsDst.Fields("KEY_HIN_NO") & " " & _
						 rsDst.Fields("KEY_MUKE_CODE") & " " & _
						 rsDst.Fields("MUKE_NAME") _
						)
			strIdNoMax = rsDst.Fields("KEY_ID_NO")
		else
			call DispMsg("  " & _
						 rsDst.Fields("ID_NO") & " " & _
						 rsDst.Fields("SYUKA_YMD") & " " & _
						 rsDst.Fields("HIN_NO") & " " & _
						 rsDst.Fields("MUKE_CODE") & " " & _
						 rsDst.Fields("MUKE_NAME") _
						)
			strIdNoMax = rsDst.Fields("ID_NO")
		end if
		rsDst.MoveNext
	loop

	strSql = GetSql(strIdNoSrc)
	call DispMsg("Sql:" & strSql)
	rsSrc.Open strSql, objDb, adOpenKeyset, adLockOptimistic
	do while rsSrc.EOF = False
		strIdNoMax = NextId(strIdNoMax)
		Call UpdateIdNo(rsSrc,strIdNoMax)
		rsSrc.MoveNext
	loop
	rsSrc.Close
	set rsSrc = Nothing

	call DispMsg("Close:" & strDbName)
	objDb.Close
	set objDb = Nothing
End Function

Function UpdateIdNo(rsSrc _
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
					 rsSrc.Fields("HIN_NO") & " " & _
					 rsSrc.Fields("MUKE_CODE") & " " & _
					 rsSrc.Fields("MUKE_NAME") _
					)
		strUpdate = ""
		if WScript.Arguments.Named.Exists("update") then
			rsSrc.Fields("ID_NO") = strIdNoMax
'			rsSrc.UpdateBatch
			rsSrc.Update
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
	if strIdNo = "" then
		lngIdNo = 1
	else
		lngIdNo = cLng(replace(strIdNo,"%","")) + 1
	end if
	NextId = "0" & lngIdNo
End Function

Function GetSql(byval strIdNo)
	dim	strSql

	if WScript.Arguments.Named.Exists("y_syuka") then
		strSql = "select distinct *"
		strSql = strSql & " from y_syuka"
		strSql = strSql & " where KEY_ID_NO like '" & strIdNo & "'"
		strSql = strSql & " order by convert(KEY_ID_NO,sql_decimal)"
	else
		strSql = "select distinct *"
		strSql = strSql & " from y_syuka_h"
		strSql = strSql & " where ID_NO like '" & strIdNo & "'"
		strSql = strSql & " order by convert(ID_NO,sql_decimal)"
	end if
	
	GetSql = strSql
End Function

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub
