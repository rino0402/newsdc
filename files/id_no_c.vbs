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
    Wscript.Echo "id_no_c.vbs <ID_NO> [変更ID_NO]"
    Wscript.Echo "<例>"
    Wscript.Echo "id_no_c.vbs 0073805 0495148"
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
	dim	rsYSyuka
	dim	rsYSyukaH
	dim	strSql
	dim	strIdNoMax

	strDbName = "newsdc-osk"
	call DispMsg("ChangeIdNo(" & strIdNoSrc & "," & strIdNoDst & ")")

	call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	call DispMsg("Open:" & strDbName)
	objDb.Open strDbName

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsYSyuka = Wscript.CreateObject("ADODB.Recordset")

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsYSyukaH = Wscript.CreateObject("ADODB.Recordset")
	strSql = GetSql(strIdNoSrc,"y_syuka_h")
	call DispMsg("Dst:" & strSql)
'	Set rsYSyukaH = objDb.Execute(strSql)
	rsYSyukaH.Open strSql, objDb, adOpenKeyset, adLockBatchOptimistic
	dim	strNewID
	do while rsYSyukaH.EOF = False
		call DispMsg("  " & _
					 rsYSyukaH.Fields("ID_NO") & " " & _
					 rsYSyukaH.Fields("CANCEL_F") & " " & _
					 rsYSyukaH.Fields("SYUKA_YMD") & " " & _
					 rsYSyukaH.Fields("MUKE_CODE") & " " & _
					 rsYSyukaH.Fields("MUKE_NAME") _
					)
		strNewID = rtrim(rsYSyukaH.Fields("ID_NO"))
		strNewID = "9" & right(strNewID,len(strNewID)-1)
		strSql = GetSql(rtrim(rsYSyukaH.Fields("ID_NO")),"y_syuka")
'		Set rsYSyuka = objDb.Execute(strSql)
		rsYSyuka.Open strSql, objDb, adOpenKeyset, adLockBatchOptimistic
		call DispMsg("  " & _
					 rsYSyuka.Fields("KEY_ID_NO") & " " & _
					 " " & " " & _
					 rsYSyuka.Fields("KEY_SYUKA_YMD") & " " & _
					 rsYSyuka.Fields("KEY_MUKE_CODE") & " " & _
					 rsYSyuka.Fields("MUKE_NAME") _
					)
		if WScript.Arguments.Named.Exists("update") then
			rsYSyuka.Fields("KEY_ID_NO") = strNewID
			rsYSyuka.Fields("ID_NO") = strNewID
			rsYSyuka.UpdateBatch
			rsYSyukaH.Fields("ID_NO") = strNewID
			rsYSyukaH.Fields("CANCEL_F") = ""
			rsYSyukaH.UpdateBatch
		end if
		rsYSyuka.Close
		call DispMsg("  " & _
					 strNewID _
					)
		rsYSyukaH.MoveNext
	loop

	call DispMsg("Close:" & strDbName)
	objDb.Close
	set objDb = Nothing
End Function

Function UpdateIdNo(byval rsYSyuka _
				   ,byval strIdNoNew _
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

Function GetSql(byval strIdNo,byval strTbl)
	dim	strSql

	if strTbl = "y_syuka" then
		strSql = "select *"
		strSql = strSql & " from y_syuka"
		strSql = strSql & " where KEY_ID_NO = '" & strIdNo & "'"
		strSql = strSql & " order by convert(KEY_ID_NO,sql_numeric)"
	else
		strSql = "select *"
		strSql = strSql & " from y_syuka_h"
		strSql = strSql & " where ID_NO like '" & strIdNo & "'"
		strSql = strSql & " and (SEK_KEN_NO + SEK_HIN_NO) in"
		strSql = strSql & " (select distinct (KEN_NO + HIN_NO) from y_syuka_tei where SND_YMD = '20120425')"
		strSql = strSql & " order by convert(ID_NO,sql_numeric)"
	end if
	
	GetSql = strSql
End Function

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub
