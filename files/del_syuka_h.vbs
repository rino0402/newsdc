Option Explicit
' del_syuka 出荷予定データ復旧
' 2011.05.20 新規作成
' 2011.07.21 大阪PC対応
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

Call Main()
WScript.Quit 0

function usage()
	Wscript.Echo "出荷予定データ削除"
	Wscript.Echo "del_syuka_h.vbs [/DB:<database>] [/DOMOVE] [/DEL] [/DT:yyyymmdd]"
	Wscript.Echo " /DT:yyyymmdd"
	Wscript.Echo " /BIN:01"
	Wscript.Echo " /DOMOVE"
	Wscript.Echo " /DEL"
	Wscript.Echo "sc32 del_syuka_h.vbs /db:newsdc5 /DT:20140402 /del"
	Wscript.Echo "sc32 del_syuka_h.vbs /db:newsdc5 /domove"
	Wscript.Echo "GetDate(Now())=" & GetDate(Now())
end function

Function Main()
	dim	strArg

	for each strArg in WScript.Arguments.UnNamed
		select case strArg
		case else
			call usage()
			exit Function
		end select
	next
	for each strArg in WScript.Arguments.Named
		select case ucase(strArg)
		case "DT"
		case "BIN"
		case "DB"
		case "DOMOVE"
		case "DEL"
		case else
			call usage()
			exit function
		end select
	next
	Call MoveYSyuka()
'	call ySyukaRecober(strDbName)
End Function

Function MoveYSyuka()
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	strDbName
	strDbName = GetOption("DB","newsdc")
	dim	objDb
	DispMsg "OpenAdodb(" & strDbName & ")"
	Set objDb = OpenAdodb(strDbName)
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	tblDel
	DispMsg "OpenDelSyuka()"
	set tblDel = OpenRs(objDb,"del_syuka")
	dim	tblDelH
	DispMsg "OpenDelSyukaH()"
	set tblDelH = OpenRs(objDb,"del_syuka_h")
	'-------------------------------------------------------------------
	'SQL問合せ
	'-------------------------------------------------------------------
	dim	strSql
	strSql = GetSql("syuka_ymd","","")
	DispMsg "SQL:" & strSql
	dim	objRs
	set objRs = objDb.Execute(strSql)
	'-------------------------------------------------------------------
	'出荷日別件数表示
	'-------------------------------------------------------------------
	dim	iDay
	dim	strSyukaYmd
	strSyukaYmd = ""
	iDay = 4
	do while not objRs.Eof
		dim	strSyukaYmdTmp
		strSyukaYmdTmp = GetFieldValue(objRs,"key_syuka_ymd")
		DispMsg strSyukaYmdTmp _
		& " " & GetFieldValue(objRs,"cnt") _
		& ""
		if strSyukaYmdTmp <= GetDate(Now()) then
			iDay = iDay - 1
		end if
		if iDay = 0 then
			strSyukaYmd = strSyukaYmdTmp
		end if
		objRs.MoveNext
	loop
	objRs.Close
	'-------------------------------------------------------------------
	'明細
	'-------------------------------------------------------------------
	strSyukaYmd = GetOption("DT",strSyukaYmd)
	if strSyukaYmd <> "" then
		strSql = GetSql("y_syuka",strSyukaYmd,"")
		DispMsg "SQL:" & strSql
		set objRs = objDb.Execute(strSql)
		do while not objRs.Eof
			dim	strIdNo
			strSyukaYmd	= GetFieldValue(objRs,"key_syuka_ymd")
			strIdNo		= GetFieldValue(objRs,"key_id_no")
			DispMsg "" _
			& "y:" _
			& " " & strSyukaYmd _
			& "   " _
			& " " & strIdNo _
			& " " & GetFieldValue(objRs,"KAN_KBN") _
			& " " & GetFieldValue(objRs,"KEY_MUKE_CODE") _
			& " " & GetFieldValue(objRs,"MUKE_NAME") _
			& ""
			dim	objDelH
			strSql = GetSql("y_syuka_h",strSyukaYmd,strIdNo)
			set objDelH = objDb.Execute(strSql)
			if not objDelH.Eof then
				DispMsg "" _
				& "h:" _
				& " " &  GetFieldValue(objDelH,"syuka_ymd") _
				& " " & GetFieldValue(objDelH,"INS_BIN") _
				& " " & GetFieldValue(objDelH,"id_no") _
				& " " & GetFieldValue(objDelH,"CANCEL_F") _
				& " " & GetFieldValue(objDelH,"OKURISAKI") _
				& " " & GetFieldValue(objDelH,"MUKE_CODE") _
				& " " & GetFieldValue(objDelH,"MUKE_NAME") _
				& ""
				if DelYSyuka(objDelH,tblDelH) = "OK" then
					strSql = GetSql("del_syuka_h",strSyukaYmd,strIdNo)
					Call objDb.Execute(strSql)
				end if
			end if
			if DelYSyuka(objRs,tblDel) = "OK" then
				strSql = GetSql("del_syuka",strSyukaYmd,strIdNo)
				Call objDb.Execute(strSql)
			end if
			objRs.MoveNext
		loop
		objRs.Close
	end if
	'-------------------------------------------------------------------
	'データベースの後処理
	'-------------------------------------------------------------------
	set tblDel = CloseRs(tblDel)
	set tblDelH = CloseRs(tblDelH)
	DispMsg "CloseAdodb()"
	set objDb = CloseAdodb(objDb)
End Function

Function DelYSyuka(objRs _
		  ,tblDel _
		   )
	dim	strRet
	strRet = ""
	if WScript.Arguments.Named.Exists("DOMOVE") then
		tblDel.Addnew
		dim	objFld
		for each objFld in objRs.Fields
'			Wscript.Echo objFld.Name & ":" & RTrim(objFld)
			dim	txtFld
			txtFld = RTrim(objFld)
			if txtFld = "ジュ－テック?菅勝ベニヤ・" then
				txtFld = "ジュ－テック 菅勝ベニヤ"
			end if
			if txtFld = "Ｒ?（有）タマルハウス（・" then
				txtFld = "Ｒ（有）タマルハウス（"
			end if
			if txtFld = "（株）中澤?（株）南洲建・" then
				txtFld = "（株）中澤（株）南洲建"
			end if

			tblDel.Fields(objFld.Name) = txtFld
		next
		tblDel.UpdateBatch
		strRet = "OK"
	elseif WScript.Arguments.Named.Exists("DEL") then
		strRet = "OK"
	end if
	DelYSyuka = strRet
End Function

Function GetSql(byval strType,byval strSyukaYmd,byval strIdNo)
	dim	strSql
	select case strType
	case "syuka_ymd"
		strSql = "select "
		strSql = strSql & " key_syuka_ymd"
		strSql = strSql & ",count(*) cnt"
		strSql = strSql & " from y_syuka"
		strSql = strSql & " where DT_SYU <> 'R'"
		strSql = strSql & " group by key_syuka_ymd"
		strSql = strSql & " order by key_syuka_ymd desc"
	case "y_syuka"
		strSql = "select "
		strSql = strSql & " *"
		strSql = strSql & " from y_syuka"
		dim	strDt
		strDt = GetOption("DT","")
		if strDt <> "" then
			strSql = strSql & " where key_syuka_ymd = '" & strDt & "'"
		else
			strSql = strSql & " where key_syuka_ymd <= '" & strSyukaYmd & "'"
		end if
		'-------------------------------------------------------------------
		' 便指定
		'-------------------------------------------------------------------
		dim	strBin
		strBin = GetOption("BIN","")
		if strBin <> "" then
			strSql = strSql & " and ID_NO in (select ID_NO from y_syuka_h where SYUKA_YMD = '" & strSyukaYmd & "' and INS_BIN = '" & strBin & "')"
		end if
		'-------------------------------------------------------------------
		' SK セキスイ 未出庫は移動しない
		'-------------------------------------------------------------------
		strSql = strSql & " and not (KEY_MUKE_CODE like 'SK%' and KAN_KBN = '0')"
		'-------------------------------------------------------------------
	case "del_syuka"
		strSql = "delete "
		strSql = strSql & " from y_syuka"
		strSql = strSql & " where key_syuka_ymd = '" & strSyukaYmd & "'"
		strSql = strSql & "   and key_id_no = '" & strIdNo & "'"
	case "y_syuka_h"
		strSql = "select "
		strSql = strSql & " *"
		strSql = strSql & " from y_syuka_h"
		strSql = strSql & " where syuka_ymd = '" & strSyukaYmd & "'"
		strSql = strSql & "   and id_no = '" & strIdNo & "'"
	case "del_syuka_h"
		strSql = "delete "
		strSql = strSql & " from y_syuka_h"
		strSql = strSql & " where syuka_ymd = '" & strSyukaYmd & "'"
		strSql = strSql & "   and id_no = '" & strIdNo & "'"
	end select
	GetSql = strSql
End Function


Sub ySyukaRecober(strDbName)
	dim	objDb
	dim	rsY
	dim	rsDel
	dim	fldY
	dim	strSql
	dim	strMsg
	dim	strTest

	strTest = ""

	call DispMsg("CreateObject(ADODB.Connection)")
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	
	call DispMsg("Open:" & strDbName)
	objDb.Open strDbName
	call DispMsg(objDb.ConnectionString)

	call DispMsg("CreateObject(ADODB.Recordset)")
	Set rsDel = Wscript.CreateObject("ADODB.Recordset")
	rsDel.MaxRecords = 1
	rsDel.CursorLocation = adUseServer
	call DispMsg("open:del_syuka")
	rsDel.Open "del_syuka_h", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

	strSql = GetOpenSql()
	call DispMsg("Execute:" & strSql)
	set rsY = objDb.Execute(strSql)
	call DispMsg("Execute:End")

	Do While Not rsY.EOF
'debug		call DispMsg(rsBak.Fields("KEY_ID_NO"))
		if strTest = "" then
			rsDel.Addnew
		end if
		for each fldY in rsY.Fields
'debug			call DispMsg(fldBak.Name & ":" & fldBak)
			if strTest = "" then
				rsDel.Fields(fldY.Name) = fldY
			end if
		next
		On Error Resume Next
		if strTest = "" then
			rsDel.UpdateBatch
			if Err.Number = 0 then
				strMsg = "Ok"
			else
				strMsg = "Err:" & Err.Number & " " & Err.Description
			end if
		else
			strMsg = "Test"
		end if
		Err.Clear
		On Error Goto 0
		call DispMsg(rsY.Fields("ID_NO") & ":" & strMsg)
		if strMsg = "Ok" then
			call objDb.Execute(GetDeleteSql(rsY.Fields("ID_NO")))
'			rsY.Delete
		end if
		rsY.MoveNext
	Loop

	call DispMsg("Close:y_syuka_h")
	rsDel.Close
	set rsDel = Nothing
	
	call DispMsg("Close:del_syuka_h")
	rsY.Close
	set rsY = Nothing

	call DispMsg("Close:" & strDbName)
	objDb.Close
	set objDb = Nothing
End Sub

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub

Function GetOpenSql()
	dim	strSql

	strSql = "select *"
	strSql = strSql & " from y_syuka_h"
'	strSql = strSql & " where KEY_ID_NO = '075519501'"
'	strSql = strSql & " where KEY_ID_NO = '075519501'"
'	strSql = strSql & " where ID_NO = '076729101'"
'	strSql = strSql & " where ID_NO = '076987002'"
'	strSql = strSql & " where syuka_ymd <= '20111220'"
'	strSql = strSql & " where syuka_ymd = '20120207'"
'	strSql = strSql & " and ID_NO = '0I2901001'"
'	strSql = strSql & " where syuka_ymd <= '20120630'"
	strSql = strSql & " where syuka_ymd <= '20120805'"

	GetOpenSql = strSql
End Function

Function GetDeleteSql(strId)
	dim	strSql

	strSql = "delete"
	strSql = strSql & " from y_syuka_h"
'	strSql = strSql & " where KEY_ID_NO = '075519501'"
'	strSql = strSql & " where KEY_ID_NO = '075519501'"
'	strSql = strSql & " where ID_NO in ('076729101')"
'	strSql = strSql & " where ID_NO in ('076987002')"
	strSql = strSql & " where ID_NO = '" & rtrim(strId) & "'"
'	strSql = strSql & " where syuka_ymd = '20120207'"
'	strSql = strSql & " and ID_NO = '0I2901001'"

	GetDeleteSql = strSql
End Function


