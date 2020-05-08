Option Explicit
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CommandTypeEnum Values ----
Const adCmdUnspecified	= -1	' Unspecified type of command 
Const adCmdText			= 1		' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2		' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4		' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8		' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

dim	db
dim	dbName
dim	sqlStr
dim	rsZip
dim	strFilename
dim	i
dim	strBuff
dim	objFSO
dim	objFile
dim	objLog
dim	strFind
dim	strMsg
dim	strUpdMsg
dim	lngLen
dim	lngUpd
dim	lngIns
dim	lngCnt
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const adSearchForward = 1
' ObjectStateEnum
' オブジェクトを開いているか閉じているか、データ ソースに接続中か、
' コマンドを実行中か、またはデータを取得中かどうかを表します。
Const	adStateClosed		= 0 ' オブジェクトが閉じていることを示します。 
Const	adStateOpen			= 1 ' オブジェクトが開いていることを示します。 
Const	adStateConnecting	= 2 ' オブジェクトが接続していることを示します。 
Const	adStateExecuting	= 4 ' オブジェクトがコマンドを実行中であることを示します。 
Const	adStateFetching		= 8 ' オブジェクトの行が取得されていることを示します。 

	If WScript.Arguments.Count = 0 Then
		Wscript.Echo "zip.vbs <zip filename>"
		Wscript.Echo "GetTM()=" & GetTm(Now())
		WScript.quit
	end if
	strFilename	= WScript.Arguments(0)
	Wscript.Echo "zip.vbs " & strFilename

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)

	' データベースOpen
	dbName = "newsdc"

	Set db = Wscript.CreateObject("ADODB.Connection")
	Wscript.Echo "open db : " & dbName
	db.open dbName

	' ZipテーブルOpen
	Set rsZip = Wscript.CreateObject("ADODB.Recordset")
	rsZip.MaxRecords = 1
'	Wscript.Echo "open table : pn"
	rsZip.CursorLocation = adUseServer
'	rsPn.Open "pn", db, adOpenKeyset, adLockOptimistic, adCmdTableDirect

	lngCnt	= 0
	lngUpd	= 0
	lngIns	= 0
	On Error Resume Next
	do while ( objFile.AtEndOfStream = False )
		strBuff = objFile.ReadLine()
		lngLen	= Get_LenB(strBuff)
		select case lngLen
		case 86
		case else
			Wscript.Echo "length error:" & lngLen
			lngLen = 0
		end select
		if lngLen > 0 then
			lngCnt	= lngCnt + 1
			strFind = "select * from Zip"
			strFind = strFind & " where ZipCode = '" & rtrim(Get_MidB(strBuff,  1, 7)) & "'"

			if rsZip.state <> adStateClosed then
				rsZip.Close
			end if
'			rsPn.Open strFind, db, adOpenStatic, adLockOptimistic
			rsZip.Open strFind, db, adOpenForwardOnly, adLockBatchOptimistic

			if rsZip.Eof = false then
				strMsg = "Upd:"
'				rsZip.Fields("UPD_ID")		= "pn2.vbs"
'				rsZip.Fields("UPD_TM")		= GetTm(now())
				strUpdMsg = ""
			else
				lngIns	= lngIns + 1
				if rsZip.state <> adStateClosed then
					rsZip.Close
				end if
'				rsZip.Open "pn2", db, adOpenKeyset, adLockOptimistic, adCmdTableDirect
				rsZip.Open "Zip", db, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
				strMsg = "New:"
				rsZip.Addnew
				rsZip.Fields("ZipCode")		= Get_MidB(strBuff,  1, 7)
				rsZip.Fields("Jis8Code")	= Get_MidB(strBuff,  9, 8)
				rsZip.Fields("Address")		= Get_MidB(strBuff, 18,64)
				rsZip.Fields("MiseCode")	= Get_MidB(strBuff, 83, 3)
				strUpdMsg = "-"
			end if
			strMsg = strMsg & rsZip.Fields("ZipCode")
			strMsg = strMsg & " " & rsZip.Fields("Jis8Code")
			strMsg = strMsg & " " & rsZip.Fields("Address")
			strMsg = strMsg & " " & rsZip.Fields("MiseCode")
			Wscript.Echo strMsg
			if strUpdMsg = "" then
				rsZip.CancelBatch
			else
				rsZip.UpdateBatch
				if strUpdMsg <> "-" then
					lngUpd	= lngUpd + 1
					Wscript.Echo strUpdMsg
				end if
			end if
		end if
	loop

	objFile.Close

	' ZipテーブルClose
	Wscript.Echo "close table : zip"
	rsZip.Close

	' DBClose
	Wscript.Echo "close db : " & dbName
	db.Close
	set db = nothing

	Wscript.Echo " 件数：" & lngCnt
	Wscript.Echo " 更新：" & lngUpd
	Wscript.Echo " 追加：" & lngIns

	Set objLog = objFSO.OpenTextFile("pn2.log", ForAppending, True)
	objLog.WriteLine "ファイル：" & strFilename
	objLog.WriteLine "    件数：" & lngCnt
	objLog.WriteLine "    更新：" & lngUpd
	objLog.WriteLine "    追加：" & lngIns
	objLog.Close
	Set objLog = Nothing
	set objFSO = nothing

Function GetTm(t)
	GetTm = year(t) & right("0" & month(t),2) & right("0" & day(t),2) & right("0" & hour(t),2)& right("0" & minute(t),2)
End Function

Function Get_LeftB(a_Str, a_int)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
		If iLenCount > Cint(a_int) Then
			Exit For
		Else
			iLeftStr = iLeftStr + Mid(a_Str, iCount, 1)
		End If
	Next
	Get_LeftB = iLeftStr
End Function

Function Get_MidB(a_Str,s_int, a_int)
	Dim iCount, iAscCode, iLenCount, iMidStr
	iLenCount = 0
	iMidStr = ""
	If Len(a_Str) = 0 Then
		Get_MidB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_MidB = ""
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
		if iLenCount >= s_int then
			If iLenCount > Cint(s_int) + Cint(a_int) - 1 Then
				Exit For
			Else
				iMidStr = iMidStr + Mid(a_Str, iCount, 1)
			End If
		end if
	Next
	Get_MidB = iMidStr
End Function

Function Get_LenB(a_Str)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LenB = 0
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
	Next
	Get_LenB = iLenCount
End Function


Function SetField(rsPn,strFieldName,strValue,strTitle,strUpdMsg)
	if rtrim(rsPn.Fields(strFieldName)) <> rtrim(strValue) then
		if strUpdMsg <> "-" then
			strUpdMsg = strUpdMsg & rsPn.Fields(strFieldName) & " ←" & strTitle & vbNewLine
			strUpdMsg = strUpdMsg & strValue & " ←変更" & vbNewLine
		end if
		rsPn.Fields(strFieldName) = strValue
	end if
	SetField = strUpdMsg
End Function
