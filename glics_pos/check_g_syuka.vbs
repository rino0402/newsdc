Option Explicit
' 2008.12.24 POSキャンセル漏れチェック追加
' select * from p_sagyo_log
'  where ID_NO = '700082764202'
'    and jitu_dt >= replace(left(convert(now()-30,SQL_CHAR),10),'-','')
' 2012.03.24 タイムアウトエラー対応
' 2016.10.01 参照テーブル変更 g_syuka → HMTAH015
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "check_g_syuka.vbs [option]"
	Wscript.Echo " /db:newsdc	データベース"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript check_g_syuka.vbs /db:newsdc3"
End Sub

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Class YSyuka
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	lngResult
	Private	strTm
	Private	objFSO
	Private	objOutput
	Private	strFileName
	Private	strSoko
	Private	strMsg
	Private	strDlm
	Private	strBuff
	Private	strBuffPos
	Private	strBuffAct
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		lngResult = 0

		dim	dtNow
		dtNow = now()
		strTm = year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2)
		strTm = strTm & right("0" & Hour(dtNow),2) & right("0" & Minute(dtNow),2) & "00"

		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		strFileName = "check_g_syuka.txt"
		Set objOutput = objFSO.OpenTextFile(strFileName, ForWriting, True)

		strSoko = ""
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		objOutput.Close 
		set	objOutput	= nothing
		set	objFSO		= nothing
		set objRs		= nothing
		set objDB		= nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		Call Load()
		Call CloseDb()
		Run = lngResult
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
	Private	intLoop
    Public Sub Load()
		Debug ".Load()"
		for intLoop = 1 to 3
			select case intLoop
			case 1
				strSql = GetSqlModify()
			case 2
				strSql = GetSqlNotIn(strSoko)
			case 3
				strSql = GetSqlCancel()
			end select

			Set objRs = objDb.Execute(strSql)

			strMsg = ""
			strSoko	= ""
			Do While Not objRs.EOF
				select case intLoop
				case 1
					if strSoko = "" then
						strSoko = GetField("Soko")
					end if
					strMsg = ""
					strDlm = ""
					if GetField("KEY_SYUKA_YMD") <> GetField("SyukaDt") then
						strMsg = strMsg & strDlm & "出荷日変更(" & GetField("KEY_SYUKA_YMD") & "→" & GetField("SyukaDt") & ")"
						strDlm = " "
						SetSql ""
						SetSql "update y_syuka"
						SetSql " set KEY_SYUKA_YMD='" & GetField("SyukaDt") & "'"
						SetSql "   , SYUKO_YMD='"     & GetField("SyukaDt") & "'"
						SetSql "   , SYUKA_YMD='"     & GetField("SyukaDt") & "'"
						SetSql "   , LK_SEQ_NO = ''"
						SetSql "   , UPD_NOW = '" & strTm & "'"
						SetSql " where KEY_ID_NO = '" & GetField("KEY_ID_NO") & "'"
						Call CallSql(strSql)
					end if
					if GetField("KEY_HIN_NO") <> GetField("Pn") then
						strMsg = strMsg & strDlm & "品番変更(" & GetField("KEY_HIN_NO") & "→" & GetField("Pn") & ")自動更新しません。"
						strDlm = " "
					end if
					if GetField("DEN_NO") <> GetField("DenNo") then
						strMsg = strMsg & strDlm & "伝票No変更(" & GetField("DEN_NO") & "→" & GetField("DenNo") & ")"
						strDlm = " "
						SetSql ""
						SetSql "update y_syuka"
						SetSql " set DEN_NO='" & GetField("DenNo") & "'"
						SetSql "   , LK_SEQ_NO = ''"
						SetSql "   , UPD_NOW = '" & strTm & "'"
						SetSql " where KEY_ID_NO = '" & GetField("KEY_ID_NO") & "'"
						Call CallSql(strSql)
					end if
					if CLng(GetField("SURYO")) <> CLng(GetField("Qty")) then
						strMsg = strMsg & strDlm & "数量変更(" & CLng(GetField("SURYO")) & "→" & GetField("Qty") & ")出庫済=" & GetField("JITU_SURYO") & " 完区=" & GetField("KAN_KBN")
						strDlm = " "
						SetSql ""
						SetSql "update y_syuka"
						SetSql " set SURYO='" & Right("0000000" & GetField("Qty"),7) & "'"
						SetSql "   , JITU_SURYO = '0000000'"
						SetSql "   , KAN_KBN = '0'"
						SetSql "   , LK_SEQ_NO = ''"
						SetSql "   , UPD_NOW = '" & strTm & "'"
						SetSql " where KEY_ID_NO = '" & GetField("KEY_ID_NO") & "'"
						Call CallSql(strSql)
					elseif (GetField("KAN_KBN") = "9") and (GetField("JITU_SURYO") <> GetField("SURYO")) then
						strMsg = strMsg & strDlm & "数量変更(" & GetField("JITU_SURYO") & "→" & GetField("SURYO") & ")出庫済=" & GetField("JITU_SURYO") & " 完区=" & GetField("KAN_KBN") & " 出荷予定データにより変更"
						strDlm = " "
						SetSql ""
						SetSql "update y_syuka"
						SetSql " set JITU_SURYO = '0000000'"
						SetSql "   , KAN_KBN = '0'"
						SetSql "   , LK_SEQ_NO = ''"
						SetSql " where KEY_ID_NO = '" & GetField("KEY_ID_NO") & "'"
						Call CallSql(strSql)
					end if
					strBuff = ""
					strBuff = strBuff & " " & GetField("KEY_SYUKA_YMD")
					strBuff = strBuff & " " & GetField("SYUKO_YMD")
					strBuff = strBuff & " " & GetField("KEY_CYU_KBN")
					strBuff = strBuff & " " & GetField("CHOKU_KBN")
					strBuff = strBuff & " " & GetField("KEY_MUKE_CODE")
					strBuff = strBuff & " " & GetField("MUKE_NAME")
					strBuff = strBuff & " " & GetField("DEN_NO")
					strBuff = strBuff & " " & GetField("KEY_ID_NO")
					strBuff = strBuff & " " & GetField("KEY_HIN_NO")
					strBuff = strBuff & " " & Right("       " & CLng(GetField("SURYO")),7)
					strBuff = strBuff & " " & GetField("LK_SEQ_NO")
			
					strBuffPos = strBuff

					strBuff = ""
					strBuff = strBuff & " " & GetField("SyukaDt")
					strBuff = strBuff & " " & GetField("SyukaDt")
					strBuff = strBuff & " " & GetField("CyuKbn")
					strBuff = strBuff & " " & GetField("ChokuKbn")
					strBuff = strBuff & " " & GetField("Aitesaki")
					strBuff = strBuff & " " & GetField("AitesakiName")
					strBuff = strBuff & " " & GetField("DenNo")
					strBuff = strBuff & " " & GetField("IDNo")
					strBuff = strBuff & " " & GetField("Pn")
					strBuff = strBuff & " " & Right("       " & GetField("Qty"),7)

					strBuffAct = strBuff
			
					if strMsg <> "" then
						Disp strMsg
						objOutput.WriteLine strMsg
			
						Disp "POS" & strBuffPos
						objOutput.WriteLine "POS" & strBuffPos
			
						Disp "ACT" & strBuffAct
						objOutput.WriteLine "ACT" & strBuffAct
						lngResult = lngResult + 1
					else
						Disp "POS" & strBuffPos
					end if
				case 2
					strBuff = ""
					strBuff = strBuff & " " & GetField("SyukaDt")
					strBuff = strBuff & " " & GetField("SyukaDt")
					strBuff = strBuff & " " & GetField("CyuKbn")
					strBuff = strBuff & " " & GetField("ChokuKbn")
					strBuff = strBuff & " " & GetField("Aitesaki")
					strBuff = strBuff & " " & GetField("AitesakiName")
					strBuff = strBuff & " " & GetField("DenNo")
					strBuff = strBuff & " " & GetField("IDNo")
					strBuff = strBuff & " " & GetField("Pn")
					strBuff = strBuff & " " & Right("       " & GetField("Qty"),7)
					if strMsg = "" then
						strMsg = "▼POSデータ無し(Activeキャンセル漏れ又はPOSデータ未受信)"
						Disp strMsg
						objOutput.WriteLine strMsg
					end if
					lngResult = lngResult + 1
					strBuffAct = strBuff
					Disp "ACT" & strBuffAct
					objOutput.WriteLine "ACT" & strBuffAct
				case 3
					strBuff = ""
					strBuff = strBuff & " " & GetField("KEY_SYUKA_YMD")
					strBuff = strBuff & " " & GetField("SYUKO_YMD")
					strBuff = strBuff & " " & GetField("KEY_CYU_KBN")
					strBuff = strBuff & " " & GetField("CHOKU_KBN")
					strBuff = strBuff & " " & GetField("KEY_MUKE_CODE")
					strBuff = strBuff & " " & GetField("MUKE_NAME")
					strBuff = strBuff & " " & GetField("DEN_NO")
					strBuff = strBuff & " " & GetField("KEY_ID_NO")
					strBuff = strBuff & " " & GetField("KEY_HIN_NO")
					strBuff = strBuff & " " & Right("       " & CLng(GetField("SURYO")),7)
					strBuff = strBuff & " " & GetField("LK_SEQ_NO")
					select case GetField("KAN_KBN")
					case "0"
						strBuff = strBuff & " 未出庫"
					case "9"
						strBuff = strBuff & " 出庫済"
					end select

					if strMsg = "" then
						strMsg = "▼POSキャンセル漏れ・・・出荷予定メンテナンスで削除して下さい。"
						Disp strMsg
						objOutput.WriteLine strMsg
					end if
					lngResult = lngResult + 1
					Disp "POS" & strBuff
					objOutput.WriteLine "POS" & strBuff
				end select
				objRs.MoveNext
			Loop
		Next
	End Sub

	Private Function GetSqlModify()
		SetSql ""
		SetSql "select"
		SetSql " y.KEY_SYUKA_YMD"
		SetSql ",g.SyukaDt"
		SetSql ",y.SYUKO_YMD"
		SetSql ",y.KEY_CYU_KBN"
		SetSql ",g.CyuKbn"
		SetSql ",y.CHOKU_KBN"
		SetSql ",g.ChokuKbn"
		SetSql ",y.KEY_MUKE_CODE"
		SetSql ",g.Aitesaki"
		SetSql ",y.MUKE_NAME"
		SetSql ",g.AitesakiName"
		SetSql ",y.DEN_NO"
		SetSql ",g.DenNo"
		SetSql ",y.KEY_ID_NO"
		SetSql ",g.IDNo"
		SetSql ",y.KEY_HIN_NO"
		SetSql ",g.Pn"
		SetSql ",y.SURYO"
		SetSql ",y.JITU_SURYO"
		SetSql ",y.KAN_KBN"
		SetSql ",g.Qty"
		SetSql ",y.LK_SEQ_NO"
		SetSql ",g.Soko"
		SetSql ",y.UPD_NOW"
		SetSql " from y_syuka y"
		SetSql " inner join HMTAH015 g"
		SetSql "  on (y.JGYOBA = g.JCode and y.KEY_ID_NO = g.IDno)"
	'	SetSql " where y.JGYOBA = '00036003'"
	'	SetSql "   and y.KEY_SYUKA_YMD <> g.SyukaDt"
		GetSqlModify = strSql
	End Function

	Private Function GetSqlNotIn(byval a_Soko)
		SetSql ""
		SetSql "select"
		SetSql " g.SyukaDt"
		SetSql ",g.CyuKbn"
		SetSql ",g.ChokuKbn"
		SetSql ",g.Aitesaki"
		SetSql ",g.AitesakiName"
		SetSql ",g.DenNo"
		SetSql ",g.IDNo"
		SetSql ",g.Pn"
		SetSql ",g.Qty"
		SetSql " from HMTAH015 g"
		SetSql " where soko = '" & a_Soko & "'"
		SetSql " and Stts='4'"
	'	SetSql " and IDNo not in (select distinct KEY_ID_NO from y_syuka where JGYOBA = '00036003' union select distinct KEY_ID_NO from del_syuka where JGYOBA = '00036003' and KEY_SYUKA_YMD >= replace(convert(curdate()-1,SQL_CHAR),'-',''))"
		SetSql " and IDNo not in (select distinct KEY_ID_NO from y_syuka where JGYOBA = '00036003')"
	'	SetSql " and IDNo not in (select distinct KEY_ID_NO from y_syuka union select distinct KEY_ID_NO from del_syuka)"
	'	SetSql " and IDNo not in (select distinct KEY_ID_NO from del_syuka)"
		GetSqlNotIn= strSql
	End Function

	Private Function GetSqlCancel()
		SetSql ""
		SetSql "select"
		SetSql " y.KEY_SYUKA_YMD"
		SetSql ",y.SYUKO_YMD"
		SetSql ",y.KEY_CYU_KBN"
		SetSql ",y.CHOKU_KBN"
		SetSql ",y.KEY_MUKE_CODE"
		SetSql ",y.MUKE_NAME"
		SetSql ",y.DEN_NO"
		SetSql ",y.KEY_ID_NO"
		SetSql ",y.KEY_HIN_NO"
		SetSql ",y.SURYO"
		SetSql ",y.JITU_SURYO"
		SetSql ",y.KAN_KBN"
		SetSql ",y.LK_SEQ_NO"
		SetSql ",y.UPD_NOW"
		SetSql " from y_syuka y"
		SetSql " where y.JGYOBA = '00036003'"
		SetSql "   and y.KEY_SYUKA_YMD between replace(convert(curdate()-10,SQL_CHAR),'-','') and replace(convert(curdate()-1,SQL_CHAR),'-','')"
		SetSql "   and y.KEY_ID_NO not in (select distinct IDNo from HMTAH015)"
		SetSql "   and y.LK_SEQ_NO = ''"
		SetSql "   and y.HAN_KBN in ('1')"
		SetSql "   and y.DATA_KBN in ('1')"
		GetSqlCancel = strSql
	End Function

	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Sub OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Sub
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Sub CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Sub
	'-----------------------------------------------------------------------
	'SQL文字列追加
	'-----------------------------------------------------------------------
	Private	strSql
	Public Function SetSql(byVal s)
		if s = "" then
			strSql = ""
		else
			if strSql <> "" then
				strSql = strSql & " "
			end if
			strSql = strSql & s
		end if
	End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Public Sub CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
'		on error resume next
		Disp "Execute:" & strSql
		Call objDB.Execute(strSql)
'		on error goto 0
    End Sub
	'-------------------------------------------------------------------
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		strField = RTrim(objRs.Fields(strName))
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
	End Function
	'-------------------------------------------------------------------
	'LenB()
	'-------------------------------------------------------------------
	Private Function LenB(byVal strVal)
	    Dim i, strChr
	    LenB = 0
	    If Trim(strVal) <> "" Then
	        For i = 1 To Len(strVal)
	            strChr = Mid(strVal, i, 1)
	            '２バイト文字は＋２
	            If (Asc(strChr) And &HFF00) <> 0 Then
	                LenB = LenB + 2
	            Else
	                LenB = LenB + 1
	            End If
	        Next
	    End If
	End Function
	'-------------------------------------------------------------------
	'Get_LeftB()
	'-------------------------------------------------------------------
	Private Function Get_LeftB(byVal a_Str,byVal a_int)
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
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
	'-----------------------------------------------------------------------
	Private Function GetOption(byval strName _
					  ,byval strDefault _
					  )
		dim	strValue

		if strName = "" then
			strValue = ""
			if strDefault < WScript.Arguments.UnNamed.Count then
				strValue = WScript.Arguments.UnNamed(strDefault)
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			Init = "オプションエラー:" & strArg
			Disp Init
			Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objYSyuka
	Set objYSyuka = New YSyuka
	if objYSyuka.Init() <> "" then
		call usage()
		Main = -1
		exit function
	end if
	Main = objYSyuka.Run()
End Function
