Option Explicit
' 2014.04.30 出荷日≦本日に変更
' 2014.08.25 出庫残 を表示
' 2016.10.01 参照テーブル変更 g_syuka → HMTAH015
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
Call Include("debug.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit(lngRet)
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "出荷完了チェック"
	Wscript.Echo "y_syuka_check.vbs [option]"
	Wscript.Echo " /db:newsdc"
	Wscript.Echo " /debug"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		call usage()
		Main = -1
		exit Function
	next
	'名前付きオプションチェック
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			call usage()
			Main = -1
			exit Function
		case else
			call usage()
			Main = -1
			exit Function
		end select
	next
'	call YSyukaCheck()
	Main = YSyukaCheck()
End Function

Function YSyukaCheck()
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	strSql
	strSql = GetSql()
	dim	rsYSyuka
	set rsYSyuka = objDb.Execute(strSql)

	strBuff = "                      合計"
	strBuff = "12345678 12345678 12345678 123456 123456"
	strBuff = "yyyymmdd 2        1        999999 999999 999999 999999 999999 999999"
	strBuff = "出荷日   注文区分 直送区分   件数 出庫残 検品残 送信残 実績残   数量"
	Wscript.Echo strBuff

	dim	lngCnt
	lngCnt = 0
	dim	lngPZan
	lngPZan = 0
	dim	lngKZan
	lngKZan = 0
	dim	lngSZan
	lngSZan = 0
	dim	lngJZan
	lngJZan = 0
	dim	lngQty
	lngQty = 0
	Do While Not rsYSyuka.EOF
		lngCnt	= lngCnt	+ CLng(rsYSyuka.Fields("件数"))
		lngPZan = lngPZan	+ CLng(rsYSyuka.Fields("出庫残"))
		lngKZan = lngKZan	+ CLng(rsYSyuka.Fields("検品残"))
		lngSZan = lngSZan	+ CLng(rsYSyuka.Fields("送信残"))
		lngJZan = lngJZan	+ CLng(rsYSyuka.Fields("実績残"))
		lngQty	= lngQty	+ CLng(rsYSyuka.Fields("数量"))
		dim	strBuff
		strBuff = ""
		strBuff = strBuff & left(trim(rsYSyuka.Fields("出荷日")) & "        ",8)
		strBuff = strBuff & " " & Get_LeftB(trim(rsYSyuka.Fields("注文区分")) & "        ",8)
		strBuff = strBuff & " " & Get_LeftB(trim(rsYSyuka.Fields("直送区分")) & "        ",8)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("件数")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("出庫残")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("検品残")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("送信残")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("実績残")),6)
		strBuff = strBuff & " " & right("      " & trim(rsYSyuka.Fields("数量")),6)
		Wscript.Echo strBuff
		rsYSyuka.movenext
	Loop
	strBuff = "                      合計"
	strBuff = strBuff & right("       " & lngCnt,7)
	strBuff = strBuff & right("       " & lngPZan,7)
	strBuff = strBuff & right("       " & lngKZan,7)
	strBuff = strBuff & right("       " & lngSZan,7)
	strBuff = strBuff & right("       " & lngJZan,7)
	strBuff = strBuff & right("       " & lngQty,7)
	Wscript.Echo strBuff
	Wscript.Echo ""
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsYSyuka = CloseRs(rsYSyuka)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
	'-------------------------------------------------------------------
	' リターン値セット
	'-------------------------------------------------------------------
	if lngCnt = 0 then
		' 出荷予定０件(休日)
		YSyukaCheck = -100
	else
		' 出荷実績連携の残数
		YSyukaCheck = lngJZan
	end if
End Function

Function GetSql()
	dim	sqlStr
	sqlStr = "select"
	sqlStr = sqlStr & " y.KEY_SYUKA_YMD									""出荷日"""
	sqlStr = sqlStr & ",y.KEY_CYU_KBN									""注文区分"""
	sqlStr = sqlStr & ",y.CHOKU_KBN + if(y.CHOKU_KBN='1',' 直送','')	""直送区分"""
	sqlStr = sqlStr & ",count(*)										""件数"""
	sqlStr = sqlStr & ",sum(if(RTrim(y.KAN_KBN)='0',1,0))				""出庫残"""
	sqlStr = sqlStr & ",sum(if(RTrim(y.KENPIN_TANTO_CODE)='',1,0))		""検品残"""
	sqlStr = sqlStr & ",sum(if(y.KEY_CYU_KBN = 'E' or RTrim(y.LK_SEQ_NO)<>'',0,1))				""送信残"""
	sqlStr = sqlStr & ",sum(if(g.IDno is null,0,1))						""実績残"""
	sqlStr = sqlStr & ",sum(convert(y.SURYO,SQL_NUMERIC))				""数量"""
	sqlStr = sqlStr & " from y_syuka y"
	sqlStr = sqlStr & " left outer join HMTAH015 g on (y.KEY_ID_NO = g.IDno)"
	sqlStr = sqlStr & " where y.JGYOBA = '00036003'"
	sqlStr = sqlStr & "   and y.DATA_KBN in ('1','3')"
	sqlStr = sqlStr & "   and y.KEY_SYUKA_YMD <= replace(convert(curdate(),SQL_CHAR),'-','')"
	sqlStr = sqlStr & " group by ""出荷日"",""注文区分"",""直送区分"""
	sqlStr = sqlStr & " order by ""出荷日"",""注文区分"",""直送区分"""
	GetSql = sqlStr
End Function

Function Get_LeftB(byval a_Str,byval a_int)
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
