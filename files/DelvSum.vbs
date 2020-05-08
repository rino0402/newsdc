Option Explicit
'-----------------------------------------------------------------------
'DelvSum.vbs
'出荷実績
'2016.11.28 新規
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "DelvSum.vbs [option]"
	Wscript.Echo " /db:newsdc	データベース"
	Wscript.Echo "Ex."
	Wscript.Echo "sc32//nologo DelvSum.vbs /db:newsdc5 201611"
End Sub

Class DelvSum
	Private	strDBName
	Private	objDB
	Private	objRs
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		Call Make()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Make() 才数データ更新
	'-----------------------------------------------------------------------
	Private	strPrev
	Private	strCurr
	Private	lngIns
	Private	lngUpd
    Private Function Make()
		Debug ".Make()"
		if WScript.Arguments.Named.Exists("del") then
			WScript.StdErr.Write "削除中..." & strYm
			WScript.StdErr.WriteLine ":" & Delete()
		end if
		SetSql	""
		SetSql	"select"
		SetSql	"Left(SYUKA_YMD,6)	SYUKA_YM"
		SetSql	",UNSOU_KAISHA"
		SetSql	",MUKE_CODE"
		SetSql	",MUKE_NAME"
		SetSql	",COL_OKURISAKI_CD"
		SetSql	",OKURISAKI_CD"
		SetSql	",OKURISAKI"
		SetSql	",JGYOBU"
		SetSql	",NAIGAI"
		SetSql	",HIN_NO HIN_GAI"
		SetSql	",if(BIKOU like '%ハソン%','1','0')	HASON"
		SetSql	",sum(convert(SURYO,SQL_DECIMAL))	Qty"
		SetSql	",count(*)							Cnt"
		SetSql	",count(distinct if(OKURI_NO='',SYUKA_YMD,OKURI_NO))	DlvCnt"
		SetSql	"from DEL_SYUKA_H"
		SetSql	"where CANCEL_F <> '1'"
		SetSql	GetWhere("and")
'		SetSql	"and SYUKA_YMD >= '20150401'"
		SetSql	"group by"
		SetSql	"SYUKA_YM"
		SetSql	",UNSOU_KAISHA"
		SetSql	",MUKE_CODE"
		SetSql	",MUKE_NAME"
		SetSql	",COL_OKURISAKI_CD"
		SetSql	",OKURISAKI_CD"
		SetSql	",OKURISAKI"
		SetSql	",JGYOBU"
		SetSql	",NAIGAI"
		SetSql	",HIN_GAI"
		SetSql	",HASON"
		Debug ".Make():" & strSql
		WScript.StdErr.Write "検索中..." & strYm
		set objRs = objDB.Execute(strSql)
		WScript.StdErr.WriteLine ":Eof:" & objRs.Eof
		strPrev = ""
		strCurr = ""
		lngIns = 0
		lngUpd = 0
		do while objRs.Eof = False
			DispData
			MakeData
			WScript.StdOut.WriteLine
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
		WScript.StdErr.WriteLine "Ins:" & Format(lngIns,-6)
		WScript.StdErr.WriteLine "Upd:" & Format(lngUpd,-6)
	End Function
	'-------------------------------------------------------------------
	'GetWhere()
	'-------------------------------------------------------------------
	Private Function GetWhere(byVal strAnd)
		select case left(strYm,2)
		case ">="
			GetWhere = strAnd & " SYUKA_YM >= '" & right(strYm,len(strYm)-2) & "'"
			exit function
		end select
		if inStr(strYm,"%") > 0 then
			GetWhere = strAnd & " SYUKA_YM like '" & strYm & "'"
			exit function
		end if
		GetWhere = strAnd & " SYUKA_YM = '" & strYm & "'"
	End Function
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function DispData()
		Debug ".DispData()"
		WScript.StdOut.Write GetStr("SYUKA_YM"			, 7)
		WScript.StdOut.Write GetStr("UNSOU_KAISHA"		, 7)
		WScript.StdOut.Write GetStr("MUKE_CODE"			, 9)
		WScript.StdOut.Write GetStr("MUKE_NAME"			,12)
		WScript.StdOut.Write GetStr("COL_OKURISAKI_CD"	,10)
		WScript.StdOut.Write GetStr("OKURISAKI_CD"		,10)
		WScript.StdOut.Write GetStr("OKURISAKI"			,21)
		WScript.StdOut.Write GetStr("JGYOBU"			, 2)
		WScript.StdOut.Write GetStr("NAIGAI"			, 2)
		WScript.StdOut.Write GetStr("HIN_GAI"			,16)
		WScript.StdOut.Write GetStr("HASON"				, 2)
		WScript.StdOut.Write GetStr("Qty"				,-3)
		WScript.StdOut.Write GetStr("Cnt"				,-3)
		WScript.StdOut.Write GetStr("DlvCnt"			,-3)
	End Function
	'-------------------------------------------------------------------
	'フィールド値 書式
	'-------------------------------------------------------------------
	Private Function GetStr(byVal strName,byVal intLen)
		GetStr = Format(GetField(strName),intLen)
	End Function
	'-------------------------------------------------------------------
	'書式
	'-------------------------------------------------------------------
	Private Function Format(byVal strV,byVal intLen)
		Format = strV
		if intLen > 0 then
			Format = Get_LeftB(Format & space(intLen),intLen)
			Format = Get_LeftB(Format & space(intLen),intLen)
		else
			intLen = Abs(intLen)
			Format = Right(space(intLen) & Format,intLen)
		end if
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
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData()"
		if lngIns > lngUpd then
			if Insert = 0 then
				Update
			end if
		else
			if Update = 0 then
				Insert
			end if
		end if
	End Function
	'-------------------------------------------------------------------
	'Delete
	'-------------------------------------------------------------------
	Private	Function Delete()
		Debug ".Delete()"
		SetSql ""
		SetSql "delete from DelvSum"
		SetSql	GetWhere("where")
		Debug strSql
		Delete = CallSql(strSql)
	End Function
	'-------------------------------------------------------------------
	'Update
	'-------------------------------------------------------------------
	Private	Function Update()
		Debug ".Update()"
		SetSql ""
		SetSql "update DelvSum"
		SetSql "set"
		SetSql " Cnt = " & GetField("Cnt") & ""
		SetSql ",Qty = " & GetField("Qty") & ""
		SetSql ",DlvCnt = " & GetField("DlvCnt") & ""
		SetSql "where SYUKA_YM = '" & GetField("SYUKA_YM") & "'"
		SetSql "and UNSOU_KAISHA = '" & GetField("UNSOU_KAISHA") & "'"
		SetSql "and MUKE_CODE = '" & GetField("MUKE_CODE") & "'"
		SetSql "and MUKE_NAME = '" & GetField("MUKE_NAME") & "'"
		SetSql "and COL_OKURISAKI_CD = '" & GetField("COL_OKURISAKI_CD") & "'"
		SetSql "and OKURISAKI_CD = '" & GetField("OKURISAKI_CD") & "'"
		SetSql "and OKURISAKI = '" & GetField("OKURISAKI") & "'"
		SetSql "and JGYOBU = '" & GetField("JGYOBU") & "'"
		SetSql "and NAIGAI = '" & GetField("NAIGAI") & "'"
		SetSql "and HIN_GAI = '" & GetField("HIN_GAI") & "'"
		SetSql "and HASON = '" & GetField("HASON") & "'"
		SetSql "and ( Cnt <> " & GetField("Cnt") & ""
		SetSql "   or Qty <> " & GetField("Qty") & ""
		SetSql "   or DlvCnt <> " & GetField("DlvCnt") & ""
		SetSql ")"
		WScript.StdOut.Write " Upd:"
		Debug strSql
		Update = CallSql(strSql)
		WScript.StdOut.Write Update
		if Update > 0 then
			lngUpd = lngUpd + Update
		end if
	End Function
	'-------------------------------------------------------------------
	'Insert
	'-------------------------------------------------------------------
	Private Function Insert()
		Debug ".Insert()"
		SetSql ""
		SetSql "insert into DelvSum ("
		SetSql "SYUKA_YM"
		SetSql ",UNSOU_KAISHA"
		SetSql ",MUKE_CODE"
		SetSql ",MUKE_NAME"
		SetSql ",COL_OKURISAKI_CD"
		SetSql ",OKURISAKI_CD"
		SetSql ",OKURISAKI"
		SetSql ",JGYOBU"
		SetSql ",NAIGAI"
		SetSql ",HIN_GAI"
		SetSql ",HASON"
		SetSql ",Cnt"
		SetSql ",Qty"
		SetSql ",DlvCnt"
		SetSql ") values ("
		SetSql "'" & GetField("SYUKA_YM") & "'"
		SetSql ",'" & GetField("UNSOU_KAISHA") & "'"
		SetSql ",'" & GetField("MUKE_CODE") & "'"
		SetSql ",'" & GetField("MUKE_NAME") & "'"
		SetSql ",'" & GetField("COL_OKURISAKI_CD") & "'"
		SetSql ",'" & GetField("OKURISAKI_CD") & "'"
		SetSql ",'" & GetField("OKURISAKI") & "'"
		SetSql ",'" & GetField("JGYOBU") & "'"
		SetSql ",'" & GetField("NAIGAI") & "'"
		SetSql ",'" & GetField("HIN_GAI") & "'"
		SetSql ",'" & GetField("HASON") & "'"
		SetSql "," & GetField("Cnt") & ""
		SetSql "," & GetField("Qty") & ""
		SetSql "," & GetField("DlvCnt") & ""
		SetSql ")"
		Debug strSql
		WScript.StdOut.Write " Ins:"
		Insert = CallSql(strSql)
		WScript.StdOut.Write Insert
		if Insert > 0 then
			lngIns = lngIns + Insert
		end if
	End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Function
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
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		dim	intNumber
		dim	strDescription
		objDb.Execute strSql
		intNumber = Err.Number
		strDescription	= Err.Description
		on error goto 0
		CallSql = 0
		select case RTrim(Hex(intNumber) & "")
		case "0"
			dim	objRc
			set objRc = objDb.Execute("select @@rowcount")
			CallSql = objRc.Fields(0)
			set	objRc = nothing
		case "80004005"	'重複するキー値があります(Btrieve Error 5)
		case else
			CallSql = -1
			WScript.StdOut.Write RTrim("0x" & Hex(intNumber) & " " & strDescription)
		end select
    End Function
	'-------------------------------------------------------------------
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		on error resume next
		strField = RTrim("" & objRs.Fields(strName))
		if Err.Number <> 0 then
			WScript.Echo "GetField(" & strName & "):Error:0x" & Hex(Err.Number)
			WScript.Quit
		end if
		on error goto 0
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
	End Function
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
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
	Function GetOption(byval strName _
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
	dim	strYm
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		strYm = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strYm = "" then
				strYm = strArg
			else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "del"
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
	dim	objDelvSum
	Set objDelvSum = New DelvSum
	if objDelvSum.Init() <> "" then
		call usage()
		exit function
	end if
	call objDelvSum.Run()
End Function
