Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "DelivSts.vbs [option] [送り状No]"
	Wscript.Echo " /db:newsdc1 データベース"
	Wscript.Echo " /make[:yyyymmdd]	DelivStsに登録(default)"
	Wscript.Echo " /check[:day]		配達状況チェック"
	Wscript.Echo " /recheck			配達完了再チェック"
	Wscript.Echo " /tbl:y_syuka_h	default:del_syuka_h"
	Wscript.Echo " /test			test用(更新しない)"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript DelivSts.vbs /db:newsdc5"
	Wscript.Echo "cscript DelivSts.vbs /db:newsdc5 /check:1"
End Sub
'-----------------------------------------------------------------------
'DelivSts
'2016.10.19 配達状況：福山通運
'2018.02.08 前回の配達状況を(Status1Last,Status2Last)に保存
'2018.02.14 「配達中」は前回の配達状況(Status1Last)に保存しない
'-----------------------------------------------------------------------
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

Const READYSTATE_COMPLETE	= 4

Class DelivSts
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	objIE
	Private	optTest
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		optAction = "make"
		set	objIE = nothing
		optTest		= False
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
		if not objIE is nothing then
			objIE.Quit
		end if
		set	objIE = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		if optAction = "make" then
			Call Make()
		else
			Call Check()
		end if
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Check() 
	'-----------------------------------------------------------------------
    Public Function Check()
		Debug ".Check()"
		SetSql ""
		SetSql "select"
		SetSql " *"
		SetSql "from DelivSts"
		SetSql "where CampName = '福山通運'"
'		SetSql "  and Status1 not like '配達完了%'"
		SetSql "  and SYUKA_YMD > left(replace(convert(DATEADD(day,-" & optDay & ",curdate()),sql_char),'-',''),8)"
		if optId <> "" then
			SetSql "and DelvNo = '" & optId & "'"
		end if
		SetSql "order by SYUKA_YMD desc"
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		Call objRs.Open(strSql, objDb, adOpenKeyset, adLockOptimistic)
		do while objRs.Eof = False
			bUpdate = False
			dim	strUpd
			strUpd = GetField("SYUKA_YMD")
			strUpd = strUpd & " " & GetField("CampName")
			strUpd = strUpd & " " & GetField("DelvNo")
			strUpd = strUpd & " " & GetField("Status1")
			Call CheckData()
			if bUpdate = True then
				WScript.StdOut.Write "→" & GetField("Status1")
				if optTest = True then
					WScript.StdOut.Write "(test)"
				else
					call objRs.Update
				end if
				strUpd = strUpd & "→" & GetField("Status1")
				strUpd = strUpd & " " & GetField("Br2Code")
				strUpd = strUpd & " " & GetField("Br2Name")
			else
				strUpd = ""
			end if
			WScript.StdOut.Write " " & GetField("Br2Code")
			WScript.StdOut.Write " " & GetField("Br2Name")
			WScript.StdOut.WriteLine
			if strUpd <> "" then
				WScript.StdErr.WriteLine strUpd
			end if
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = nothing
	End Function
	'-------------------------------------------------------------------
	'CheckData() 1行読込
	'-------------------------------------------------------------------
	Private	strDelvNo
	Private Function CheckData()
		Debug ".CheckData()"
'		DispLine
		WScript.StdOut.Write GetField("SYUKA_YMD")
		WScript.StdOut.Write " " & GetField("CampName")
'		if GetField("CampName") <> "福山通運" then
'			exit function
'		end if
		WScript.StdOut.Write " " & GetField("DelvNo")
		WScript.StdOut.Write " " & GetField("Status1")
		strID = GetField("DelvNo")
		strBody = GetTrackingBody()
		Debug strBody

		dim	strLine
		dim	intBr
		For Each strLine In Split(strBody, vbCrLf)
		    strLine = RTrim(strLine)
		    Debug strLine
			dim	strStat
		    Select Case strStat
		    Case ""
		        Select Case strLine
		        Case "お荷物配達状況詳細"
		            strStat = strLine
		        Case "現在の配達状況"
		            strStat = strLine
		        Case "お届け予定日/配達完了日"
		            strStat = strLine
		        Case "支店電話番号"
					intBr = 0
		            strStat = strLine
		        End Select
		    Case "お荷物配達状況詳細"
		        If strLine = "お問い合わせ番号" Then
		            strStat = "お問い合わせ番号"
		        End If
		    Case "お問い合わせ番号"
				Call GetContents(strLine, strStat)
		        strStat = "個数"
		    Case "個数"
				Call GetContents(strLine, strStat)
		        strStat = "重量"
		    Case "重量"
				Call GetContents(strLine, strStat)
		        strStat = ""
		    Case "現在の配達状況"
				Call GetContents(strLine, strStat)
		        strStat = ""
		    Case "お届け予定日/配達完了日"
		        strStat = "お届け予定日/配達完了日1"
		    Case "お届け予定日/配達完了日1"
				Call GetContents(Trim(strLine), strStat)
		        strStat = ""
		    Case "支店電話番号"
		        strStat = "受付"
		    Case "受付"
				Call GetContents(strLine, strStat)
		        strStat = "発送"
		    Case "発送"
				Call GetContents(strLine, strStat)
		        strStat = "到着"
		    Case "到着"
				Call GetContents(strLine, strStat)
		        strStat = "持出"
		    Case "持出"
				Call GetContents(strLine, strStat)
		        strStat = "配達完了"
		    Case "配達完了"
				Call GetContents(strLine, strStat)
		        strStat = "支店コード"
			case "支店コード"
		        If strLine = "支店電話番号" Then
			        strStat = "支店電話番号0"
				end if
			case "支店電話番号0"
		        strStat = "支店電話番号1"
			case "支店電話番号1"
				Call GetContents(strLine, strStat)
		        strStat = "支店電話番号2"
			case "支店電話番号2"
				Call GetContents(strLine, strStat)
		        strStat = "END"
		    End Select
		Next
	End Function
	Private Function GetContents(ByVal strLine,ByVal strStat)
		Debug strStat & ":" & strLine
	    Dim strValue
	    strValue = ""
	    Select Case strStat
	    Case "お荷物配達状況詳細"
	    Case "お問い合わせ番号"
	        strValue = Split(strLine, " ")(0)
	    Case "個数"
	        strValue = Split(strLine, " ")(0)
			SetField "Qty",strValue
	    Case "重量"
	        strValue = Split(strLine, " ")(0)
			SetField "Weight",strValue
	    Case "現在の配達状況"
	        strValue = strLine
			SetField "Status1",strValue
	    Case "お届け予定日/配達完了日1"
	        strValue = strLine
			SetField "Status2",strValue
	    Case "受付","発送","到着","持出"
	        strValue = Split(strLine, strStat)(1)
			dim	v
			dim	strDtm
			dim	strBr
			dim	strBrTel
			strDtm = ""
			strBr = ""
			strBrTel = ""
			for each v in Split(strValue," ")
				if isDate(v) then
					if strDtm = "" then
						strDtm = v
					else
						strDtm = strDtm & " " & v
					end if
				else
					if strBr = "" then
						strBr = v
					else
						strBrTel = v
					end if
				end if
			next
			select case strStat
		    Case "受付"
				SetField "UkeDTm",strDtm
				SetField "UkeBr",strBr
				SetField "UkeBrTel",strBrTel
		    Case "発送"
				SetField "HatDTm",strDtm
				SetField "HatBr",strBr
				SetField "HatBrTel",strBrTel
		    Case "到着"
				SetField "ChaDTm",strDtm
				SetField "ChaBr",strBr
				SetField "ChaBrTel",strBrTel
		    Case "持出"
				SetField "MotDTm",strDtm
				SetField "MotBr",strBr
				SetField "MotBrTel",strBrTel
			end select
	    Case "配達完了"
	        strValue = Split(strLine, strStat)(1)
			if inStr(strValue," ") > 0 then
				SetField "FinDTm",Split(strValue," ")(0) & " " & Split(strValue," ")(1)
			else
				SetField "FinDTm",strValue
			end if
	    Case "支店電話番号1","支店電話番号2"
			if strLine <> "" then
				dim	intBr
				intBr = CInt(Right(strStat,1))
		        strValue = strLine
				dim	strCode
				dim	strName
				dim	strAddress
				dim	strTel
				if intBr = 1 then
					strName	= GetField("UkeBr")
					strTel	= GetField("UkeBrTel")
				else
					strName	= GetField("ChaBr")
					if strName = "" then
						strName	= GetField("MotBr")
					end if
					strTel	= GetField("ChaBrTel")
					if strTel = "" then
						strTel	= GetField("MotBrTel")
					end if
				end if
				
				strCode	= Split(strValue,strName)(0)
				Debug "strValue:" & strValue
				Debug "strName:" & strName
				if strName <> "" then
					strAddress	= Split(strValue,strName)(1)
				else
					strAddress	= strValue
				end if
				strAddress	= Split(strAddress,strTel)(0)
				SetField "Br" & intBr & "Code",strCode
				SetField "Br" & intBr & "Name",strName
				SetField "Br" & intBr & "Address",strAddress
				SetField "Br" & intBr & "Tel",strTel
			end if
	    End Select
	    GetContents = strValue
	End Function
	'-------------------------------------------------------------------
	' 福通の問合せNoから配達状況を取得
	'-------------------------------------------------------------------
	Private	strID
	Private	strUrl
	Private	strBody
	Private Function GetTrackingBody()
	    GetTrackingBody = ""

		if GetField("CampName") <> "福山通運" then
			exit function
		end if
		if GetField("Status1") = "配達完了です" then
		if GetField("Br2Code") <> "" then
		if GetField("FinDTm") <> "" then
			if WScript.Arguments.Named.Exists("recheck") = false then
				exit function
			end if
		end if
		end if
		end if

'		strUrl = "http://thira.plavox.info/transport/api/?t=fukutsu&no=" & strID
		strUrl = "https://corp.fukutsu.co.jp/situation/tracking_no_hunt/" & strID

	    strBody = ""
		Debug "接続:" & strUrl
	    'IEの起動
		if objIE is nothing then
			Debug "InternetExplorer.Application"
			Set objIE = CreateObject("InternetExplorer.Application")
			objIE.Visible = False
		end if
'		WScript.StdOut.Write strID
        objIE.Navigate strUrl
'		WScript.StdOut.Write ":"

        ' ページが取り込まれるまで待つ
        Do While objIE.Busy or objIE.readyState <> READYSTATE_COMPLETE
			WScript.StdOut.Write "."
            WScript.Sleep 1000
        Loop
'        Do While objIE.readyState <> READYSTATE_COMPLETE
'			WScript.StdOut.Write "*"
'			Debug "読込中 " & objIE.Document.readyState
'           WScript.Sleep 3000
'        Loop
'		WScript.StdOut.WriteLine
        ' テキスト形式で出力
		strBody = objIE.Document.Body.InnerText
'		strBody = objIE.Document.Body.textContent
		' ＨＴＭＬ形式で出力
		' objIE.Document.Body.InnerHtml
	    GetTrackingBody = strBody
	End Function
	'-----------------------------------------------------------------------
	'Make() 
	'-----------------------------------------------------------------------
	Private	strSyukaYmd
    Public Function Make()
		Debug ".Make()"
		SetSql ""
		SetSql "select"
		SetSql "distinct"
		SetSql " y.SYUKA_YMD SYUKA_YMD"
		SetSql ",y.UNSOU_KAISHA CampName"
		SetSql ",y.OKURI_NO DelvNo"
		SetSql ",Max(Convert(y.KUTI_SU,sql_decimal)) yQty"
		SetSql ",Max(Convert(y.JURYO,sql_decimal)) yWeight"
		SetSql ",Max(Convert(y.SAI_SU,sql_decimal)) ySai"
		SetSql ",d.DelvNo dDelvNo"
'		SetSql "from del_syuka_h y"
		SetSql "from " & GetOption("tbl","del_syuka_h") & " y"
		SetSql "left outer join DelivSts d"
		SetSql " on (y.SYUKA_YMD = d.SYUKA_YMD"
		SetSql " and y.UNSOU_KAISHA = d.CampName"
		SetSql " and y.OKURI_NO = d.DelvNo"
		SetSql " )"
		SetSql "where y.OKURI_NO<>''"
		if inStr(strSyukaYmd,"%") > 0 then
			SetSql "and y.SYUKA_YMD like '" & strSyukaYmd & "'"
		else
			SetSql "and y.SYUKA_YMD = '" & strSyukaYmd & "'"
		end if
		SetSql "group by"
		SetSql " y.SYUKA_YMD"
		SetSql ",y.UNSOU_KAISHA"
		SetSql ",y.OKURI_NO"
		SetSql ",d.DelvNo"
		SetSql "order by"
		SetSql " y.SYUKA_YMD"
		SetSql ",y.UNSOU_KAISHA"
		SetSql ",y.OKURI_NO"
		Debug ".Make():" & strSql
		set objRs = objDB.Execute(strSql)
		do while objRs.Eof = False
			Call MakeData()
			objRs.MoveNext
		loop
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'MakeData() 1行読込
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData()"
		Call DispLine()
		if GetField("DelvNo") = GetField("dDelvNo") then
			WScript.StdOut.WriteLine ":登録済み"
			exit function
		end if
		SetSql ""
		SetSql "insert into DelivSts ("
		SetSql " SYUKA_YMD"
		SetSql ",CampName"
		SetSql ",DelvNo"
		SetSql ",yQty"
		SetSql ",yWeight"
		SetSql ",ySai"
		SetSql ",EntID"
		SetSql ") values ("
		SetSql " '" & GetField("SYUKA_YMD") & "'"
		SetSql ",'" & GetField("CampName") & "'"
		SetSql ",'" & GetField("DelvNo") & "'"
		SetSql "," & CDbl(GetField("yQty"))
		SetSql "," & CDbl(GetField("yWeight"))
		SetSql "," & CDbl(GetField("ySai"))
		SetSql ",'DelivSts.vbs'"
		SetSql ")"
		on error resume next
		Call objDB.Execute(strSql)
		WScript.StdOut.Write ":" & "0x" & Hex(Err.Number) ' & ":" & Err.Description
		on error goto 0
		WScript.StdOut.WriteLine 
	End Function
	'-------------------------------------------------------------------
	'1行表示
	'-------------------------------------------------------------------
	Private objF
	Private Function DispLine()
		Debug ".DispLine()"
		for each objF in objRs.Fields
			WScript.StdOut.Write RTrim("" & objF)
			WScript.StdOut.Write " "
		next
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
		Call objDB.Execute(strSql)
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field値
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		strField = RTrim("" & objRs.Fields(strName))
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
	End Function
	'-------------------------------------------------------------------
	'フィールドセット
	'-------------------------------------------------------------------
	Private	bUpdate
	Private Function SetField(byVal strName,byVal strValue)
		dim	intLen
		intLen = objRs.Fields(strName).DefinedSize
		Debug ".SetField():" & strName & "(" & intLen & "):" & strValue
		strValue = Get_LeftB(strValue,intLen)
		Debug ".SetField():" & strName & "(" & intLen & "):" & strValue
		Debug ".SetField():" & strName & "(" & intLen & "):" & RTrim("" & objRs.Fields(strName))
		if strName = "Weight" then
			if CCur(objRs.Fields(strName)) = CCur(strValue) then
				exit function
			end if
		else
			if RTrim("" & objRs.Fields(strName)) = RTrim("" & strValue) then
				exit function
			end if
		end if
		Debug ".SetField():更新"
'		WScript.StdOut.Write strName & ":" & objRs.Fields(strName) & "→" & strValue & " "
		bUpdate	= True
		select case strName
		case "Status1"
			if Left(objRs.Fields("Status1"),3) <> "配達中" then
				objRs.Fields("Status1Last") = objRs.Fields("Status1")
			end if
		case "Status2"
			objRs.Fields("Status2Last") = objRs.Fields("Status2")
		end select
		objRs.Fields(strName) = strValue
		objRs.Fields("UpdID")	= "DelivSts.vbs"
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
	Function GetOption(byval strName ,byval strDefault)
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
	Private	optAction
	Private	optDay
	Private	optId
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		optId = ""
		For Each strArg In WScript.Arguments.UnNamed
			optId = strArg
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "test"
				optTest	= True
			case "make"
				optAction = "make"
			case "check"
				optAction = "check"
			case "recheck"
			case "tbl"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
		strSyukaYmd = GetOption("make","")
		if strSyukaYmd = "" then
			dim	dtTmp
			dtTmp = DateAdd("d",-1,Now())
			strSyukaYmd = year(dtTmp) & right("0" & month(dtTmp),2) & Right("0" & day(dtTmp),2)
		end if
		optDay = GetOption("check","10")
	End Function
End Class
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objDelivSts
	Set objDelivSts = New DelivSts
	if objDelivSts.Init() <> "" then
		call usage()
		exit function
	end if
	call objDelivSts.Run()
	Set objDelivSts = Nothing
End Function
