Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "DelivSts.vbs [option] [�����No]"
	Wscript.Echo " /db:newsdc1 �f�[�^�x�[�X"
	Wscript.Echo " /make[:yyyymmdd]	DelivSts�ɓo�^(default)"
	Wscript.Echo " /check[:day]		�z�B�󋵃`�F�b�N"
	Wscript.Echo " /recheck			�z�B�����ă`�F�b�N"
	Wscript.Echo " /tbl:y_syuka_h	default:del_syuka_h"
	Wscript.Echo " /test			test�p(�X�V���Ȃ�)"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript DelivSts.vbs /db:newsdc5"
	Wscript.Echo "cscript DelivSts.vbs /db:newsdc5 /check:1"
End Sub
'-----------------------------------------------------------------------
'DelivSts
'2016.10.19 �z�B�󋵁F���R�ʉ^
'2018.02.08 �O��̔z�B�󋵂�(Status1Last,Status2Last)�ɕۑ�
'2018.02.14 �u�z�B���v�͑O��̔z�B��(Status1Last)�ɕۑ����Ȃ�
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
	'Run() ���s����
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
		SetSql "where CampName = '���R�ʉ^'"
'		SetSql "  and Status1 not like '�z�B����%'"
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
				WScript.StdOut.Write "��" & GetField("Status1")
				if optTest = True then
					WScript.StdOut.Write "(test)"
				else
					call objRs.Update
				end if
				strUpd = strUpd & "��" & GetField("Status1")
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
	'CheckData() 1�s�Ǎ�
	'-------------------------------------------------------------------
	Private	strDelvNo
	Private Function CheckData()
		Debug ".CheckData()"
'		DispLine
		WScript.StdOut.Write GetField("SYUKA_YMD")
		WScript.StdOut.Write " " & GetField("CampName")
'		if GetField("CampName") <> "���R�ʉ^" then
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
		        Case "���ו��z�B�󋵏ڍ�"
		            strStat = strLine
		        Case "���݂̔z�B��"
		            strStat = strLine
		        Case "���͂��\���/�z�B������"
		            strStat = strLine
		        Case "�x�X�d�b�ԍ�"
					intBr = 0
		            strStat = strLine
		        End Select
		    Case "���ו��z�B�󋵏ڍ�"
		        If strLine = "���₢���킹�ԍ�" Then
		            strStat = "���₢���킹�ԍ�"
		        End If
		    Case "���₢���킹�ԍ�"
				Call GetContents(strLine, strStat)
		        strStat = "��"
		    Case "��"
				Call GetContents(strLine, strStat)
		        strStat = "�d��"
		    Case "�d��"
				Call GetContents(strLine, strStat)
		        strStat = ""
		    Case "���݂̔z�B��"
				Call GetContents(strLine, strStat)
		        strStat = ""
		    Case "���͂��\���/�z�B������"
		        strStat = "���͂��\���/�z�B������1"
		    Case "���͂��\���/�z�B������1"
				Call GetContents(Trim(strLine), strStat)
		        strStat = ""
		    Case "�x�X�d�b�ԍ�"
		        strStat = "��t"
		    Case "��t"
				Call GetContents(strLine, strStat)
		        strStat = "����"
		    Case "����"
				Call GetContents(strLine, strStat)
		        strStat = "����"
		    Case "����"
				Call GetContents(strLine, strStat)
		        strStat = "���o"
		    Case "���o"
				Call GetContents(strLine, strStat)
		        strStat = "�z�B����"
		    Case "�z�B����"
				Call GetContents(strLine, strStat)
		        strStat = "�x�X�R�[�h"
			case "�x�X�R�[�h"
		        If strLine = "�x�X�d�b�ԍ�" Then
			        strStat = "�x�X�d�b�ԍ�0"
				end if
			case "�x�X�d�b�ԍ�0"
		        strStat = "�x�X�d�b�ԍ�1"
			case "�x�X�d�b�ԍ�1"
				Call GetContents(strLine, strStat)
		        strStat = "�x�X�d�b�ԍ�2"
			case "�x�X�d�b�ԍ�2"
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
	    Case "���ו��z�B�󋵏ڍ�"
	    Case "���₢���킹�ԍ�"
	        strValue = Split(strLine, " ")(0)
	    Case "��"
	        strValue = Split(strLine, " ")(0)
			SetField "Qty",strValue
	    Case "�d��"
	        strValue = Split(strLine, " ")(0)
			SetField "Weight",strValue
	    Case "���݂̔z�B��"
	        strValue = strLine
			SetField "Status1",strValue
	    Case "���͂��\���/�z�B������1"
	        strValue = strLine
			SetField "Status2",strValue
	    Case "��t","����","����","���o"
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
		    Case "��t"
				SetField "UkeDTm",strDtm
				SetField "UkeBr",strBr
				SetField "UkeBrTel",strBrTel
		    Case "����"
				SetField "HatDTm",strDtm
				SetField "HatBr",strBr
				SetField "HatBrTel",strBrTel
		    Case "����"
				SetField "ChaDTm",strDtm
				SetField "ChaBr",strBr
				SetField "ChaBrTel",strBrTel
		    Case "���o"
				SetField "MotDTm",strDtm
				SetField "MotBr",strBr
				SetField "MotBrTel",strBrTel
			end select
	    Case "�z�B����"
	        strValue = Split(strLine, strStat)(1)
			if inStr(strValue," ") > 0 then
				SetField "FinDTm",Split(strValue," ")(0) & " " & Split(strValue," ")(1)
			else
				SetField "FinDTm",strValue
			end if
	    Case "�x�X�d�b�ԍ�1","�x�X�d�b�ԍ�2"
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
	' ���ʂ̖⍇��No����z�B�󋵂��擾
	'-------------------------------------------------------------------
	Private	strID
	Private	strUrl
	Private	strBody
	Private Function GetTrackingBody()
	    GetTrackingBody = ""

		if GetField("CampName") <> "���R�ʉ^" then
			exit function
		end if
		if GetField("Status1") = "�z�B�����ł�" then
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
		Debug "�ڑ�:" & strUrl
	    'IE�̋N��
		if objIE is nothing then
			Debug "InternetExplorer.Application"
			Set objIE = CreateObject("InternetExplorer.Application")
			objIE.Visible = False
		end if
'		WScript.StdOut.Write strID
        objIE.Navigate strUrl
'		WScript.StdOut.Write ":"

        ' �y�[�W����荞�܂��܂ő҂�
        Do While objIE.Busy or objIE.readyState <> READYSTATE_COMPLETE
			WScript.StdOut.Write "."
            WScript.Sleep 1000
        Loop
'        Do While objIE.readyState <> READYSTATE_COMPLETE
'			WScript.StdOut.Write "*"
'			Debug "�Ǎ��� " & objIE.Document.readyState
'           WScript.Sleep 3000
'        Loop
'		WScript.StdOut.WriteLine
        ' �e�L�X�g�`���ŏo��
		strBody = objIE.Document.Body.InnerText
'		strBody = objIE.Document.Body.textContent
		' �g�s�l�k�`���ŏo��
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
	'MakeData() 1�s�Ǎ�
	'-------------------------------------------------------------------
	Private Function MakeData()
		Debug ".MakeData()"
		Call DispLine()
		if GetField("DelvNo") = GetField("dDelvNo") then
			WScript.StdOut.WriteLine ":�o�^�ς�"
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
	'1�s�\��
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
	'SQL������ǉ�
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
	'Sql���s
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		Call objDB.Execute(strSql)
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Field�l
	'-------------------------------------------------------------------
	Private Function GetField(byVal strName)
		dim	strField
		strField = RTrim("" & objRs.Fields(strName))
		Debug ".GetField():" & strName & ":" & strField
		GetField = strField
	End Function
	'-------------------------------------------------------------------
	'�t�B�[���h�Z�b�g
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
		Debug ".SetField():�X�V"
'		WScript.StdOut.Write strName & ":" & objRs.Fields(strName) & "��" & strValue & " "
		bUpdate	= True
		select case strName
		case "Status1"
			if Left(objRs.Fields("Status1"),3) <> "�z�B��" then
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
			'** Asc�֐��ŕ����R�[�h�擾
			iAscCode = Asc(Mid(a_Str, iCount, 1))
			'** ���p�͕����R�[�h�̒�����2�A�S�p��4(2�ȏ�)�Ƃ��Ĕ��f
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
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'���b�Z�[�W�\��
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'�I�v�V�����擾
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
	'Init() �I�v�V�����`�F�b�N
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
				Init = "�I�v�V�����G���[:" & strArg
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
'���C��
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
