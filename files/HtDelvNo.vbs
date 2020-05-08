Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "HtDelvNo.vbs [option]"
	Wscript.Echo " /db:newsdc1 �f�[�^�x�[�X"
	Wscript.Echo " /make:ssx   ssx(default)"
	Wscript.Echo " /make:b2    b2"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript HtDelvNo.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'HtDelvNo
'2016.10.26 �V�K �Y�@������o�^
'2016.10.27 EntTm ���Z�b�g����悤�ɕύX(�g���K�[�폜)
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

Class HtDelvNo
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strAction	' make/csv
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		optNew = "new"
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
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		Call Make()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Make() 
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		SetSql ""
		if GetOption("make","ssx") = "ssx" then
'			SetSql ""
'			SetSql "delete from HtDelvNo where CampName='SSX'"
'			Disp strSql
'			Call objDB.Execute(strSql)
			SetSql ""
			SetSql "select"
			SetSql "distinct"
			SetSql " 'SSX' CampName"
			SetSql ",RTrim(ID)+'0' DelvNo"
			SetSql ",SCode ChoCode"
			SetSql ",SName1 ChoName"
			SetSql "from RSmile"
			SetSql "where (RTrim(ID)+'0') not in (select distinct RTrim(DelvNo) from HtDelvNo where CampName='SSX')"
		else
			SetSql ""
			SetSql "delete from HtDelvNo where CampName='���R�ʉ^'"
			Disp strSql
			Call objDB.Execute(strSql)
			SetSql ""
			SetSql "select"
			SetSql "distinct"
			SetSql " '���R�ʉ^' CampName"
			SetSql ",RTrim(c04) DelvNo"
			SetSql ",RTrim(c08) ChoCode"
			SetSql ",RTrim(c14)+RTrim(c15)+RTrim(c16) ChoName"
			SetSql "from b2excel"
'			SetSql "where idRow > 1"
			SetSql "where RTrim(c01) not in ('','���q�l�Ǘ��ԍ�')"
			SetSql "and RTrim(c04) <> ''"
'			SetSql "and RTrim(c04) not in (select distinct RTrim(DelvNo) from HtDelvNo where CampName='���R�ʉ^')"
		end if
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
		SetSql ""
		SetSql "insert into HtDelvNo"
		SetSql "("
		SetSql " CampName"
		SetSql ",DelvNo"
		SetSql ",ChoCode"
		SetSql ",ChoName"
		SetSql ",EntID"
		SetSql ",EntTM"
		SetSql ") values ("
		SetSql " '" & GetField("CampName") & "'"
		SetSql ",'" & GetField("DelvNo") & "'"
		SetSql ",'" & GetField("ChoCode") & "'"
		SetSql ",'" & LeftB(GetField("ChoName"),20) & "'"
		SetSql ",'HtDelvNo.vbs'"
		SetSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
		SetSql ")"
		Call CallSql(strSql)
	End Function
	'-------------------------------------------------------------------
	'1�s�\��
	'-------------------------------------------------------------------
	Private objF
	Private Function DispLine()
		Debug ".DispLine()"
		WScript.StdOut.Write GetField("CampName")
		WScript.StdOut.Write " " & GetField("DelvNo")
		WScript.StdOut.Write " " & Left(GetField("ChoCode") & Space(10),10)
		WScript.StdOut.Write GetField("ChoName")
		WScript.StdOut.WriteLine
'		for each objF in objRs.Fields
'			WScript.StdOut.Write RTrim(objF)
'			WScript.StdOut.Write " "
'		next
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
'		on error resume next
		Call objDB.Execute(strSql)
'		on error goto 0
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
	'-----------------------------------------------------------------------
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
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
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
	Private	optNew
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			Init = "�I�v�V�����G���[:" & strArg
			Disp Init
			Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "make"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-------------------------------------------------------------------
	'LeftB()
	'-------------------------------------------------------------------
	Private Function LeftB(byVal a_Str,byVal a_int)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			LeftB = ""
			Exit Function
		End If
		If a_int = 0 Then
			LeftB = ""
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
		LeftB = iLeftStr
	End Function
End Class
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objHtDelvNo
	Set objHtDelvNo = New HtDelvNo
	if objHtDelvNo.Init() <> "" then
		call usage()
		exit function
	end if
	call objHtDelvNo.Run()
End Function
