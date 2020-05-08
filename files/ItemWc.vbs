Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "ItemWc.vbs [option]"
	Wscript.Echo " /db:newsdc1	�f�[�^�x�[�X"
	Wscript.Echo " /update:off"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo ItemWc.vbs /db:newsdc1"
End Sub
'-----------------------------------------------------------------------
'ItemSize(�ː��e�[�u��)�X�V
'2016.10.27 �V�K
'2016.10.29 Insert/Update�� @@rowcount �\��
'2016.10.31 �O���ː��^�����̏ꍇ �����S���l�̌ܓ�
'-----------------------------------------------------------------------
Class Item
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	optUpdate
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		optUpdate = GetOption("update","on")
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
		OpenDB
		Make
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Make() �d����WC �X�V
	'-----------------------------------------------------------------------
    Public Function Make()
		Debug ".Make()"
		SetSql	""
		SetSql	"select"
		SetSql	"distinct"
		SetSql	" y.JGYOBU"
		SetSql	",y.HIN_NO"
		SetSql	",y.SHIIRE_WORK_CENTER"
		SetSql	",i.TORI_SHIIRE_WORK_CTR"
		SetSql	",Max(y.SYUKO_YMD) SYUKO_YMD"
		SetSql	"from y_glics y"
		SetSql	"inner join item i on (i.JGYOBU = y.JGYOBU and i.NAIGAI = '1' and i.HIN_GAI = y.HIN_NO)"
		SetSql	"where y.SHIIRE_WORK_CENTER <> ''"
		if strJgyobu <> "" then
			SetSql	"  and y.JGYOBU = '" & strJgyobu & "'"
		end if
		SetSql	"group by"
		SetSql	" y.JGYOBU"
		SetSql	",y.HIN_NO"
		SetSql	",y.SHIIRE_WORK_CENTER"
		SetSql	",i.TORI_SHIIRE_WORK_CTR"
		SetSql	"order by"
		SetSql	" y.JGYOBU"
		SetSql	",y.HIN_NO"
		SetSql	",SYUKO_YMD desc"
		SetSql	",y.SHIIRE_WORK_CENTER"
		SetSql	",i.TORI_SHIIRE_WORK_CTR"
		Debug ".Make():" & strSql
		WScript.StdErr.Write "������..."
		set objRs = objDB.Execute(strSql)
		WScript.StdErr.WriteLine "Eof:" & objRs.Eof
		prvJgyobu = ""
		prvPn = ""
		do while objRs.Eof = False
			SetWc
			objRs.MoveNext
		loop
		WScript.StdOut.WriteLine
		objRs.Close
		set objRs = Nothing
	End Function
	'-------------------------------------------------------------------
	'SetWc() 1�s�Ǎ�
	'-------------------------------------------------------------------
	dim	prvJgyobu
	dim	prvPn
	dim	curJgyobu
	dim	curPn
	Private Function SetWc()
		Debug ".SetWc()"
		curJgyobu = GetField("JGYOBU")
		curPn = GetField("HIN_NO")
		if curJgyobu = prvJgyobu and curPn = prvPn then
'			WScript.StdOut.Write "."
			exit function
		end if
		prvJgyobu = curJgyobu
		prvPn = curPn
		if GetField("SHIIRE_WORK_CENTER") = GetField("TORI_SHIIRE_WORK_CTR") then
			exit function
		end if
		WScript.StdOut.WriteLine
		WScript.StdOut.Write T(GetField("JGYOBU"),-2)
		WScript.StdOut.Write T(GetField("HIN_NO"),-20)
		WScript.StdOut.Write T(GetField("SHIIRE_WORK_CENTER"),-9)
		WScript.StdOut.Write T(GetField("TORI_SHIIRE_WORK_CTR"),-9)
		WScript.StdOut.Write T(GetField("SYUKO_YMD"),-9)
		Update
	End Function
	'-------------------------------------------------------------------
	'Update
	'-------------------------------------------------------------------
	Private	Function Update()
		Debug ".Update()"
		if optUpdate = "off" then
			WScript.StdOut.Write " Upd:off"
			exit function
		end if
		SetSql ""
		SetSql "update Item"
		SetSql "set TORI_SHIIRE_WORK_CTR = '" & GetField("SHIIRE_WORK_CENTER") & "'"
		SetSql ",UPD_TANTO = 'ItemWc'"
		SetSql ",UPD_DATETIME = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
		SetSql "where JGYOBU = '" & GetField("JGYOBU") & "'"
		SetSql "and NAIGAI = '1'"
		SetSql "and HIN_GAI = '" & GetField("HIN_NO") & "'"
		WScript.StdOut.Write " Upd:" '& strSql
		CallSql strSql
	End Function
	'-----------------------------------------------------------------------
	'T() ������
	'-----------------------------------------------------------------------
	Private Function T(byVal v,byVal i)
		if i > 0 then
			T = right(space(i) & v,i)
		else
			i = i * -1
			T = LeftB(v & space(i),i)
		end if
	End Function
	'-----------------------------------------------------------------------
	'LeftB() ������
	'-----------------------------------------------------------------------
	Private Function LeftB(byVal a_Str, byVal a_int)
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
		if LenB(iLeftStr) < a_int then
			iLeftStr = iLeftStr & space(a_int - LenB(iLeftStr))
		end if
		LeftB = iLeftStr
	End Function
	'-----------------------------------------------------------------------
	'LenB() ������
	'-----------------------------------------------------------------------
	Function LenB(byVal a_Str)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			LenB = 0
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
		Next
		LenB = iLenCount
	End Function
	'-------------------------------------------------------------------
	'GetDbl
	'-------------------------------------------------------------------
	Private Function GetDbl(byVal v)
		GetDbl = 0
		if v = "" then exit function
		GetDbl = CDbl(v)
	End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
		objDB.Open strDbName
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		objDB.Close
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
		objDb.Execute strSql
		on error goto 0
		dim	intNumver
		dim	strDescription
		intNumver = Err.Number
		strDescription	= Err.Description
		if intNumver = 0 then
			dim	objRc
			set objRc = objDb.Execute("select @@rowcount")
			WScript.StdOut.Write objRc.Fields(0)
		else
			WScript.StdOut.Write RTrim("0x" & Hex(intNumver) & " " & strDescription)
		end if
    End Function
	'-------------------------------------------------------------------
	'Field�l
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
	Private	strJgyobu
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		strJgyobu = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strJgyobu = "" then
				strJgyobu = strArg
			else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "update"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
End Class
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objItem
	Set objItem = New Item
	if objItem.Init() <> "" then
		call usage()
		exit function
	end if
	call objItem.Run()
End Function
