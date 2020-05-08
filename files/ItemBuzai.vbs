Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "ItemBuzai.vbs [option] <filename>"
	Wscript.Echo " /db:newsdc5"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo ItemBuzai.vbs \\w5\y\���i��\���i������\������\49��\HEG�Ɩ�201603.xlsx"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objItemBuzai
	Set objItemBuzai = New ItemBuzai
	if objItemBuzai.Init() <> "" then
		call usage()
		exit function
	end if
	call objItemBuzai.Run()
End Function
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset		= 1
Const adOpenDynamic		= 2
Const adOpenStatic		= 3

'---- LockTypeEnum Values ----
Const adLockReadOnly 		= 1
Const adLockPessimistic 	= 2
Const adLockOptimistic 		= 3
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

Const adStateClosed		= 0 ' �I�u�W�F�N�g�����Ă���

Const ForReading		= 1
Const ForWriting		= 2
Const ForAppending		= 8
Const adSearchForward	= 1

Const xlUp = -4162

'-----------------------------------------------------------------------
'�o�c����
'-----------------------------------------------------------------------
Class ItemBuzai
	Private	strDBName
	Private	objDB
	Private	objSrcRs
	Private	strFilename
	Private	strPassword
	'-----------------------------------------------------------------------
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Private Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.Echo strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'���b�Z�[�W�\��
	'-----------------------------------------------------------------------
	Private Sub Disp(byVal strMsg)
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
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strFilename = "" then
				strFilename = strArg
			elseif strPassword = "" then
				strPassword = strArg
			else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		if strFilename = "" then
			Init = "." & strArg
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db","newsdc5")
		set objDB = nothing
		set objSrcRs = nothing
		set	objXL = nothing
		set	objBk = nothing
        strFilename = ""
		strPassword = ""
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		if not objBk is nothing then
			Debug ".Class_Terminate():Close:" & objBk.Name
			call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call CreateExcelApp()
		Call OpenExcel()
		Call OpenDB()
		Call LoadExcel()
		Call CloseDB()
	End Function
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	Private	objXL
	Private Function CreateExcelApp()
		Debug ".CreateExcelApp()"
		if objXL is nothing then
			Set objXL = WScript.CreateObject("Excel.Application")
			Set	objBk = nothing
		end if
	end function
	'-------------------------------------------------------------------
	'GetAbsPath
	'-------------------------------------------------------------------
	Private Function GetAbsPath(byVal strPath)
		Dim objFileSys
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		strPath = objFileSys.GetAbsolutePathName(strPath)
		Set objFileSys = Nothing
		GetAbsPath = strPath
	End Function
	'-------------------------------------------------------------------
	'GetScriptPath
	'-------------------------------------------------------------------
	Private Function GetScriptPath()
		GetScriptPath = Left(WScript.ScriptFullName,Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
	End Function
	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	Private	objBk
	Private Function OpenExcel()
		Debug ".OpenExcel():" & strFilename
		if strFilename = "" then
			exit function
		end if
		if objBk is nothing then
			Set objBk = objXL.Workbooks.Open(GetAbsPath(strFilename),False,True,,strPassword)
		end if
	end function
	'-------------------------------------------------------------------
	'�Ǎ����� Book
	'-------------------------------------------------------------------
	Private	objSt
	Private Function LoadExcel()
		if objBk is nothing then
			exit function
		end if
		Debug ".LoadExcel():" & objBk.Name
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
	end function
	'-------------------------------------------------------------------
	'�Ǎ����� Sheet
	'-------------------------------------------------------------------
	Private Function LoadXls()
		Debug ".LoadXls():" & objSt.Name
		if LoadXlsFile() then
			exit function
		end if
	end function
	'-------------------------------------------------------------------
	'�Ǎ����� FILE
	'-------------------------------------------------------------------
	Private Function LoadXlsFile()
		LoadXlsFile = False
		if objSt.Name <> "���i�\" then
			exit function
		end if
		Debug ".LoadXlsFile():" & objSt.Name
		Call OpenRs()
		dim	lngRow
		dim	lngRowMax
		lngRowMax = objSt.Range("A65535").End(xlUp).Row
		for lngRow = 1 to lngRowMax
			if AddRecord(lngRow) then
				exit for
			end if
		next
		LoadXlsFile = True
	end function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDb = Wscript.CreateObject("ADODB.Connection")
'		objDb.CursorLocation = adUseClient
		Call objDb.Open(strDbName)
		Set objRs = nothing
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		if not objRs is nothing then
			Call objRs.Close()
		end if
		Call objDb.Close()
		set objDb = Nothing
    End Function
	'-------------------------------------------------------------------
	'OpenRs
	'-------------------------------------------------------------------
	Private	strTable
	Private	objRs
	Private	objTable
	Private	strHIN_GAI
	Private Function OpenRs()
		strTable = "ItemBuzai"
		Debug ".OpenRs():" & strTable
'		Set objTable = Wscript.CreateObject("ADODB.Recordset")
'		objTable.Open strTable, objDb , adOpenDynamic, adLockOptimistic , adCmdTableDirect
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
'		objRs.Open strTable, objDb , adOpenKeyset, adLockOptimistic , adCmdTableDirect
	End Function
	'-------------------------------------------------------------------
	'���R�[�h����(Find)
	'-------------------------------------------------------------------
	Private Function GetRs()
		Debug ".GetRs():" & strHIN_GAI
		dim	strSql
		strSql = "select * from ItemBuzai" & vbCrLf
		strSql = strSql & "where HIN_GAI='" & strHIN_GAI & "'"
		if objRs.State <> adStateClosed then 
			objRs.Close
		end if
		objRs.Open strSql, objDb , adOpenKeyset, adLockOptimistic
		if objRs.EOF then
			objRs.Close
			objRs.Open strTable, objDb , adOpenKeyset, adLockOptimistic , adCmdTableDirect
			objRs.AddNew
			objRs.Fields("HIN_GAI")	= strHIN_GAI
		end if
		Debug ".GetRs():" & objRs.Fields("HIN_GAI") & " " & objRs.Fields("Kubun")
	End Function
	'-------------------------------------------------------------------
	'�l
	'-------------------------------------------------------------------
	Private Function GetValue(byVal v)
		GetValue = 0
		if Trim(v) = "" then
			exit function
		end if
		GetValue = CCur(v)
	End Function
	'-------------------------------------------------------------------
	'���R�[�h�ǉ�
	'-------------------------------------------------------------------
	Private Function AddRecord(byVal lngRow)
		AddRecord = False
		Debug ".AddRecord():" & lngRow
		dim	strKubun
		strHIN_GAI	= RTrim(objSt.Range("A" & lngRow) & "")
		strKubun	= RTrim(objSt.Range("B" & lngRow) & "")

		Disp lngRow & ":" & strHIN_GAI & " " & strKubun
		if strHIN_GAI = "" then
			exit function
		end if
		if strHIN_GAI = "�i��" then
			exit function
		end if
		Call GetRs()
		objRs.Fields("Kubun")		= strKubun
		objRs.Fields("Price")		= GetValue(objSt.Range("D" & lngRow) & "")		'���ϒP��
		objRs.Fields("PrcHanbai")	= GetValue(objSt.Range("L" & lngRow) & "")     '�̔�
		objRs.Fields("PrcShizai")	= GetValue(objSt.Range("M" & lngRow) & "")     '������
		objRs.Fields("PrcNaka")		= GetValue(objSt.Range("P" & lngRow) & "")     '�����Y��
		objRs.Fields("PrcSKosyo")	= GetValue(objSt.Range("S" & lngRow) & "")     '�H��
		objRs.Fields("PrcPP")		= GetValue(objSt.Range("V" & lngRow) & "")     'PF�ǉ��H
		objRs.Fields("PrcPE")		= GetValue(objSt.Range("Y" & lngRow) & "")     'PE�ǉ��H
		objRs.Fields("PrcPEs")		= GetValue(objSt.Range("AB" & lngRow) & "")    'PE�Ǖ�����
		objRs.Fields("PrcBasyo")	= GetValue(objSt.Range("AE" & lngRow) & "")    '��Əꏊ
		dim	strMsg
		strMsg = ""
		on error resume next
'			objRs.UpdateBatch
			objRs.Update
			select case Err.Number
			case 0
			case &h80004005
				strMsg = "����d�o�^��"
			case else
				strMsg = "0x" & Hex(Err.Number) & " " & Err.Description
				AddSyushi = True
			end select
			if strMsg <> "" then
				Call objRs.CancelUpdate
				Disp strMsg
			end if
		on error goto 0
	End Function
End Class
