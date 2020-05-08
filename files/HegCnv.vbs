Option Explicit
Const xlUp = -4162
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "HegCnv.vbs [option] �V�[�g�� �t�@�C����"
	Wscript.Echo " /db:newsdc5	�f�[�^�x�[�X"
'	Wscript.Echo " /s:10000	    �J�n�s"
'	Wscript.Echo " /l:100       �ǂݍ��ލs��"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript HegCnv.vbs /db:newsdc5 ���i�\ \\w5\y\�G���i��\���i������\HEG�Ɩ�201607Check.xlsx"
End Sub

'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objHegCnv
	Set objHegCnv = New HegCnv
	if objHegCnv.Init() <> "" then
		call usage()
		exit function
	end if
	call objHegCnv.Run()
End Function

'-----------------------------------------------------------------------
'HegCnv
'-----------------------------------------------------------------------
Class HegCnv
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	strAction
	Private	strSheetName
	Private	strFileName
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
			if strSheetName = "" then
				strSheetName = strArg
			elseif strFileName = "" then
				strFileName = strArg
			else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		if strSheetName = "" then
			Init = "�V�[�g�����w��"
			Disp Init
			Exit Function
		end if
		if strFileName = "" then
			Init = "�t�@�C�������w��"
			Disp Init
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "j"
			case "s"
			case "l"
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
		strSheetName	= ""
		strFileName		= ""
		strDBName = GetOption("db"	,"newsdc")
		set objDB = nothing
		set objRs = nothing
		set objBk = nothing
		set objXL = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		if not objBk is nothing then
			Call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
		set objRs = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		Call Load()
		Call CloseDb()
	End Function
	'-----------------------------------------------------------------------
	'Load() �Ǎ�
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load():" & strFileName
		select case FileType()
		case "excel"
			Call CreateExcelApp()
			Call OpenExcel()
			if LoadExcel() <> strSheetName then
				Wscript.Echo "�V�[�g�Ȃ�:" & strSheetName
			end if
		case "csv"
'			Call OpenCsv()
'			Call LoadCsv()
'			Call CloseCsv()
		end select
	End Function
	'-------------------------------------------------------------------
	'�t�@�C���̎��
	'-------------------------------------------------------------------
	Private Function FileType()
		FileType = ""
		select case lcase(fileExt(strFileName))
		case "xls","xlsx"	FileType = "excel"
		case "csv"			FileType = "csv"
		end select
		Debug(".FileType():" & FileType)
	End Function
	'-------------------------------------------------------------------
	'�g���q
	'-------------------------------------------------------------------
	Private Function fileExt(byVal f)
		dim	fobj
		set fobj = CreateObject("Scripting.FileSystemObject")
		dim	strExt
		strExt = fobj.GetextensionName(f)
		set fobj = Nothing
		fileExt = strExt
	End Function
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	Private	objXL
	Private Function CreateExcelApp()
		Debug(".CreateExcelApp()")
		if objXL is nothing then
			Debug(".CreateExcelApp():CreateObject(Excel.Application)")
			Set objXL = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	Private	objBk
	Private Function OpenExcel()
		Debug(".OpenExcel()")
		if objBk is nothing then
			Debug(".OpenExcel().Open=" & GetAbsPath(strFileName))
			Set objBk = objXL.Workbooks.Open(GetAbsPath(strFileName),False,True,,"")
		end if
	end function
	'-------------------------------------------------------------------
	'��΃p�X
	'-------------------------------------------------------------------
	Private Function GetAbsPath(byVal strPath)
		Dim objFileSys
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		strPath = objFileSys.GetAbsolutePathName(strPath)
		Set objFileSys = Nothing
		GetAbsPath = strPath
	End Function
	'-------------------------------------------------------------------
	'�Ǎ�����
	'-------------------------------------------------------------------
	Public	objSt
	Private Function LoadExcel()
		Debug ".LoadExcel()"
		LoadExcel = ""
		if objBk is nothing then
			exit function
		end if
		if strSheetName <> "" then
			For each objSt in objBk.Worksheets
				if objSt.Name = strSheetName then
					Call LoadXls()
					LoadExcel = strSheetName
					exit function
				end if
			Next
		end if
	end function
	'-------------------------------------------------------------------
	'�Ǎ�����(�V�[�g)
	'-------------------------------------------------------------------
	Private Function LoadXls()
		Debug ".LoadXls()"
		if objSt is nothing then
			exit function
		end if
		Call LoadData()
	end function
	'-------------------------------------------------------------------
	'�V�[�g�Ǎ�
	'-------------------------------------------------------------------
	Private	clsData
	Private Function LoadData()
		Debug ".LoadData():" & objSt.Name
		' �i�ڃJ�e�S���[/���уJ�e�S���[
'		Set clsData = New BoHinmoku
'		LoadData = clsData.Init(me)
'		Set clsData = Nothing
'		if LoadData then
'			exit function
'		end if
		' �i�����X�g�⑫�f�[�^
'		Set clsData = New BoHosoku
'		LoadData = clsData.Init(me)
'		Set clsData = Nothing
'		if LoadData then
'			exit function
'		end if
		' �T�t�@�C�A�[���\��
'		Set clsData = New SaDelv
'		LoadData = clsData.Init(me)
'		Set clsData = Nothing
'		if LoadData then
'			exit function
'		end if
		' Active�f�[�^
'		Set clsData = New AcData
'		LoadData = clsData.Init(me)
'		Set clsData = Nothing
'		if LoadData then
'			exit function
'		end if
		' Heg���i�\
		Set clsData = New HegPrice
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
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
	'Field��
	'-------------------------------------------------------------------
	Public Function GetFields(byVal strTable)
		Debug ".GetFields():" & strTable
		dim	strFields
		strFields = ""
		dim	objRs
		set objRS = objDB.Execute("select top 1 * from " & strTable)
		dim	objF
		for each objF in objRS.Fields
			if strFields <> "" then
				strFields = strFields & ","
			end if
			strFields = strFields & objF.Name
		next
		set objRs = nothing
		GetFields = strFields
	End Function
End Class

' �^�C�g��������
Private Function getTitle(byVal strT)
	getTitle = Replace(strT,vbLf,"")
End Function

' �^�C�g����r
Private Function CompTitle(byVal strS,byVal strD)
	CompTitle = true
	if getTitle(strS) = getTitle(strD) then
		CompTitle = false
	end if
End Function

' Excel�ŏI�s
Private Function excelGetMaxRow(objSt,byVal strCol,byVal lngRow)
	dim lngRowMax
	lngRowMax = objSt.rows.count
	lngRowMax = objSt.Range(strCol & lngRowMax).End(xlUp).Row
	if lngRow > lngRowMax then
		lngRowMax = lngRow
	end if
	excelGetMaxRow = lngRowMax
End Function

' �i�ڃJ�e�S���[/���уJ�e�S���[
Class BoHinmoku
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		'PN����_�i�ڔԍ�	�i��_�i�ڃR�[�h	�i��_�i�ڃJ�e�S���[��
		'PN����_�i�ڔԍ�	�i��_�i�ڃR�[�h	�i��_�i�ڃJ�e�S���[�ʖ�
		if	getTitle(objSt.Range("A4")) <> "" then
			exit function
		end if
		if	getTitle(objSt.Range("B4")) <> "PN����_�i�ڔԍ�" then
			exit function
		end if
		if	getTitle(objSt.Range("C4")) <> "�i��_�i�ڃR�[�h" then
			exit function
		end if
		select case getTitle(objSt.Range("D4"))
		case "�i��_�i�ڃJ�e�S���[��","�i��_�i�ڃJ�e�S���[�ʖ�"
		case else
			exit function
		end select
		Call Load()
		Init = true
	End Function
	Public Function Load()
		lngRowTop = objParent.GetOption("s",5)
		lngRowEnd = excelGetMaxRow(objSt,"B",5)
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
		next
	End Function
	Private	strJGYOBU
	Private	strHIN_GAI
	Private	strHinmokuCode
	Private	strHinmokuName
	Public Function Disp()
		strJGYOBU		= objParent.strJGYOBU
		strHIN_GAI		= objSt.Range("B" & lngRow)
		strHinmokuCode	= objSt.Range("C" & lngRow)
		strHinmokuName	= objSt.Range("D" & lngRow)
		objParent.Disp	lngRow & "/" & lngRowEnd	_
				& " " & strJGYOBU	_
				& " " & strHIN_GAI	_
				& " " & strHinmokuCode	_
				& " " & strHinmokuName
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into PnHinmoku"
		strSql = strSql & " (JGYOBU"
		strSql = strSql & " ,HIN_GAI"
		strSql = strSql & " ,HinmokuCode"
		strSql = strSql & " ,EntID"
		strSql = strSql & " ) values ("
		strSql = strSql & "  '" & strJGYOBU & "'" 
		strSql = strSql & " ,'" & strHIN_GAI & "'"
		strSql = strSql & " ,'" & strHinmokuCode & "'"
		strSql = strSql & " ,'BoCnv'"
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
		strSql = strSql & "insert into Hinmoku"
		strSql = strSql & " (JGYOBU"
		strSql = strSql & " ,HinmokuCode"
		strSql = strSql & " ,HinmokuName"
		strSql = strSql & " ,EntID"
		strSql = strSql & " ) values ("
		strSql = strSql & "  '" & strJGYOBU & "'" 
		strSql = strSql & " ,'" & strHinmokuCode & "'"
		strSql = strSql & " ,'" & strHinmokuName & "'"
		strSql = strSql & " ,'BoCnv'"
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
	End Function
End Class

' �i�����X�g�⑫�f�[�^
Class BoHosoku
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		'PN����_���Y�Ǘ����Ə�R�[�h	PN����_�i�ڔԍ�	PN����_�i�ږ�	PN���ʁiPN)_��\�@��i�ڃR�[�h	�@��i��_�@��i�ڃJ�e�S���[��	PN����_�i�ڃR�[�h	�i��_�i�ڃJ�e�S���[��	�o�m����_���������J�n�N��	�o�m����_�������Y�ŐؔN��	PN����_�����������i�敪	�o�m����_�A�o�����J�n�N��	�o�m����_�A�o���Y�ŐؔN��	PN����_�C�O�������i�敪	PN����_���l��
		if	getTitle(objSt.Range("A4")) <> "" then
			exit function
		end if
		if	getTitle(objSt.Range("B4")) <> "PN����_���Y�Ǘ����Ə�R�[�h" then
			exit function
		end if
		if	getTitle(objSt.Range("C4")) <> "PN����_�i�ڔԍ�" then
			exit function
		end if
		if	getTitle(objSt.Range("D4")) <> "PN����_�i�ږ�" then
			exit function
		end if
		if	getTitle(objSt.Range("E4")) <> "PN���ʁiPN)_��\�@��i�ڃR�[�h" then
			exit function
		end if
		if	getTitle(objSt.Range("F4")) <> "�@��i��_�@��i�ڃJ�e�S���[��" then
			exit function
		end if
		if	getTitle(objSt.Range("G4")) <> "PN����_�i�ڃR�[�h" then
			exit function
		end if
		if	getTitle(objSt.Range("H4")) <> "�i��_�i�ڃJ�e�S���[��" then
			exit function
		end if
		if	getTitle(objSt.Range("I4")) <> "�o�m����_���������J�n�N��" then
			exit function
		end if
		if	getTitle(objSt.Range("J4")) <> "�o�m����_�������Y�ŐؔN��" then
			exit function
		end if
		if	getTitle(objSt.Range("K4")) <> "PN����_�����������i�敪" then
			exit function
		end if
		if	getTitle(objSt.Range("L4")) <> "�o�m����_�A�o�����J�n�N��" then
			exit function
		end if
		if	getTitle(objSt.Range("M4")) <> "�o�m����_�A�o���Y�ŐؔN��" then
			exit function
		end if
		if	getTitle(objSt.Range("N4")) <> "PN����_�C�O�������i�敪" then
			exit function
		end if
		if	getTitle(objSt.Range("O4")) <> "PN����_���l��" then
			exit function
		end if
		' �f�[�^�Ǎ�
		Call Load()
		Init = true
	End Function
	Private	lngLimit
	Private	lngCount
	Public Function Limit()
		objParent.Debug ".Limit():" & lngCount & "/" & lngLimit
		Limit = false
		lngCount = lngCount + 1
		if lngLimit <> 0 then
			if lngCount >= lngLimit then
				Limit = true
			end if
		end if
	End Function
	Private	strFields
	Public Function Load()
		strFields = objParent.GetFields("PnHosoku")
		lngRowTop = objParent.GetOption("s",5)
		lngRowEnd = excelGetMaxRow(objSt,"B",5)
		lngCount  = 0
		lngLimit  = CLng(objParent.GetOption("l",0))
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
			if Limit() then
				exit for
			end if
		next
	End Function
	Private	aryCell(255)
	Public Function Disp()
		dim	i
		dim	strMsg
		strMsg = lngRow & "/" & lngRowEnd
		dim	strValues
		strValues = ""
		for i = 1 to 14
			aryCell(i) = objSt.Range("A" & lngRow).Offset(0,i)
			strMsg = strMsg & " " & aryCell(i)
			if strValues <> "" then
				strValues = strValues & ","
			end if
			strValues = strValues & "'" & aryCell(i) & "'"
		next
		objParent.Disp	strMsg
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into PnHosoku"
		strSql = strSql & " ("
		strSql = strSql & "  ShisanJCode" 	
		strSql = strSql & " ,Pn" 			
		strSql = strSql & " ,PName" 			
		strSql = strSql & " ,DModel" 		
		strSql = strSql & " ,DModelName"		
		strSql = strSql & " ,Hinmoku" 		
		strSql = strSql & " ,HinmokuName"	
		strSql = strSql & " ,NaiSupplyYm"	
		strSql = strSql & " ,NaiBldOutYm"	
		strSql = strSql & " ,NaiKbn" 		
		strSql = strSql & " ,GaiSupplyYm"	
		strSql = strSql & " ,GaiBldOutYm"	
		strSql = strSql & " ,GaiKbn" 		
		strSql = strSql & " ,Biko" 			
		strSql = strSql & " ) values ("
		strSql = strSql & strValues
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
'		Call objParent.CallSql(strSql)
	End Function
End Class

'�Ώ۔N��	�N��
'�o�ɐ�/���o��	91H
'�\��/����	�\��
'�w����t	2016�N05��01�� �`2016�N12��31��
'����/��	�����̂�
'�ߏ�	�S��
'����/���z	�S��
'���i�R�[�h	���i��	���B�N���x	���B��	�\����t	AMPM	����	���ѓ��t	�`�[��	����R�[�h	�d��/�x����	�[���ꏊ	�\�萔	���ѐ�	�o�ɐ�/���o��	�ʉ݃R�[�h	�d��/�x���P��	���z	���藝�R1	������	�ۊǏꏊ	�݌ɋ敪	�Ή��`�[��	���ރ��C��	�ʉ݃R�[�h	�����x���P��	���藝�R2	�������t	�����敪	���v�敪	��C�o�C���[	������i�R�[�h	���敪	�\��ԍ�	�\�񎞒��B�N��	�\�񎞒��B��	�\�񎞕ύX��	Daily�敪	�[�������N	���ЃR�[�h	�s���{���R�[�h	���Y���R�[�h	PDM�@��R�[�h	�F�敪	�d�l�敪	���`���j�b�g�敪	���[�U�[	���[�U�[��	�X�V���t
' �T�t�@�C�A�[���\��
Class SaDelv
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		dim	strTitle
		strTitle = "���i�R�[�h	���i��	���B�N���x	���B��	�\����t	AMPM	����	���ѓ��t	�`�[��	����R�[�h	�d��/�x����	�[���ꏊ	�\�萔	���ѐ�	�o�ɐ�/���o��	�ʉ݃R�[�h	�d��/�x���P��	���z	���藝�R1	������	�ۊǏꏊ	�݌ɋ敪	�Ή��`�[��	���ރ��C��	�ʉ݃R�[�h	�����x���P��	���藝�R2	�������t	�����敪	���v�敪	��C�o�C���[	������i�R�[�h	���敪	�\��ԍ�	�\�񎞒��B�N��	�\�񎞒��B��	�\�񎞕ύX��	Daily�敪	�[�������N	���ЃR�[�h	�s���{���R�[�h	���Y���R�[�h	PDM�@��R�[�h	�F�敪	�d�l�敪	���`���j�b�g�敪	���[�U�[	���[�U�[��	�X�V���t"
		dim	aryTitle
		aryTitle = Split(strTitle,vbTab)
		dim	objR
		for each objR in objSt.Range("A10:AW10")
			objParent.Debug ".Init():" & objR & ":" & objR.Column & ":" & aryTitle(objR.Column - 1)
			if CompTitle(objR,aryTitle(objR.Column - 1)) then
				exit function
			end if
		next
		' �f�[�^�Ǎ�
		Call Load()
		Init = true
	End Function
	Private	lngLimit
	Private	lngCount
	Public Function Limit()
		objParent.Debug ".Limit():" & lngCount & "/" & lngLimit
		Limit = false
		lngCount = lngCount + 1
		if lngLimit <> 0 then
			if lngCount >= lngLimit then
				Limit = true
			end if
		end if
	End Function
	Private	strFields
	Public Function Load()
		strFields = objParent.GetFields("SaDelv")
		strFields = Replace(strFields,",EntID,EntTm,UpdID,UpdTm","")
		lngRowTop = objParent.GetOption("s",11)
		lngRowEnd = excelGetMaxRow(objSt,"A",11)
		lngCount  = 0
		lngLimit  = CLng(objParent.GetOption("l",0))
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
			if Limit() then
				exit for
			end if
		next
	End Function
	Private	aryCell(255)
	Public Function Disp()
		dim	i
		dim	strMsg
		strMsg = lngRow & "/" & lngRowEnd
		dim	strValues
		strValues = ""
		dim	objR
		for each objR in objSt.Range("A" & lngRow & ":AW" & lngRow)
			strMsg = strMsg & " " & objR
			if strValues <> "" then
				strValues = strValues & ","
			end if
			strValues = strValues & "'" & objR & "'"
		next
		objParent.Disp	strMsg
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into SaDelv"
		strSql = strSql & " ("
		strSql = strSql & strFields
		strSql = strSql & " ) values ("
		strSql = strSql & strValues
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
'		Call objParent.CallSql(strSql)
	End Function
End Class

' Heg���i�\
Class HegPrice
	Private	lngRow
	Private	lngRowTop
	Private	lngRowEnd
	Private objSt
	Private	objParent
	Public Function Init(oParent)
		Init = false
		set objParent = oParent
		set objSt	  = oParent.objSt
		dim	strTitle
		strTitle = "�i��	�敪	�d���P��	�P������	�P���v�Z	����	�Ǘ����	�Ǘ���v�Z	����	�P��	�Ǘ���(5%)	�Ǘ����	�P��	�Ǘ���(5%)	�Ǘ����	�P��	�Ǘ���(5%)	�Ǘ����	�P��	�Ǘ���(5%)	�Ǘ����	�P��	�Ǘ���(5%)	�Ǘ����	�P��	�Ǘ���(5%)	�Ǘ����	�P��	�Ǘ���(5%)	�Ǘ����	�P��	�Ǘ���(5%)	�Ǘ����	111�G�R�L���[�g(���i���H��)	112�L���r�l�b�g(���i���H��)	11�R���̑�(���i���H��)	114�_�N�g���i��(���i���H��)	120�G�R�L���[�g(PF�ǉ��H)	160�G�R�L���[�g(PE�ǉ��H)	211�G�R�L���[�g(�̔�)	212���̑�(�̔�)	240�L���r�l�b�g/IH(�̔�)	251�G�R�L���[�g(������)	252�L���r�l�b�g/IH(������)	253���̑�(������)	254�_�N�g(������)	270���ރZ���^�[�_�N�g���ށi�̔��j	170���ރZ���^�[�_�N�g���H�i���i���H���j"

		dim	aryTitle
		aryTitle = Split(strTitle,vbTab)
		dim	objR
		for each objR in objSt.Range("A3:AV3")
			objParent.Debug ".Init():" & objR & ":" & objR.Column & ":" & aryTitle(objR.Column - 1)
			if CompTitle(objR,aryTitle(objR.Column - 1)) then
				exit function
			end if
		next
		' �f�[�^�Ǎ�
		Call Load()
		Init = true
	End Function
	Private	lngLimit
	Private	lngCount
	Public Function Limit()
		objParent.Debug ".Limit():" & lngCount & "/" & lngLimit
		Limit = false
		lngCount = lngCount + 1
		if lngLimit <> 0 then
			if lngCount >= lngLimit then
				Limit = true
			end if
		end if
	End Function
	Private	strFields
	Private	aryFields
	Private Function Message(byVal strMsg)
		Wscript.Echo "HegPrice:" & strMsg
	End Function
	Public Function Load()
		strFields = objParent.GetFields("HegPrice")
		strFields = Replace(strFields,",EntTm,UpdID,UpdTm","")
		aryFields = Split(strFields,",")
		lngRowTop = objParent.GetOption("s",4)
		lngRowEnd = excelGetMaxRow(objSt,"A",4)
		lngCount  = 0
		lngLimit  = CLng(objParent.GetOption("l",0))
		dim	strSql
		strSql = ""
		strSql = strSql & "delete from HegPrice"
		Call Message(strSql)
		Call objParent.CallSql(strSql)
		for lngRow = lngRowTop to lngRowEnd
			objParent.Debug ".Load():" & objSt.Name & " " & lngRow & "/" & lngRowEnd
			Call Disp()
			if Limit() then
				exit for
			end if
		next
	End Function
	Private	aryCell(255)
	Public Function Disp()
		dim	i
		dim	strMsg
		strMsg = lngRow & "/" & lngRowEnd & " " & objSt.Range("A" & lngRow) & " " & objSt.Range("B" & lngRow)
		Call Message(strMsg)
		dim	strValues
		strValues = ""
		i = 0
		strValues = strValues & " '" & objSt.Range("A" & lngRow) & "'"
		strValues = strValues & ",'" & objSt.Range("B" & lngRow) & "'"
		strValues = strValues & "," & CDbl(objSt.Range("L" & lngRow)) & ""	'//�̔�
		strValues = strValues & "," & CDbl(objSt.Range("M" & lngRow)) & ""	'//������
		strValues = strValues & "," & CDbl(objSt.Range("P" & lngRow)) & ""	'//�����Y��
		strValues = strValues & "," & CDbl(objSt.Range("S" & lngRow)) & ""	'//���i���H��
		strValues = strValues & "," & CDbl(objSt.Range("V" & lngRow)) & ""	'//PF�ǉ��H
		strValues = strValues & "," & CDbl(objSt.Range("Y" & lngRow)) & ""	'//PE�ǉ��H
		strValues = strValues & "," & CDbl(objSt.Range("AB" & lngRow)) & ""	'//PE�Ǖ�����
		strValues = strValues & "," & CDbl(objSt.Range("AE" & lngRow)) & ""	'//��Əꏊ����
		strValues = strValues & "," & CDbl(objSt.Range("AH" & lngRow)) & ""	'//a111 �G�R�L���[�g(���i���H��)
		strValues = strValues & "," & CDbl(objSt.Range("AI" & lngRow)) & ""	'//a112 �L���r�l�b�g(���i���H��)
		strValues = strValues & "," & CDbl(objSt.Range("AJ" & lngRow)) & ""	'//a113 ���̑�(���i���H��)
		strValues = strValues & "," & CDbl(objSt.Range("AK" & lngRow)) & ""	'//a114 �_�N�g���i��(���i���H��)
		strValues = strValues & "," & CDbl(objSt.Range("AL" & lngRow)) & ""	'//d120 �G�R�L���[�g(PF�ǉ��H)
		strValues = strValues & "," & CDbl(objSt.Range("AM" & lngRow)) & ""	'//d160 �G�R�L���[�g(PE�ǉ��H) 
		strValues = strValues & "," & CDbl(objSt.Range("AN" & lngRow)) & ""	'//b211 �G�R�L���[�g(�̔�)
		strValues = strValues & "," & CDbl(objSt.Range("AO" & lngRow)) & ""	'//b212 ���̑�(�̔�)
		strValues = strValues & "," & CDbl(objSt.Range("AP" & lngRow)) & ""	'//b240 �L���r�l�b�g/IH(�̔�)
		strValues = strValues & "," & CDbl(objSt.Range("AQ" & lngRow)) & ""	'//c251 �G�R�L���[�g(������)
		strValues = strValues & "," & CDbl(objSt.Range("AR" & lngRow)) & ""	'//c252 �L���r�l�b�g/IH(������)
		strValues = strValues & "," & CDbl(objSt.Range("AS" & lngRow)) & ""	'//c253 ���̑�(������)
		strValues = strValues & "," & CDbl(objSt.Range("AT" & lngRow)) & ""	'//c254 �_�N�g(������)
		strValues = strValues & "," & CDbl(objSt.Range("AU" & lngRow)) & ""	'//b270 ���ރZ���^�[�_�N�g���ށi�̔��j
		strValues = strValues & "," & CDbl(objSt.Range("AV" & lngRow)) & ""	'//a170 ���ރZ���^�[�_�N�g���H�i���i���H���j
		strValues = strValues & ",'HegCnv'"
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into HegPrice"
		strSql = strSql & " ("
		strSql = strSql & strFields
		strSql = strSql & " ) values ("
		strSql = strSql & strValues
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
	End Function
End Class

