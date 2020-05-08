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
	Wscript.Echo "BoCnv.vbs [option]"
	Wscript.Echo " /db:newsdc1	�f�[�^�x�[�X"
	Wscript.Echo " /j:4			���ƕ�"
	Wscript.Echo " /s:10000		�J�n�s"
	Wscript.Echo " /l:100		�ǂݍ��ލs��"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc1 /j:4 bo\���уJ�e�S���[.xls"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc1 /j:5 bo\�i�ڃJ�e�S���[.xls"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc3 bo\�i�����X�g�⑫�f�[�^.xlsx"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc3 bo\�T�t�@�C�A.xlsx"
	Wscript.Echo "cscript BoCnv.vbs /db:newsdc4 00025800.xls"
End Sub

'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objBoCnv
	Set objBoCnv = New BoCnv
	if objBoCnv.Init() <> "" then
		call usage()
		exit function
	end if
	call objBoCnv.Run()
End Function

'-----------------------------------------------------------------------
'BoCnv
'-----------------------------------------------------------------------
Class BoCnv
	Private	strDBName
	Private	objDB
	Private	objRs
	Public	strJGYOBU
	Private	strAction
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
		strAction = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
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
	'CheckFunction
	'-----------------------------------------------------------------------
	Private Function CheckFunction(byval strA)
		Debug ".CheckFunction():" & strA
		CheckFunction = False
		if strAction = "" then
			exit function
		end if
		if WScript.Arguments.Named.Exists(strA) then
			exit function
		end if
		CheckFunction = True
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"newsdc")
		strJGYOBU = GetOption("j"	,"4")
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
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			Debug ".Run():" & strArg
			strFileName = strArg
			Call Load()
		Next
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
			Call LoadExcel()
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
		if objBk is nothing then
			exit function
		end if
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
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
		Set clsData = New BoHinmoku
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
		end if
		' �i�����X�g�⑫�f�[�^
		Set clsData = New BoHosoku
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
		end if
		' �T�t�@�C�A�[���\��
		Set clsData = New SaDelv
		LoadData = clsData.Init(me)
		Set clsData = Nothing
		if LoadData then
			exit function
		end if
		' Active�f�[�^
		Set clsData = New AcData
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

' Active�f�[�^
Class AcData
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
		strTitle = "ID-NO	���Ə�R�[�h	���Y�Ǘ����Ə�R�[�h	�̔��敪	�i�ڔԍ�	���Ӑ�R�[�h	���Ӑ旪��	���������R�[�h	�i���@�T�[�r�X�f�[�^�i���敪	�����@�݌Ɏ��x������	�o�ח\�萔	�O��[���񓚔N����	�w���S���[���񓚔N����	�w���S�����R�g�p��	�[���񓚃f�[�^���M�敪	�[���񓚃f�[�^���M��	�o�ח\��N����	�o�ɗ\��N����	�󒍔N����	�w��[�����@�w��[���N����	�o�׎w��N����	�[���񓚓��@�[���񓚔N����	���[�敪	���[�敪�ĕt�^	�I�[�_�[NO	ITEM-NO	�`�[�ԍ�	�T�[�r�X�̔����[�g�R�[�h	�����敪	�d���惏�[�N�Z���^�[�R�[�h	�w���S���@�w���S���҃R�[�h	�[���񓚃f�[�^�����敪	�[���񓚎����t�^�敪	�[���񓚎��R�g�p��	�o�^���[�U�ID	�o�^���t	�o�^����	�X�V���[�U�ID	�X�V���t	�X�V����"
		dim	aryTitle
		aryTitle = Split(strTitle,vbTab)
		dim	objR
		for each objR in objSt.Range("A1:AN1")
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
	Public Function Load()
		strFields = objParent.GetFields("A_Data")
		strFields = Replace(strFields,",EntID,EntTm,UpdID,UpdTm","")
		aryFields = Split(strFields,",")
		lngRowTop = objParent.GetOption("s",3)
		lngRowEnd = excelGetMaxRow(objSt,"A",3)
		lngCount  = 0
		lngLimit  = CLng(objParent.GetOption("l",0))
		dim	strSql
		strSql = ""
		strSql = strSql & "delete from A_Data"
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
		strMsg = lngRow & "/" & lngRowEnd
		dim	strValues
		strValues = ""
		dim	objR
		i = 0
		for each objR in objSt.Range("A" & lngRow & ":AN" & lngRow)
			objParent.Debug ".Load():" & aryFields(i) & ":" & objR
			strMsg = strMsg & " " & objR
			if strValues <> "" then
				strValues = strValues & ","
			end if
			dim	strV
			strV = objR
			strV = Replace(strV,"'","")
			if Right(aryFields(i),2) = "Dt" then
				strV = Replace(strV,"/","")
			end if
			dim	strQ
			strQ = "'"
			if Right(aryFields(i),3) = "Qty" then
				strQ = ""
			end if
			strValues = strValues & strQ & strV & strQ
			i = i + 1
		next
		objParent.Disp	strMsg
		dim	strSql
		strSql = ""
		strSql = strSql & "insert into A_Data"
		strSql = strSql & " ("
		strSql = strSql & strFields
		strSql = strSql & " ) values ("
		strSql = strSql & strValues
		strSql = strSql & " )"
		Call objParent.CallSql(strSql)
		strSql = ""
	End Function
End Class

