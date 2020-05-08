Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "makexls.vbs [option]"
	Wscript.Echo " /db:newsdc	:�f�[�^�x�[�X"
	Wscript.Echo " /a:10		:�ǉ��pdummy����(default:0)"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript makexls.vbs /db:newsdc7 l164157.csv /a:10"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objJcsJoin
	Set objJcsJoin = New JcsJoin
	if objJcsJoin.Init() <> "" then
		call usage()
		exit function
	end if
	call objJcsJoin.Run()
End Function
'-----------------------------------------------------------------------
'JcsJoin
'-----------------------------------------------------------------------
Const xlEdgeTop		=	8
Const xlContinuous	=	1
Const xlThin		=	2
Class JcsJoin
	Private	strDBName
	Private	objDB
	Private	objRs
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
	Private strScriptPath
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
		Next
		if WScript.Arguments.UnNamed.Count = 0 then
			Init = "�t�@�C�����w��"
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "a"
			case "debug"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Disp Init
				Exit Function
			end select
		Next
		strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	End Function
	'-----------------------------------------------------------------------
	'CheckFunction
	'-----------------------------------------------------------------------
	Private Function CheckFunction(byval strA)
		Debug ".CheckFunction():" & strA
		CheckFunction = False
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
		set objDB = nothing
		set objRs = nothing
		set	objExcel = nothing
		set	objBook = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objRs = nothing
		set objDB = nothing
		set	objBook = nothing
		set	objExcel = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Call OpenDB()
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			strFileName = strArg
			Call Load()
		Next
		Call CloseDb()
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
	'Load() �Ǎ�
	'-----------------------------------------------------------------------
	Private	strBookName
    Public Function Load()
		Debug ".Load():" & strFileName
		strBookName = GetBaseName(strFileName) & ".xls"
		Debug ".Load():" & strBookName
		Call CreateExcel()
		Call OpenBook("���惊�X�g.xls")
		Call MakeBook()
		Call SaveBook("csv\" & strBookName)
		Call CloseBook()
	End Function
	'-----------------------------------------------------------------------
	'���惊�X�g�쐬
	'-----------------------------------------------------------------------
	Private	objSheet
    Public Function MakeBook()
		Debug ".MakeBook():" & objBook.Name
		set objSheet = objBook.ActiveSheet
		objSheet.Name = "���惊�X�g " & GetBaseName(strFilename)
		Debug ".MakeBook():" & objSheet.Name
		'�s�폜
		objSheet.Range("2:65536").Delete
		'�⍇��
		Call OpenRs()
		'���R�[�h�Z�b�g
		Call SetData()
		'�⍇��Close
		Call CloseRs()
		'�t�B���^�[ ����<>0
'		Call objSheet.Range("K1").AutoFilter(1, "<>0")
		'�t�B���^�[��̏����Z�b�g
'		prvPn = ""
'		curPn = ""
'		lngRow = 2
'		do while true
'			curPn = objSheet.Range("G" & lngRow)
'			if curPn = "" then
'				exit do
'			end if
'			if objSheet.Rows(lngRow).hidden = false then
'				Call FormatRow()
'				prvPn = curPn
'			end if
'			lngRow = lngRow + 1
'		loop
	End Function
	'-----------------------------------------------------------------------
	'���R�[�h�Z�b�g
	'-----------------------------------------------------------------------
	Private	curPn
	Private	prvPn
	Private	lngRow
    Public Function SetData()
		Debug ".SetData()"
		if objRs is nothing then
			exit function
		end if
		prvPn = ""
		curPn = ""
		lngRow = 2
		do while objRs.Eof = false
			Call SetDataRow()
			curPn = GetField("Pn")
			Call FormatRow()
			Call YNyuka()
			prvPn = curPn
			lngRow = lngRow + 1
			objRs.Movenext
		loop
		curPn = ""
		'�ǉ��pDummy
		dim	lngDummy
		lngDummy = GetOption("a",0) - 1
		dim	i
		for i=0 to lngDummy
			Call SetDummy(i)
			Call FormatRow()
			Call YNyukaDummy(i)
			lngRow = lngRow + 1
		next
		Call FormatRow()
	End Function
	'-----------------------------------------------------------------------
	'�ǉ��p�_�~�[�s�Z�b�g
	'-----------------------------------------------------------------------
    Public Function SetDummy(byVal i)
		Debug ".SetDummy():" & lngRow & ":" & i
		i = i + 1
		dim	strID
		strID = "J"
		strID = strID & strSyukaYmd
		strID = strID & Right(GetBaseName(strFilename),6)
		strID = strID & "A"
		strID = strID & Right("00" & i,2)
		objSheet.Range("Q" & lngRow) = "*" & strID & "*"			 '���ID
	End Function
	'-----------------------------------------------------------------------
	'YNyukaInsert
	'-----------------------------------------------------------------------
	Private	strIdNo
    Public Function YNyukaDummy(byVal i)
		Debug ".YNyukaDummy()"
		i = i + 1
		strTextNo = Right(GetBaseName(strFilename),6)
		strTextNo = strTextNo & "A"
		strTextNo = strTextNo & Right("00" & i,2)
		strIdNo = UCase(GetBaseName(strFilename))
		strIdNo = strIdNo & "A"
		strIdNo = strIdNo & Right("0000" & i,4)
		strSql = "insert into Y_NYUKA"
		strSql = strSql & " ("
		strSql = strSql & " DT_SYU"
		strSql = strSql & ",JGYOBU"
		strSql = strSql & ",NAIGAI"
		strSql = strSql & ",TEXT_NO"
		strSql = strSql & ",ID_NO"
		strSql = strSql & ",ID_NO2"
		strSql = strSql & ",SYUKO_YMD"
		strSql = strSql & ",SYUKA_YMD"
		strSql = strSql & ",MAEGARI_SURYO"
		strSql = strSql & ",INS_TANTO"
		strSql = strSql & " ) values ( "
		strSql = strSql & " '1'"			'DT_SYU"
		strSql = strSql & ",'J'"       		'JGYOBU"
		strSql = strSql & ",'1'"        	'NAIGAI"
		strSql = strSql & ",'" & strTextNo & "'"
		strSql = strSql & ",'" & strIdNo & "'"
		strSql = strSql & ",'" & strIdNo & "'"			'ID_NO2
		strSql = strSql & ",'" & strSyukaYmd & "'"
		strSql = strSql & ",'" & strSyukaYmd & "'"
		strSql = strSql & ",'00000000'"		'MAEGARI_SURYO"
		strSql = strSql & ",'makex'"		'INS_TANTO"
		strSql = strSql & " )"
		Debug ".YNyukaDummy():" & strSql
		on error resume next
			Call objDb.Execute(strSql)
			Debug ".YNyukaDummy():0x" & Hex(Err.Number) & ":" & Err.Description
		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'YNyuka
	'-----------------------------------------------------------------------
	Private	lngQty
	Private	strSyukaYmd
	Private	strTextNo
    Public Function YNyuka()
		Debug ".YNyuka()"
		if objRs is nothing then
			exit function
		end if
		dim	lngRcptQty
		lngRcptQty = CLng(GetField("RcptQty"))
		if curPn = prvPn then
			lngQty = lngQty + lngRcptQty
			Call YNyukaUpdate()
		else
			strSyukaYmd	= GetField(".SYUKA_YMD")
			strTextNo	= GetField(".TEXT_NO")
			lngQty = lngRcptQty
			Call YNyukaInsert()
		end if
	End Function
	'-----------------------------------------------------------------------
	'YNyukaUpdate
	'-----------------------------------------------------------------------
    Public Function YNyukaUpdate()
		Debug ".YNyukaUpdate()"
		strSql = "update Y_NYUKA"
		strSql = strSql & " set SURYO='" & lngQty & "'"
		strSql = strSql & " ,UPD_TANTO='makex'"
		strSql = strSql & " where JGYOBU='J'"
		strSql = strSql & "   and SYUKA_YMD='" & strSyukaYmd &  "'"
		strSql = strSql & "   and TEXT_NO='" & strTextNo &  "'"
		Debug ".YNyukaUpdate():" & strSql
'		on error resume next
			Call objDb.Execute(strSql)
			Debug ".YNyukaUpdate():0x" & Hex(Err.Number) & ":" & Err.Description
'		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'YNyukaInsert
	'-----------------------------------------------------------------------
    Public Function YNyukaInsert()
		Debug ".YNyukaInsert()"
		strSql = "insert into Y_NYUKA"
		strSql = strSql & " ("
		strSql = strSql & " DT_SYU"
		strSql = strSql & ",JGYOBU"
		strSql = strSql & ",NAIGAI"
		strSql = strSql & ",TEXT_NO"
		strSql = strSql & ",ID_NO"
		strSql = strSql & ",ID_NO2"
		strSql = strSql & ",HIN_NO"
		strSql = strSql & ",HIN_NAI"
		strSql = strSql & ",DEN_NO"
		strSql = strSql & ",SURYO"
		strSql = strSql & ",MUKE_CODE"
		strSql = strSql & ",SYUKO_YMD"
		strSql = strSql & ",SYUKA_YMD"
		strSql = strSql & ",HIN_NAME"
		strSql = strSql & ",NOUKI_YMD"
		strSql = strSql & ",SHIIRE_WORK_CENTER"
		strSql = strSql & ",MAEGARI_SURYO"
		strSql = strSql & ",INS_TANTO"
		strSql = strSql & " ) values ( "
		strSql = strSql & " '1'"			'DT_SYU"
		strSql = strSql & ",'J'"       		'JGYOBU"
		strSql = strSql & ",'1'"        	'NAIGAI"
		strSql = strSql & ",'" & GetField(".TEXT_NO") & "'"
		strSql = strSql & ",'" & GetField(".ID_NO") & "'"
		strSql = strSql & ",'" & GetField(".ID_NO") & "'"			'ID_NO2
		strSql = strSql & ",'" & GetField(".HIN_NO") & "'"
		strSql = strSql & ",'" & GetField(".HIN_NAI") & "'"
		strSql = strSql & ",'" & GetField(".DEN_NO") & "'"
		strSql = strSql & ",'" & lngQty & "'"
		strSql = strSql & ",'" & GetField(".MUKE_CODE") & "'"
		strSql = strSql & ",'" & GetField(".SYUKO_YMD") & "'"
		strSql = strSql & ",'" & GetField(".SYUKA_YMD") & "'"
		strSql = strSql & ",'" & GetField(".HIN_NAME") & "'"
		strSql = strSql & ",'" & GetField(".NOUKI_YMD") & "'"
		strSql = strSql & ",'" & GetField(".SHIIRE_WORK_CENTER") & "'"
		strSql = strSql & ",'00000000'"		'MAEGARI_SURYO"
		strSql = strSql & ",'makex'"		'INS_TANTO"
		strSql = strSql & " )"
		Debug ".YNyukaInsert():" & strSql
		on error resume next
			Call objDb.Execute(strSql)
			Debug ".YNyukaInsert():0x" & Hex(Err.Number) & ":" & Err.Description
		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'�s����
	'-----------------------------------------------------------------------
    Public Function FormatRow()
		objSheet.Range("A" & lngRow).RowHeight = 60
		if curPn = prvPn then
			objSheet.Range("Q" & lngRow) = ""
			exit function
		end if
		if curPn <> "" then
			if objSheet.Range("Q" & lngRow) = "" then
				dim	i
				i = 1
				do while objSheet.Range("Q" & lngRow) = ""
					objSheet.Range("Q" & lngRow) = objSheet.Range("Q" & lngRow - i)
					i = i - 1
				loop
			end if
		end if
	    objSheet.Range("A" & lngRow & ":Q" & lngRow).Borders(xlEdgeTop).LineStyle = xlContinuous
	    objSheet.Range("A" & lngRow & ":Q" & lngRow).Borders(xlEdgeTop).Weight = xlThin
	End Function
	'-----------------------------------------------------------------------
	'���R�[�h�Z�b�g�s
	'-----------------------------------------------------------------------
    Public Function SetDataRow()
		Debug ".SetData():" & lngRow
		if objRs is nothing then
			exit function
		end if
		objSheet.Range("A" & lngRow) = GetField("CstDlvNo")	'�ڋq�[���w���ԍ�
		objSheet.Range("B" & lngRow) = GetField("CstPn")	'�ڋq�i��
		objSheet.Range("C" & lngRow) = GetField("CstDlvDt")	'�ڋq�[��
		objSheet.Range("D" & lngRow) = GetField("CstQty")	'�ڋq�[������
		objSheet.Range("E" & lngRow) = GetField("OrderNo")	'�����ԍ�
		objSheet.Range("F" & lngRow) = GetField("TestKb")	'����
		objSheet.Range("G" & lngRow) = GetField("Pn")		'�i��
		objSheet.Range("H" & lngRow) = GetField("PName")	'����
		objSheet.Range("I" & lngRow) = GetField("DlvDt")	'�[����
		objSheet.Range("J" & lngRow) = GetField("DlvTm")	'����
		objSheet.Range("K" & lngRow) = GetField("RcptQty")	'����
		objSheet.Range("L" & lngRow) = GetField("CenterCd")	'���_
		objSheet.Range("M" & lngRow) = GetField("Location")	'�[���ꏊ
		objSheet.Range("N" & lngRow) = GetField("ClientNo")	'�����
		objSheet.Range("O" & lngRow) = GetField("DlvMdfDt")	'�[���ύX��
		objSheet.Range("P" & lngRow) = GetField("SdcStkQty")	'SDC�݌ɐ���
		objSheet.Range("Q" & lngRow) = "*" & GetField(".ID") & "*"			 '���ID
'		objSheet.Range("R" & lngRow) = GetField("SdcStkQty")	'�o�[�R�[�h
	End Function
	'-----------------------------------------------------------------------
	'Fields �l
	'-----------------------------------------------------------------------
    Public Function GetField(byVal strFldNm)
		Debug ".GetField():" & strFldNm
		if objRs is nothing then
			exit function
		end if
		if left(strFldNm,1) <> "." then
			GetField = RTrim(objRs.Fields(strFldNm))
		else
			GetField = "."
		end if
		if GetField <> "" then
			select case strFldNm
			case "CstDlvDt"	'�ڋq�[��	03/16 D
				GetField = Right(GetField,4)
				GetField = Left(GetField,2) & "/" & Right(GetField,2)
				GetField = GetField & " " & RTrim(objRs.Fields("CstDlvSft"))
			case "DlvDt"	'�[����	03/16 D
				GetField = Right(GetField,4)
				GetField = Left(GetField,2) & "/" & Right(GetField,2)
				GetField = GetField & " " & RTrim(objRs.Fields("DlvSft"))
			case "DlvTm"	'����	17:30
				GetField = Left(GetField,4)
				GetField = Left(GetField,2) & ":" & Right(GetField,2)
			case "DlvMdfDt"	'�[���ύX��	03/16
					GetField = Right(GetField,4)
					GetField = Left(GetField,2) & "/" & Right(GetField,2)
							' 123456789							
			case ".TEXT_NO"	'L151121
				dim	strTextNo
				strTextNo = Right(GetBaseName(strFilename),6)
				strTextNo = strTextNo & Right("000" & GetField("Row"),3)
				GetField = strTextNo
			case ".ID_NO"
				dim	strIdNo
				strIdNo = UCase(GetBaseName(strFilename))
				strIdNo = strIdNo & Right("00000" & GetField("Row"),5)
				GetField = strIdNo 
			case ".ID"
				dim	strId
				strId = "J"
				strId = strId & GetField(".SYUKA_YMD")
				strId = strId & GetField(".TEXT_NO")
				GetField = strId 
			case ".HIN_NO"
				GetField = RTrim(objRs.Fields("Pn"))
			case ".HIN_NAI"
				GetField = RTrim(objRs.Fields("CstPn"))
			case ".DEN_NO"
				GetField = RTrim(objRs.Fields("OrderNo"))
			case ".SURYO"
				GetField = RTrim(objRs.Fields("RcptQty"))
			case ".MUKE_CODE"
				GetField = RTrim(objRs.Fields("Location"))
			case ".SYUKO_YMD",".SYUKA_YMD"
				GetField = RTrim(objRs.Fields("DlvMdfDt"))
				if GetField = "" then
					GetField = RTrim(objRs.Fields("DlvDt"))
				end if
			case ".HIN_NAME"
				GetField = RTrim(objRs.Fields("PName"))
			case ".NOUKI_YMD"
				GetField = RTrim(objRs.Fields("DlvDt"))
			case ".SHIIRE_WORK_CENTER"
				GetField = RTrim(objRs.Fields("ClientNo"))
			end select
		end if
	End Function
	'-----------------------------------------------------------------------
	'�⍇��
	'-----------------------------------------------------------------------
	Private strSql
    Public Function OpenRs()
		Debug ".OpenRs()"
		if objDb is nothing then
			exit function
		end if
		strSql = "select"
		strSql = strSql & " *"
		strSql = strSql & " from JcsTakeOn"
		strSql = strSql & " where Filename = '" & strFilename & "'"
		strSql = strSql & "   and RcptQty <> 0"
		strSql = strSql & " order by"
		strSql = strSql & "  Pn"
		strSql = strSql & " ,Row"
		set objRs = objDb.Execute(strSql)
	End Function
	'-----------------------------------------------------------------------
	'�⍇��Close
	'-----------------------------------------------------------------------
    Public Function CloseRs()
		Debug ".CloseRs()"
		if objRs is nothing then
			exit function
		end if
		Call objRs.Close()
		set objRs = nothing
	End Function
	'-------------------------------------------------------------------
	'�t�@�C����(�p�X�A�g���q ����)
	'-------------------------------------------------------------------
	Private Function GetBaseName(byVal f)
		dim	fobj
		set fobj = CreateObject("Scripting.FileSystemObject")
		dim	strBaseName
		strBaseName = fobj.GetBaseName(f)
		set fobj = Nothing
		GetBaseName = strBaseName
	End Function
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	Private	objExcel
	Private Function CreateExcel()
		Debug(".CreateExcel()")
		if objExcel is nothing then
			Debug(".CreateExcel():CreateObject(Excel.Application)")
			Set objExcel = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	Private	objBook
	Private Function OpenBook(byVal strBkName)
		Debug(".OpenBook()")
		if objBook is nothing then
			strBkName = strScriptPath & strBkName
			Debug(".OpenBook().Open:" & strBkName)
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
		end if
	end function
	'-------------------------------------------------------------------
	'Excel �t�@�C�����O��t���ĕۑ�
	'-------------------------------------------------------------------
	Private Function SaveBook(byVal strBkName)
		Debug(".SaveBook()")
		if not objBook is nothing then
			strBkName = strScriptPath & strBkName
			Debug(".SaveBook().Save:" & strBkName)
			objExcel.DisplayAlerts = False
			Call objBook.SaveAs(strBkName)
			objExcel.DisplayAlerts = True
		end if
	end function
	'-------------------------------------------------------------------
	'Excel �t�@�C���N���[�Y
	'-------------------------------------------------------------------
	Private Function CloseBook()
		Debug(".CloseBook()")
		if not objBook is nothing then
			Debug(".CloseBook().Close:" & objBook.Name)
			Call objBook.Close(False)
			set objBook = nothing
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
End Class
