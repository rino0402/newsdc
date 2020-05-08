Option Explicit
'-----------------------------------------------------------------------
'���C���ďo���C���N���[�h
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
WScript.Quit lngRet
'-----------------------------------------------------------------------
Function GetCD()
	Dim objWshShell
	'�@WScript.Shell�I�u�W�F�N�g�̍쐬
	Set objWshShell = CreateObject("WScript.Shell")
	'�J�����g�f�B���N�g����\��
	dim	strCD
	strCD = objWshShell.CurrentDirectory
	Set objWshShell = Nothing
	GetCD = strCD
End Function

Function GetAbsPath(byVal strPath)
	Dim objFileSys
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	strPath = objFileSys.GetAbsolutePathName(strPath)
	Set objFileSys = Nothing
	GetAbsPath = strPath
End Function

Function GetDate2(byVal v)
	dim	strDate
	strDate = ""
	if isDate(v) then
		strDate = year(v) & Right(00 & month(v), 2) & Right(00 & day(v), 2)
	end if
	GetDate2 = strDate
End Function

Function GetScriptPath()
	GetScriptPath = Left(WScript.ScriptFullName,Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
End Function

Function GetFileName(byVal strFullName)
	dim	strFileName
	strFileName = strFullName
	dim	c
	for each c in split(strFileName,"\")
		Call Debug("GetFileName():" & c)
		if c <> "" then
			strFileName = c
		end if
	next
	GetFileName = strFileName
End Function

Function GetTab(ByVal s)
    Dim r
	r = Split(s,vbTab)
	GetTab = r
End Function

Function GetTrim(byval c)
	if left(c,1) = """" then
		if right(c,1) = """" then
			c = Right(c,Len(c) -1 )
			c = Left(c,Len(c) -1 )
		end if
	end if
	GetTrim = c
End Function

'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "�퐶�̔����㖾�׃f�[�^"
	Wscript.Echo "YSales.vbs [option] <�t�@�C����>"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo "CurrentDirectory=" & GetCD()
	Wscript.Echo "GetAbsPath()=" & GetAbsPath("YSales.vbs")
	Wscript.Echo "WScript.Path=" & WScript.Path
	Wscript.Echo "WScript.ScriptFullName=" & WScript.ScriptFullName
	Wscript.Echo "WScript.ScriptName=" & WScript.ScriptName
	Wscript.Echo "GetScriptPath()=" & GetScriptPath()
	Wscript.Echo "GetFileName()=" & GetFileName(WScript.ScriptFullName)
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	'���O�����I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		else
			strFilename = ""
		end if
	next
	'���O�t���I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			strFilename = ""
		case else
			strFilename = ""
		end select
	next
	if strFilename = "" then
		usage()
		Main = 1
		exit Function
	end if
	call LoadYSales(strFilename)
	Main = 0
End Function

Function LoadYSales(byVal strFilename)
	Call Debug("LoadYSales(" & strFilename & ")")
	'-------------------------------------------------------------------
	'Excel�t�@�C����
	'-------------------------------------------------------------------
	if strFileName = "" then
		Call DispMsg("�t�@�C�������w�肵�ĉ�����")
		Exit Function
	end if
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	dim	objXL
	Set objXL = WScript.CreateObject("Excel.Application")
	Call Debug("CreateObject(Excel.Application)")
	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	dim	strPassword
	strPassword = ""
	dim	objBk
	Set objBk = objXL.Workbooks.Open(strFilename,False,True,,strPassword)
	Call Debug("Workbooks.Open=" & objBk.Name)
	'-------------------------------------------------------------------
	'�Ǎ�����
	'-------------------------------------------------------------------
	Call LoadYSalesXls(objXL,objBk)
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadYSales():End")
End Function

Function LoadYSalesXls(objXL,objBk)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'�e�[�u���I�[�v��
	'-------------------------------------------------------------------
	dim	rsYSales
	set rsYSales = OpenRs(objDb,"YSales")

	Call Debug("LoadYSalesXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadYSalesXls():SheetName=" & strShtName)
		Call LoadYSalesXst(objXL,objBk,objSt,objDb,rsYSales)
	Next
	
	'-------------------------------------------------------------------
	'�e�[�u���̃N���[�Y
	'-------------------------------------------------------------------
	set rsYSales = CloseRs(rsYSales)
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̃N���[�Y
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadYSalesXst(objXL,objBk,objSt,objDb,rsYSales)
	Call Debug("LoadYSalesXst():SheetName=" & objSt.Name)
	Dim		aryYM()
	ReDim	aryYM(0)
	dim	lngMaxRow
	lngMaxRow = objSt.rows.count
	dim	lngRow
	For lngRow = 6 to lngMaxRow
		if LoadYSalesRow(aryYM,objXL,objBk,objSt,objDb,rsYSales,lngRow) = 0 then
			Exit For
		end if
	Next
End Function

Function LoadYSalesRow(aryYM,objXL,objBk,objSt,objDb,rsYSales,byVal lngRow)
	Call DispMsg("LoadYSalesRow():" & objSt.Name & ":" & lngRow)
	dim	strB
	strB = objSt.Range("B" & lngRow)
	if strB = "" then
		LoadYSalesRow = 0
		exit function
	end if
	dim	strYM
	strYM = Left(Replace(strB,"/",""),6)
	if chkYM(aryYM,strYM) <> "" then
		Call DispMsg("delete SDt:" & strYM)
		Call ExecuteAdodb(objDb,"delete from YSales where Left(SDt,6) = '" & strYM & "'")
		ReDim Preserve aryYM(UBound(aryYM) + 1)
		aryYM(UBound(aryYM)) = strYM
	end if
	rsYSales.AddNew
	Call SetField(rsYSales,objSt,"SDt",		"B" , lngRow)	'		Char( 8) default '' not null	// �����
	Call SetField(rsYSales,objSt,"DenNo",	"C" , lngRow)	'		Char(12) default '' not null	// �`�[�ԍ�
	Call SetField(rsYSales,objSt,"TrKb",	"D" , lngRow)	'		Char(20) default '' not null	// ����敪
	Call SetField(rsYSales,objSt,"Shime",	"E" , lngRow)	'		Char(20) default '' not null	// ����
	Call SetField(rsYSales,objSt,"TkCd",	"F" , lngRow)	'		Char(10) default '' not null	// ���Ӑ�R�[�h
	Call SetField(rsYSales,objSt,"TkName",	"G" , lngRow)	'		Char(80) default '' not null	// ���Ӑ於
	Call SetField(rsYSales,objSt,"PaySft",	"H" , lngRow)	'		Char(20) default '' not null	// �œ]��
	Call SetField(rsYSales,objSt,"NnCd",	"I" , lngRow)	'		Char(10) default '' not null	// �[����R�[�h
	Call SetField(rsYSales,objSt,"NnName",	"J" , lngRow)	'		Char(80) default '' not null	// �[���於
	Call SetField(rsYSales,objSt,"TnCd",	"K" , lngRow)	'		Char(10) default '' not null	// �S���҃R�[�h
	Call SetField(rsYSales,objSt,"TnName",	"L" , lngRow)	'		Char(40) default '' not null	// �S���Җ�
	Call SetField(rsYSales,objSt,"Uchi",	"M" , lngRow)	'		Char(20) default '' not null	// ����
	Call SetField(rsYSales,objSt,"Syuka",	"N" , lngRow)	'		Char(20) default '' not null	// �o��
	Call SetField(rsYSales,objSt,"ShCd",	"O" , lngRow)	'		Char(20) default '' not null	// ���i�R�[�h
	Call SetField(rsYSales,objSt,"ShName",	"P" , lngRow)	'		Char(80) default '' not null	// ���i��/�E�v
	Call SetField(rsYSales,objSt,"Pcs",		"Q" , lngRow)	'		Char(10) default '' not null	// �P��
	Call SetField(rsYSales,objSt,"PerCase",	"R" , lngRow)	'	CURRENCY default  0 not null	// ����
	Call SetField(rsYSales,objSt,"QtyCase",	"S" , lngRow)	'	CURRENCY default  0 not null	// �P�[�X
	Call SetField(rsYSales,objSt,"Qty",		"T" , lngRow)	'		CURRENCY default  0 not null	// ����
	Call SetField(rsYSales,objSt,"GPrice",	"U" , lngRow)	'		CURRENCY default  0 not null	// ���P��
	Call SetField(rsYSales,objSt,"UPrice",	"V" , lngRow)	'		CURRENCY default  0 not null	// �P��
	Call SetField(rsYSales,objSt,"Gross",	"W" , lngRow)	'		CURRENCY default  0 not null	// �e���v
	Call SetField(rsYSales,objSt,"Amount",	"X" , lngRow)	'		CURRENCY default  0 not null	// ���z
	Call SetField(rsYSales,objSt,"PayKb",	"Y" , lngRow)	'		Char(20) default '' not null	// �ېŋ敪
	Call SetField(rsYSales,objSt,"Biko",	"Z" , lngRow)	'		Char(80) default '' not null	// ���l
	Call SetField(rsYSales,objSt,"MtNo",	"AA" , lngRow)	'		Char(20) default '' not null	// ���ϔԍ�
	Call SetField(rsYSales,objSt,"JcNo",	"AB" , lngRow)	'		Char(20) default '' not null	// �󒍔ԍ�
	rsYSales.UpdateBatch
	LoadYSalesRow = lngRow
End Function

Private Function chkYM(aryYM(),byVal strYM)
	dim	a
'	for each a in aryJCode
	dim	i
'	Call Debug("chkJCode:" & LBound(aryJCode) & " to " & UBound(aryJCode))
	for i = LBound(aryYM) to UBound(aryYM)
		a = aryYM(i)
'		Call Debug("chkJCode:" & a & "=" & strJCode)
		if a = strYM then
			strYM = ""
			exit for
		end if
	next
	chkYM = strYM
End Function

Function SetField(objRs,objSt,byVal strField,byVal strCol,byVal lngRow)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & objSt.Range(strCol & lngRow))
	dim	v
	v = RTrim(objSt.Range(strCol & lngRow))
	select case strField
	case "SDt"
		v = Replace(v,"/","")
	end select
	objRs.Fields(strField) = v
End Function

Function FindLotNo(objDb,rsLotNo,strModel,strPLotNo)
	dim	strSql
	strSql = "select * from LotNo"
	strSql = makeWhere(strSql,"Model"	,strModel	,"")
	strSql = makeWhere(strSql,"PLotNo"	,strPLotNo	,"")
	FindLotNo = UpdateOpenRs(objDb,rsLotNo,strSql)
End Function

