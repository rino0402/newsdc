Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Call Main()
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objExcel
	Set objExcel = New Excel
	objExcel.Run
	Set objExcel = Nothing
End Function
'-----------------------------------------------------------------------
'Excel
'2017.05.19 �V�K
'-----------------------------------------------------------------------
Const xlUp = -4162

Class Excel
	'-----------------------------------------------------------------------
	'�g�p���@
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "RfYotei.vbs [option] <*.xlsx>"
		Echo "/db:newsdc4"
		Echo "/debug"
		Echo "Ex."
		Echo "cscript//nologo RfYotei.vbs �①�ɏ��i���c��20170519.xlsx /db:newsdc4"
		Echo ""
		Echo "  :" & StrDate("")
		Echo "  :" & StrDate(Date())
		Echo "-1:" & StrDate(WorkDay(Date(),-1))
		Echo " 0:" & StrDate(WorkDay(Date(), 0))
		Echo " 1:" & StrDate(WorkDay(Date(), 1))
		Echo "14:" & StrDate(WorkDay(Date()+14,0))
		Echo "xx:" & StrDate(WorkDay("2017/05/11",1))
	End Sub
	'-----------------------------------------------------------------------
	'Private �ϐ�
	'-----------------------------------------------------------------------
	Private	strDBName
	Private	objDB
	Private	strFileName
	Private	strBookName
	Private	objExcel
	Private	objBook
	Private	objSheet
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName		= GetOption("db","newsdc")
		set objDB		= nothing
		set	objExcel	= nothing
		set	objBook		= nothing
		set	objSheet	= nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB		= nothing
		set	objSheet	= nothing
		set	objBook		= nothing
		set	objExcel	= nothing
    End Sub
	'-----------------------------------------------------------------------
	'Quit() �����I��
	'-----------------------------------------------------------------------
	Private Function Quit()
		Debug ".Quit()"
		Wscript.Quit
	End Function
	'-----------------------------------------------------------------------
	'Echo()
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		Wscript.Echo s
	End Sub
	'-----------------------------------------------------------------------
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
	Private Function Init()
		Debug ".Init()"
		dim	strArg
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Echo "Error:�I�v�V����:" & strArg
				Disp Init
				Usage
				Quit
			end if
		Next
		if strFileName = "" then
			Echo "Error:�t�@�C�����w��."
			Usage
			Quit
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case else
				Echo "Error:�I�v�V����:" & strArg
				Usage
				Quit
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		Init
		OpenDb
		Load
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Load() �Ǎ�
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load():" & strFileName
		CreateExcel
		OpenBook strFileName
		LoadBook
		CloseBook
	End Function
	'-------------------------------------------------------------------
	'LoadBook()
	'-------------------------------------------------------------------
	Private	lngTopRow
	Private	intCol
	Private	strMaxCol
    Private Function LoadBook()
		Debug ".LoadBook()"
		for each objSheet in objBook.Worksheets
			Write objSheet.Name & ":"
			select case SheetType()
			case "�[���w�聕��Ɨ\��"
				WriteLine "Load"
				LoadSheet
			case else
				WriteLine "skip"
			end select
		next
    End Function
	'-------------------------------------------------------------------
	'SheetType()
	'-------------------------------------------------------------------
	Private	Function SheetType()
		SheetType = ""
		if Trim(objSheet.Name) = "�[���w�聕��Ɨ\��" then
			SheetType = Trim(objSheet.Name)
			exit function
		end if
	End Function
	'-------------------------------------------------------------------
	'LoadSheet()
	'-------------------------------------------------------------------
	Private	lngMaxRow
	Private	lngRow
    Private Function LoadSheet()
		Debug ".LoadSheet():" & objBook.Name & ":" & objSheet.Name
		lngMaxRow = objSheet.Range("D65535").End(xlUp).Row
		Debug ".LoadSheet():MaxRow=" & lngMaxRow
		CallSql "delete from PLN_S_YOTEI"
		for lngRow = 4 to lngMaxRow
			Write objSheet.Name & ":" & lngRow & "/" & lngMaxRow
			LoadLine
			WriteLine ""
		next
    End Function
	'-------------------------------------------------------------------
	'LoadLine()
	'-------------------------------------------------------------------
    Private Function LoadLine()
		Debug ".LoadLine()"
		dim	strPn
		strPn = GetValue(objSheet.Range("D" & lngRow))
		Write " " & strPn
		Insert
    End Function
	'-------------------------------------------------------------------
	'Insert()
	'-------------------------------------------------------------------
    Private Function Insert()
		Debug ".Insert()"
		AddSql ""
		AddSql "insert into PLN_S_YOTEI"
		AddSql "(TORIKOMI_DT"
		AddSql ",JGYOBU"
		AddSql ",NAIGAI"
		AddSql ",HIN_GAI"
		AddSql ",YOTEI_DT"
		AddSql ",YOTEI_QTY"
		AddSql ",S_KOUSU"
		AddSql ",S_JIKAN"
		AddSql ",S_LIST_DateTime"
		AddSql ",SASIZU_DateTime"
		AddSql ",S_KAN_DateTime"
		AddSql ",TENKAI_DateTime"
		AddSql ",TOTAL_CNT"
		AddSql ",TOTAL_AVE_CNT"
		AddSql ",S_SYUKA_QTY1"
		AddSql ",S_SYUKA_CNT1"
		AddSql ",S_AVE_SYUKA_QTY1"
		AddSql ",S_AVE_SYUKA_CNT1"
		AddSql ",S_SYUKA_QTY2"
		AddSql ",S_SYUKA_CNT2"
		AddSql ",S_AVE_SYUKA_QTY2"
		AddSql ",S_AVE_SYUKA_CNT2"
		AddSql ",Z_QTY_MI"
		AddSql ",Z_QTY_S"
		AddSql ",JIZEN"
		AddSql ",NYUKA_YOTEI_DT"
		AddSql ",NYUKA_YOTEI_QTY"
		AddSql ",S_KOUSU_X"
		AddSql ",S_JIKAN_X"
		AddSql ",YOTEI_DT_X"
		AddSql ",YOTEI_QTY_X"
		AddSql ",SIZAI"
		AddSql ",GAISO_HINBAN"
		AddSql ",GAISO_MAISU"
		AddSql ",ST_SOKO"
		AddSql ",ST_RETU"
		AddSql ",ST_REN"
		AddSql ",ST_DAN"
		AddSql ",BETU1_SOKO"
		AddSql ",BETU1_RETU"
		AddSql ",BETU1_REN"
		AddSql ",BETU1_DAN"
		AddSql ",BETU1_QTY"
		AddSql ",BETU2_SOKO"
		AddSql ",BETU2_RETU"
		AddSql ",BETU2_REN"
		AddSql ",BETU2_DAN"
		AddSql ",BETU2_QTY"
		AddSql ",JIZEN_NEEDS_QTY"
		AddSql ",JITU_KOUSU"
		AddSql ",SAGYOU_KOUSU"
		AddSql ",NAI_BUHIN"
		AddSql ",GAI_BUHIN"
		AddSql ",TEHAISAKI"
		AddSql ",KEY_NO"
		AddSql ",INP_NYUKA_YOTEI_DT"
		AddSql ",INP_NYUKA_YOTEI_QTY"
		AddSql ",Y_NYUKA_KEY_NO"
		AddSql ",FILLER"
		AddSql ",INS_TANTO"
		AddSql ",Ins_DateTime"
		AddSql ",UPD_TANTO"
		AddSql ",UPD_DATETIME"
		AddSql ") values ("
		AddSql " replace(convert(curdate(),sql_char),'-','')"	'// �捞�ݓ��t
		AddSql ",'R'"	'// ���ƕ��敪
		AddSql ",'1'"	'// �����O
		AddSql ",'" & GetValue(objSheet.Range("D" & lngRow)) & "'"	'// �i�ԁi�O���j

		dim	dtKan	'������(�Z���^�[�q��)
		dtKan = GetValue(objSheet.Range("G" & lngRow))
		dim	dtKan2	'������(BD:tmp)
		dtKan2 = GetValue(objSheet.Range("J" & lngRow))
		if dtKan < dtKan2 then
			dtKan = dtKan2
		end if

		dim	dtNyuka	'���ח\���
		dtNyuka = GetValue(objSheet.Range("M" & lngRow))

		dim	dtYotei		'���i���\���
		dtYotei = GetValue(objSheet.Range("A" & lngRow))
		if isDate(dtYotei) = False then
			if isDate(dtKan) = True then
				'������
				dtYotei = dtKan
			else'������
				if isDate(dtNyuka) = True then
					'���׍� ������\��
					dtYotei = WorkDay(dtNyuka,1)
				else'������	14��������\��
					dtYotei = WorkDay(Date()+14,0)
				end if
				if dtYotei < Date() then
					dtYotei = WorkDay(Date(),1)
				end if
			end if
		end if
		AddSql ",'" & StrDate(dtYotei) & "'"	'// ���i���\����t
		dim	strQty
		strQty = CCur(GetValue(objSheet.Range("F" & lngRow))) + CCur(GetValue(objSheet.Range("I" & lngRow)))
		AddSql ",'" & strQty & "'"	'// ���i���\�萔
		AddSql ",'0'"	'// ���i���@�W���H��
		AddSql ",'0'"	'// ���i���@�W������ YOTEI_QTY �~ S_KOUSU
		AddSql ",''"	'// ���i���\�胊�X�g�������
		AddSql ",'" & StrDateTm(dtKan) & "'"	'// ���i���w�}�[�������
		AddSql ",'" & StrDateTm(dtKan) & "'"	'// ���i�������o�^����
		AddSql ",''"	'// ���v�ʓW�J����
		AddSql ",'0'"	'// ���o�׌���
		AddSql ",'0'"	'// ���ϑ��o�׌���
		AddSql ",'0'"	'// ���Y�v��o�א�(1)
		AddSql ",'0'"	'// ���Y�v��o�׌���(1)
		AddSql ",'0'"	'// ���ϐ��Y�v��o�א�(1)
		AddSql ",'0'"	'// ���ϐ��Y�v��o�׌���(1)
		AddSql ",'0'"	'// ���Y�v��o�א�(2)
		AddSql ",'0'"	'// ���Y�v��o�׌���(2)
		AddSql ",'0'"	'// ���ϐ��Y�v��o�א�(2)
		AddSql ",'0'"	'// ���ϐ��Y�v��o�׌���(2)
		AddSql ",'0'"	'// �݌Ɂi�����i�j
		AddSql ",'0'"	'// �݌Ɂi���i���ς݁j
		AddSql ",'0'"	'// ���O���i����
		AddSql ",'" & StrDate(dtNyuka) & "'"	'// ���i���p���i���ח\���
		AddSql ",'" & strQty & "'"	'// ���i���p���i���ח\�萔
		AddSql ",'0'"	'// ���ύH��
		AddSql ",'0'"	'// ���i���@�W������ YOTEI_QTY �~ S_KOUSU
		AddSql ",''"	'// ���i���\����t
		AddSql ",'0'"	'// ���i���\�萔
		AddSql ",''"	'// ���ށi�����j
		AddSql ",''"	'// �O���i��
		AddSql ",'0'"	'// �O���g�p����
		AddSql ",''"	'// �W�����ɑq�� �q��
		AddSql ",''"	'//  ��
		AddSql ",''"	'//  �A
		AddSql ",''"	'//  �i
		AddSql ",''"	'// �ʒu�P �q��
		AddSql ",''"	'//  ��
		AddSql ",''"	'//  �A
		AddSql ",''"	'//  �i
		AddSql ",''"	'// �ʒu���ʇ@
		AddSql ",''"	'// �ʒu�Q �q��
		AddSql ",''"	'//  ��
		AddSql ",''"	'//  �A
		AddSql ",''"	'//  �i
		AddSql ",'0'"	'// �ʒu���ʇ@
		AddSql ",'0'"	'// ���O���i���K�v��
		AddSql ",'0'"	'// ���эH��
		AddSql ",'0'"	'// ��ƍH��
		AddSql ",''"	'// �����������i�敪
		AddSql ",''"	'// �C�O�������i�敪
		AddSql ",''"	'// ���i��������z��
		AddSql ",'" & lngRow & "'"	'KEY_NO
		AddSql ",'" & StrDate(dtNyuka) & "'"	'// ���i���p���i���ח\���(����)
		AddSql ",'" & strQty & "'"	'// ���i���p���i���ח\�萔(����)
		AddSql ",''"	'// ���ח\��KEYNO
		AddSql ",''"	'// FILLER
		AddSql ",'RfYotei'"	'// �ǉ��@�S����
		AddSql ",left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"	'// �ǉ��@���� YYYYMMDDhhmmss
		AddSql ",''"	'// �X�V�@�S����
		AddSql ",''"	'// �X�V�@���� YYYYMMDDhhmmss
		AddSql ")"
		CallSql strSql
	End Function
	'-------------------------------------------------------------------
	'StrDate()
	'-------------------------------------------------------------------
	Private	Function StrDate(byVal vDt)
		StrDate = ""
		if isDate(vDt) = False then
			exit function
		end if
		StrDate = Replace(vDt,"/","")
	End Function
	'-------------------------------------------------------------------
	'StrDateTm()
	'-------------------------------------------------------------------
	Private	Function StrDateTm(byVal vDt)
		StrDateTm = ""
		if isDate(vDt) = False then
			exit function
		end if
		StrDateTm = Replace(vDt,"/","") & "000000"
	End Function
	'-------------------------------------------------------------------
	'WorkDay()
	'-------------------------------------------------------------------
	Private	Function WorkDay(byVal vDt,byVal vDays)
		Debug ".WorkDay():" & vDt & ":" & vDays
		WorkDay = vDt
		if isDate(vDt) = False then
			exit function
		end if
		vDt = CDate(vDt) + vDays
		do while true
			select case WeekDay(vDt)
			case 1,7	'��,�y
				if vDays > 0 then
					vDt = vDt + 1
				else
					vDt = vDt - 1
				end if
			case else	'���`��
				exit do
			end select
		loop
		WorkDay = vDt
	End Function
	'-------------------------------------------------------------------
	'CCur()
	'-------------------------------------------------------------------
	Private	Function CCur(byVal v)
		CCur = 0
		if isNumeric(v) = false then
			exit function
		end if
		CCur = CLng(v)
	End Function
	'-------------------------------------------------------------------
	'GetValue()
	'-------------------------------------------------------------------
	Private	Function GetValue(objR)
		dim	strValue
		on error resume next
		strValue = Trim(objR)
		on error goto 0
		if Err.Number <> 0 then
'			Wscript.StdOut.WriteLine ".GetValue():0x" & Hex(Err.Number) & ":" & Err.Description
'			Wscript.StdOut.WriteLine
'			Wscript.StdOut.WriteLine objR.Address & ":(" & objR.Text & ")"
'			Wscript.Quit
			strValue = Trim(objR.Text)
		end if
		if strValue <> "" then
			if Asc(strValue) = 0 then
				strValue = ""
			end if
		end if
		strValue = Replace(strValue,vbCr,"")
		strValue = Replace(strValue,vbLf,"")
'		GetValue = Replace(GetValue,vbCrLf,"")
'		GetValue = Replace(GetValue,0,"")
		Debug "GetValue():" & objR.Address & ":" & strValue & ":"
		GetValue = strValue
	End Function
	'-------------------------------------------------------------------
	'Sql���s
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		select case Err.Number
'		case -2147467259	'�d��
		case 0,500
		case else
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine Err.Number & "(0x" & Hex(Err.Number) & "):" & Err.Description
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end select
		on error goto 0
'		on error resume next
'		Call objDB.Execute(strSql)
'		on error goto 0
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
	'������ǉ� strSql
	'-------------------------------------------------------------------
	dim	strSql
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
	End Function
	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	Private Function CreateExcel()
		Debug ".CreateExcel()"
		if objExcel is nothing then
			Debug ".CreateExcel():CreateObject(Excel.Application)"
			Set objExcel = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'AbsPath() ��΃p�X
	'-------------------------------------------------------------------
	Private	Function AbsPath(byVal strPath)
		dim	objFso
		Set objFso = CreateObject("Scripting.FileSystemObject")
		AbsPath = objFso.GetAbsolutePathName(strPath)
		Set objFso = Nothing
	End Function
	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	Private Function OpenBook(byVal strBkName)
		Debug ".OpenBook()"
		if objBook is nothing then
			strBkName = AbsPath(strBkName)
			Write strBkName & " :"
			on error resume next
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,"")
			WriteLine Err.Number
			if Err.Number <> 0 then
				WriteLine
				WriteLine "0x" & Hex(Err.Number) & ":" & Err.Description
				Quit
			end if
			on error goto 0
		end if
	end function
	'-------------------------------------------------------------------
	'Excel �t�@�C���N���[�Y
	'-------------------------------------------------------------------
	Private Function CloseBook()
		Debug ".CloseBook()"
		if not objBook is nothing then
			Debug ".CloseBook().Close:" & objBook.Name
			Call objBook.Close(False)
			set objBook = nothing
		end if
	end function
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
	'-------------------------------------------------------------------
	'Write
	'-------------------------------------------------------------------
	Private	Sub Write(byVal s)
		Wscript.StdOut.Write s
	End Sub
	'-------------------------------------------------------------------
	'WriteLine
	'-------------------------------------------------------------------
	Private	Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine s
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
End Class
