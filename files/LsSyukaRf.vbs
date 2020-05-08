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
Call Include("get_b.vbs")
Call Include("file.vbs")
Call Include("excel.vbs")
Call Include("debug.vbs")
dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "Bo�o�׃f�[�^�ϊ�"
	Wscript.Echo "LsSyukaRf.vbs [option] <filename>"
	Wscript.Echo " /db:dns"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript LsSyukaRf.vbs /db:newsdc9 ""I:\0SDC_honsya\���ƕ��ʏ��i���o�׋��z�܂Ƃ�\RF NPL����̏o�׎���\�y00021259�z�̔��݌Ɏ���_201510.xlsx"""
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	objLsSyukaRf
	Set objLsSyukaRf = New LsSyukaRf
	if objLsSyukaRf.Init() <> "" then
		call usage()
		exit function
	end if
	call objLsSyukaRf.Run()
End Function
'-----------------------------------------------------------------------
'Bo�o�׃f�[�^�ϊ�
'-----------------------------------------------------------------------
Class LsSyukaRf
	Private	strDBName
	Private	objDB
	Private	objRs
	Private	objXL
	Private	objBk
	Private	objSt
	Private	strFilename
	Private	strSql
	Private obj103
	Private strFunction
	Private lngRow
	Private lngMaxRow
	Private	strTable
	Private	strYM
	Private	strJCd
	Private	strSoko
	Private	strType
	Private	strDeleteSql

    Private Sub Class_Initialize
		Call Debug("LsSyukaRf.Class_Initialize()")
		strDBName = GetOption("db","newsdc")
		set objDB = nothing
		set objRs = nothing
		set objXL = nothing
		set objBk = nothing
		set objSt = nothing
		strFilename = ""
        strFunction = "check"
		strTable = "LsSyuka"
		strYM = ""
		strJCd = ""
		strSoko = ""
		strDeleteSql = ""
    End Sub

    Private Sub Class_Terminate
		Call Debug("LsSyukaRf.Class_Terminate()")
'		Call Close()
		if not objBk is nothing then
			call objBk.Close(False)
			set objBk = nothing
		end if
		set objXL = nothing
    End Sub

    Public Function Init()
		Call Debug("LsSyukaRf.Init()")
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
	    	select case strArg
			case else
				if strFilename <> "" then
					Init = "filename error"
					exit Function
				end if
				strFilename = strArg
				Call Debug("strFilename=" & strFilename)
			end select
		Next
		if strFilename = "" then
			Init = "filename error"
			exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "debug"
			case "c","check"
                strFunction = "check"
			case "i","info"
                strFunction = "info"
			case else
				Init = "unknown option:" & strArg
				Exit Function
			end select
		Next
	End Function

    Public Function Run()
		Call Debug("LsSyukaRf.Run()")
		Call CreateExcelApp()
		Call OpenExcel()
		Call LoadExcel()
	End Function

	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	Private Function CreateExcelApp()
		Call Debug("LsSyukaRf.CreateExcelApp()")
		if objXL is nothing then
			Call Debug("	CreateObject(Excel.Application)")
			Set objXL = WScript.CreateObject("Excel.Application")
		end if
	end function

	'-------------------------------------------------------------------
	'Excel �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	Private Function OpenExcel()
		Call Debug("LsSyukaRf.OpenExcel()")
		if objBk is nothing then
			Call Debug("	Workbooks.Open=" & strFilename)
			Set objBk = objXL.Workbooks.Open(strFilename,False,True,,"")
			Call Debug("	    objBk.Path=" & objBk.Path)
			Call Debug("	    objBk.Name=" & objBk.Name)
		end if
	end function
	'-------------------------------------------------------------------
	'�Ǎ�����
	'-------------------------------------------------------------------
	Private Function LoadExcel()
		Call Debug("LsSyukaRf.LoadExcel()")
		if objBk is nothing then
			exit function
		end if
		For each objSt in objBk.Worksheets
			Call LoadXls()
		Next
	end function

	Private Function DataType()
		Call Debug("LsSyukaRf.DataType()")
		Call DispMsg(objSt.Name)
		Call DispMsg( ""_
					& " " & objSt.Range("A1") _
					& " " & objSt.Range("B1") _
					& " " & objSt.Range("C1") _
					& " " & objSt.Range("D1") _
					& " " & objSt.Range("E1") _
					)

		DataType = ""
		strType = ""
		if objSt is nothing then
			exit function
		end if
		if	objSt.Range("A1") = "NO" and _
			objSt.Range("B1") = "���Ə�CD" and _
			objSt.Range("C1") = "�i�ڔԍ�" and _
			objSt.Range("D1") = "�̔�����" then
			strType = "Row4"
		'NO	���Ə�CD	�i�ڔԍ�	�̔��敪	�̔�����
		elseif	objSt.Range("A1") = "NO" and _
			objSt.Range("B1") = "���Ə�CD" and _
			objSt.Range("C1") = "�i�ڔԍ�" and _
			objSt.Range("D1") = "�̔��敪" and _
			objSt.Range("E1") = "�̔�����" then
			strType = "Row5"
		'���Ə�CD	�i�ڔԍ�	�̔�����
		elseif	objSt.Range("A1") = "���Ə�CD" and _
			objSt.Range("B1") = "�i�ڔԍ�" and _
			objSt.Range("C1") = "�̔�����" then
			strType = "Row3"
		else
			exit function
		end if
		DataType = strType
	end Function

	Private Function LoadXls()
		Call Debug("LsSyukaRf.LoadXls()")
		if objSt is nothing then
			exit function
		end if
		Call Debug("	objSt.Name=" & objSt.Name)
		strYM = right(split(objBk.name,".")(0),6)
		if DataType() = "" then
			exit function
		end if
		select case strType
		case "Row3"
			strJCd = objSt.Range("A2")
			select case objSt.Name
			case "�����̔�����"
				strSoko = "NAI"
			case "OEM�E�C�O�̔�����"
				strSoko = "GAI"
			case else
				strSoko = "---"
			end select
		case "Row4"
			strJCd = objSt.Range("B2")
			select case objSt.Name
			case "�i�O���j�����̔�����"
				strSoko = "NAI"
			case "�i�O���jOEM�̔�����"
				strSoko = "OEM"
			case "�i�O���j�C�O�̔�����"
				strSoko = "GAI"
			case else
				strSoko = "---"
			end select
		case "Row5"
			strJCd = objSt.Range("B2")
			select case objSt.Name
			case "�����̔��iOEM�܂ށj"
				strSoko = "NAI"
			case "�C�O�̔�����"
				strSoko = "GAI"
			case else
				strSoko = "---"
			end select
		end select
		Call OpenDB()
		Call OpenRs()
		Call DeleteRs()
		lngMaxRow = excelGetMaxRow(objSt,"A",2)
		for lngRow = 2 to lngMaxRow
			dim	strMsg
			strMsg = AddRecord()
			Call DispMsg(lngRow & "/" & lngMaxRow _
						& " " & strType _
						& " " & objSt.Name _
						& " " & strYM _
						& " " & strSoko _
						& " " & objSt.Range("A" & lngRow) _
						& " " & objSt.Range("B" & lngRow) _
						& " " & objSt.Range("C" & lngRow) _
						& " " & objSt.Range("D" & lngRow) _
						& " " & objSt.Range("E" & lngRow) _
						& " " & strMsg _
						)
		next
		Call CloseRs()
		Call CloseDB()
	end function

    Private Function OpenDB()
		Call Debug("LsSyukaRf.OpenDB():" & strDBName)
		set objDb = OpenAdodb(strDBName)
    End Function

    Private Function CloseDB()
		Call Debug("LsSyukaRf.CloseDB():" & strDBName)
		Call objDb.Close()
		set objDb = Nothing
    End Function

	Private Function OpenRs()
		Call Debug("LsSyukaRf.OpenRs()")
		if strDeleteSql = "" then
			dim	strSql
			strSql = "delete from " & strTable
			strSql = strSql & " where YM = '" & strYM & "'"
			strSql = strSql & "   and JCd = '" & strJCd & "'"
'			strSql = strSql & "   and Soko = '" & strSoko & "'"
			Call DispMsg("strSql=" & strSql)
			Call objDb.Execute(strSql)
			strDeleteSql = strSql
		end if
	End Function

	Private Function DeleteRs()
		Call Debug("LsSyukaRf.DeleteRs()")
		Set objRs = Wscript.CreateObject("ADODB.Recordset")
		objRs.Open strTable, objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	End Function

	Private Function CloseRs()
		Call Debug("LsSyukaRf.CloseRs()")
		Call Debug("Table=" & strTable)
		Call objRs.Close()
		set objRs = Nothing
	End Function

	Private Function AddRecord()
		Call Debug("LsSyukaRf.AddRecord()")
		if objRs is Nothing then
			exit function
		end if
		objRs.AddNew
		select case strType
		case "Row3"
			objRs.Fields("YM")		= strYM
			objRs.Fields("No")		= lngRow
			objRs.Fields("Soko")	= strSoko
			objRs.Fields("JCd")		= objSt.Range("A" & lngRow)
			objRs.Fields("Pn")		= objSt.Range("B" & lngRow)
			objRs.Fields("Qty")		= CLng(objSt.Range("C" & lngRow))
		case "Row4"
			objRs.Fields("YM")		= strYM
			objRs.Fields("No")		= objSt.Range("A" & lngRow)
			objRs.Fields("Soko")	= strSoko
			objRs.Fields("JCd")		= objSt.Range("B" & lngRow)
			objRs.Fields("Pn")		= objSt.Range("C" & lngRow)
			objRs.Fields("Qty")		= CLng(objSt.Range("D" & lngRow))
		case "Row5"
			objRs.Fields("YM")		= strYM
			objRs.Fields("No")		= objSt.Range("A" & lngRow)
			objRs.Fields("Soko")	= strSoko
			objRs.Fields("JCd")		= objSt.Range("B" & lngRow)
			objRs.Fields("Pn")		= objSt.Range("C" & lngRow)
			objRs.Fields("Qty")		= CLng(objSt.Range("E" & lngRow))
		end select
		on error resume next
			objRs.UpdateBatch
			select case Err.Number
			case &h80004005
				AddRecord = "����d�o�^��"
				Call objRs.CancelUpdate
			case 0
			case else
				AddRecord = "0x" & Hex(Err.Number) & " " & Err.Description
				Call objRs.CancelUpdate
			end select
		on error goto 0
	End Function

End Class
