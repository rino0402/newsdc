Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objCsv
	Set objCsv = New Csv
	objCsv.Run
	Set objCsv = Nothing
End Function
'-----------------------------------------------------------------------
'Cnv
'-----------------------------------------------------------------------
Const ForReading = 1
Class Csv
	'-----------------------------------------------------------------------
	'�g�p���@
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Wscript.Echo "b2csv.vbs [option]"
		Wscript.Echo "Ex."
		Wscript.Echo "cscript//nologo b2csv.vbs b2data.csv"
	End Sub
	'-----------------------------------------------------------------------
	'�ϐ�
	'-----------------------------------------------------------------------
	Private	strFileName
	Private	objFileSys
	Private	objFile
	'-----------------------------------------------------------------------
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strFileName = "" then
				strFileName = strArg
			else
				Init = "�t�@�C�������w��:" & strArg
				Echo Init
				Exit Function
			end if
		Next
		if strFileName = "" then
			Init = "�t�@�C�����w��"
			Echo Init
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "debug"
			case else
				Init = "�I�v�V�����G���[:" & strArg
				Echo Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		set	objFile	= nothing
		Set objFileSys	= WScript.CreateObject("Scripting.FileSystemObject")
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set	objFile	= nothing
		set objFileSys	= nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		if Init() <> "" then
			exit function
		end if
		OpenCsv
		LoadCsv
		CloseCsv
	End Function
	'-----------------------------------------------------------------------
	'LoadCsv() �Ǎ�
	'-----------------------------------------------------------------------
	Private	intNum
    Public Function LoadCsv()
		Debug ".LoadCsv():" & strFileName
		dim	lngRow
		lngRow = 0
		intNum = 0
		dim	strBuff
		do while ( objFile.AtEndOfStream = False )
			strBuff = objFile.ReadLine()
			lngRow = lngRow + 1
			Line lngRow,strBuff
		loop
		WriteLine "���v:" & intNum & "��"
	End Function
	'-------------------------------------------------------------------
	'Line()
	'-------------------------------------------------------------------
	Private Function Line(byVal lngRow,byVal strBuff)
		Debug ".Line(" & lngRow & "):" & strBuff
		dim	aryCsv
		aryCsv = GetCSV(strBuff)
		dim	i
		for i = LBound(aryCsv) to UBound(aryCsv)
			Debug ".Line(" & lngRow & "," & i & "):" & aryCsv(i)
		next
		if lngRow = 1 then
			exit function
		end if
'20170707-164239-12 S5YN30 �Y�@�Q�n�߰󒍊Ǘ� 18��
'20170718-092723-3 5YD18K000 �k�C���e�N�j�J���T�[�r�X 3��
'20170718-115828-5 S5YH10 ��ſƯ��Y�@���ѽ� �ߋE�x�X 2��
'20170718-141030-7  2922100E ��ſƯ��Y�@���ѽސÉ��c�Ə� 1��
'12345678901234567891234567891234567890123456789012345678
		Write Format(lngRow - 1,-3) & " "
		Write Format(aryCsv(2),19)
		Write Format(aryCsv(3),10)
		Write Format(aryCsv(5),28)
		dim	strNum
		strNum = Format(Split(aryCsv(15),"�F")(1),-4)
		Write strNum
		WriteLine ""
		intNum = intNum + CInt(Replace(strNum,"��",""))
	end function
	'-------------------------------------------------------------------
	'����
	'-------------------------------------------------------------------
	Private Function Format(byVal strV,byVal intLen)
		Format = strV
		if intLen > 0 then
			Format = LeftB(Format & space(intLen),intLen)
		else
			intLen = Abs(intLen)
			Format = Right(space(intLen) & Format,intLen)
		end if
	End Function
	'-------------------------------------------------------------------
	'Get_LeftB()
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
	'-------------------------------------------------------------------
	'GetCSV()
	'-------------------------------------------------------------------
	Private Function GetCSV(ByVal s)
	    Const One = 1
	    ReDim r(0)

	    Const sUndef = 11 ' ���m��(�J���}���_�u���N�H�[�e�[�V�������u�X�y�[�X�ȊO�̕����v��҂��)
	    Const sQuot = 22 ' �_�u���N�H�[�e�[�V�����ň͂܂ꂽ���Ƃ��J�n���Ă��܂������(�_�u���N�H�[�e�[�V��������т��̌�̃J���}�҂�)
	    Const sPlain = 33 ' �_�u���N�H�[�e�[�V�����Ȃ��̂��Ƃ��J�n���Ă��܂������(�J���}�҂�)
	    Const sTerm = 44 ' �_�u���N�H�[�e�[�V�����ň͂܂ꂽ���Ƃ��I�����Ă��܂������(�J���}�҂�)
	    Const sEsc = 55 ' �_�u���N�H�[�e�[�V�����ň͂܂ꂽ���Ƃ��J�n���Ă��܂�����ԂŁA���_�u���N�H�[�e�[�V�������o��������ԁB
	    Dim w
	    w = sUndef

	    Dim a
	    a = ""
	    Dim i
	    For i = 0 To Len(s) - One + 1
	        Dim c
	        c = Mid(s, i + One, 1)
	        If c = """" Then
	            If w = sUndef Then
	                a = ""
	                w = sQuot
	            ElseIf w = sQuot Then
	                w = sEsc
	            ElseIf w = sPlain Then ' �G���[
	                ReDim r(0)
	                Exit For
	            ElseIf w = sTerm Then ' �G���[
	                ReDim r(0)
	                Exit For
	            ElseIf w = sEsc Then
	                a = a & c
	                w = sQuot
	            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
	            End If
	        ElseIf c = "," Then
	            If w = sUndef Then
	                ReDim Preserve r(UBound(r) + 1)
	                r(UBound(r)) = ""
	            ElseIf w = sQuot Then
	                a = a & c
	            ElseIf w = sPlain Then
	                ReDim Preserve r(UBound(r) + 1)
	                r(UBound(r)) = RTrim(a)
	                a = ""
	                w = sUndef
	            ElseIf w = sTerm Then
	                a = ""
	                w = sUndef
	            ElseIf w = sEsc Then
	                ReDim Preserve r(UBound(r) + 1)
	                r(UBound(r)) = a
	                a = ""
	                w = sUndef
	            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
	            End If
	        ElseIf c = " " Then
	            If w = sUndef Then
	                ' do nothing.
	            ElseIf w = sQuot Then
	                a = a & c
	            ElseIf w = sPlain Then
	                a = a & c
	            ElseIf w = sTerm Then
	                ' do nothing
	            ElseIf w = sEsc Then
	                ReDim Preserve r(UBound(r) + 1)
	                r(UBound(r)) = a
	                a = ""
	                w = sTerm
	            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
	            End If
	        ElseIf c = "" Then ' �ŏI���[�v�̂�
	            If w = sUndef Then
	                ReDim Preserve r(UBound(r) + 1)
	                r(UBound(r)) = ""
	            ElseIf w = sQuot Then
	                ReDim r(0)
	                Exit For
	            ElseIf w = sPlain Then
	                ReDim Preserve r(UBound(r) + 1)
	                r(UBound(r)) = RTrim(a)
	                a = ""
	                w = sUndef
	            ElseIf w = sTerm Then
	                ' do nothing
	            ElseIf w = sEsc Then
	                ReDim Preserve r(UBound(r) + 1)
	                r(UBound(r)) = RTrim(a)
	                a = ""
	                w = sUndef
	            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
	            End If
	        Else
	            If w = sUndef Then
	                a = a & c
	                w = sPlain
	            ElseIf w = sQuot Then
	                a = a & c
	            ElseIf w = sPlain Then
	                a = a & c
	            ElseIf w = sTerm Then
	                ReDim r(0)
	                Exit For
	            ElseIf w = sEsc Then
	                ReDim r(0)
	                Exit For
	            Else ' �����ɗ��邱�Ƃ͂Ȃ��B
	            End If
	        End If
	    Next

	    GetCSV = r
	End Function

	'-------------------------------------------------------------------
	'CloseCsv() �t�@�C���N���[�Y
	'-------------------------------------------------------------------
	Private Function CloseCsv()
		Debug ".CloseCsv()"
		objFile.Close
		set objFile		= nothing
	end function
	'-------------------------------------------------------------------
	'OpenCsv() �t�@�C���I�[�v��
	'-------------------------------------------------------------------
	Private Function OpenCsv()
		Debug ".OpenCsv():" & GetAbsPath(strFileName)
		Set objFile	= objFileSys.OpenTextFile(GetAbsPath(strFileName), ForReading, False)
	end function
	'-------------------------------------------------------------------
	'��΃p�X
	'-------------------------------------------------------------------
	Private Function GetAbsPath(byVal strPath)
		strPath		= objFileSys.GetAbsolutePathName(strPath)
		GetAbsPath	= strPath
	End Function
	'-----------------------------------------------------------------------
	'Write
	'-----------------------------------------------------------------------
	Public Sub Write(byVal strMsg)
		Wscript.StdOut.Write strMsg
	End Sub
	'-----------------------------------------------------------------------
	'WriteLine
	'-----------------------------------------------------------------------
	Public Sub WriteLine(byVal strMsg)
		Wscript.StdOut.WriteLine strMsg
	End Sub
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
	Public Sub Echo(byVal strMsg)
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
End Class
