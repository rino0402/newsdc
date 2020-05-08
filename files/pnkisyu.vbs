Option Explicit
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
End Function
Call Include("const.vbs")

Call Main()

Function usage()
    Wscript.Echo "PN代表機種一括登録(2011.12.20)"
    Wscript.Echo "pnkisyu.vbs <PN代表機種.csv>"
    Wscript.Echo "<例>"
    Wscript.Echo "pnkisyu.vbs PN代表機種.csv"
End Function

Function Main()
	dim	strFilename
	dim	objFSO
	dim	objFile
	dim	objDb
	dim	strDbName

	strFilename = ""
	strDbName	= "newsdc-kst"
	If WScript.Arguments.Count > 0 Then
		strFilename	= WScript.Arguments(0)
	end if
	
	if strFilename = "" then
		Call usage()
		Exit Function
	end if


	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)

	Set objDb = Wscript.CreateObject("ADODB.Connection")
	Call objDb.Open(strDbName)

	Call LoadPnKisyu(objDb,objFile)

	objDb.Close
	set objDb = Nothing

	objFile.Close
	set objFSO = nothing
End Function

Function LoadPnKisyu(byval objDb, _
				   byVal objFile _
				  )
	dim	strBuff
	dim	strPn
	dim	strPnKisyu1
	dim	strPnKisyu2
	dim	strPnTanto
	dim	strBfr
	dim	strAft
	dim	lngCnt
	dim	aryBuff

	strPn = ""

	lngCnt = 0
	do while ( objFile.AtEndOfStream = False )
		strBuff = objFile.ReadLine()
		lngCnt = lngCnt + 1
'		aryBuff = split(strBuff,",")
		aryBuff = GetCSV(strBuff)

		Wscript.Echo lngCnt & ":" & strBuff
		Wscript.Echo lngCnt & ": " & aryBuff(1) & "   " & aryBuff(2) & "   " & aryBuff(3) & "   " & aryBuff(4)

		strPn		= TrimCsv(aryBuff(1))
		strPnKisyu1 = TrimCsv(aryBuff(2))
		strPnKisyu2 = TrimCsv(aryBuff(3))
		strPnTanto	= TrimCsv(aryBuff(4))

		if lngCnt = 1 then
			call DeletePnKisyu(objDb)
		else
			Wscript.Echo lngCnt & ":" & strPn
			call InsertPnKisyu(objDb,strPn,strPnKisyu1,strPnKisyu2,strPnTanto)
		end if
'		if lngCnt > 100 then
'			exit do
'		end if
	loop
End Function

Function DeletePnKisyu(byval objDb _
					 )
	dim	rsPn
	dim	strSql

	strSql = "delete from PnKisyu"
	Wscript.Echo strSql
	set rsPn = objDb.Execute(strSql)
End Function


Function InsertPnKisyu(byval objDb, _
					  byval strPn, _
					  byval strPnKisyu1, _
					  byval strPnKisyu2, _
					  byval strPnTanto _
					 )
	dim	strSql
	dim	rsPn

	strSql = "insert into PnKisyu"
	strSql = strSql & " (JGYOBU"
	strSql = strSql & " ,NAIGAI"
	strSql = strSql & " ,HIN_GAI"
	strSql = strSql & " ,PnKisyu1"
	strSql = strSql & " ,PnKisyu2"
	strSql = strSql & " ,PnTanto"
	strSql = strSql & " ) values ("
	strSql = strSql & "  'A'"
	strSql = strSql & " ,'1'"
	strSql = strSql & " ,'" & strPn & "'"
	strSql = strSql & " ,'" & strPnKisyu1 & "'"
	strSql = strSql & " ,'" & strPnKisyu2 & "'"
	strSql = strSql & " ,'" & strPnTanto & "'"
	strSql = strSql & " )"
	Wscript.Echo strSql
	set rsPn = objDb.Execute(strSql)
End Function

Function TrimCsv(byval strCsv)
	dim	strTrim

	strTrim = rtrim(strCsv)
'	strTrim = right(strTrim,len(strTrim)-1)
'	strTrim = left(strTrim,len(strTrim)-1)
	TrimCsv = strTrim
End Function

Function GetCSV(ByVal s)
    Const One = 1
    ReDim r(0)

    Const sUndef = 11 ' 未確定(カンマかダブルクォーテーションか「スペース以外の文字」を待つ状態)
    Const sQuot = 22 ' ダブルクォーテーションで囲まれたことが開始してしまった状態(ダブルクォーテーションおよびその後のカンマ待ち)
    Const sPlain = 33 ' ダブルクォーテーションなしのことが開始してしまった状態(カンマ待ち)
    Const sTerm = 44 ' ダブルクォーテーションで囲まれたことが終了してしまった状態(カンマ待ち)
    Const sEsc = 55 ' ダブルクォーテーションで囲まれたことが開始してしまった状態で、かつダブルクォーテーションが出現した状態。
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
            ElseIf w = sPlain Then ' エラー
                ReDim r(0)
                Exit For
            ElseIf w = sTerm Then ' エラー
                ReDim r(0)
                Exit For
            ElseIf w = sEsc Then
                a = a & c
                w = sQuot
            Else ' ここに来ることはない。
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
            Else ' ここに来ることはない。
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
            Else ' ここに来ることはない。
            End If
        ElseIf c = "" Then ' 最終ループのみ
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
            Else ' ここに来ることはない。
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
            Else ' ここに来ることはない。
            End If
        End If
    Next

    GetCSV = r
End Function
