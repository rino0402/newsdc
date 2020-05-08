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
    Wscript.Echo "原産国登録(2011.12.13)"
    Wscript.Echo "rfgensan.vbs <xxx.csv>"
    Wscript.Echo "<例>"
    Wscript.Echo "rggensan.vbs xxx.csv"
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

	Call LoadGensan(objDb,objFile)

	objDb.Close
	set objDb = Nothing

	objFile.Close
	set objFSO = nothing
End Function

Function LoadGensan(byval objDb, _
				   byVal objFile _
				  )
	dim	strBuff
	dim	strPn
	dim	strBfr
	dim	strAft
	dim	lngCnt
	dim	aryBuff
	dim	strCountryCode

	strPn = ""

	lngCnt = 0
	do while ( objFile.AtEndOfStream = False )
		strBuff = objFile.ReadLine()
		lngCnt = lngCnt + 1
'		aryBuff = split(strBuff,",")
		aryBuff = GetCSV(strBuff)

		Wscript.Echo lngCnt & ":" & strBuff
		Wscript.Echo lngCnt & ": " & aryBuff(1) & "   " & aryBuff(2) & "   " & aryBuff(3) & "   " & aryBuff(4) & "   " & aryBuff(5)

		strPn = TrimCsv(aryBuff(1))
		strCountryCode = TrimCsv(aryBuff(5))
		Wscript.Echo lngCnt & ":" & strPn & ":" & strCountryCode
		call UpdateGensanPn(objDb,strPn,strCountryCode)
'		if lngCnt > 100 then
'			exit do
'		end if
	loop
End Function

Function UpdateGensanPn(byval objDb, _
					  byval strPn, _
					  byval strCountryCode _
					 )
	dim	strSql
	dim	rsPn
	dim	strTM

	strTM = GetDateTime(now())
 	
	strSql = "update Pn3 set"
	strSql = strSql & " MadeInCode = '" & strCountryCode & "'"
	strSql = strSql & ",UPD_ID = 'RfGen'"
	strSql = strSql & ",UPD_TM = '" & strTM & "'"
	strSql = strSql & " where JCode = '00021259'"
	strSql = strSql & "   and Pn = '" & strPn & "'"
	strSql = strSql & "   and MadeInCode <> '" & strCountryCode & "'"
	Wscript.Echo strSql
	set rsPn = objDb.Execute(strSql)
End Function

Function UpdateGensanItem(byval objDb, _
					  byval strPn, _
					  byval strCountryCode _
					 )
	dim	strSql
	dim	strName
	dim	rsCountry
	dim	rsItem

	strName = strCountryCode
	strSql = "select * from Country"
	strSql = strSql & " where CountryCode = '" & strCountryCode & "'"
	set rsCountry = objDb.Execute(strSql)
	if rsCountry.EOF = False then
		strName = rtrim(rsCountry("CountryName2"))
	end if

	strSql = "update item set"
	strSql = strSql & " TORI_GENSANKOKU = '" & strName & "'"
	strSql = strSql & ",UPD_TANTO = 'RfGen'"
	strSql = strSql & ",UPD_DATETIME = '20111213110000'"
	strSql = strSql & " where JGYOBU = 'R'"
	strSql = strSql & "   and NAIGAI = '1'"
	strSql = strSql & "   and HIN_GAI = '" & strPn & "'"
	strSql = strSql & "   and TORI_GENSANKOKU <> '" & strName & "'"
	Wscript.Echo strSql
	set rsItem = objDb.Execute(strSql)
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
