Private Function CheckHead(byval aryHead())
	Call DispMsg("CheckHead()")
	dim	strTable
	strTable = ""
	select case aryHead(0)
	case """PN共通_資産管理事業場コード"""
		strTable = "BoTana"
	case """在庫収支_倉庫コード"""
		strTable = "BoZaiko"
	case """PN倉庫（収支)_倉庫コード"""
		strTable = "BoZaiko"
	end select
	Call DispMsg("CheckHead():" & aryHead(0) & ":" & strTable)
	CheckHead = strTable
End Function

Function Load(byVal strDb,byval strFilename)
	dim	strRet
	strRet = ""
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db",strDb) & ")")
	set objDb = OpenAdodb(GetOption("db",strDb))
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	dim	cnt
	cnt = 0
	dim	objRs
	dim	lngOk
	dim	lngDp
	dim	lngNk
	lngOk = 0
	lngDp = 0
	lngNg = 0
	do while ( objFile.AtEndOfStream = False )
		cnt = cnt + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		Call DispMsg("読込中..." & cnt & " OK:" & lngOk & " DP:" & lngDp)
		Call Debug(strBuff)
		dim	aryBuff
		aryBuff = GetTab(strBuff)
		if cnt = 1 then
			dim	aryTop
			aryTop = aryBuff
			dim	strTable
			strTable = CheckHead(aryBuff)
			if strTable = "" then
				strRet = strRet & "Error:対応していないファイルです。" & vbCrlf
				exit do
			end if
			call DispMsg("OpenRs(" & strTable & ")")
			set objRs = OpenRs(objDb,strTable)
			call DispMsg("ExecuteAdodb(" & strTable & ")")
			objDb.CommandTimeout = 0
			Call ExecuteAdodb(objDb,"delete from " & strTable)
		else
			on error resume next
			objRs.AddNew
			on error goto 0
			dim	i
			i = 0
			dim	c
			for each c in (aryBuff)
				c = GetTrim(c)
				Call Debug(i & ":" & objRs.Fields(i).Name & ":" & c)
				if strTable = "BoZaiko" then
					select case i
					case 8
						objRs.Fields(9) = c
					case 9
						objRs.Fields(8) = c
					case else
						objRs.Fields(i) = c
					end select
				else
					objRs.Fields(i) = c
				end if
				i = i + 1
			next
			on error resume next
			Call objRs.UpdateBatch
			select case Err.Number
			case &h80004005
				Call objRs.CancelUpdate
				lngDp = lngDp + 1
				Call DispMsg("■二重登録■")
			case 0
				lngOk = lngOk + 1
			case else
				strRet = strRet & strBuff & vbCrlf
				strRet = strRet & "エラー発生：" & cnt & vbCrlf
				strRet = strRet & "Error.Number:0x" & Hex(Err.Number) & vbCrlf
				strRet = strRet & "Error.Description:" & Err.Description & vbCrlf
				lngNg = lngNg + 1
				Call objRs.CancelUpdate
				Exit Do
			end select
			on error goto 0
		end if
	loop
	objFile.Close
	set objFile = nothing
	set objFSO = nothing

	set objRs = CloseRs(objRs)
	set objDb = nothing
	strRet = strRet & "テーブル名：" & strTable & vbCrlf
	strRet = strRet & "　　　正常：" & Right(Space(10) & formatnumber(lngOk,0,,-1),10) & "件" & vbCrlf
	strRet = strRet & "　　　重複：" & Right(Space(10) & formatnumber(lngDp,0,,-1),10) & "件" & vbCrlf
	strRet = strRet & "　　エラー：" & Right(Space(10) & formatnumber(lngNg,0,,-1),10) & "件" & vbCrlf
	Load = strRet
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

Private Function makeSql()
	dim	strSql
	dim	strTop
	strTop = GetOption("top","")
	if strTop <> "" then
		strTop = " top " & strTop
	end if
	strSql = "select" & strTop
	strSql = strSql & " *"
	strSql = strSql & " from BoTana"
	makeSql = strSql
End Function

Function GetTab(ByVal s)
    Dim r
	r = Split(s,vbTab)
	GetTab = r
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
