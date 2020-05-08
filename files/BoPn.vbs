Option Explicit
'-----------------------------------------------------------------------
'メイン呼出＆インクルード
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
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "BOPN共通"
	Wscript.Echo "BoPn.vbs [option]"
	Wscript.Echo " /db:<dbname>      : Ex.newsdc-nar"
	Wscript.Echo " /list"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
	Wscript.Echo "bopn.vbs /db:newsdc-nar /load ""I:\pos\PPSC奈良\bo\16399722_20130716.csv"""
	Wscript.Echo "bopn.vbs /db:newsdc-nar /load ""I:\pos\PPSC奈良\bo\PN共通（機種品目追加）.csv"""
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	For Each strArg In WScript.Arguments.UnNamed
    	select case strArg
		case else
			if strFilename <> "" then
				usage()
				Main = 1
				exit Function
			end if
			strFilename = strArg
		end select
	Next
	For Each strArg In WScript.Arguments.Named
    	select case lcase(strArg)
		case "db"
		case "list"
		case "load"
		case "top"
		case else
			usage()
			Main = 1
			exit Function
		end select
	Next
	select case GetFunction()
	case "list"
		Call List()
	case "load"
		Call Load(strFilename)
	end select
	Main = 0
End Function

Private Function GetFunction()
	GetFunction = "list"
	if WScript.Arguments.Named.Exists("load") then
		GetFunction = "load"
	end if
End Function

Private Function Load(byval strFilename)
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	dim	objRs
	set objRs = OpenRs(objDb,"BoPn")
'	Call ExecuteAdodb(objDb,"delete from BoPn")
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	dim	cnt
	cnt = 0
	Dim		aryJCode()
	ReDim	aryJCode(0)
	do while ( objFile.AtEndOfStream = False )
		cnt = cnt + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		Call DispMsg(strBuff)
		if cnt > 1 then
			dim	aryBuff
			aryBuff = GetTab(strBuff)
			objRs.AddNew
			dim	i
			i = 0
			dim	c
			for each c in (aryBuff)
				c = GetTrim(c)
				Call DispMsg(i & ":" & c)
				select case i
				case 2	' 資産管理事業場コード
					if chkJCode(aryJCode,c) <> "" then
						Call Debug("delete jCode=" & c )
						Call ExecuteAdodb(objDb,"delete from BoPn where ShisanJCode = '" & c & "'")
						Call Debug("aryJCode(" & UBound(aryJCode) & ")=" & aryJCode(UBound(aryJCode)))
		                ReDim Preserve aryJCode(UBound(aryJCode) + 1)
           			    aryJCode(UBound(aryJCode)) = c
						Call Debug("aryJCode(" & UBound(aryJCode) & ")=" & aryJCode(UBound(aryJCode)))
					end if
				case else
				end select
				objRs.Fields(i) = c
				i = i + 1
			next
			On Error Resume Next
			dim	strMsg
			Call objRs.UpdateBatch
			if Err.Number = 0 then
				strMsg = "Ok"
			else
				strMsg = "Err:" & Err.Number & " " & Err.Description
				Call objRs.CancelUpdate
			end if
			Call DispMsg(strMsg)
			Err.Clear
			On Error Goto 0
		end if
	loop
	objFile.Close
	set objFile = nothing
	set objFSO = nothing

	set objRs = CloseRs(objRs)
	set objDb = nothing
End Function

Private Function chkJCode(byVal aryJCode(),byVal strJCode)
	dim	a
'	for each a in aryJCode
	dim	i
	Call Debug("chkJCode:" & LBound(aryJCode) & " to " & UBound(aryJCode))
	for i = LBound(aryJCode) to UBound(aryJCode)
		a = aryJCode(i)
		Call Debug("chkJCode:" & a & "=" & strJCode)
		if a = strJCode then
			strJCode = ""
			exit for
		end if
	next
	chkJCode = strJCode
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

Private Function List()
	dim	objDb
	Call DispMsg("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	Call DispMsg("objDb.Execute(" & strSql & ")")
	set rsList = objDb.Execute(strSql)

	do while rsList.Eof = False
		Call DispMsg("■" _
			 & " " & rsList.Fields("Soko") _
			 & " " & rsList.Fields("JCode") _
			 & " " & rsList.Fields("ShisanJCode") _
			 & " " & rsList.Fields("Pn") _
			 & " " & rsList.Fields("PnBetsu") _
			 & " " & rsList.Fields("PNameEngA") _
			 & " " & rsList.Fields("PName") _
			 & " " & rsList.Fields("KobaiTanto") _
			 & " " & rsList.Fields("NaiKbn") _
			 & " " & rsList.Fields("GaiKbn") _
			 & " " & rsList.Fields("UntKbn") _
			 & " " & rsList.Fields("NaiDisconYm") _
			 & " " & rsList.Fields("GaiDisconYm") _
			 & " " & rsList.Fields("UntDisconYm") _
					)
		Call rsList.MoveNext
	loop

	Call DispMsg("CloseAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = Nothing
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
	strSql = strSql & " from BoPn"
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
