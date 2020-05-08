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
	Wscript.Echo "Active注文データ"
	Wscript.Echo "a_order.vbs [option]"
	Wscript.Echo " /list"
	Wscript.Echo " /debug"
	Wscript.Echo " /load <filename>"
	Wscript.Echo " /top:<num>"
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
		case "debug"
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
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))
	dim	objRs
	set objRs = OpenRs(objDb,"a_order")
	Call ExecuteAdodb(objDb,"delete from a_order where jCode = ''")
	dim	objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	dim	objFile
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	dim	cnt
	cnt = 0
	dim	cntAdd
	cntAdd = 0
	Dim		aryJCode()
	ReDim	aryJCode(0)
	do while ( objFile.AtEndOfStream = False )
		cnt = cnt + 1
		dim	strBuff
		strBuff = objFile.ReadLine()
		Call Debug(strBuff)
		dim		aryTitle
		dim		aryBuff
		if cnt = 1 then
			aryTitle = GetCSV(strBuff)
		elseif cnt > 1 then
			cntAdd = cntAdd + 1
			aryBuff = GetCSV(strBuff)
			objRs.AddNew
			dim	i
			i = 0
			dim	c
			for i = LBound(aryBuff) to UBound(aryBuff)
				c = aryBuff(i)
				dim	strFName
				strFName = GetFName(aryTitle(i))
				Call Debug(i & ":" & strFName & ":" & c)
				if strFName <> "" then
					select case strFName
					case "JCode"
						if chkJCode(aryJCode,c) <> "" then
							Call Debug("delete jCode=" & c )
							Call ExecuteAdodb(objDb,"delete from a_order where jCode = '" & c & "'")
							Call Debug("aryJCode(" & UBound(aryJCode) & ")=" & aryJCode(UBound(aryJCode)))
			                ReDim Preserve aryJCode(UBound(aryJCode) + 1)
            			    aryJCode(UBound(aryJCode)) = c
							Call Debug("aryJCode(" & UBound(aryJCode) & ")=" & aryJCode(UBound(aryJCode)))
						end if
					case "strFName" _
						,"SrvDtSts" _
						,"ChuKbn"
						c = Left(c,1)
'					case 5,8,14,15,18
'						c = Left(c,1)
'					case 20,21,22
'						c = Replace(c,"/","")
					end select
					objRs.Fields(strFName) = c
				end if
			next
			Call objRs.UpdateBatch
		end if
	loop
	objFile.Close
	set objFile = nothing
	set objFSO = nothing

	set objRs = CloseRs(objRs)
	set objDb = nothing
	Call DispMsg("読込件数：" & cnt)
	Call DispMsg("登録件数：" & cntAdd)
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

Private Function GetFName(byval strTitle)
	dim	strFName
	strFName = ""
						'strFName = "BTKbn"		' 部品取引区分
	select case strTitle
	case "取引区分"
	case "サービスデータ進捗区分"
						strFName = "SrvDtSts"	' サービスデータ進捗区分
	case "資産管理事業場コード"
						strFName = "JCode"		' 資産事業　資産管理事業場コード
	case "品目番号"
						strFName = "Pn"			' 出荷品目番号
	case "グローバル品目番号"
	case "サービス品目番号"
	case "受付品目番号"
						strFName = "PnRcv"		' 受注品目番号
	case "相手先コード"
	case "相手先名"
	case "数量"
'						strFName = "QtyRcv"		' 受注実績数
						strFName = "QtySnd"		' 受注実績数
	case "単価"
						strFName = "Price"		' 単価　実際単価    9999999.0000
	case "実際金額"
						strFName = "Amount"		' 実際金額
	case "オーダーNo."
						strFName = "OrderNo"	' オーダーNO
	case "ITEM-No."
	case "伝票番号"
						strFName = "DenNo"		' 伝票番号
	case "ID-No."
						strFName = "IDNo"		' ID-NO
	case "在庫収支略式名"
						strFName = "ZSyushiRk"	' 在庫収支略式名
	case "在庫収支コード"
	case "資産管理在庫収支コード"
	case "補助在庫収支コード"
	case "帳端区分"
						strFName = "CHKbn"		' 帳端区分
	case "値差区分"
						strFName = "NSKbn"		' 値差区分
	case "返品区分"
	case "実績日(予定日)"
						strFName = "SalesDt"	' 売上予定年月日 yyyymmdd
	case "受発注年月日"
						strFName = "RcvDt"		' 受注年月日
	case "出庫年月日"
						strFName = "PckDt"		' 出庫予定年月日
	case "出荷年月日"
						strFName = "SndDt"		' 出荷予定年月日
	case "発送年月日"
	case "出荷指定年月日"
	case "指定納期年月日"
						strFName = "DlvDt"		' 指定納期日　指定納期年月日
	case "納期回答年月日"
						strFName = "AnsDt"		' 納期回答日　納期回答年月日
	case "受注出荷・販売区分"
	case "受注出荷・直送先コード"
	case "受注出荷・注文区分"
						strFName = "ChuKbn"		' 注文区分
	end select
	GetFName = strFName
End Function

Private Function List()
	dim	objDb
	Call Debug("OpenAdodb(" & GetOption("db","newsdc") & ")")
	set objDb = OpenAdodb(GetOption("db","newsdc"))

	dim	strSql
	strSql = makeSql()

	dim	rsList
	Call DispMsg("objDb.Execute(" & strSql & ")")
	set rsList = objDb.Execute(strSql)

	do while rsList.Eof = False
		Call DispMsg("■" _
			 & " " & rsList.Fields("IDNo") _
			 & " " & rsList.Fields("JCode") _
			 & " " & rsList.Fields("Pn") _
			 & " " & rsList.Fields("PnRcv") _
			 & " " & rsList.Fields("BTKbn") _
			 & " " & rsList.Fields("TKCode") _
			 & " " & rsList.Fields("ChokuCode") _
			 & " " & rsList.Fields("SrvDtSts") _
					)
		Call rsList.MoveNext
	loop

	Call Debug("CloseAdodb(" & GetOption("db","newsdc") & ")")
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
	strSql = strSql & " from a_order"
	makeSql = strSql
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
