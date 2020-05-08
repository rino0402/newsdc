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
Call Include("excel.vbs")
Call Include("get_b.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
Function GetCD()
	Dim objWshShell
	'①WScript.Shellオブジェクトの作成
	Set objWshShell = CreateObject("WScript.Shell")
	'カレントディレクトリを表示
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
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "弥生得意先マスター"
	Wscript.Echo "YCustomer.vbs [option] <ファイル名>"
	Wscript.Echo " /db:fhd"
	Wscript.Echo " /debug"
	Wscript.Echo "ex."
	Wscript.Echo "sc32 YCustomer.vbs /db:fhd /debug ""F:\it\得意先リスト.xlsx"""
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	strFilename = ""
	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		if strFilename = "" then
			strFilename = strArg
		else
			strFilename = ""
		end if
	next
	'名前付きオプションチェック
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
	call LoadYCustomer(strFilename)
	Main = 0
End Function

Function LoadYCustomer(byVal strFilename)
	Call Debug("LoadYCustomer(" & strFilename & ")")
	'-------------------------------------------------------------------
	'Excelファイル名
	'-------------------------------------------------------------------
	if strFileName = "" then
		Call DispMsg("ファイル名を指定して下さい")
		Exit Function
	end if
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	dim	objXL
	Set objXL = WScript.CreateObject("Excel.Application")
	Call Debug("CreateObject(Excel.Application)")
	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	dim	strPassword
	strPassword = ""
	dim	objBk
	Set objBk = objXL.Workbooks.Open(strFilename,False,True,,strPassword)
	Call Debug("Workbooks.Open=" & objBk.Name)
	'-------------------------------------------------------------------
	'読込処理
	'-------------------------------------------------------------------
	Call LoadYCustomerXls(objXL,objBk)
	'-------------------------------------------------------------------
	'Excelの後処理
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
	Call Debug("LoadYCustomer():End")
End Function

Function LoadYCustomerXls(objXL,objBk)
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","fhd"))
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsYCustomer
	set rsYCustomer = OpenRs(objDb,"YCustomer")

	Call Debug("LoadYCustomerXls():BookName=" & objBk.Name)
	Dim	objSt
	For each objSt in objBk.Worksheets
		Dim	strShtName
		strShtName = objSt.Name
		Call Debug("LoadYCustomerXls():SheetName=" & strShtName)
		Call LoadYCustomerXst(objXL,objBk,objSt,objDb,rsYCustomer)
	Next
	
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsYCustomer = CloseRs(rsYCustomer)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function LoadYCustomerXst(objXL,objBk,objSt,objDb,rsYCustomer)
	Call Debug("LoadYCustomerXst():SheetName=" & objSt.Name)

	Call Debug("delete from YCustomer")
	Call ExecuteAdodb(objDb,"delete from YCustomer")

	dim	lngMaxRow
	lngMaxRow = excelGetMaxRow(objSt,"B",0)
	dim	lngRow
	For lngRow = 6 to lngMaxRow
		if LoadYCustomerRow(objXL,objBk,objSt,objDb,rsYCustomer,lngRow) = 0 then
			Exit For
		end if
	Next
End Function

Function LoadYCustomerRow(objXL,objBk,objSt,objDb,rsYCustomer,byVal lngRow)
	Call Debug("LoadYCustomerRow():" & objSt.Name & ":" & lngRow)
	dim	strB
	strB = objSt.Range("B" & lngRow)
	if strB = "" then
		LoadYCustomerRow = 0
		exit function
	end if
	rsYCustomer.AddNew
	Call SetField(rsYCustomer,objSt,"Code",		"B" , lngRow)	'		//  コード
	Call SetField(rsYCustomer,objSt,"Name",		"C" , lngRow)	'		//	名称
	Call SetField(rsYCustomer,objSt,"NameK",	"D" , lngRow)	'		//	フリガナ
	Call SetField(rsYCustomer,objSt,"NameR",	"E" , lngRow)	'		//	略称
	Call SetField(rsYCustomer,objSt,"Zip",		"F" , lngRow)	'		//	郵便番号
	Call SetField(rsYCustomer,objSt,"Address1",	"G" , lngRow)	'		//	住所１
	Call SetField(rsYCustomer,objSt,"Address2",	"H" , lngRow)	'		//	住所２
	Call SetField(rsYCustomer,objSt,"Section",	"I" , lngRow)	'		//	部署名
	Call SetField(rsYCustomer,objSt,"Post",		"J" , lngRow)	'		//	役職名
	Call SetField(rsYCustomer,objSt,"Person",	"K" , lngRow)	'		//	ご担当者
	Call SetField(rsYCustomer,objSt,"Prefix",	"L" , lngRow)	'		//	敬称
	Call SetField(rsYCustomer,objSt,"Tel",		"M" , lngRow)	'		//	TEL
	Call SetField(rsYCustomer,objSt,"Fax",		"N" , lngRow)	'		//	FAX
	Call SetField(rsYCustomer,objSt,"Mail",		"O" , lngRow)	'		//	メールアドレス
	Call SetField(rsYCustomer,objSt,"Url",		"P" , lngRow)	'		//	ホームページ
	Call SetField(rsYCustomer,objSt,"Tanto",	"Q" , lngRow)	'		//	担当者
	Call SetField(rsYCustomer,objSt,"TantoName","R" , lngRow)	'		//	担当者名
	Call SetField(rsYCustomer,objSt,"TKbn",		"S" , lngRow)	'		//	取引区分
	Call SetField(rsYCustomer,objSt,"TkS",		"T" , lngRow)	'		//	単価種類
	Call SetField(rsYCustomer,objSt,"Rate",		"U" , lngRow)	'		//	掛率
	Call SetField(rsYCustomer,objSt,"Bill1",	"V" , lngRow)	'		//	請求先
	Call SetField(rsYCustomer,objSt,"Bill2",	"W" , lngRow)	'		//	
	Call SetField(rsYCustomer,objSt,"SGrp1",	"X" , lngRow)	'		//	締グループ
	Call SetField(rsYCustomer,objSt,"SGrp2",	"Y" , lngRow)	'		//	
	Call SetField(rsYCustomer,objSt,"s1",		"Z" , lngRow)	'		//	金額端数処理
	Call SetField(rsYCustomer,objSt,"s2",		"AA" , lngRow)	'		//	税端数処理
	Call SetField(rsYCustomer,objSt,"s3",		"AB" , lngRow)	'		//	税転嫁
	Call SetField(rsYCustomer,objSt,"s4",		"AC" , lngRow)	'		//	与信限度額
	Call SetField(rsYCustomer,objSt,"s5",		"AD" , lngRow)	'		//	売掛残高
	Call SetField(rsYCustomer,objSt,"s6",		"AE" , lngRow)	'		//	回収方法
	Call SetField(rsYCustomer,objSt,"s7",		"AH" , lngRow)	'		//	回収サイクル
	Call SetField(rsYCustomer,objSt,"s8",		"AI" , lngRow)	'		//	回収日
	Call SetField(rsYCustomer,objSt,"s9",		"AJ" , lngRow)	'		//	手数料負担
	Call SetField(rsYCustomer,objSt,"s10",		"AK" , lngRow)	'		//	サイト
	Call SetField(rsYCustomer,objSt,"s11",		"AL" , lngRow)	'		//	指定売上伝票
	Call SetField(rsYCustomer,objSt,"s12",		"AM" , lngRow)	'		//	指定請求書
	Call SetField(rsYCustomer,objSt,"s13",		"AN" , lngRow)	'		//	宛名ラベル
	Call SetField(rsYCustomer,objSt,"s14",		"AO" , lngRow)	'		//	企業コード
	Call SetField(rsYCustomer,objSt,"Cate1",	"AP" , lngRow)	'		//	分類１
	Call SetField(rsYCustomer,objSt,"CateNate1","AQ" , lngRow)	'		//	分類１名称
	Call SetField(rsYCustomer,objSt,"Cate2",	"AR" , lngRow)	'		//	分類２
	Call SetField(rsYCustomer,objSt,"CateNate2","AS" , lngRow)	'		//	分類２名称
	Call SetField(rsYCustomer,objSt,"Cate3",	"AT" , lngRow)	'		//	分類３
	Call SetField(rsYCustomer,objSt,"CateNate3","AU" , lngRow)	'		//	分類３名称
	Call SetField(rsYCustomer,objSt,"Memo",		"AV" , lngRow)	'		//	メモ欄
	Call SetField(rsYCustomer,objSt,"s19",		"AW" , lngRow)	'		//	参照表示
	Call SetField(rsYCustomer,objSt,"s20",		"AX" , lngRow)	'		//	更新日
	Call SetField(rsYCustomer,objSt,"s21",		"AY" , lngRow)	'		//	請求書合算
	rsYCustomer.UpdateBatch
	LoadYCustomerRow = lngRow
End Function

Function SetField(objRs,objSt,byVal strField,byVal strCol,byVal lngRow)
	Call Debug("SetField():" & lngRow & ":" & strField & ":" & objSt.Range(strCol & lngRow))
	dim	v
	v = RTrim(objSt.Range(strCol & lngRow))
	select case strField
	case "SDt"
		v = Replace(v,"/","")
	end select
	dim	dsize
	dsize = objRs.Fields(strField).DefinedSize
	v = Get_LeftB(v,dsize)
	Call Debug("SetField():" & lngRow & ":" & strField & "(" & dsize & ")=" & v)
	objRs.Fields(strField) = v
End Function
