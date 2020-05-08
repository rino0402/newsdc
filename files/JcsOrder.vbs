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
Call Include("file.vbs")
Call Include("get_b.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "JcsOrder.vbs [option]"
	Wscript.Echo " /db:newsdc7"
	Wscript.Echo " /debug"
	Wscript.Echo "cscript JcsOrder.vbs /debug"
End Sub
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	strArg
	dim	strFilename
	dim	strSheetname
	strFilename = ""
	strSheetname = ""
	'名前無しオプションチェック
	for each strArg in WScript.Arguments.UnNamed
		usage()
		exit function
	next
	'名前付きオプションチェック
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "debug"
		case "?"
			usage()
			exit function
		case else
			usage()
			exit function
		end select
	next
	call LoadJcsOrder()
	Main = 0
End Function

Function LoadJcsOrder()
	Call Debug("LoadJcsOrder()")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'テーブルオープン(Order)
	'-------------------------------------------------------------------
	dim	strSql
	strSql = ""
	strSql = strSql & "delete from PLN_S_YOTEI"
	Call Debug("Execute():" & strSql)
	Call objDb.Execute(strSql)

	dim	rsOrder
	set rsOrder = Wscript.CreateObject("ADODB.Recordset")
	rsOrder.MaxRecords = 1
	rsOrder.Open "PLN_S_YOTEI", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsJcsOrder
	set rsJcsOrder = objDb.Execute("select * from JcsOrder")

	'-------------------------------------------------------------------
	'JcsItem → Item
	'-------------------------------------------------------------------
	do while not rsJcsOrder.Eof
'		Call Debug(rsJcsItem.Fields("Pn"))
		Call DispMsg(GetFieldValue(rsJcsOrder,"XlsRow") & " " & GetFieldValue(rsJcsOrder,"Pn"))
		Call SetOrder(objDb,rsOrder,rsJcsOrder)
		rsJcsOrder.MoveNext
	loop
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsJcsOrder	= CloseRs(rsJcsOrder)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function SetOrder(objDb,rsOrder,rsJcsOrder)
	Call Debug("SetOrder()")
	rsOrder.AddNew
	Call SetField(rsOrder,"TORIKOMI_DT"	,Left( GetDateTime( Now()) ,8 ) )
	Call SetField(rsOrder,"JGYOBU"		,"J")
	Call SetField(rsOrder,"NAIGAI"		,"1")
	Call SetField(rsOrder,"HIN_GAI"		,GetFieldValue(rsJcsOrder,"MazdaPn"))
	Call SetField(rsOrder,"YOTEI_DT"	,GetFieldValue(rsJcsOrder,"DlvDt"))
	Call SetField(rsOrder,"YOTEI_QTY"	,GetFieldValue(rsJcsOrder,"Qty"))
	Call SetField(rsOrder,"S_KOUSU"					,"0")
	Call SetField(rsOrder,"S_JIKAN"					,"0")
	Call SetField(rsOrder,"TOTAL_CNT"				,"0")
	Call SetField(rsOrder,"TOTAL_AVE_CNT"			,"0")
	Call SetField(rsOrder,"S_SYUKA_QTY1"			,"0")
	Call SetField(rsOrder,"S_SYUKA_CNT1"			,"0")
	Call SetField(rsOrder,"S_AVE_SYUKA_QTY1"		,"0")
	Call SetField(rsOrder,"S_AVE_SYUKA_CNT1"		,"0")
	Call SetField(rsOrder,"S_SYUKA_QTY2"			,"0")
	Call SetField(rsOrder,"S_SYUKA_CNT2"			,"0")
	Call SetField(rsOrder,"S_AVE_SYUKA_QTY2"		,"0")
	Call SetField(rsOrder,"S_AVE_SYUKA_CNT2"		,"0")
	Call SetField(rsOrder,"Z_QTY_MI"				,"0")
	Call SetField(rsOrder,"Z_QTY_S"					,"0")
	Call SetField(rsOrder,"JIZEN"					,"0")
	Call SetField(rsOrder,"NYUKA_YOTEI_QTY"			,"0")
	Call SetField(rsOrder,"S_KOUSU_X"				,"0")
	Call SetField(rsOrder,"S_JIKAN_X"				,"0")
	Call SetField(rsOrder,"YOTEI_QTY_X"				,"0")
	Call SetField(rsOrder,"GAISO_MAISU"				,"0")
	Call SetField(rsOrder,"BETU1_QTY"				,"0")
	Call SetField(rsOrder,"BETU2_QTY"				,"0")
	Call SetField(rsOrder,"JIZEN_NEEDS_QTY"			,"0")
	Call SetField(rsOrder,"JITU_KOUSU"				,"0")
	Call SetField(rsOrder,"SAGYOU_KOUSU"			,"0")
	Call SetField(rsOrder,"INP_NYUKA_YOTEI_QTY"		,"0")
	Call SetField(rsOrder,"S_LIST_DateTime"		,"")
	Call SetField(rsOrder,"SASIZU_DateTime"		,"")
	Call SetField(rsOrder,"S_KAN_DateTime"		,"")
	Call SetField(rsOrder,"TENKAI_DateTime"		,"")
	Call SetField(rsOrder,"INS_TANTO"			,"Order")
	Call SetField(rsOrder,"Ins_DateTime"		,GetDateTime( Now() ) )
	Call SetField(rsOrder,"KEY_NO"		,GetFieldValue(rsJcsOrder,"XlsRow"))
	rsOrder.UpdateBatch
End Function


Function SetField(objRs,strField,byVal v)
	Call Debug("SetField():" & strField & ":" & v)
	select case strField
	case "LAST_SYU_DT","YOTEI_DT"
		v = Replace(v,"/","")
	case "SDt","DDt"	'//  I:J 売上日 '//  K:L 納入日
		v = v & "/" & objSt.Range(strCol & lngRow).Offset(0,1)
		if isDate(v) then
			v = CDate(v)
			v = Replace(v,"/","")
		else
			v = ""
		end if
	case "Amount","AmountEHN"
		if isNumeric(v) <> True then
			v = 0
		end if
		v = CCur(v)
	end select
	v = Get_LeftB(v,objRs.Fields(strField).DefinedSize)
	Call Debug("SetField():" & strField & ":" & v)
	objRs.Fields(strField) = v
End Function

