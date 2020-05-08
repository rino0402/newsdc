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
	Wscript.Echo "JCS MF(Excel)データ変換"
	Wscript.Echo "JcsItem.vbs [option]"
	Wscript.Echo " /db:newsdc7"
	Wscript.Echo " /debug"
	Wscript.Echo "cscript JcsItem.vbs /debug"
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
	call LoadJcsItem()
	Main = 0
End Function

Function LoadJcsItem()
	Call Debug("LoadJcsItem()")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'テーブルオープン(Item)
	'-------------------------------------------------------------------
	dim	rsItem
	set rsItem = Wscript.CreateObject("ADODB.Recordset")
	rsItem.MaxRecords = 1
'	rsItem.CursorLocation = adUseServer
	dim	rsPCompo
	set rsPCompo = Wscript.CreateObject("ADODB.Recordset")
	rsPCompo.MaxRecords = 1
	dim	rsPCompoK
	set rsPCompoK = Wscript.CreateObject("ADODB.Recordset")
	rsPCompoK.MaxRecords = 1
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsJcsItem
	set rsJcsItem = objDb.Execute("select * from JcsItem")
	'-------------------------------------------------------------------
	'JcsItem → Item
	'-------------------------------------------------------------------
	do while not rsJcsItem.Eof
'		Call Debug(rsJcsItem.Fields("Pn"))
		Call DispMsg(rsJcsItem.Fields("Pn"))
		Call SetItem(objDb,rsItem,rsJcsItem)
		Call SetPCompo(objDb,rsPCompo,rsJcsItem)
		Call SetPCompoK(objDb,rsPCompoK,rsJcsItem)
		rsJcsItem.MoveNext
	loop
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsItem		= CloseRs(rsItem)
	set rsPCompo	= CloseRs(rsPCompo)
	set rsPCompoK	= CloseRs(rsPCompoK)
	set rsJcsItem	= CloseRs(rsJcsItem)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function SetPCompoK(objDb,rsPCompoK,rsJcsItem)
	Call Debug("SetPCompoK()")
	dim	strPn
	strPn = GetFieldValue(rsJcsItem,"MazdaPn")
	dim	strSql
	strSql = ""
	strSql = strSql & "delete from p_compo_k"
	strSql = strSql & " where SHIMUKE_CODE = '01'"
	strSql = strSql & "   and JGYOBU = 'J'"
	strSql = strSql & "   and NAIGAI = '1'"
	strSql = strSql & "   and HIN_GAI = '" & strPn & "'"
	strSql = strSql & "   and DATA_KBN = '2'"
	Call Debug("SetPCompoK():" & strSql)
	Call objDb.Execute(strSql)
	if rsPCompoK.state <> adStateClosed then
		Call Debug("SetPCompoK():rsPCompoK.Close")
		rsPCompoK.Close
	end if
	rsPCompoK.Open "p_compo_k", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
	dim	i
	dim	iSeq
	iSeq = 0
	for i = 1 to 2
		dim	strKo
		dim	strQty
		if i = 1 then
			strKo = GetFieldValue(rsJcsItem,"GPn")
			strQty = GetFieldValue(rsJcsItem,"GQty")
		else
			strKo = GetFieldValue(rsJcsItem,"LastPn")
			strQty = GetFieldValue(rsJcsItem,"LastQty")
		end if
		if strKo <> "" then
			iSeq = iSeq + 1
			rsPCompoK.AddNew
			Call SetField(rsPCompoK,"SHIMUKE_CODE"	,"01")
			Call SetField(rsPCompoK,"JGYOBU"		,"J")
			Call SetField(rsPCompoK,"NAIGAI"		,"1")
			Call SetField(rsPCompoK,"HIN_GAI"		,strPn)
			Call SetField(rsPCompoK,"DATA_KBN"		,"2")
			Call SetField(rsPCompoK,"SEQNO"			,"0" & iSeq & "0")
			'KO_SYUBETSU	KO_JGYOBU	KO_NAIGAI	KO_HIN_GAI	KO_QTY	KO_BIKOU	CLASS_CODE
			Call SetField(rsPCompoK,"KO_SYUBETSU"		,"")
			Call SetField(rsPCompoK,"KO_JGYOBU"			,"S")
			Call SetField(rsPCompoK,"KO_NAIGAI"			,"1")
			Call SetField(rsPCompoK,"KO_HIN_GAI"		,strKo)
			Call SetField(rsPCompoK,"KO_QTY"			,strQty)
			Call SetField(rsPCompoK,"KO_BIKOU"			,"")
			Call SetField(rsPCompoK,"CLASS_CODE"		,"")
			Call SetField(rsPCompoK,"UPD_TANTO"			,"JcsItem"								)
			Call SetField(rsPCompoK,"UPD_DATETIME"		,Left(GetDateTime(now()),12)			)
			rsPCompoK.UpdateBatch
		end if
	next
End Function

Function SetPCompo(objDb,rsPCompo,rsJcsItem)
	Call Debug("SetPCompo()")
	if rsPCompo.state <> adStateClosed then
		Call Debug("SetPCompo():rsPCompo.Close")
		rsPCompo.Close
	end if
	dim	strPn
	strPn = GetFieldValue(rsJcsItem,"MazdaPn")
	dim	strSql
	strSql = ""
	strSql = strSql & "select * from p_compo"
	strSql = strSql & " where SHIMUKE_CODE = '01'"
	strSql = strSql & "   and JGYOBU = 'J'"
	strSql = strSql & "   and NAIGAI = '1'"
	strSql = strSql & "   and HIN_GAI = '" & strPn & "'"
	strSql = strSql & "   and DATA_KBN = '0'"
	Call Debug("SetPCompo():" & strSql)
	rsPCompo.Open strSql, objDb, adOpenForwardOnly, adLockBatchOptimistic
	if rsPCompo.Eof then
		Call Debug("SetPCompo():rsPCompo.Eof")
		if rsPCompo.state <> adStateClosed then
			Call Debug("SetPCompo():rsPCompo.Close")
			rsPCompo.Close
		end if
		rsPCompo.Open "p_compo", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
		rsPCompo.AddNew
		Call SetField(rsPCompo,"SHIMUKE_CODE"	,"01")
		Call SetField(rsPCompo,"JGYOBU"			,"J")
		Call SetField(rsPCompo,"NAIGAI"			,"1")
		Call SetField(rsPCompo,"HIN_GAI"		,strPn)
		Call SetField(rsPCompo,"DATA_KBN"		,"0")
		Call SetField(rsPCompo,"SEQNO"			,"000")
	end if
	'SEQNO	CLASS_CODE	BIKOU	F_CLASS_CODE	N_CLASS_CODE	FILLER	UPD_TANTO	UPD_DATETIME
	dim	strBikou
	strBikou = ""
	strBikou = strBikou & GetFieldValue(rsJcsItem,"CheckConf") & vbCrLf
	strBikou = strBikou & GetFieldValue(rsJcsItem,"Sagyo")
	Call SetField(rsPCompo,"BIKOU"				,strBikou								)
	Call SetField(rsPCompo,"CLASS_CODE"			,GetFieldValue(rsJcsItem,"SSpec")		)
	Call SetField(rsPCompo,"F_CLASS_CODE"		,GetFieldValue(rsJcsItem,"SType")		)
	Call SetField(rsPCompo,"N_CLASS_CODE"		,""										)
	Call SetField(rsPCompo,"UPD_TANTO"			,"JcsItem"								)
	Call SetField(rsPCompo,"UPD_DATETIME"		,Left(GetDateTime(now()),12)			)
	rsPCompo.UpdateBatch
End Function

Function SetItem(objDb,rsItem,rsJcsItem)
	Call Debug("SetItem()")
	if rsItem.state <> adStateClosed then
		Call Debug("SetItem():rsItem.Close")
		rsItem.Close
	end if
	dim	strPn
	strPn = GetFieldValue(rsJcsItem,"MazdaPn")
	dim	strSql
	strSql = ""
	strSql = strSql & "select * from Item"
	strSql = strSql & " where JGYOBU = 'J'"
	strSql = strSql & "   and NAIGAI = '1'"
	strSql = strSql & "   and HIN_GAI = '" & strPn & "'"
	Call Debug("SetItem():" & strSql)
	rsItem.Open strSql, objDb, adOpenForwardOnly, adLockBatchOptimistic
	if rsItem.Eof then
		Call Debug("SetItem():rsItem.Eof")
		if rsItem.state <> adStateClosed then
			Call Debug("SetItem():rsItem.Close")
			rsItem.Close
		end if
		rsItem.Open "Item", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
		rsItem.AddNew
		Call SetField(rsItem,"JGYOBU"	,"J")
		Call SetField(rsItem,"NAIGAI"	,"1")
		Call SetField(rsItem,"HIN_GAI"	,strPn)
	end if
	Call SetField(rsItem,"L_HIN_NAME_E"		,GetFieldValue(rsJcsItem,"NameE")		)
	Call SetField(rsItem,"HIN_NAI"			,GetFieldValue(rsJcsItem,"Pn")			)
	Call SetField(rsItem,"HIN_NAME"			,GetFieldValue(rsJcsItem,"NameJ")		)
	Call SetField(rsItem,"L_BIKOU"			,GetFieldValue(rsJcsItem,"Color")		)
	Call SetField(rsItem,"L_KAISHA_CODE"	,GetFieldValue(rsJcsItem,"LabelType")	)
	Call SetField(rsItem,"L_MAISU"			,GetFieldValue(rsJcsItem,"LabelCut")	)
	Call SetField(rsItem,"BIKOU_TANA"		,GetFieldValue(rsJcsItem,"Location")	)
	Call SetField(rsItem,"K_KEITAI"			,GetFieldValue(rsJcsItem,"SType")		)
	Call SetField(rsItem,"SHIYOU_NO"		,GetFieldValue(rsJcsItem,"SSpec")		)
'	Call SetField(rsItem,""					,GetFieldValue(rsJcsItem,"GPn")			)
'	Call SetField(rsItem,""					,GetFieldValue(rsJcsItem,"GQty")		)
'	Call SetField(rsItem,""					,GetFieldValue(rsJcsItem,"LastPn")		)
	Call SetField(rsItem,"GAISO_IRI_QTY"	,GetFieldValue(rsJcsItem,"LastQty")		)
	Call SetField(rsItem,"L_SAGYO_SHIJI_1"	,GetFieldValue(rsJcsItem,"CheckConf")	)
	Call SetField(rsItem,"GLICS1_TANA"		,GetFieldValue(rsJcsItem,"Location1")	)
	Call SetField(rsItem,"GLICS2_TANA"		,GetFieldValue(rsJcsItem,"Location2")	)
'	Call SetField(rsItem,""					,GetFieldValue(rsJcsItem,"ShareNum")	)
'	Call SetField(rsItem,""					,GetFieldValue(rsJcsItem,"ShareNo")		)
	Call SetField(rsItem,"Ins_DateTime"		,GetFieldValue(rsJcsItem,"AlterDate")	)
	Call SetField(rsItem,"INS_TANTO"		,GetFieldValue(rsJcsItem,"AlterPerson")	)
	Call SetField(rsItem,"LAST_SYU_DT"		,GetFieldValue(rsJcsItem,"LastShipDate"))
	Call SetField(rsItem,"G_LAST_SYUKA_QTY"	,GetFieldValue(rsJcsItem,"LastShipQty")	)
''	Call SetField(rsItem,""					,GetFieldValue(rsJcsItem,"LastShipPn")	)
'	Call SetField(rsItem,""					,GetFieldValue(rsJcsItem,"CoStockDate")	)
	Call SetField(rsItem,"G_ZEN_ZAIKO_QTY"	,GetFieldValue(rsJcsItem,"CoStockQty")	)
	Call SetField(rsItem,"UPD_TANTO"		,"JcsItem"								)
	Call SetField(rsItem,"UPD_DateTime"		,Left(GetDateTime(now()),12)			)
	rsItem.UpdateBatch
End Function

Function SetField(objRs,strField,byVal v)
	Call Debug("SetField():" & strField & ":" & v)
	select case strField
	case "LAST_SYU_DT"
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

