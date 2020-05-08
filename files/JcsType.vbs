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
	Wscript.Echo "JcsType.vbs [option]"
	Wscript.Echo " /db:newsdc7"
	Wscript.Echo " /debug"
	Wscript.Echo "cscript JcsType.vbs /debug"
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
	call LoadJcsType()
	Main = 0
End Function

Function LoadJcsType()
	Call Debug("LoadJcsType()")
	'-------------------------------------------------------------------
	'データベースの準備
	'-------------------------------------------------------------------
	dim	objDb
	Set objDb = OpenAdodb(GetOption("db","newsdc7"))
	'-------------------------------------------------------------------
	'テーブルオープン(Item)
	'-------------------------------------------------------------------
	dim	rsPCompoK
	set rsPCompoK = Wscript.CreateObject("ADODB.Recordset")
	rsPCompoK.MaxRecords = 1
	'-------------------------------------------------------------------
	'テーブルオープン
	'-------------------------------------------------------------------
	dim	rsJcsType
	set rsJcsType = objDb.Execute("select * from JcsType")
	'-------------------------------------------------------------------
	'JcsItem → Item
	'-------------------------------------------------------------------
	do while not rsJcsType.Eof
		Call Debug(rsJcsType.Fields("SType"))
		Call SetPCompoK(objDb,rsJcsType)
		rsJcsType.MoveNext
	loop
	'-------------------------------------------------------------------
	'テーブルのクローズ
	'-------------------------------------------------------------------
	set rsPCompoK	= CloseRs(rsPCompoK)
	set rsJcsType	= CloseRs(rsJcsType)
	'-------------------------------------------------------------------
	'データベースのクローズ
	'-------------------------------------------------------------------
	set objDb = CloseAdodb(objDb)
End Function

Function SetPCompoK(objDb,rsJcsItem)
	Call Debug("SetPCompoK()")
	'XlsRow	SType SPn1 SQty1 SPn2 SQty2 SPn3 SQty3 SPn4 SQty4
	'PUnit	MCP1	MCP2	PCP1	PCP2	OCP1	OCP2	AlterDate
	dim	strSType
	strSType	= GetFieldValue(rsJcsItem,"SType")
	dim	strSql
	strSql = ""
	strSql = strSql & "delete p_compo_k"
	strSql = strSql & " where SHIMUKE_CODE='01'"
	strSql = strSql & "   and DATA_KBN='1'"
	strSql = strSql & "   and CLASS_CODE='" & strSType & "'"
	Call Debug(strSql)
	objDb.Execute(strSql)
	dim	i
	dim	iSeq
	iSeq = 0
	for i = 1 to 4
		dim	strPn
		strPn		= GetFieldValue(rsJcsItem,"SPn" & i)
		dim	strQty
		strQty		= GetFieldValue(rsJcsItem,"SQty" & i)
		Call Debug(strSType & " " & strPn & " " & strQty)
		if strPn <> "" then
			iSeq = iSeq + 1
			strSql = ""
			strSql = strSql & "insert into p_compo_k ("
			strSql = strSql & " SHIMUKE_CODE"
			strSql = strSql & ",JGYOBU"
			strSql = strSql & ",NAIGAI"
			strSql = strSql & ",HIN_GAI"
			strSql = strSql & ",DATA_KBN"
			strSql = strSql & ",SEQNO"
			strSql = strSql & ",KO_SYUBETSU"
			strSql = strSql & ",KO_JGYOBU"
			strSql = strSql & ",KO_NAIGAI"
			strSql = strSql & ",KO_HIN_GAI"
			strSql = strSql & ",KO_QTY"
			strSql = strSql & ",KO_BIKOU"
			strSql = strSql & ",CLASS_CODE"
			strSql = strSql & ",FILLER"
			strSql = strSql & ",UPD_TANTO"
			strSql = strSql & ",UPD_DATETIME"
			strSql = strSql & ")"
			strSql = strSql & " select"
			strSql = strSql & " '01'"					'SHIMUKE_COD
			strSql = strSql & ",JGYOBU"					'JGYOBU"	
			strSql = strSql & ",NAIGAI"                 'NAIGAI"    
			strSql = strSql & ",HIN_GAI"                'HIN_GAI"   
			strSql = strSql & ",'1'"                    'DATA_KBN"  
			strSql = strSql & ",'0" & iSeq & "0'"       'SEQNO"     
			strSql = strSql & ",''"                     'KO_SYUBETSU
			strSql = strSql & ",'S'"                    'KO_JGYOBU" 
			strSql = strSql & ",'1'"                    'KO_NAIGAI" 
			strSql = strSql & ",'" & strPn & "'"        'KO_HIN_GAI"
			strSql = strSql & ",'" & strQty & "'"       'KO_QTY"    
			strSql = strSql & ",''"                     'KO_BIKOU"  
			strSql = strSql & ",'" & strSType & "'"     'CLASS_CODE"
			strSql = strSql & ",''"                     'FILLER"    
			strSql = strSql & ",'JType'"  		        'UPD_TANTO" 
			strSql = strSql & ",left(replace(replace(replace(convert(CURRENT_TIMESTAMP(),SQL_CHAR),':',''),'-',''),' ',''),14)"     'UPD_DATETIME"
			strSql = strSql & " from Item"                                                                                          
			strSql = strSql & " where JGYOBU='J'"
			strSql = strSql & "   and NAIGAI='1'"
			strSql = strSql & "   and K_KEITAI='" & strSType & "'"
			Call Debug(strSql)
			objDb.Execute(strSql)
		end if
	next
End Function
