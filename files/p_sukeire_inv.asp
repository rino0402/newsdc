<%@ LANGUAGE="VBScript" CODEPAGE=65001 %>
<%
Option Explicit
Session.CodePage=65001
Response.Expires = -1
Response.AddHeader "Cache-Control", "No-Cache"
Response.AddHeader "Pragma", "No-Cache"
Response.ContentType = "application/json; charset=UTF-8"

Const	adStateClosed		= 0	'オブジェクトが閉じている
Const	adStateOpen			= 1 'オブジェクトが開いている
Const	adStateConnecting	= 2 'オブジェクトが接続している
Const	adStateExecuting	= 4 'オブジェクトがコマンドを実行中
Const	adStateFetching		= 8 'オブジェクトの行が取得されている

Call Main()

Private Function Main()
	dim	objPSInv
	set objPSInv = new PSInv
	objPSInv.Run
	set objPSInv = nothing
End Function

Class PSInv
	Private	optDebug
	Private	strDt1
	Private	strDt2
	Private	strShimuke
	Private	strSTanto
	Private	strCheck
	Private	strDBName
	Private	objDB
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		optDebug = False
'		optDebug = True		' デバッグモード
		Debug ".Class_Initialize()"
		strDBName	= GetDbName()
		set objDB	= nothing
		set	objRs	= Nothing

		strDt1		= Request.QueryString("UKEIRE_DT1")
		strDt2		= Request.QueryString("UKEIRE_DT2")
		strShimuke	= Request.QueryString("SHIMUKE_CODE")
		strCheck	= Request.QueryString("CHECK")
		LogOutput "init"
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB	= nothing
		LogOutput "term"
    End Sub
	'-----------------------------------------------------------------------
	'Run
	'-----------------------------------------------------------------------
	Public Function	Run()
		Debug ".Run()"
		OpenDb
		List
		ListJson
		set	objRs = Nothing
		CloseDB
	End Function
	'-------------------------------------------------------------------
	'ListJson() 結果Json出力
	'-------------------------------------------------------------------
    Private Function ListJson()
		Debug ".ListJson()"
		JsonSet "{","",""
		JsonSet " ","status","success"
		JsonSet ",","dbName",strDbName
		JsonSet ",","UKEIRE_DT1",strDt1
		JsonSet ",","UKEIRE_DT2",strDt2
		JsonSet ",","SHIMUKE_CODE"	,strShimuke
		JsonSet ",","S_TANTO",strSTanto
		JsonSet ",","CHECK",strCheck
		JsonSet ",","SQL",strSql
		JsonSet ",","LIST","["
		dim	strCma
		strCma = ""
		if not objRs is nothing then
			do while objRs.Eof = False
				JsonSet strCma & "{","",""
				dim	objF
				dim	strCma2
				strCma2 = " "
				for each objF in objRs.Fields
					JsonSet strCma2,objF.Name,GetField(objF.Name)
					strCma2 = ","
				next
				JsonSet "}","",""
				objRs.MoveNext
				strCma = ","
			loop
		end if
		JsonSet "]","",""
		JsonSet "}","",""
	End Function
	'-------------------------------------------------------------------
	'GetField()
	'-------------------------------------------------------------------
    Private Function GetField(byVal f)
		GetField = RTrim(objRs.Fields(f))
	End Function
	'-------------------------------------------------------------------
	'JsonSet()
	'-------------------------------------------------------------------
    Private Function JsonSet(byVal d,byVal t,byVal v)
		if t = "" then
			JsonWrite d
			exit function
		end if
		if v = "[" then
			JsonWrite d & """" & t & """:" & v
			exit function
		end if
		JsonWrite d & """" & t & """:""" & JsonValue(v) & """"
	End Function
	'-------------------------------------------------------------------
	'List()
	'-------------------------------------------------------------------
    Private Function List()
		Debug ".List()"
	    AddSql "select"
	    AddSql " u.SHIJI_NO      SNo"
	    AddSql ",u.SEQNO         QNo"
	    AddSql ",''              KNo"
	    AddSql ",CANCEL_F        C"
	    AddSql ",u.SHIMUKE_CODE  Shimuke"
	    AddSql ",u.UKEIRE_DT       UkeDt"
	    AddSql ",o.HIN_GAI         Pn"
	    AddSql ",o.S_CLASS_CODE    SClass"
	    AddSql ",o.F_CLASS_CODE    FClass"
	    AddSql ",o.N_CLASS_CODE    NClass"
	    AddSql ",o.SHIJI_QTY       SQty"
	    AddSql ",convert(u.UKEIRE_QTY,sql_decimal)    UQty"
	    AddSql ",''              KOPn"
	    AddSql ",''              KOQty"
	    AddSql ",o.S_TANTO       STanto"
	    AddSql ",o.JGYOBU        JGyobu"
	    AddSql ",o.NAIGAI        Naigai"
	    AddSql ",u.SHIMUKE_CODE  SHIMUKE_CODE"
	    AddSql ",i.HIN_NAME          HIN_NAME"
	    AddSql ",convert(i.S_KOUSU_BAIKA,sql_decimal)	KrT"
	    AddSql ",convert(u.UKEIRE_QTY,sql_decimal)"
	    AddSql "*convert(i.S_KOUSU_BAIKA,sql_decimal)	KrK"
	    AddSql ",convert(i.S_SHIZAI_BAIKA,sql_decimal)	SzT"
	    AddSql ",convert(u.UKEIRE_QTY,sql_decimal)"
	    AddSql "*convert(i.S_SHIZAI_BAIKA,sql_decimal)	SzK"
	    AddSql ",Round(convert(i.S_GAISO_TANKA,sql_decimal),2)	GsT"
	    AddSql ",convert(u.UKEIRE_QTY,sql_decimal)"
	    AddSql "*convert(i.S_GAISO_TANKA,sql_decimal)	GsK"
	    AddSql ",convert(i.S_PPSC_KAKO_KOSU,sql_decimal)	KkT"
	    AddSql ",convert(u.UKEIRE_QTY,sql_decimal)"
	    AddSql "*convert(i.S_PPSC_KAKO_KOSU,sql_decimal)	KkK"

	    AddSql ",convert(i.S_BU_KAKO_KOSU,sql_decimal)	BuT"
	    AddSql ",convert(u.UKEIRE_QTY,sql_decimal)"
	    AddSql "*convert(i.S_BU_KAKO_KOSU,sql_decimal)	BuK"

	    AddSql ",i.S_PPSC_KAKO_KOSU  S_PPSC_KAKO_KOSU"
	    AddSql ",i.S_BU_KAKO_KOSU    S_BU_KAKO_KOSU"
	    AddSql ",i.L_KISHU1          DModel"
	    AddSql ",i.BEF_S_KOUSU_BAIKA     BEF_S_KOUSU_BAIKA"
	    AddSql ",i.BEF_S_SHIZAI_BAIKA    BEF_S_SHIZAI_BAIKA"
	    AddSql ",i.BEF_S_GAISO_TANKA     BEF_S_GAISO_TANKA"
	    AddSql ",i.BEF_S_PPSC_KAKO_KOSU  BEF_S_PPSC_KAKO_KOSU"
	    AddSql ",i.BEF_S_BU_KAKO_KOSU    BEF_S_BU_KAKO_KOSU"
	    AddSql ",i.TANKA_KIRIKAE_DT      TANKA_KIRIKAE_DT"
		if strCheck <> "" then
		    AddSql ",convert(i.MAIN_KOUTEI_02,sql_decimal)	sKoso0"		'//個装作業(log)
		    AddSql ",c.sKoso"		'//個装作業(log)
		    AddSql ",c.sSyugo"		'//梱包作業(log)
		    AddSql ",c.sDokon"		'//同梱件数(log)
		    AddSql ",c.sKako"		'//加工作業(log)
		    AddSql ",round("
		    AddSql " CEILING(("
			AddSql " convert(i.BEF_KOUTEI_10,sql_decimal)"	'//
			AddSql "+round(("								'//作業時間:計算
			AddSql "+(convert(i.SEI_LABEL_QTY,sql_decimal) * 4)"	'//ラベル貼り
'			AddSql "+convert(i.MAIN_KOUTEI_01,sql_decimal)"	'//ラベル貼り
			AddSql "+c.sKoso"								'//個装作業
'			AddSql "+convert(i.MAIN_KOUTEI_02,sql_decimal)"	//個装作業
			AddSql "+(c.sDokon * 4)"						'//同梱作業
'			AddSql "+convert(i.MAIN_KOUTEI_03,sql_decimal)"	'//同梱作業
			AddSql "+c.sKako"								'//加工作業
'			AddSql "+convert(i.MAIN_KOUTEI_04,sql_decimal)"	'//加工作業
			AddSql "+c.sSyugo"								'//集合梱包
'			AddSql "+convert(i.MAIN_KOUTEI_05,sql_decimal)"	'//集合梱包
'			AddSql "+convert(i.MAIN_KOUTEI_06,sql_decimal)"	//
'			AddSql "+convert(i.MAIN_KOUTEI_07,sql_decimal)"	//
'			AddSql "+convert(i.MAIN_KOUTEI_08,sql_decimal)"	//
'			AddSql "+convert(i.MAIN_KOUTEI_09,sql_decimal)"	//
			AddSql ")*1.15,0)"								'//
'			AddSql "+convert(i.MAIN_KOUTEI_10,sql_decimal)"	//作業時間(秒)
			AddSql "+convert(i.AFT_KOUTEI_10,sql_decimal)"	'//
			AddSql "+convert(i.PLUS_KOUSU,sql_decimal)"		'//
		    AddSql ")/6)/10"
'			AddSql " convert(i.S_KOUSU,sql_decimal)"	//工数(分)
		    AddSql "*convert(i.SEI_RATE,sql_decimal)"	'//分レート"
		    AddSql ",2)"
		    AddSql " cKrT"								'//工料
		    AddSql ",ifNull(c.SzT,0)	cSzT"			'//箱代
		    AddSql ",ifNull(Round(c.GsT,2),0)	cGsT"	'//外装
		    AddSql ",ifNull(Round(c.KkT,2),0)	cKkT"	'//加工
		    AddSql ",ifNull(Round(c.BuT,2),0)	cBuT"	'//Bu加工
		end if
	    AddSql " From P_SUKEIRE u"
	    AddSql " left outer join P_SSHIJI_O o on (u.shiji_no=o.shiji_no)"
	    AddSql " left outer join Item i on (o.JGYOBU = i.JGYOBU and o.NAIGAI = i.NAIGAI and o.HIN_GAI = i.HIN_GAI)"
		if strCheck <> "" then
			AddSql " left outer join ("
			AddSql " select"
			AddSql " k.SHIMUKE_CODE"
			AddSql ",k.JGYOBU"
			AddSql ",k.NAIGAI"
			AddSql ",k.HIN_GAI"
'			AddSql ",sum(if(k.DATA_KBN='1' or k.KO_SYUBETSU in ('03','90'),convert(s.G_ST_URITAN,sql_decimal) * convert(k.KO_QTY,SQL_DECIMAL),0)) SzT"
			AddSql ",sum(if((k.DATA_KBN='1' or k.KO_SYUBETSU in ('03','90')) and s.SEI_KBN not in ('1','2'),convert(s.G_ST_URITAN,sql_decimal) * convert(k.KO_QTY,SQL_DECIMAL),0)) SzT"
			AddSql ",sum(if(k.DATA_KBN='1' or k.KO_SYUBETSU in ('03','90'),convert(s.S_KOUSU,sql_decimal) * convert(k.KO_QTY,SQL_DECIMAL),0)) sKoso"
			AddSql ",sum(if(k.DATA_KBN='1' or k.KO_SYUBETSU in ('03','90'),convert(s.SEI_SYU_KON,sql_decimal),0)) sSyugo"
			AddSql ",sum(if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('01'),convert(k.KO_QTY,SQL_DECIMAL),0)) sDokon"
			AddSql ",sum(if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('05'),convert(s.S_KOUSU,sql_decimal) * convert(k.KO_QTY,SQL_DECIMAL),0)) sKako"
'　商品化単価見積
'　①Excel【副資材費】外装の算出式変更
'　　数量＝１÷員数(入数) 小数3桁未満四捨五入
'　　　　　<例> 40入→1/40=0.025
'　　　　　　　 30入→1/30=0.033
'　　　　　　　 60入→1/60=0.017
'　　単価＝販売＠
'　　金額＝数量×単価 小数2桁未満四捨五入
'　　　　　<例> 0.025×232= 5.80
'　　　　　　　 0.033×232= 7.66
'　　　　　　　 0.017×232= 3.94
'  ②見積画面 外装の算出方法変更
'　　外装＝( 1/員数(小数桁未満四捨五入)×販売＠ )小数2桁未満四捨五入
'　③種別：08 付帯作業 を追加し、作業時間として集計、Excel【副資材費】には載せない。
			AddSql ",sum("
			AddSql "if((k.DATA_KBN='2' or k.KO_SYUBETSU in ('91')) and s.SEI_KBN not in ('1','2')"
			AddSql ",Gaiso(convert(k.KO_QTY,SQL_DECIMAL),convert(s.G_ST_URITAN,sql_decimal))"
			AddSql ",0)"		'if
			AddSql ") GsT"		'sum
			'加工
			AddSql ",sum("
			AddSql "if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('06','04')"
			AddSql ",Kako(convert(k.KO_QTY,SQL_DECIMAL),convert(s.S_KOUSU,sql_decimal))"
			AddSql ",0)"		'if
			AddSql ") KkT"		'sum
			'Bu加工
			AddSql ",sum("
			AddSql "if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('07')"
			AddSql ",Kako(convert(k.KO_QTY,SQL_DECIMAL),convert(s.S_KOUSU,sql_decimal))"
			AddSql ",0)"		'if
			AddSql ") BuT"		'sum
			'from
			AddSql " From P_COMPO_K k"
			AddSql " left outer join ITEM s"
			AddSql " on (k.KO_JGYOBU =s.JGYOBU and k.KO_NAIGAI =s.NAIGAI and k.KO_HIN_GAI =s.HIN_GAI)"
'			AddSql " where k.DATA_KBN='1' or (k.DATA_KBN='3' and k.KO_SYUBETSU in ('03','90'))"
'			AddSql " where (k.DATA_KBN='3' and k.KO_SYUBETSU in ('03','90'))"
'			AddSql "    or k.DATA_KBN='1'"
			'where
			AddSql " where k.DATA_KBN in ('1','2','3')"
'検索速度をあげる為---
			AddSql " and k.HIN_GAI in ("
			AddSql 		" select distinct HIN_GAI from P_SSHIJI_O"
			AddSql 		" where shiji_no in ("
			AddSql 			" select distinct shiji_no from P_SUKEIRE"
			dim	strW
			strW = ""
			strW = MakeWhere(strW,"UKEIRE_DT"		,strDt1 & "," & strDt2)
			strW = MakeWhere(strW,"SHIMUKE_CODE"	,strShimuke)
			AddSql 			strW
			AddSql 			")"
			AddSql 		")"
'検索速度をあげる為---
			'group
			AddSql " group by"
			AddSql " k.SHIMUKE_CODE"
			AddSql ",k.JGYOBU"
			AddSql ",k.NAIGAI"
			AddSql ",k.HIN_GAI"
			AddSql ") c on (o.SHIMUKE_CODE = c.SHIMUKE_CODE and o.JGYOBU = c.JGYOBU and o.NAIGAI = c.NAIGAI and o.HIN_GAI = c.HIN_GAI)"
		end if
		AddSql " " & GetWhere()
	    AddSql " order by u.UKEIRE_DT,o.HIN_GAI,u.SHIJI_NO,u.SEQNO"
		set objRs = GetRs()
	End Function
	Private	objRs
	'-------------------------------------------------------------------
	'GetWhere():Sql
	'-------------------------------------------------------------------
	Public Function GetWhere()
		Debug ".GetWhere()"
		dim	strWhere
		strWhere = ""
		strWhere = MakeWhere(strWhere,"u.UKEIRE_DT"		,strDt1 & "," & strDt2)
		strWhere = MakeWhere(strWhere,"u.SHIMUKE_CODE"	,strShimuke)
		strWhere = MakeWhere(strWhere,"i.S_TANTO"		,strSTanto)
		select case strCheck
		case "2"
			strWhere = MakeWhere(strWhere,""	,"(KrT <> cKrT or SzT <> cSzT or GsT <> cGsT or KkT <> cKkT or BuT <> cBuT)")
		case "3"
			strWhere = MakeWhere(strWhere,""	,"(KrT <> cKrT)")
		case "4"
			strWhere = MakeWhere(strWhere,""	,"(SzT <> cSzT)")
		case "5"
			strWhere = MakeWhere(strWhere,""	,"(GsT <> cGsT)")
		case "6"
			strWhere = MakeWhere(strWhere,""	,"(KkT <> cKkT)")
		case "7"
			strWhere = MakeWhere(strWhere,""	,"(BuT <> cBuT)")
		end select
		GetWhere = strWhere
    End Function
	'-------------------------------------------------------------------
	'MakeWhere():Sql
	'-------------------------------------------------------------------
	Public Function MakeWhere(byVal strWhere,byVal f,byVal v)
		Debug ".MakeWhere():" & v
		MakeWhere = strWhere
		if v = "" then
			exit function
		end if
		dim	strCmp
		strCmp = " = "
		if Left(v,1) = ">" then
			strCmp = " >= "
			v = Right(v,Len(v)-1)
		end if
		if strWhere = "" then
			strWhere = " where "
		else
			strWhere = strWhere & " and "
		end if
		select case f
		case "u.UKEIRE_DT","UKEIRE_DT"
			if Right(v,1) <> "," then
				strCmp = " between "
				v = "'" & Replace(v,",","' and '") & "'"
			else
				v = Left(v,len(v)-1)
			end if
		case ""
			strCmp = ""
		case else
			v = "'" & v & "'"
		end select
		strWhere = strWhere & f & strCmp & v
		MakeWhere = strWhere
    End Function
	'-------------------------------------------------------------------
	'GetRs():Sql実行 レコードセットを返す
	'-------------------------------------------------------------------
	Public Function GetRs()
		Debug ".GetRs():" & strSql
		on error resume next
		set GetRs = objDb.Execute(strSql)
		ErrCheck strSql
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Server.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
		on error resume next
		objDB.Open strDbName
'		Err.Raise 1,"エラーテスト"
		ErrCheck "OpenDB()"
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'ErrCheck() エラーの場合は中断
	'-------------------------------------------------------------------
	Private Function ErrCheck(byVal s)
		if Err.Number = 0 then
			exit function
		end if
		JsonWrite "{"
		JsonWrite " ""status"":""error"""
		JsonWrite ",""dbName"":""" 		& strDbName		& """"
		JsonWrite ",""ErrPlace"":""" 	& JsonValue(s) & """"
		JsonWrite ",""ErrNumber"":""" 	& Err.Number & "(0x" & Hex(Err.Number) & ")" 	& """"
		JsonWrite ",""ErrDescription"":"""	& Err.Description & """"
		JsonWrite ",""ErrSource"":"""	& Err.Source & """"
		JsonWrite "}"
		on error goto 0
		Class_Terminate
		Response.End
	End Function
	'-------------------------------------------------------------------
	'GetDbName
	'-------------------------------------------------------------------
	Private Function GetDbName()
		dim	strDbName
		strDbName = Request.QueryString("dbName")
		if strDbName = "" then
			strDbName = lcase(Split(Request.ServerVariables("PATH_TRANSLATED"),"\")(1))
		end if
		if strDbName = "it" then
			strDbName = "newsdc4"
		end if
		GetDbName = strDbName
	End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		if objDB is nothing then
			exit function
		end if
		if objDB.State <> adStateClosed then
			objDB.Close
		end if
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'文字列追加 strSql
	'-------------------------------------------------------------------
	Private	strSql
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
	End Function
	'-------------------------------------------------------------------
	'ErrMsg() エラーメッセージ出力
	'-------------------------------------------------------------------
    Private Function ErrMsg(byVal strMsg)
		Response.Write strMsg & vbCrLf
	End Function
	'-------------------------------------------------------------------
	'Json出力
	'-------------------------------------------------------------------
    Private Function JsonWrite(byVal strJson)
		Response.Write strJson & vbCrLf
	End Function
	'-------------------------------------------------------------------
	'Json 文字列エスケープ
	'-------------------------------------------------------------------
    Private Function JsonValue(byVal strJson)
'		strJson = Replace(strJson,"\","\\")
'		strJson = Replace(strJson,",","\,")
		strJson = Replace(strJson,"""","\""")
		strJson = Replace(strJson,vbTab," ")
'		strJson = Replace(strJson,vbCrLf,"\n")
		JsonValue = strJson
	End Function
	'-------------------------------------------------------------------
	'全Trim
	'-------------------------------------------------------------------
    Private Function ATrim(byVal strTrim)
		ATrim = Replace(Replace(Trim(strTrim)," ",""),"　","")
	End Function
	'-------------------------------------------------------------------
	'StrDate()
	'-------------------------------------------------------------------
	Private	Function StrDate(byVal vDt)
		StrDate = ""
		if isDate(vDt) = False then
			exit function
		end if
		StrDate = Year(vDt) & Right("0" & Month(vDt),2) & Right("0" & Day(vDt),2)
	End Function
	'-------------------------------------------------------------------
	'CCur()
	'-------------------------------------------------------------------
	Private	Function CCur(byVal v)
		CCur = 0
		if isNumeric(v) = false then
			exit function
		end if
		CCur = CLng(v)
	End Function
	'-----------------------------------------------------------------------
	'Debug
	'-----------------------------------------------------------------------
	Private Sub Debug(byVal strMsg)
		if optDebug then Response.Write strMsg & vbCrLf
	End Sub
	'-----------------------------------------------------------------------
	'LogTime
	'-----------------------------------------------------------------------
	Private Function LogTime()
		LogTime = year(now()) & "/" & right("0" & month(now()),2)  & "/" & right("0" & day(now()),2)
		LogTime = LogTime & " " & right("0" & hour(now()),2) & ":" & right("0" & minute(now()),2)  & ":" & right("0" & second(now()),2)
	End Function
	'-----------------------------------------------------------------------
	'Browser
	'-----------------------------------------------------------------------
	Private Function Browser()
		dim	userAgent
		userAgent = lcase(Request.ServerVariables("HTTP_USER_AGENT"))
		if inStr(userAgent,"msie") > 0 then
			Browser = "IEolder"
		elseif inStr(userAgent,"trident") > 0 then
			Browser = "IE11"
		elseif inStr(userAgent,"edge") > 0 then
			Browser = "Edge"
		elseif inStr(userAgent,"chrome") > 0 then
			Browser = "Chrome"
		elseif inStr(userAgent,"firefox") > 0 then
			Browser = "FireFox"
		elseif inStr(userAgent,"opera") > 0 then
			Browser = "Opera"
		else
			Browser = "unknown"
		end if
	End Function
	'-----------------------------------------------------------------------
	'LogTime
	'-----------------------------------------------------------------------
	Private Function LogOutput(byVal t)
		dim	objFS
		dim	ts
		dim	strFileName
		const	ForAppending	= 8
		const	ForReading	= 1
		const	ForWriting	= 2

		strFileName = "p_sukeire_inv.log"
		Set objFS = Server.CreateObject("Scripting.FileSystemObject")
		Set ts = objFS.OpenTextFile(Server.MapPath(strFileName),ForAppending, True)
		ts.WriteLine LogTime() & " " & Request.ServerVariables("REMOTE_ADDR") & " " & Browser() & ":" & t & ":" & Request.ServerVariables("QUERY_STRING")
		ts.close
		set ts = nothing
		set objFS = nothing
	End function
End Class
%>
