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

function usage()
'    Wscript.Echo "福山通運引渡データ出力(2010.03.01) 重量ありの時、才数=000 をセット"
'    Wscript.Echo "福山通運引渡データ出力(2010.03.16) 1件目キャンセルの対応"
'    Wscript.Echo "福山通運引渡データ出力(2010.04.05) TelNo対応(2便より) ※郵便番号はデータ項目にない"
'    Wscript.Echo "福山通運引渡データ出力(2010.06.11) IDNoに出荷日を追加"
'    Wscript.Echo "福山通運引渡データ出力(2011.10.05) 才数を最低1になるように変更"
    Wscript.Echo "福山通運引渡データ出力(2019.10.25) 送り状：310% を除外"
	Wscript.Echo "y_syuka_h.vbs [option] <yyyymmdd>"
	Wscript.Echo "               -del : del_syuka_h を参照(デフォルト y_syuka_h)"
	Wscript.Echo "               -b1  : 1便"
	Wscript.Echo "               -b2  : 2便"
	Wscript.Echo "               -b3  : 3便"
	Wscript.Echo "               -label : 荷札ラベルデータ出力"
    Wscript.Echo "               -?"
end function

Sub Main()
	dim	db
	dim	dbName
	dim	strSql
	dim	rsList
	dim	strFilename
	dim	i
	dim	strBuff
	dim	objFSO
	dim	objFile
	dim	objLog
	dim	strFind
	dim	strMsg
	dim	strUpdMsg
	dim	lngCnt			' 送り状件数
	dim	lngQty			' 口数
	dim	lngSai			' 才数
	dim	lngWait			' 重量
	dim	lngQty100		' 口数 100以上の件数
	dim	strDt
	dim	strNinushi
	dim	strBukasyo
	dim	strIdNo
	dim	strHDt
	dim	strONo
	dim	strNoS
	dim	strNoE
	dim	strHKbn
	dim	strMKbn
	dim	strQty
	dim	strSai
	dim	strWait
	dim	strHoken
	dim	strAddress1
	dim	strAddress2
	dim	strName1
	dim	strName2
	dim	strTel
	dim	strKiji1
	dim	strKiji2
	dim	strKiji3
	dim	strKiji4
	dim	strKiji5
	dim	strYobi
	dim	strYSyukaH
	dim	strBin
	dim	strLabel
	dim	strWork
	dim	strErr

	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const adSearchForward = 1
	' ObjectStateEnum
	' オブジェクトを開いているか閉じているか、データ ソースに接続中か、
	' コマンドを実行中か、またはデータを取得中かどうかを表します。
	Const	adStateClosed		= 0 ' オブジェクトが閉じていることを示します。 
	Const	adStateOpen			= 1 ' オブジェクトが開いていることを示します。 
	Const	adStateConnecting	= 2 ' オブジェクトが接続していることを示します。 
	Const	adStateExecuting	= 4 ' オブジェクトがコマンドを実行中であることを示します。 
	Const	adStateFetching		= 8 ' オブジェクトの行が取得されていることを示します。 

	strYSyukaH	= "y_syuka_h"
	strDt		= ""
	strBin		= ""
	strLabel	= ""
	strWork		= ""
	strErr		= ""
	for i = 0 to WScript.Arguments.count - 1
	    select case lcase(WScript.Arguments(i))
	    case "-del"
			strYSyukaH = "del_syuka_h"
	    case "-b1"
			strBin = "01"
	    case "-b2"
			strBin = "02"
	    case "-b3"
			strBin = "03"
	    case "-label"
			strLabel = "label"
		case "-work"
			strWork		= "work"
		case "-err"
			strErr		= "err"
	    case "-?"
			usage()
			Wscript.Quit
	    case else
			if strDt = "" then
				strDt = WScript.Arguments(i)
			else
				usage()
				Wscript.Quit
			end if
	    end select
	next
	if strDt = "" then
		usage()
		Wscript.Quit
	end if
	Wscript.Echo "y_syuka_h.vbs " & strDt & " " & strYSyukaH & " " & strBin

	' データベースOpen
	dbName = "newsdc"

	Set db = Wscript.CreateObject("ADODB.Connection")
	Wscript.Echo "open db : " & dbName
	db.open dbName

	' PNテーブルOpen
	Set rsList = Wscript.CreateObject("ADODB.Recordset")
	strSql = "select * from " & strYSyukaH
'	strSql = strSql & " where SYUKA_YMD like '" & strDt & "'"
	if strYSyukaH = "del_syuka_h" then
		strSql = strSql & " where length(rtrim(OKURI_NO)) = 11"
		strSql = strSql & " and SYUKA_YMD = '" & strDt & "'"
		strSql = strSql & " and left(KENPIN_NOW,8) <> '" & strDt & "'"
	else
		strSql = strSql & " where length(rtrim(OKURI_NO)) = 11"
	end if
'	strSql = strSql & " and SEQ_NO = '1'"
	strSql = strSql & " and CANCEL_F = ''"
	strSql = strSql & " and UNSOU_KAISHA = '福山通運'"
	strSql = strSql & " and OKURI_NO not like '310%'"
	if strErr = "" then
		strSql = strSql & " and (convert(SAI_SU,sql_numeric) > 0"
		strSql = strSql & "  or  convert(JURYO,sql_numeric) > 0)"
	else
		strSql = strSql & " and (convert(SAI_SU,sql_numeric) = 0"
		strSql = strSql & "  and convert(JURYO,sql_numeric) = 0)"
	end if
	if strBin <> "" then
		strSql = strSql & " and INS_BIN = '" & strBin & "'"
	end if
	if strWork = "" then
		strSql = strSql & " and OKURI_NO not in"
		strSql = strSql & "  (select distinct OKURI_NO from y_syuka_h"
		strSql = strSql &   " where SYUKA_YMD = '" & strDt & "'"
		strSql = strSql &   " and length(rtrim(OKURI_NO)) = 11"
		strSql = strSql &   " and CANCEL_F = ''"
		strSql = strSql &   " and UNSOU_KAISHA = '福山通運'"
		strSql = strSql &   " and (convert(SAI_SU,sql_numeric) = 0"
		strSql = strSql &   " and  convert(JURYO,sql_numeric) = 0)"
		strSql = strSql &   ")"
	end if
	strSql = strSql & " order by OKURI_NO,ID_NO"
	rsList.Open strSql, db, adOpenForwardOnly, adLockBatchOptimistic
	strONo 		= ""
	lngCnt 		= 0		' 送り状件数
	lngQty		= 0		' 口数
	lngSai		= 0		' 才数
	lngWait	 	= 0		' 重量
	lngQty100	= 0		' 口数 100以上の件数
	do while ( rsList.Eof = False )
		if rtrim(strONo) <> rtrim(rsList.Fields("OKURI_NO")) then
			strNinushi		= "072874606S"
			strBukasyo		= "      "
'			strHDt			= Get_Buff(right(rsList.Fields("SYUKA_YMD"),6),6)
'			strHDt			= Get_Buff(right(strDt,6),6)
			strHDt			= Get_Buff(right(left(rsList.Fields("KENPIN_NOW"),8),6),6)
			strIdNo			= Get_Buff(strHDt & rsList.Fields("ID_NO"),20)
'			strHDt			= Get_Buff(right(rsList.Fields("SYUKA_YMD"),6),6)
			strONo			= Get_Buff(rsList.Fields("OKURI_NO"),11)
			strNoS			= Get_Buff(Left(RTrim(rsList.Fields("OKURI_NO")),11) & "01",13)
			strNoE			= Get_Buff(Left(RTrim(rsList.Fields("OKURI_NO")),11) & Right(RTrim(rsList.Fields("KUTI_SU")),2),13)
			strHKbn			= "1"
			strMKbn			= "1"
			strQty			= Get_Buff(Right(RTrim(rsList.Fields("KUTI_SU")),3),3)
'			strSai			= Get_Buff(Right("000"&round(cdbl("0"&RTrim(rsList.Fields("SAI_SU"))),0),3),3)	' "000"
			strSai			= Get_Buff(Right("000"&GetSaisu(rsList.Fields("SAI_SU")),3),3)			' "000"
			strWait			= Get_Buff(Right("0000"&round(cdbl("0"&RTrim(rsList.Fields("JURYO"))),0),4),4)	' "0000"
			if strWait <> "0000" then
				strSai		= "000"
			end if
			strHoken		= "0000"
			strAddress1		= Get_BuffZ(RTrim(rsList.Fields("JYUSHO")),80)
			strAddress2		= ""
'			strAddress1		= Get_BuffZ("荷受人住所１",40)
'			strAddress2		= Get_BuffZ("荷受人住所２",40)
			strName1		= Get_BuffZ(rsList.Fields("OKURISAKI"),40)
			strName2		= ""
			if rsList.Fields("OKURISAKI") <> rsList.Fields("MUKE_NAME") then
				strName2 = rsList.Fields("MUKE_NAME")
			end if
			strName2		= Get_BuffZ(strName2,40)				' MUKE_NAME			Char(40)	
'			strTel			= Get_Buff("00-0000-0000",15)
			strTel			= Get_Buff(rsList.Fields("TEL_No"),15)
			strKiji1		= Get_BuffZ(rsList.Fields("BIKOU"),200)	' BIKOU				Char(100)
			strKiji2		= ""
			if rsList.Fields("SYUKA_YMD") <> strDt then
				strKiji1	= Get_BuffZ("出荷日変更：" & rsList.Fields("SYUKA_YMD"),40)
				strKiji2	= Get_BuffZ(rsList.Fields("BIKOU"),160)	' BIKOU				Char(100)
			end if
			strKiji3		= ""
			strKiji4		= ""
			strKiji5		= ""
'			strKiji2		= Get_BuffZ("記事欄２",40)
'			strKiji3		= Get_BuffZ("記事欄３",40)
'			strKiji4		= Get_BuffZ("記事欄４",40)
'			strKiji5		= Get_BuffZ("記事欄５",40)
			strYobi			= Get_BuffZ("",40)
			if strLabel <> "" then
				Wscript.Echo  "JOB"
				WScript.Echo "DEF MK=1,DK=8,MD=1,PW=384,PH=344,XO=8,UM=8"
				WScript.Echo "START"
				WScript.Echo "FONT TP=3,CS=0"
				WScript.Echo "TEXT X=33,Y=0,L=1,NS=12,NE=2,NZ=0"
				WScript.Echo strONo & "01"
				WScript.Echo "TEXT X=275,Y=0,L=1,NS=1,NE=3,NZ=1"
				WScript.Echo "001/" & GetQty(strQty," ")
				WScript.Echo "BCD TP=6,X=0,Y=22,HT=40,HR=0,NS=12,NE=2,NZ=0"
				WScript.Echo strONo & "01"
				WScript.Echo "FONT TP=7,CS=0,LG=36,WD=18,LS=0"
				WScript.Echo "TEXT X=574,Y=65,L=1"
				WScript.Echo "着店:000"
				WScript.Echo "TEXT X=0,Y=65,L=7"
				WScript.Echo strTel
				WScript.Echo Get_LeftB(strAddress1,40)
				WScript.Echo Get_MidB(strAddress1,41,40)
				WScript.Echo strName1
				WScript.Echo strName2
				WScript.Echo "                          20" & Get_MidB(strHDt,1,2) & "年" & Get_MidB(strHDt,3,2) & "月" & Get_MidB(strHDt,5,2) & "日"
				WScript.Echo "        (株)エスディーシィー　部材流通Ｃ"
				WScript.Echo "QTY P=" & GetQty(strQty,"")
				WScript.Echo "END"
				WScript.Echo "JOBE"
			else
				strMsg = ""
				strMsg = strMsg & strNinushi
				strMsg = strMsg & strBukasyo	
				strMsg = strMsg & strIdNo		
				strMsg = strMsg & strHDt		
				strMsg = strMsg & strONo		
				strMsg = strMsg & strNoS		
				strMsg = strMsg & strNoE		
				strMsg = strMsg & strHKbn		
				strMsg = strMsg & strMKbn		
				strMsg = strMsg & strQty		
				strMsg = strMsg & strSai		
				strMsg = strMsg & strWait		
				strMsg = strMsg & strHoken		
				strMsg = strMsg & strAddress1	
				strMsg = strMsg & strAddress2	
				strMsg = strMsg & strName1		
				strMsg = strMsg & strName2		
				strMsg = strMsg & strTel		
				strMsg = strMsg & strKiji1		
				strMsg = strMsg & strKiji2		
				strMsg = strMsg & strKiji3		
				strMsg = strMsg & strKiji4		
				strMsg = strMsg & strKiji5		
				strMsg = strMsg & strYobi		
				Wscript.Echo strMsg
			end if
			lngCnt	= lngCnt  + 1			' 送り状件数
			lngQty	= lngQty  + clng(strQty)		' 口数
			lngSai	= lngSai  + clng(strSai)	' 才数
			lngWait	= lngWait + clng(strWait)	' 重量
			if clng(strQty) >= 100 then
				lngQty100	= lngQty100  + 1		' 口数(>100)
			end if
		end if
		rsList.movenext
	loop
	Wscript.Echo "出荷日：" & strDt
	Wscript.Echo "送り状：" & right("        " & lngCnt ,6)
	Wscript.Echo "  口数：" & right("        " & lngQty ,6)
	Wscript.Echo "  才数：" & right("        " & lngSai ,6)
	Wscript.Echo "  重量：" & right("        " & lngWait,6)
	Wscript.Echo "口数≧100の件数：" & lngQty100 & " 件"

	' テーブルClose
	Wscript.Echo "close table : " & strYSyukaH
	rsList.Close

	' DBClose
	Wscript.Echo "close db : " & dbName
	db.Close
	set db = nothing
End Sub

Function GetTm(t)
	GetTm = year(t) & right("0" & month(t),2) & right("0" & day(t),2) & right("0" & hour(t),2)& right("0" & minute(t),2)
End Function

Function Get_Buff(a_Str,a_int)
	dim	strRet

	strRet = a_Str & space(a_int)
	strRet = Get_LeftB(strRet,a_int)
	Get_Buff = strRet
End Function

Function Get_BuffZ(a_Str,a_int)
	dim	strRet

	strRet = StrConvWide(rtrim(a_Str)) & string(a_int,"　")
	strRet = Get_LeftB(strRet,a_int)
	Get_BuffZ = strRet
End Function

Function Get_LeftB(a_Str, a_int)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_LeftB = ""
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
		If iLenCount > Cint(a_int) Then
			Exit For
		Else
			iLeftStr = iLeftStr + Mid(a_Str, iCount, 1)
		End If
	Next
	Get_LeftB = iLeftStr
End Function

Function Get_MidB(a_Str,s_int, a_int)
	Dim iCount, iAscCode, iLenCount, iMidStr
	iLenCount = 0
	iMidStr = ""
	If Len(a_Str) = 0 Then
		Get_MidB = ""
		Exit Function
	End If
	If a_int = 0 Then
		Get_MidB = ""
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
		if iLenCount >= s_int then
			If iLenCount > Cint(s_int) + Cint(a_int) - 1 Then
				Exit For
			Else
				iMidStr = iMidStr + Mid(a_Str, iCount, 1)
			End If
		end if
	Next
	Get_MidB = iMidStr
End Function

Function Get_LenB(a_Str)
	Dim iCount, iAscCode, iLenCount, iLeftStr
	iLenCount = 0
	iLeftStr = ""
	If Len(a_Str) = 0 Then
		Get_LenB = 0
		Exit Function
	End If
	For iCount = 1 to Len(a_Str)
		'** Asc関数で文字コード取得
		iAscCode = Asc(Mid(a_Str, iCount, 1))
		'** 半角は文字コードの長さが2、全角は4(2以上)として判断
		If Len(Hex(iAscCode)) > 2 Then
			iLenCount = iLenCount + 2
		Else
			iLenCount = iLenCount + 1
		End If
	Next
	Get_LenB = iLenCount
End Function


Function SetField(rsPn,strFieldName,strValue,strTitle,strUpdMsg)
	if rtrim(rsPn.Fields(strFieldName)) <> rtrim(strValue) then
		if strUpdMsg <> "-" then
			strUpdMsg = strUpdMsg & rsPn.Fields(strFieldName) & " ←" & strTitle & vbNewLine
			strUpdMsg = strUpdMsg & strValue & " ←変更" & vbNewLine
		end if
		rsPn.Fields(strFieldName) = strValue
	end if
	SetField = strUpdMsg
End Function


'**********************************************************************************
' Script関数名①  : DBCS_Convert(変換する文字列) Ver1.0
' Script関数名②  : SBCS_Convert(変換する文字列) Ver1.0
' Script関数名③  : SBCS_DBCS_Check(チェックする１文字) Ver1.0
' 機能概要  : ①文字列中の半角文字を全角に変換します
'           : ②文字列中の全角文字を半角に変換します
'           : ③文字を全角か半角か判定します
' Made By   : Copyright(C) 2008 T.Tokunaga All right reserved
'           : このプログラムは日本国著作権法および国際条約により保護されています。
'           : このプログラムを転載する場合は著作権所有者の許可が必要となります｡
'**********************************************************************************

'濁点・半濁点文字（半角）
'ｶﾞｷﾞｸﾞｹﾞｺﾞｻﾞｼﾞｽﾞｾﾞｿﾞﾀﾞﾁﾞﾂﾞﾃﾞﾄﾞﾊﾞﾊﾟﾋﾞﾋﾟﾌﾞﾌﾟﾍﾞﾍﾟﾎﾞﾎﾟｳﾞ
Public CNDakutenSBCS
CNDakutenSBCS = "ｶﾞｷﾞｸﾞｹﾞｺﾞｻﾞｼﾞｽﾞｾﾞｿﾞﾀﾞﾁﾞﾂﾞﾃﾞﾄﾞﾊﾞﾊﾟﾋﾞﾋﾟﾌﾞﾌﾟﾍﾞﾍﾟﾎﾞﾎﾟｳﾞ"

'濁点・半濁点文字（全角）
'ガギグゲゴザジズゼゾダヂヅデドバパビピブプベペボポヴ
Public CNDakutenDBCS
CNDakutenDBCS = "ガギグゲゴザジズゼゾダヂヅデドバパビピブプベペボポヴ"

'半角文字
' !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ
Public CNConvSBCS
CNConvSBCS = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ"

'全角文字
'　！”＃＄％＆’（）＊＋，－．／０１２３４５６７８９：；＜＝＞？＠ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ［￥］＾＿‘ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛｜｝～。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜
Public CNConvDBCS
CNConvDBCS = "　！”＃＄％＆’（）＊＋，－．／０１２３４５６７８９：；＜＝＞？＠ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ［￥］＾＿‘ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛｜｝～。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜"

'*************************************
'関数呼び出し例
'
'Check_after = DBCS_Convert(CNDakutenSBCS)
'
'MsgBox Check_after
'
'Check_after = DBCS_Convert(CNConvSBCS)
'
'MsgBox Check_after
'
'Check_after = SBCS_Convert(CNDakutenDBCS)
'
'MsgBox Check_after
'
'Check_after = SBCS_Convert(CNConvDBCS)
'
'MsgBox Check_after

'*************************************

Private Function DBCS_Convert(Check_String)
dim	DBCS_Convert_Temp_Data
dim	DBCS_Convert_j
dim	DBCS_Convert_jj
dim	DBCS_Convert_AscChk_Data
dim	DBCS_Convert_Sarch_SBCS
dim	DBCS_Convert_Conv_Data

  If  Len(Check_String) > 0 Then
    DBCS_Convert_Temp_Data = ""
    For DBCS_Convert_j = 1 To Len(Check_String)
      If SBCS_DBCS_Check(Mid(Check_String,DBCS_Convert_j,1)) = 1 Then
        DBCS_Convert_Temp_Data = DBCS_Convert_Temp_Data & Mid(Check_String,DBCS_Convert_j,1)
      Else
        DBCS_Convert_jj = DBCS_Convert_j + 1
        If DBCS_Convert_jj <= Len(Check_String) Then
          DBCS_Convert_AscChk_Data = Mid(Check_String,DBCS_Convert_jj,1)
          If Mid(Check_String,DBCS_Convert_jj,1) <> "ﾟ" And Mid(Check_String,DBCS_Convert_jj,1) <> "ﾞ" Then
            DBCS_Convert_AscChk_Data = ""
          Else 
            DBCS_Convert_AscChk_Data = Mid(Check_String,DBCS_Convert_j,1)&Mid(Check_String,DBCS_Convert_jj,1)
          End If
        Else
          DBCS_Convert_AscChk_Data = ""
        End If
        If DBCS_Convert_AscChk_Data = "" Then
          DBCS_Convert_Sarch_SBCS = InStr(1,CNConvSBCS,Mid(Check_String,DBCS_Convert_j,1),vbBinaryCompare)
          If DBCS_Convert_Sarch_SBCS = "" Or  DBCS_Convert_Sarch_SBCS = 0 Then
            DBCS_Convert_Conv_Data = "①"	' Mid(Check_String,DBCS_Convert_j,1)
          Else
            DBCS_Convert_Conv_Data = Mid(CNConvDBCS,DBCS_Convert_Sarch_SBCS,1)
          End If
        Else
          DBCS_Convert_Sarch_SBCS = InStr(1,CNDakutenSBCS,DBCS_Convert_AscChk_Data,vbBinaryCompare)
          If DBCS_Convert_Sarch_SBCS = "" Or  DBCS_Convert_Sarch_SBCS = 0 Then
            DBCS_Convert_Sarch_SBCS = InStr(1,CNConvSBCS,Mid(Check_String,DBCS_Convert_j,1),vbBinaryCompare)
            If DBCS_Convert_Sarch_SBCS = "" Or  DBCS_Convert_Sarch_SBCS = 0 Then
              DBCS_Convert_Conv_Data = "②"
            Else
              DBCS_Convert_Conv_Data = Mid(CNConvDBCS,DBCS_Convert_Sarch_SBCS,1)
            End If
          Else
            DBCS_Convert_Conv_Data = Mid(CNDakutenDBCS,(DBCS_Convert_Sarch_SBCS + 1 ) / 2, 1)
            DBCS_Convert_j = DBCS_Convert_j + 1
          End If
        End If
        DBCS_Convert_Temp_Data = DBCS_Convert_Temp_Data & DBCS_Convert_Conv_Data
      End If
    Next
    DBCS_Convert = DBCS_Convert_Temp_Data
  Else
    DBCS_Convert = Check_String
  End If

End Function

Private Function SBCS_Convert(Check_String)
	dim	SBCS_Convert_Temp_Data
	dim	SBCS_Convert_j
	dim	SBCS_Convert_Sarch_DBCS
	dim	SBCS_Convert_Conv_Data

  If  Len(Check_String) > 0 Then
    SBCS_Convert_Temp_Data = ""
    For SBCS_Convert_j = 1 To Len(Check_String)
      If SBCS_DBCS_Check(Mid(Check_String,SBCS_Convert_j,1)) = 0 Then
        SBCS_Convert_Temp_Data = SBCS_Convert_Temp_Data & Mid(Check_String,SBCS_Convert_j,1)
      Else
        SBCS_Convert_Sarch_DBCS = InStr(1,CNDakutenDBCS,Mid(Check_String,SBCS_Convert_j,1),vbBinaryCompare)
        If SBCS_Convert_Sarch_DBCS = "" Or  SBCS_Convert_Sarch_DBCS = 0 Then
          SBCS_Convert_Sarch_DBCS = InStr(1,CNConvDBCS,Mid(Check_String,SBCS_Convert_j,1),vbBinaryCompare)
          If SBCS_Convert_Sarch_DBCS = "" Or  SBCS_Convert_Sarch_DBCS = 0 Then
            SBCS_Convert_Conv_Data = "?"
          Else
            SBCS_Convert_Conv_Data = Mid(CNConvSBCS,SBCS_Convert_Sarch_DBCS,1)
          End If
        Else
          SBCS_Convert_Conv_Data = Mid(CNDakutenSBCS,(SBCS_Convert_Sarch_DBCS * 2) - 1, 2)
        End If
        SBCS_Convert_Temp_Data = SBCS_Convert_Temp_Data & SBCS_Convert_Conv_Data
      End If
    Next
    SBCS_Convert = SBCS_Convert_Temp_Data
  Else
    SBCS_Convert = Check_String
  End If
End Function

Private Function SBCS_DBCS_Check(Check_Word)
  If ASC(Check_Word) > 255 Or ASC(Check_Word) < 0 Then
     'DBCS(全角)と認識
    SBCS_DBCS_Check = 1
  Else
     'SBCS(半角)と認識
    SBCS_DBCS_Check = 0
  End If
End Function

'#################################################################
'# StrConv Clone For VBScript
'#  author: Yasuhiro Matsumoto
'#  url: http://www.ac.cyberhome.ne.jp/~mattn/cgi-bin/blosxom.cgi
'#  mailto: mattn.jp@gmai.com
'#  see: http://www.ac.cyberhome.ne.jp/~mattn/AcrobatASP/1.html
'#################################################################

'***************************************************
' StrConvUpperCase
'---------------------------------------------------
' 用途 : StrConv(sInp,vbUpperCase) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvUpperCase(sInp)
	StrConvUpperCase = UCase(sInp)
End Function

'***************************************************
' StrConvLowerCase
'---------------------------------------------------
' 用途 : StrConv(sInp,vbLowerCase) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvLowerCase(sInp)
	StrConvLowerCase = LCase(sInp)
End Function

'***************************************************
' StrConvProperCase
'---------------------------------------------------
' 用途 : StrConv(sInp,vbProperCase) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvProperCase(sInp)
	Dim nPos
	Dim nSpc

	nPos = 1
	Do While InStr(nPos, sInp, " ", 1) <> 0
		nSpc = InStr(nPos, sInp, " ", 1)
		StrConvProperCase = StrConvProperCase & UCase(Mid(sInp, nPos, 1))
		StrConvProperCase = StrConvProperCase & LCase(Mid(sInp, nPos + 1, nSpc - nPos))
		nPos = nSpc + 1
	Loop

	StrConvProperCase = StrConvProperCase & UCase(Mid(sInp, nPos, 1))
	StrConvProperCase = StrConvProperCase & LCase(Mid(sInp, nPos + 1))
	StrConvProperCase = StrConvProperCase
End Function

'***************************************************
' StrConvWide
'---------------------------------------------------
' 用途 : StrConv(s,vbWide) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvWide(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("ﾞﾟ", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case " "
			StrConvWide = StrConvWide & "　"
		Case "!"
			StrConvWide = StrConvWide & "！"
		Case """"
			StrConvWide = StrConvWide & "＂"
		Case "#"
			StrConvWide = StrConvWide & "＃"
		Case "$"
			StrConvWide = StrConvWide & "＄"
		Case "%"
			StrConvWide = StrConvWide & "％"
		Case "&"
			StrConvWide = StrConvWide & "＆"
		Case "'"
			StrConvWide = StrConvWide & "＇"
		Case "("
			StrConvWide = StrConvWide & "（"
		Case ")"
			StrConvWide = StrConvWide & "）"
		Case "*"
			StrConvWide = StrConvWide & "＊"
		Case "+"
			StrConvWide = StrConvWide & "＋"
		Case ","
			StrConvWide = StrConvWide & "，"
		Case "-"
			StrConvWide = StrConvWide & "－"
		Case "."
			StrConvWide = StrConvWide & "．"
		Case "/"
			StrConvWide = StrConvWide & "／"
		Case "0"
			StrConvWide = StrConvWide & "０"
		Case "1"
			StrConvWide = StrConvWide & "１"
		Case "2"
			StrConvWide = StrConvWide & "２"
		Case "3"
			StrConvWide = StrConvWide & "３"
		Case "4"
			StrConvWide = StrConvWide & "４"
		Case "5"
			StrConvWide = StrConvWide & "５"
		Case "6"
			StrConvWide = StrConvWide & "６"
		Case "7"
			StrConvWide = StrConvWide & "７"
		Case "8"
			StrConvWide = StrConvWide & "８"
		Case "9"
			StrConvWide = StrConvWide & "９"
		Case ":"
			StrConvWide = StrConvWide & "："
		Case ";"
			StrConvWide = StrConvWide & "；"
		Case "<"
			StrConvWide = StrConvWide & "＜"
		Case "="
			StrConvWide = StrConvWide & "＝"
		Case ">"
			StrConvWide = StrConvWide & "＞"
		Case "?"
			StrConvWide = StrConvWide & "？"
		Case "@"
			StrConvWide = StrConvWide & "＠"
		Case "A"
			StrConvWide = StrConvWide & "Ａ"
		Case "B"
			StrConvWide = StrConvWide & "Ｂ"
		Case "C"
			StrConvWide = StrConvWide & "Ｃ"
		Case "D"
			StrConvWide = StrConvWide & "Ｄ"
		Case "E"
			StrConvWide = StrConvWide & "Ｅ"
		Case "F"
			StrConvWide = StrConvWide & "Ｆ"
		Case "G"
			StrConvWide = StrConvWide & "Ｇ"
		Case "H"
			StrConvWide = StrConvWide & "Ｈ"
		Case "I"
			StrConvWide = StrConvWide & "Ｉ"
		Case "J"
			StrConvWide = StrConvWide & "Ｊ"
		Case "K"
			StrConvWide = StrConvWide & "Ｋ"
		Case "L"
			StrConvWide = StrConvWide & "Ｌ"
		Case "M"
			StrConvWide = StrConvWide & "Ｍ"
		Case "N"
			StrConvWide = StrConvWide & "Ｎ"
		Case "O"
			StrConvWide = StrConvWide & "Ｏ"
		Case "P"
			StrConvWide = StrConvWide & "Ｐ"
		Case "Q"
			StrConvWide = StrConvWide & "Ｑ"
		Case "R"
			StrConvWide = StrConvWide & "Ｒ"
		Case "S"
			StrConvWide = StrConvWide & "Ｓ"
		Case "T"
			StrConvWide = StrConvWide & "Ｔ"
		Case "U"
			StrConvWide = StrConvWide & "Ｕ"
		Case "V"
			StrConvWide = StrConvWide & "Ｖ"
		Case "W"
			StrConvWide = StrConvWide & "Ｗ"
		Case "X"
			StrConvWide = StrConvWide & "Ｘ"
		Case "Y"
			StrConvWide = StrConvWide & "Ｙ"
		Case "Z"
			StrConvWide = StrConvWide & "Ｚ"
		Case "["
			StrConvWide = StrConvWide & "［"
		Case "]"
			StrConvWide = StrConvWide & "］"
		Case "^"
			StrConvWide = StrConvWide & "＾"
		Case "_"
			StrConvWide = StrConvWide & "＿"
		Case "`"
			StrConvWide = StrConvWide & "｀"
		Case "a"
			StrConvWide = StrConvWide & "ａ"
		Case "b"
			StrConvWide = StrConvWide & "ｂ"
		Case "c"
			StrConvWide = StrConvWide & "ｃ"
		Case "d"
			StrConvWide = StrConvWide & "ｄ"
		Case "e"
			StrConvWide = StrConvWide & "ｅ"
		Case "f"
			StrConvWide = StrConvWide & "ｆ"
		Case "g"
			StrConvWide = StrConvWide & "ｇ"
		Case "h"
			StrConvWide = StrConvWide & "ｈ"
		Case "i"
			StrConvWide = StrConvWide & "ｉ"
		Case "j"
			StrConvWide = StrConvWide & "ｊ"
		Case "k"
			StrConvWide = StrConvWide & "ｋ"
		Case "l"
			StrConvWide = StrConvWide & "ｌ"
		Case "m"
			StrConvWide = StrConvWide & "ｍ"
		Case "n"
			StrConvWide = StrConvWide & "ｎ"
		Case "o"
			StrConvWide = StrConvWide & "ｏ"
		Case "p"
			StrConvWide = StrConvWide & "ｐ"
		Case "q"
			StrConvWide = StrConvWide & "ｑ"
		Case "r"
			StrConvWide = StrConvWide & "ｒ"
		Case "s"
			StrConvWide = StrConvWide & "ｓ"
		Case "t"
			StrConvWide = StrConvWide & "ｔ"
		Case "u"
			StrConvWide = StrConvWide & "ｕ"
		Case "v"
			StrConvWide = StrConvWide & "ｖ"
		Case "w"
			StrConvWide = StrConvWide & "ｗ"
		Case "x"
			StrConvWide = StrConvWide & "ｘ"
		Case "y"
			StrConvWide = StrConvWide & "ｙ"
		Case "z"
			StrConvWide = StrConvWide & "ｚ"
		Case "{"
			StrConvWide = StrConvWide & "｛"
		Case "|"
			StrConvWide = StrConvWide & "｜"
		Case "}"
			StrConvWide = StrConvWide & "｝"
		Case "~"
			StrConvWide = StrConvWide & "～"
		Case "｡"
			StrConvWide = StrConvWide & "。"
		Case "｢"
			StrConvWide = StrConvWide & "「"
		Case "｣"
			StrConvWide = StrConvWide & "」"
		Case "､"
			StrConvWide = StrConvWide & "、"
		Case "･"
			StrConvWide = StrConvWide & "・"
		Case "ｦ"
			StrConvWide = StrConvWide & "ヲ"
		Case "ｧ"
			StrConvWide = StrConvWide & "ァ"
		Case "ｨ"
			StrConvWide = StrConvWide & "ィ"
		Case "ｩ"
			StrConvWide = StrConvWide & "ゥ"
		Case "ｪ"
			StrConvWide = StrConvWide & "ェ"
		Case "ｫ"
			StrConvWide = StrConvWide & "ォ"
		Case "ｬ"
			StrConvWide = StrConvWide & "ャ"
		Case "ｭ"
			StrConvWide = StrConvWide & "ュ"
		Case "ｮ"
			StrConvWide = StrConvWide & "ョ"
		Case "ｯ"
			StrConvWide = StrConvWide & "ッ"
		Case "ｰ"
			StrConvWide = StrConvWide & "ー"
		Case "ｱ"
			StrConvWide = StrConvWide & "ア"
		Case "ｲ"
			StrConvWide = StrConvWide & "イ"
		Case "ｳ"
			StrConvWide = StrConvWide & "ウ"
		Case "ｳﾞ"
			StrConvWide = StrConvWide & "ヴ"
		Case "ｴ"
			StrConvWide = StrConvWide & "エ"
		Case "ｵ"
			StrConvWide = StrConvWide & "オ"
		Case "ｶ"
			StrConvWide = StrConvWide & "カ"
		Case "ｶﾞ"
			StrConvWide = StrConvWide & "ガ"
		Case "ｷ"
			StrConvWide = StrConvWide & "キ"
		Case "ｷﾞ"
			StrConvWide = StrConvWide & "ギ"
		Case "ｸ"
			StrConvWide = StrConvWide & "ク"
		Case "ｸﾞ"
			StrConvWide = StrConvWide & "グ"
		Case "ｹ"
			StrConvWide = StrConvWide & "ケ"
		Case "ｹﾞ"
			StrConvWide = StrConvWide & "ゲ"
		Case "ｺ"
			StrConvWide = StrConvWide & "コ"
		Case "ｺﾞ"
			StrConvWide = StrConvWide & "ゴ"
		Case "ｻ"
			StrConvWide = StrConvWide & "サ"
		Case "ｻﾞ"
			StrConvWide = StrConvWide & "ザ"
		Case "ｼ"
			StrConvWide = StrConvWide & "シ"
		Case "ｼﾞ"
			StrConvWide = StrConvWide & "ジ"
		Case "ｽ"
			StrConvWide = StrConvWide & "ス"
		Case "ｽﾞ"
			StrConvWide = StrConvWide & "ズ"
		Case "ｾ"
			StrConvWide = StrConvWide & "セ"
		Case "ｾﾞ"
			StrConvWide = StrConvWide & "ゼ"
		Case "ｿ"
			StrConvWide = StrConvWide & "ソ"
		Case "ｿﾞ"
			StrConvWide = StrConvWide & "ゾ"
		Case "ﾀ"
			StrConvWide = StrConvWide & "タ"
		Case "ﾀﾞ"
			StrConvWide = StrConvWide & "ダ"
		Case "ﾁ"
			StrConvWide = StrConvWide & "チ"
		Case "ﾁﾞ"
			StrConvWide = StrConvWide & "ヂ"
		Case "ﾂ"
			StrConvWide = StrConvWide & "ツ"
		Case "ﾂﾞ"
			StrConvWide = StrConvWide & "ヅ"
		Case "ﾃ"
			StrConvWide = StrConvWide & "テ"
		Case "ﾃﾞ"
			StrConvWide = StrConvWide & "デ"
		Case "ﾄ"
			StrConvWide = StrConvWide & "ト"
		Case "ﾄﾞ"
			StrConvWide = StrConvWide & "ド"
		Case "ﾅ"
			StrConvWide = StrConvWide & "ナ"
		Case "ﾆ"
			StrConvWide = StrConvWide & "ニ"
		Case "ﾇ"
			StrConvWide = StrConvWide & "ヌ"
		Case "ﾈ"
			StrConvWide = StrConvWide & "ネ"
		Case "ﾉ"
			StrConvWide = StrConvWide & "ノ"
		Case "ﾊ"
			StrConvWide = StrConvWide & "ハ"
		Case "ﾊﾞ"
			StrConvWide = StrConvWide & "バ"
		Case "ﾊﾟ"
			StrConvWide = StrConvWide & "パ"
		Case "ﾋ"
			StrConvWide = StrConvWide & "ヒ"
		Case "ﾋﾞ"
			StrConvWide = StrConvWide & "ビ"
		Case "ﾋﾟ"
			StrConvWide = StrConvWide & "ピ"
		Case "ﾌ"
			StrConvWide = StrConvWide & "フ"
		Case "ﾌﾞ"
			StrConvWide = StrConvWide & "ブ"
		Case "ﾌﾟ"
			StrConvWide = StrConvWide & "プ"
		Case "ﾍ"
			StrConvWide = StrConvWide & "ヘ"
		Case "ﾍﾞ"
			StrConvWide = StrConvWide & "ベ"
		Case "ﾍﾟ"
			StrConvWide = StrConvWide & "ペ"
		Case "ﾎ"
			StrConvWide = StrConvWide & "ホ"
		Case "ﾎﾞ"
			StrConvWide = StrConvWide & "ボ"
		Case "ﾎﾟ"
			StrConvWide = StrConvWide & "ポ"
		Case "ﾏ"
			StrConvWide = StrConvWide & "マ"
		Case "ﾐ"
			StrConvWide = StrConvWide & "ミ"
		Case "ﾑ"
			StrConvWide = StrConvWide & "ム"
		Case "ﾒ"
			StrConvWide = StrConvWide & "メ"
		Case "ﾓ"
			StrConvWide = StrConvWide & "モ"
		Case "ﾔ"
			StrConvWide = StrConvWide & "ヤ"
		Case "ﾕ"
			StrConvWide = StrConvWide & "ユ"
		Case "ﾖ"
			StrConvWide = StrConvWide & "ヨ"
		Case "ﾗ"
			StrConvWide = StrConvWide & "ラ"
		Case "ﾘ"
			StrConvWide = StrConvWide & "リ"
		Case "ﾙ"
			StrConvWide = StrConvWide & "ル"
		Case "ﾚ"
			StrConvWide = StrConvWide & "レ"
		Case "ﾛ"
			StrConvWide = StrConvWide & "ロ"
		Case "ﾜ"
			StrConvWide = StrConvWide & "ワ"
		Case "ﾝ"
			StrConvWide = StrConvWide & "ン"
		Case "ﾞ"
			StrConvWide = StrConvWide & "゛"
		Case "ﾟ"
			StrConvWide = StrConvWide & "゜"
		Case " "
			StrConvWide = StrConvWide & "　"
		Case "!"
			StrConvWide = StrConvWide & "！"
		Case """"
			StrConvWide = StrConvWide & "＂"
		Case "#"
			StrConvWide = StrConvWide & "＃"
		Case "$"
			StrConvWide = StrConvWide & "＄"
		Case "%"
			StrConvWide = StrConvWide & "％"
		Case "&"
			StrConvWide = StrConvWide & "＆"
		Case "'"
			StrConvWide = StrConvWide & "＇"
		Case "("
			StrConvWide = StrConvWide & "（"
		Case ")"
			StrConvWide = StrConvWide & "）"
		Case "*"
			StrConvWide = StrConvWide & "＊"
		Case "+"
			StrConvWide = StrConvWide & "＋"
		Case ","
			StrConvWide = StrConvWide & "，"
		Case "-"
			StrConvWide = StrConvWide & "－"
		Case "."
			StrConvWide = StrConvWide & "．"
		Case "/"
			StrConvWide = StrConvWide & "／"
		Case "0"
			StrConvWide = StrConvWide & "０"
		Case "1"
			StrConvWide = StrConvWide & "１"
		Case "2"
			StrConvWide = StrConvWide & "２"
		Case "3"
			StrConvWide = StrConvWide & "３"
		Case "4"
			StrConvWide = StrConvWide & "４"
		Case "5"
			StrConvWide = StrConvWide & "５"
		Case "6"
			StrConvWide = StrConvWide & "６"
		Case "7"
			StrConvWide = StrConvWide & "７"
		Case "8"
			StrConvWide = StrConvWide & "８"
		Case "9"
			StrConvWide = StrConvWide & "９"
		Case ":"
			StrConvWide = StrConvWide & "："
		Case ";"
			StrConvWide = StrConvWide & "；"
		Case "<"
			StrConvWide = StrConvWide & "＜"
		Case "="
			StrConvWide = StrConvWide & "＝"
		Case ">"
			StrConvWide = StrConvWide & "＞"
		Case "?"
			StrConvWide = StrConvWide & "？"
		Case "@"
			StrConvWide = StrConvWide & "＠"
		Case "A"
			StrConvWide = StrConvWide & "Ａ"
		Case "B"
			StrConvWide = StrConvWide & "Ｂ"
		Case "C"
			StrConvWide = StrConvWide & "Ｃ"
		Case "D"
			StrConvWide = StrConvWide & "Ｄ"
		Case "E"
			StrConvWide = StrConvWide & "Ｅ"
		Case "F"
			StrConvWide = StrConvWide & "Ｆ"
		Case "G"
			StrConvWide = StrConvWide & "Ｇ"
		Case "H"
			StrConvWide = StrConvWide & "Ｈ"
		Case "I"
			StrConvWide = StrConvWide & "Ｉ"
		Case "J"
			StrConvWide = StrConvWide & "Ｊ"
		Case "K"
			StrConvWide = StrConvWide & "Ｋ"
		Case "L"
			StrConvWide = StrConvWide & "Ｌ"
		Case "M"
			StrConvWide = StrConvWide & "Ｍ"
		Case "N"
			StrConvWide = StrConvWide & "Ｎ"
		Case "O"
			StrConvWide = StrConvWide & "Ｏ"
		Case "P"
			StrConvWide = StrConvWide & "Ｐ"
		Case "Q"
			StrConvWide = StrConvWide & "Ｑ"
		Case "R"
			StrConvWide = StrConvWide & "Ｒ"
		Case "S"
			StrConvWide = StrConvWide & "Ｓ"
		Case "T"
			StrConvWide = StrConvWide & "Ｔ"
		Case "U"
			StrConvWide = StrConvWide & "Ｕ"
		Case "V"
			StrConvWide = StrConvWide & "Ｖ"
		Case "W"
			StrConvWide = StrConvWide & "Ｗ"
		Case "X"
			StrConvWide = StrConvWide & "Ｘ"
		Case "Y"
			StrConvWide = StrConvWide & "Ｙ"
		Case "Z"
			StrConvWide = StrConvWide & "Ｚ"
		Case "["
			StrConvWide = StrConvWide & "［"
		Case "]"
			StrConvWide = StrConvWide & "］"
		Case "^"
			StrConvWide = StrConvWide & "＾"
		Case "_"
			StrConvWide = StrConvWide & "＿"
		Case "`"
			StrConvWide = StrConvWide & "｀"
		Case "a"
			StrConvWide = StrConvWide & "ａ"
		Case "b"
			StrConvWide = StrConvWide & "ｂ"
		Case "c"
			StrConvWide = StrConvWide & "ｃ"
		Case "d"
			StrConvWide = StrConvWide & "ｄ"
		Case "e"
			StrConvWide = StrConvWide & "ｅ"
		Case "f"
			StrConvWide = StrConvWide & "ｆ"
		Case "g"
			StrConvWide = StrConvWide & "ｇ"
		Case "h"
			StrConvWide = StrConvWide & "ｈ"
		Case "i"
			StrConvWide = StrConvWide & "ｉ"
		Case "j"
			StrConvWide = StrConvWide & "ｊ"
		Case "k"
			StrConvWide = StrConvWide & "ｋ"
		Case "l"
			StrConvWide = StrConvWide & "ｌ"
		Case "m"
			StrConvWide = StrConvWide & "ｍ"
		Case "n"
			StrConvWide = StrConvWide & "ｎ"
		Case "o"
			StrConvWide = StrConvWide & "ｏ"
		Case "p"
			StrConvWide = StrConvWide & "ｐ"
		Case "q"
			StrConvWide = StrConvWide & "ｑ"
		Case "r"
			StrConvWide = StrConvWide & "ｒ"
		Case "s"
			StrConvWide = StrConvWide & "ｓ"
		Case "t"
			StrConvWide = StrConvWide & "ｔ"
		Case "u"
			StrConvWide = StrConvWide & "ｕ"
		Case "v"
			StrConvWide = StrConvWide & "ｖ"
		Case "w"
			StrConvWide = StrConvWide & "ｗ"
		Case "x"
			StrConvWide = StrConvWide & "ｘ"
		Case "y"
			StrConvWide = StrConvWide & "ｙ"
		Case "z"
			StrConvWide = StrConvWide & "ｚ"
		Case "{"
			StrConvWide = StrConvWide & "｛"
		Case "|"
			StrConvWide = StrConvWide & "｜"
		Case "}"
			StrConvWide = StrConvWide & "｝"
		Case "~"
			StrConvWide = StrConvWide & "～"
		Case "うﾞ"
			StrConvWide = StrConvWide & "ヴ"
		Case "かﾞ"
			StrConvWide = StrConvWide & "が"
		Case "きﾞ"
			StrConvWide = StrConvWide & "ぎ"
		Case "くﾞ"
			StrConvWide = StrConvWide & "ぐ"
		Case "けﾞ"
			StrConvWide = StrConvWide & "げ"
		Case "こﾞ"
			StrConvWide = StrConvWide & "ご"
		Case "さﾞ"
			StrConvWide = StrConvWide & "ざ"
		Case "しﾞ"
			StrConvWide = StrConvWide & "じ"
		Case "すﾞ"
			StrConvWide = StrConvWide & "ず"
		Case "せﾞ"
			StrConvWide = StrConvWide & "ぜ"
		Case "そﾞ"
			StrConvWide = StrConvWide & "ぞ"
		Case "たﾞ"
			StrConvWide = StrConvWide & "だ"
		Case "ちﾞ"
			StrConvWide = StrConvWide & "ぢ"
		Case "つﾞ"
			StrConvWide = StrConvWide & "づ"
		Case "てﾞ"
			StrConvWide = StrConvWide & "で"
		Case "とﾞ"
			StrConvWide = StrConvWide & "ど"
		Case "はﾞ"
			StrConvWide = StrConvWide & "ば"
		Case "はﾟ"
			StrConvWide = StrConvWide & "ぱ"
		Case "ひﾞ"
			StrConvWide = StrConvWide & "び"
		Case "ひﾟ"
			StrConvWide = StrConvWide & "ぴ"
		Case "ふﾞ"
			StrConvWide = StrConvWide & "ぶ"
		Case "ふﾟ"
			StrConvWide = StrConvWide & "ぷ"
		Case "へﾞ"
			StrConvWide = StrConvWide & "べ"
		Case "へﾟ"
			StrConvWide = StrConvWide & "ぺ"
		Case "ほﾞ"
			StrConvWide = StrConvWide & "ぼ"
		Case "ほﾟ"
			StrConvWide = StrConvWide & "ぽ"
		Case "ウﾞ"
			StrConvWide = StrConvWide & "ヴ"
		Case "カﾞ"
			StrConvWide = StrConvWide & "ガ"
		Case "キﾞ"
			StrConvWide = StrConvWide & "ギ"
		Case "クﾞ"
			StrConvWide = StrConvWide & "グ"
		Case "ケﾞ"
			StrConvWide = StrConvWide & "ゲ"
		Case "コﾞ"
			StrConvWide = StrConvWide & "ゴ"
		Case "サﾞ"
			StrConvWide = StrConvWide & "ザ"
		Case "シﾞ"
			StrConvWide = StrConvWide & "ジ"
		Case "スﾞ"
			StrConvWide = StrConvWide & "ズ"
		Case "セﾞ"
			StrConvWide = StrConvWide & "ゼ"
		Case "ソﾞ"
			StrConvWide = StrConvWide & "ゾ"
		Case "タﾞ"
			StrConvWide = StrConvWide & "ダ"
		Case "チﾞ"
			StrConvWide = StrConvWide & "ヂ"
		Case "ツﾞ"
			StrConvWide = StrConvWide & "ヅ"
		Case "テﾞ"
			StrConvWide = StrConvWide & "デ"
		Case "トﾞ"
			StrConvWide = StrConvWide & "ド"
		Case "ハﾞ"
			StrConvWide = StrConvWide & "バ"
		Case "ハﾟ"
			StrConvWide = StrConvWide & "パ"
		Case "ヒﾞ"
			StrConvWide = StrConvWide & "ビ"
		Case "ヒﾟ"
			StrConvWide = StrConvWide & "ピ"
		Case "フﾞ"
			StrConvWide = StrConvWide & "ブ"
		Case "フﾟ"
			StrConvWide = StrConvWide & "プ"
		Case "ヘﾞ"
			StrConvWide = StrConvWide & "ベ"
		Case "ヘﾟ"
			StrConvWide = StrConvWide & "ペ"
		Case "ホﾞ"
			StrConvWide = StrConvWide & "ボ"
		Case "ホﾟ"
			StrConvWide = StrConvWide & "ポ"
		Case "｡"
			StrConvWide = StrConvWide & "。"
		Case "｢"
			StrConvWide = StrConvWide & "「"
		Case "｣"
			StrConvWide = StrConvWide & "」"
		Case "､"
			StrConvWide = StrConvWide & "、"
		Case "･"
			StrConvWide = StrConvWide & "・"
		Case "ｦ"
			StrConvWide = StrConvWide & "ヲ"
		Case "ｧ"
			StrConvWide = StrConvWide & "ァ"
		Case "ｨ"
			StrConvWide = StrConvWide & "ィ"
		Case "ｩ"
			StrConvWide = StrConvWide & "ゥ"
		Case "ｪ"
			StrConvWide = StrConvWide & "ェ"
		Case "ｫ"
			StrConvWide = StrConvWide & "ォ"
		Case "ｬ"
			StrConvWide = StrConvWide & "ャ"
		Case "ｭ"
			StrConvWide = StrConvWide & "ュ"
		Case "ｮ"
			StrConvWide = StrConvWide & "ョ"
		Case "ｯ"
			StrConvWide = StrConvWide & "ッ"
		Case "ｰ"
			StrConvWide = StrConvWide & "ー"
		Case "ｱ"
			StrConvWide = StrConvWide & "ア"
		Case "ｲ"
			StrConvWide = StrConvWide & "イ"
		Case "ｳ"
			StrConvWide = StrConvWide & "ウ"
		Case "ｳﾞ"
			StrConvWide = StrConvWide & "ヴ"
		Case "ｴ"
			StrConvWide = StrConvWide & "エ"
		Case "ｵ"
			StrConvWide = StrConvWide & "オ"
		Case "ｶ"
			StrConvWide = StrConvWide & "カ"
		Case "ｶﾞ"
			StrConvWide = StrConvWide & "ガ"
		Case "ｷ"
			StrConvWide = StrConvWide & "キ"
		Case "ｷﾞ"
			StrConvWide = StrConvWide & "ギ"
		Case "ｸ"
			StrConvWide = StrConvWide & "ク"
		Case "ｸﾞ"
			StrConvWide = StrConvWide & "グ"
		Case "ｹ"
			StrConvWide = StrConvWide & "ケ"
		Case "ｹﾞ"
			StrConvWide = StrConvWide & "ゲ"
		Case "ｺ"
			StrConvWide = StrConvWide & "コ"
		Case "ｺﾞ"
			StrConvWide = StrConvWide & "ゴ"
		Case "ｻ"
			StrConvWide = StrConvWide & "サ"
		Case "ｻﾞ"
			StrConvWide = StrConvWide & "ザ"
		Case "ｼ"
			StrConvWide = StrConvWide & "シ"
		Case "ｼﾞ"
			StrConvWide = StrConvWide & "ジ"
		Case "ｽ"
			StrConvWide = StrConvWide & "ス"
		Case "ｽﾞ"
			StrConvWide = StrConvWide & "ズ"
		Case "ｾ"
			StrConvWide = StrConvWide & "セ"
		Case "ｾﾞ"
			StrConvWide = StrConvWide & "ゼ"
		Case "ｿ"
			StrConvWide = StrConvWide & "ソ"
		Case "ｿﾞ"
			StrConvWide = StrConvWide & "ゾ"
		Case "ﾀ"
			StrConvWide = StrConvWide & "タ"
		Case "ﾀﾞ"
			StrConvWide = StrConvWide & "ダ"
		Case "ﾁ"
			StrConvWide = StrConvWide & "チ"
		Case "ﾁﾞ"
			StrConvWide = StrConvWide & "ヂ"
		Case "ﾂ"
			StrConvWide = StrConvWide & "ツ"
		Case "ﾂﾞ"
			StrConvWide = StrConvWide & "ヅ"
		Case "ﾃ"
			StrConvWide = StrConvWide & "テ"
		Case "ﾃﾞ"
			StrConvWide = StrConvWide & "デ"
		Case "ﾄ"
			StrConvWide = StrConvWide & "ト"
		Case "ﾄﾞ"
			StrConvWide = StrConvWide & "ド"
		Case "ﾅ"
			StrConvWide = StrConvWide & "ナ"
		Case "ﾆ"
			StrConvWide = StrConvWide & "ニ"
		Case "ﾇ"
			StrConvWide = StrConvWide & "ヌ"
		Case "ﾈ"
			StrConvWide = StrConvWide & "ネ"
		Case "ﾉ"
			StrConvWide = StrConvWide & "ノ"
		Case "ﾊ"
			StrConvWide = StrConvWide & "ハ"
		Case "ﾊﾞ"
			StrConvWide = StrConvWide & "バ"
		Case "ﾊﾟ"
			StrConvWide = StrConvWide & "パ"
		Case "ﾋ"
			StrConvWide = StrConvWide & "ヒ"
		Case "ﾋﾞ"
			StrConvWide = StrConvWide & "ビ"
		Case "ﾋﾟ"
			StrConvWide = StrConvWide & "ピ"
		Case "ﾌ"
			StrConvWide = StrConvWide & "フ"
		Case "ﾌﾞ"
			StrConvWide = StrConvWide & "ブ"
		Case "ﾌﾟ"
			StrConvWide = StrConvWide & "プ"
		Case "ﾍ"
			StrConvWide = StrConvWide & "ヘ"
		Case "ﾍﾞ"
			StrConvWide = StrConvWide & "ベ"
		Case "ﾍﾟ"
			StrConvWide = StrConvWide & "ペ"
		Case "ﾎ"
			StrConvWide = StrConvWide & "ホ"
		Case "ﾎﾞ"
			StrConvWide = StrConvWide & "ボ"
		Case "ﾎﾟ"
			StrConvWide = StrConvWide & "ポ"
		Case "ﾏ"
			StrConvWide = StrConvWide & "マ"
		Case "ﾐ"
			StrConvWide = StrConvWide & "ミ"
		Case "ﾑ"
			StrConvWide = StrConvWide & "ム"
		Case "ﾒ"
			StrConvWide = StrConvWide & "メ"
		Case "ﾓ"
			StrConvWide = StrConvWide & "モ"
		Case "ﾔ"
			StrConvWide = StrConvWide & "ヤ"
		Case "ﾕ"
			StrConvWide = StrConvWide & "ユ"
		Case "ﾖ"
			StrConvWide = StrConvWide & "ヨ"
		Case "ﾗ"
			StrConvWide = StrConvWide & "ラ"
		Case "ﾘ"
			StrConvWide = StrConvWide & "リ"
		Case "ﾙ"
			StrConvWide = StrConvWide & "ル"
		Case "ﾚ"
			StrConvWide = StrConvWide & "レ"
		Case "ﾛ"
			StrConvWide = StrConvWide & "ロ"
		Case "ﾜ"
			StrConvWide = StrConvWide & "ワ"
		Case "ﾝ"
			StrConvWide = StrConvWide & "ン"
		Case "ﾞ"
			StrConvWide = StrConvWide & "゛"
		Case "ﾟ"
			StrConvWide = StrConvWide & "゜"
		Case Else
			StrConvWide = StrConvWide & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvNarrow
'---------------------------------------------------
' 用途 : StrConv(s,vbNarrow) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvNarrow(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("ﾞﾟ", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "　"
			StrConvNarrow = StrConvNarrow & " "
		Case "、"
			StrConvNarrow = StrConvNarrow & "､"
		Case "。"
			StrConvNarrow = StrConvNarrow & "｡"
		Case "，"
			StrConvNarrow = StrConvNarrow & ","
		Case "．"
			StrConvNarrow = StrConvNarrow & "."
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "："
			StrConvNarrow = StrConvNarrow & ":"
		Case "；"
			StrConvNarrow = StrConvNarrow & ";"
		Case "？"
			StrConvNarrow = StrConvNarrow & "?"
		Case "！"
			StrConvNarrow = StrConvNarrow & "!"
		Case "゛"
			StrConvNarrow = StrConvNarrow & "ﾞ"
		Case "゜"
			StrConvNarrow = StrConvNarrow & "ﾟ"
		Case "｀"
			StrConvNarrow = StrConvNarrow & "`"
		Case "＾"
			StrConvNarrow = StrConvNarrow & "^"
		Case "＿"
			StrConvNarrow = StrConvNarrow & "_"
		Case "ー"
			StrConvNarrow = StrConvNarrow & "ｰ"
		Case "／"
			StrConvNarrow = StrConvNarrow & "/"
		Case "～"
			StrConvNarrow = StrConvNarrow & "~"
		Case "｜"
			StrConvNarrow = StrConvNarrow & "|"
		Case "‘"
			StrConvNarrow = StrConvNarrow & "'"
		Case "’"
			StrConvNarrow = StrConvNarrow & "'"
		Case "“"
			StrConvNarrow = StrConvNarrow & """"
		Case "”"
			StrConvNarrow = StrConvNarrow & """"
		Case "（"
			StrConvNarrow = StrConvNarrow & "("
		Case "）"
			StrConvNarrow = StrConvNarrow & ")"
		Case "［"
			StrConvNarrow = StrConvNarrow & "["
		Case "］"
			StrConvNarrow = StrConvNarrow & "]"
		Case "｛"
			StrConvNarrow = StrConvNarrow & "{"
		Case "｝"
			StrConvNarrow = StrConvNarrow & "}"
		Case "「"
			StrConvNarrow = StrConvNarrow & "｢"
		Case "」"
			StrConvNarrow = StrConvNarrow & "｣"
		Case "＋"
			StrConvNarrow = StrConvNarrow & "+"
		Case "－"
			StrConvNarrow = StrConvNarrow & "-"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "＝"
			StrConvNarrow = StrConvNarrow & "="
		Case "＜"
			StrConvNarrow = StrConvNarrow & "<"
		Case "＞"
			StrConvNarrow = StrConvNarrow & ">"
		Case "￥"
			StrConvNarrow = StrConvNarrow & "\"
		Case "＄"
			StrConvNarrow = StrConvNarrow & "$"
		Case "％"
			StrConvNarrow = StrConvNarrow & "%"
		Case "＃"
			StrConvNarrow = StrConvNarrow & "#"
		Case "＆"
			StrConvNarrow = StrConvNarrow & "&"
		Case "＊"
			StrConvNarrow = StrConvNarrow & "*"
		Case "＠"
			StrConvNarrow = StrConvNarrow & "@"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "０"
			StrConvNarrow = StrConvNarrow & "0"
		Case "１"
			StrConvNarrow = StrConvNarrow & "1"
		Case "２"
			StrConvNarrow = StrConvNarrow & "2"
		Case "３"
			StrConvNarrow = StrConvNarrow & "3"
		Case "４"
			StrConvNarrow = StrConvNarrow & "4"
		Case "５"
			StrConvNarrow = StrConvNarrow & "5"
		Case "６"
			StrConvNarrow = StrConvNarrow & "6"
		Case "７"
			StrConvNarrow = StrConvNarrow & "7"
		Case "８"
			StrConvNarrow = StrConvNarrow & "8"
		Case "９"
			StrConvNarrow = StrConvNarrow & "9"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "Ａ"
			StrConvNarrow = StrConvNarrow & "A"
		Case "Ｂ"
			StrConvNarrow = StrConvNarrow & "B"
		Case "Ｃ"
			StrConvNarrow = StrConvNarrow & "C"
		Case "Ｄ"
			StrConvNarrow = StrConvNarrow & "D"
		Case "Ｅ"
			StrConvNarrow = StrConvNarrow & "E"
		Case "Ｆ"
			StrConvNarrow = StrConvNarrow & "F"
		Case "Ｇ"
			StrConvNarrow = StrConvNarrow & "G"
		Case "Ｈ"
			StrConvNarrow = StrConvNarrow & "H"
		Case "Ｉ"
			StrConvNarrow = StrConvNarrow & "I"
		Case "Ｊ"
			StrConvNarrow = StrConvNarrow & "J"
		Case "Ｋ"
			StrConvNarrow = StrConvNarrow & "K"
		Case "Ｌ"
			StrConvNarrow = StrConvNarrow & "L"
		Case "Ｍ"
			StrConvNarrow = StrConvNarrow & "M"
		Case "Ｎ"
			StrConvNarrow = StrConvNarrow & "N"
		Case "Ｏ"
			StrConvNarrow = StrConvNarrow & "O"
		Case "Ｐ"
			StrConvNarrow = StrConvNarrow & "P"
		Case "Ｑ"
			StrConvNarrow = StrConvNarrow & "Q"
		Case "Ｒ"
			StrConvNarrow = StrConvNarrow & "R"
		Case "Ｓ"
			StrConvNarrow = StrConvNarrow & "S"
		Case "Ｔ"
			StrConvNarrow = StrConvNarrow & "T"
		Case "Ｕ"
			StrConvNarrow = StrConvNarrow & "U"
		Case "Ｖ"
			StrConvNarrow = StrConvNarrow & "V"
		Case "Ｗ"
			StrConvNarrow = StrConvNarrow & "W"
		Case "Ｘ"
			StrConvNarrow = StrConvNarrow & "X"
		Case "Ｙ"
			StrConvNarrow = StrConvNarrow & "Y"
		Case "Ｚ"
			StrConvNarrow = StrConvNarrow & "Z"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "ａ"
			StrConvNarrow = StrConvNarrow & "a"
		Case "ｂ"
			StrConvNarrow = StrConvNarrow & "b"
		Case "ｃ"
			StrConvNarrow = StrConvNarrow & "c"
		Case "ｄ"
			StrConvNarrow = StrConvNarrow & "d"
		Case "ｅ"
			StrConvNarrow = StrConvNarrow & "e"
		Case "ｆ"
			StrConvNarrow = StrConvNarrow & "f"
		Case "ｇ"
			StrConvNarrow = StrConvNarrow & "g"
		Case "ｈ"
			StrConvNarrow = StrConvNarrow & "h"
		Case "ｉ"
			StrConvNarrow = StrConvNarrow & "i"
		Case "ｊ"
			StrConvNarrow = StrConvNarrow & "j"
		Case "ｋ"
			StrConvNarrow = StrConvNarrow & "k"
		Case "ｌ"
			StrConvNarrow = StrConvNarrow & "l"
		Case "ｍ"
			StrConvNarrow = StrConvNarrow & "m"
		Case "ｎ"
			StrConvNarrow = StrConvNarrow & "n"
		Case "ｏ"
			StrConvNarrow = StrConvNarrow & "o"
		Case "ｐ"
			StrConvNarrow = StrConvNarrow & "p"
		Case "ｑ"
			StrConvNarrow = StrConvNarrow & "q"
		Case "ｒ"
			StrConvNarrow = StrConvNarrow & "r"
		Case "ｓ"
			StrConvNarrow = StrConvNarrow & "s"
		Case "ｔ"
			StrConvNarrow = StrConvNarrow & "t"
		Case "ｕ"
			StrConvNarrow = StrConvNarrow & "u"
		Case "ｖ"
			StrConvNarrow = StrConvNarrow & "v"
		Case "ｗ"
			StrConvNarrow = StrConvNarrow & "w"
		Case "ｘ"
			StrConvNarrow = StrConvNarrow & "x"
		Case "ｙ"
			StrConvNarrow = StrConvNarrow & "y"
		Case "ｚ"
			StrConvNarrow = StrConvNarrow & "z"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "ァ"
			StrConvNarrow = StrConvNarrow & "ｧ"
		Case "ア"
			StrConvNarrow = StrConvNarrow & "ｱ"
		Case "ィ"
			StrConvNarrow = StrConvNarrow & "ｨ"
		Case "イ"
			StrConvNarrow = StrConvNarrow & "ｲ"
		Case "ゥ"
			StrConvNarrow = StrConvNarrow & "ｩ"
		Case "ウ"
			StrConvNarrow = StrConvNarrow & "ｳ"
		Case "ェ"
			StrConvNarrow = StrConvNarrow & "ｪ"
		Case "エ"
			StrConvNarrow = StrConvNarrow & "ｴ"
		Case "ォ"
			StrConvNarrow = StrConvNarrow & "ｫ"
		Case "オ"
			StrConvNarrow = StrConvNarrow & "ｵ"
		Case "カ"
			StrConvNarrow = StrConvNarrow & "ｶ"
		Case "ガ"
			StrConvNarrow = StrConvNarrow & "ｶﾞ"
		Case "キ"
			StrConvNarrow = StrConvNarrow & "ｷ"
		Case "ギ"
			StrConvNarrow = StrConvNarrow & "ｷﾞ"
		Case "ク"
			StrConvNarrow = StrConvNarrow & "ｸ"
		Case "グ"
			StrConvNarrow = StrConvNarrow & "ｸﾞ"
		Case "ケ"
			StrConvNarrow = StrConvNarrow & "ｹ"
		Case "ゲ"
			StrConvNarrow = StrConvNarrow & "ｹﾞ"
		Case "コ"
			StrConvNarrow = StrConvNarrow & "ｺ"
		Case "ゴ"
			StrConvNarrow = StrConvNarrow & "ｺﾞ"
		Case "サ"
			StrConvNarrow = StrConvNarrow & "ｻ"
		Case "ザ"
			StrConvNarrow = StrConvNarrow & "ｻﾞ"
		Case "シ"
			StrConvNarrow = StrConvNarrow & "ｼ"
		Case "ジ"
			StrConvNarrow = StrConvNarrow & "ｼﾞ"
		Case "ス"
			StrConvNarrow = StrConvNarrow & "ｽ"
		Case "ズ"
			StrConvNarrow = StrConvNarrow & "ｽﾞ"
		Case "セ"
			StrConvNarrow = StrConvNarrow & "ｾ"
		Case "ゼ"
			StrConvNarrow = StrConvNarrow & "ｾﾞ"
		Case "ソ"
			StrConvNarrow = StrConvNarrow & "ｿ"
		Case "ゾ"
			StrConvNarrow = StrConvNarrow & "ｿﾞ"
		Case "タ"
			StrConvNarrow = StrConvNarrow & "ﾀ"
		Case "ダ"
			StrConvNarrow = StrConvNarrow & "ﾀﾞ"
		Case "チ"
			StrConvNarrow = StrConvNarrow & "ﾁ"
		Case "ヂ"
			StrConvNarrow = StrConvNarrow & "ﾁﾞ"
		Case "ッ"
			StrConvNarrow = StrConvNarrow & "ｯ"
		Case "ツ"
			StrConvNarrow = StrConvNarrow & "ﾂ"
		Case "ヅ"
			StrConvNarrow = StrConvNarrow & "ﾂﾞ"
		Case "テ"
			StrConvNarrow = StrConvNarrow & "ﾃ"
		Case "デ"
			StrConvNarrow = StrConvNarrow & "ﾃﾞ"
		Case "ト"
			StrConvNarrow = StrConvNarrow & "ﾄ"
		Case "ド"
			StrConvNarrow = StrConvNarrow & "ﾄﾞ"
		Case "ナ"
			StrConvNarrow = StrConvNarrow & "ﾅ"
		Case "ニ"
			StrConvNarrow = StrConvNarrow & "ﾆ"
		Case "ヌ"
			StrConvNarrow = StrConvNarrow & "ﾇ"
		Case "ネ"
			StrConvNarrow = StrConvNarrow & "ﾈ"
		Case "ノ"
			StrConvNarrow = StrConvNarrow & "ﾉ"
		Case "ハ"
			StrConvNarrow = StrConvNarrow & "ﾊ"
		Case "バ"
			StrConvNarrow = StrConvNarrow & "ﾊﾞ"
		Case "パ"
			StrConvNarrow = StrConvNarrow & "ﾊﾟ"
		Case "ヒ"
			StrConvNarrow = StrConvNarrow & "ﾋ"
		Case "ビ"
			StrConvNarrow = StrConvNarrow & "ﾋﾞ"
		Case "ピ"
			StrConvNarrow = StrConvNarrow & "ﾋﾟ"
		Case "フ"
			StrConvNarrow = StrConvNarrow & "ﾌ"
		Case "ブ"
			StrConvNarrow = StrConvNarrow & "ﾌﾞ"
		Case "プ"
			StrConvNarrow = StrConvNarrow & "ﾌﾟ"
		Case "ヘ"
			StrConvNarrow = StrConvNarrow & "ﾍ"
		Case "ベ"
			StrConvNarrow = StrConvNarrow & "ﾍﾞ"
		Case "ペ"
			StrConvNarrow = StrConvNarrow & "ﾍﾟ"
		Case "ホ"
			StrConvNarrow = StrConvNarrow & "ﾎ"
		Case "ボ"
			StrConvNarrow = StrConvNarrow & "ﾎﾞ"
		Case "ポ"
			StrConvNarrow = StrConvNarrow & "ﾎﾟ"
		Case "マ"
			StrConvNarrow = StrConvNarrow & "ﾏ"
		Case "ミ"
			StrConvNarrow = StrConvNarrow & "ﾐ"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "ム"
			StrConvNarrow = StrConvNarrow & "ﾑ"
		Case "メ"
			StrConvNarrow = StrConvNarrow & "ﾒ"
		Case "モ"
			StrConvNarrow = StrConvNarrow & "ﾓ"
		Case "ャ"
			StrConvNarrow = StrConvNarrow & "ｬ"
		Case "ヤ"
			StrConvNarrow = StrConvNarrow & "ﾔ"
		Case "ュ"
			StrConvNarrow = StrConvNarrow & "ｭ"
		Case "ユ"
			StrConvNarrow = StrConvNarrow & "ﾕ"
		Case "ョ"
			StrConvNarrow = StrConvNarrow & "ｮ"
		Case "ヨ"
			StrConvNarrow = StrConvNarrow & "ﾖ"
		Case "ラ"
			StrConvNarrow = StrConvNarrow & "ﾗ"
		Case "リ"
			StrConvNarrow = StrConvNarrow & "ﾘ"
		Case "ル"
			StrConvNarrow = StrConvNarrow & "ﾙ"
		Case "レ"
			StrConvNarrow = StrConvNarrow & "ﾚ"
		Case "ロ"
			StrConvNarrow = StrConvNarrow & "ﾛ"
		Case "ワ"
			StrConvNarrow = StrConvNarrow & "ﾜ"
		Case "ヲ"
			StrConvNarrow = StrConvNarrow & "ｦ"
		Case "ン"
			StrConvNarrow = StrConvNarrow & "ﾝ"
		Case "ヴ"
			StrConvNarrow = StrConvNarrow & "ｳﾞ"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "￤"
			StrConvNarrow = StrConvNarrow & "|"
		Case "＇"
			StrConvNarrow = StrConvNarrow & "'"
		Case "＂"
			StrConvNarrow = StrConvNarrow & """"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case "￤"
			StrConvNarrow = StrConvNarrow & "|"
		Case "＇"
			StrConvNarrow = StrConvNarrow & "'"
		Case "＂"
			StrConvNarrow = StrConvNarrow & """"
		Case "・"
			StrConvNarrow = StrConvNarrow & "･"
		Case Else
			StrConvNarrow = StrConvNarrow & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvKatakana
'---------------------------------------------------
' 用途 : StrConv(s,vbKatakana) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvKatakana(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("ﾞﾟ", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case "ゝ"
			StrConvKatakana = StrConvKatakana & "ヽ"
		Case "ゞ"
			StrConvKatakana = StrConvKatakana & "ヾ"
		Case "ぁ"
			StrConvKatakana = StrConvKatakana & "ァ"
		Case "あ"
			StrConvKatakana = StrConvKatakana & "ア"
		Case "ぃ"
			StrConvKatakana = StrConvKatakana & "ィ"
		Case "い"
			StrConvKatakana = StrConvKatakana & "イ"
		Case "ぅ"
			StrConvKatakana = StrConvKatakana & "ゥ"
		Case "う"
			StrConvKatakana = StrConvKatakana & "ウ"
		Case "ぇ"
			StrConvKatakana = StrConvKatakana & "ェ"
		Case "え"
			StrConvKatakana = StrConvKatakana & "エ"
		Case "ぉ"
			StrConvKatakana = StrConvKatakana & "ォ"
		Case "お"
			StrConvKatakana = StrConvKatakana & "オ"
		Case "か"
			StrConvKatakana = StrConvKatakana & "カ"
		Case "が"
			StrConvKatakana = StrConvKatakana & "ガ"
		Case "き"
			StrConvKatakana = StrConvKatakana & "キ"
		Case "ぎ"
			StrConvKatakana = StrConvKatakana & "ギ"
		Case "く"
			StrConvKatakana = StrConvKatakana & "ク"
		Case "ぐ"
			StrConvKatakana = StrConvKatakana & "グ"
		Case "け"
			StrConvKatakana = StrConvKatakana & "ケ"
		Case "げ"
			StrConvKatakana = StrConvKatakana & "ゲ"
		Case "こ"
			StrConvKatakana = StrConvKatakana & "コ"
		Case "ご"
			StrConvKatakana = StrConvKatakana & "ゴ"
		Case "さ"
			StrConvKatakana = StrConvKatakana & "サ"
		Case "ざ"
			StrConvKatakana = StrConvKatakana & "ザ"
		Case "し"
			StrConvKatakana = StrConvKatakana & "シ"
		Case "じ"
			StrConvKatakana = StrConvKatakana & "ジ"
		Case "す"
			StrConvKatakana = StrConvKatakana & "ス"
		Case "ず"
			StrConvKatakana = StrConvKatakana & "ズ"
		Case "せ"
			StrConvKatakana = StrConvKatakana & "セ"
		Case "ぜ"
			StrConvKatakana = StrConvKatakana & "ゼ"
		Case "そ"
			StrConvKatakana = StrConvKatakana & "ソ"
		Case "ぞ"
			StrConvKatakana = StrConvKatakana & "ゾ"
		Case "た"
			StrConvKatakana = StrConvKatakana & "タ"
		Case "だ"
			StrConvKatakana = StrConvKatakana & "ダ"
		Case "ち"
			StrConvKatakana = StrConvKatakana & "チ"
		Case "ぢ"
			StrConvKatakana = StrConvKatakana & "ヂ"
		Case "っ"
			StrConvKatakana = StrConvKatakana & "ッ"
		Case "つ"
			StrConvKatakana = StrConvKatakana & "ツ"
		Case "づ"
			StrConvKatakana = StrConvKatakana & "ヅ"
		Case "て"
			StrConvKatakana = StrConvKatakana & "テ"
		Case "で"
			StrConvKatakana = StrConvKatakana & "デ"
		Case "と"
			StrConvKatakana = StrConvKatakana & "ト"
		Case "ど"
			StrConvKatakana = StrConvKatakana & "ド"
		Case "な"
			StrConvKatakana = StrConvKatakana & "ナ"
		Case "に"
			StrConvKatakana = StrConvKatakana & "ニ"
		Case "ぬ"
			StrConvKatakana = StrConvKatakana & "ヌ"
		Case "ね"
			StrConvKatakana = StrConvKatakana & "ネ"
		Case "の"
			StrConvKatakana = StrConvKatakana & "ノ"
		Case "は"
			StrConvKatakana = StrConvKatakana & "ハ"
		Case "ば"
			StrConvKatakana = StrConvKatakana & "バ"
		Case "ぱ"
			StrConvKatakana = StrConvKatakana & "パ"
		Case "ひ"
			StrConvKatakana = StrConvKatakana & "ヒ"
		Case "び"
			StrConvKatakana = StrConvKatakana & "ビ"
		Case "ぴ"
			StrConvKatakana = StrConvKatakana & "ピ"
		Case "ふ"
			StrConvKatakana = StrConvKatakana & "フ"
		Case "ぶ"
			StrConvKatakana = StrConvKatakana & "ブ"
		Case "ぷ"
			StrConvKatakana = StrConvKatakana & "プ"
		Case "へ"
			StrConvKatakana = StrConvKatakana & "ヘ"
		Case "べ"
			StrConvKatakana = StrConvKatakana & "ベ"
		Case "ぺ"
			StrConvKatakana = StrConvKatakana & "ペ"
		Case "ほ"
			StrConvKatakana = StrConvKatakana & "ホ"
		Case "ぼ"
			StrConvKatakana = StrConvKatakana & "ボ"
		Case "ぽ"
			StrConvKatakana = StrConvKatakana & "ポ"
		Case "ま"
			StrConvKatakana = StrConvKatakana & "マ"
		Case "み"
			StrConvKatakana = StrConvKatakana & "ミ"
		Case "む"
			StrConvKatakana = StrConvKatakana & "ム"
		Case "め"
			StrConvKatakana = StrConvKatakana & "メ"
		Case "も"
			StrConvKatakana = StrConvKatakana & "モ"
		Case "ゃ"
			StrConvKatakana = StrConvKatakana & "ャ"
		Case "や"
			StrConvKatakana = StrConvKatakana & "ヤ"
		Case "ゅ"
			StrConvKatakana = StrConvKatakana & "ュ"
		Case "ゆ"
			StrConvKatakana = StrConvKatakana & "ユ"
		Case "ょ"
			StrConvKatakana = StrConvKatakana & "ョ"
		Case "よ"
			StrConvKatakana = StrConvKatakana & "ヨ"
		Case "ら"
			StrConvKatakana = StrConvKatakana & "ラ"
		Case "り"
			StrConvKatakana = StrConvKatakana & "リ"
		Case "る"
			StrConvKatakana = StrConvKatakana & "ル"
		Case "れ"
			StrConvKatakana = StrConvKatakana & "レ"
		Case "ろ"
			StrConvKatakana = StrConvKatakana & "ロ"
		Case "ゎ"
			StrConvKatakana = StrConvKatakana & "ヮ"
		Case "わ"
			StrConvKatakana = StrConvKatakana & "ワ"
		Case "ゐ"
			StrConvKatakana = StrConvKatakana & "ヰ"
		Case "ゑ"
			StrConvKatakana = StrConvKatakana & "ヱ"
		Case "を"
			StrConvKatakana = StrConvKatakana & "ヲ"
		Case "ん"
			StrConvKatakana = StrConvKatakana & "ン"
		Case Else
			StrConvKatakana = StrConvKatakana & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvHiragana
'---------------------------------------------------
' 用途 : StrConv(s,vbHiragana) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvHiragana(s)
	Dim nCnt
	Dim nLen
	Dim sChr
	Dim sMud

	nLen = Len(s)
	For nCnt = 1 To nLen
		sChr = Mid(s,nCnt,1)
		sMud = Mid(s,nCnt+1,1) 
		If InStr("ﾞﾟ", sMud) Then
			sChr = sChr & sMud
			nCnt = nCnt + 1
		End If
		Select Case sChr
		Case "ヽ"
			StrConvHiragana = StrConvHiragana & "ゝ"
		Case "ヾ"
			StrConvHiragana = StrConvHiragana & "ゞ"
		Case "ァ"
			StrConvHiragana = StrConvHiragana & "ぁ"
		Case "ア"
			StrConvHiragana = StrConvHiragana & "あ"
		Case "ィ"
			StrConvHiragana = StrConvHiragana & "ぃ"
		Case "イ"
			StrConvHiragana = StrConvHiragana & "い"
		Case "ゥ"
			StrConvHiragana = StrConvHiragana & "ぅ"
		Case "ウ"
			StrConvHiragana = StrConvHiragana & "う"
		Case "ェ"
			StrConvHiragana = StrConvHiragana & "ぇ"
		Case "エ"
			StrConvHiragana = StrConvHiragana & "え"
		Case "ォ"
			StrConvHiragana = StrConvHiragana & "ぉ"
		Case "オ"
			StrConvHiragana = StrConvHiragana & "お"
		Case "カ"
			StrConvHiragana = StrConvHiragana & "か"
		Case "ガ"
			StrConvHiragana = StrConvHiragana & "が"
		Case "キ"
			StrConvHiragana = StrConvHiragana & "き"
		Case "ギ"
			StrConvHiragana = StrConvHiragana & "ぎ"
		Case "ク"
			StrConvHiragana = StrConvHiragana & "く"
		Case "グ"
			StrConvHiragana = StrConvHiragana & "ぐ"
		Case "ケ"
			StrConvHiragana = StrConvHiragana & "け"
		Case "ゲ"
			StrConvHiragana = StrConvHiragana & "げ"
		Case "コ"
			StrConvHiragana = StrConvHiragana & "こ"
		Case "ゴ"
			StrConvHiragana = StrConvHiragana & "ご"
		Case "サ"
			StrConvHiragana = StrConvHiragana & "さ"
		Case "ザ"
			StrConvHiragana = StrConvHiragana & "ざ"
		Case "シ"
			StrConvHiragana = StrConvHiragana & "し"
		Case "ジ"
			StrConvHiragana = StrConvHiragana & "じ"
		Case "ス"
			StrConvHiragana = StrConvHiragana & "す"
		Case "ズ"
			StrConvHiragana = StrConvHiragana & "ず"
		Case "セ"
			StrConvHiragana = StrConvHiragana & "せ"
		Case "ゼ"
			StrConvHiragana = StrConvHiragana & "ぜ"
		Case "ソ"
			StrConvHiragana = StrConvHiragana & "そ"
		Case "ゾ"
			StrConvHiragana = StrConvHiragana & "ぞ"
		Case "タ"
			StrConvHiragana = StrConvHiragana & "た"
		Case "ダ"
			StrConvHiragana = StrConvHiragana & "だ"
		Case "チ"
			StrConvHiragana = StrConvHiragana & "ち"
		Case "ヂ"
			StrConvHiragana = StrConvHiragana & "ぢ"
		Case "ッ"
			StrConvHiragana = StrConvHiragana & "っ"
		Case "ツ"
			StrConvHiragana = StrConvHiragana & "つ"
		Case "ヅ"
			StrConvHiragana = StrConvHiragana & "づ"
		Case "テ"
			StrConvHiragana = StrConvHiragana & "て"
		Case "デ"
			StrConvHiragana = StrConvHiragana & "で"
		Case "ト"
			StrConvHiragana = StrConvHiragana & "と"
		Case "ド"
			StrConvHiragana = StrConvHiragana & "ど"
		Case "ナ"
			StrConvHiragana = StrConvHiragana & "な"
		Case "ニ"
			StrConvHiragana = StrConvHiragana & "に"
		Case "ヌ"
			StrConvHiragana = StrConvHiragana & "ぬ"
		Case "ネ"
			StrConvHiragana = StrConvHiragana & "ね"
		Case "ノ"
			StrConvHiragana = StrConvHiragana & "の"
		Case "ハ"
			StrConvHiragana = StrConvHiragana & "は"
		Case "バ"
			StrConvHiragana = StrConvHiragana & "ば"
		Case "パ"
			StrConvHiragana = StrConvHiragana & "ぱ"
		Case "ヒ"
			StrConvHiragana = StrConvHiragana & "ひ"
		Case "ビ"
			StrConvHiragana = StrConvHiragana & "び"
		Case "ピ"
			StrConvHiragana = StrConvHiragana & "ぴ"
		Case "フ"
			StrConvHiragana = StrConvHiragana & "ふ"
		Case "ブ"
			StrConvHiragana = StrConvHiragana & "ぶ"
		Case "プ"
			StrConvHiragana = StrConvHiragana & "ぷ"
		Case "ヘ"
			StrConvHiragana = StrConvHiragana & "へ"
		Case "ベ"
			StrConvHiragana = StrConvHiragana & "べ"
		Case "ペ"
			StrConvHiragana = StrConvHiragana & "ぺ"
		Case "ホ"
			StrConvHiragana = StrConvHiragana & "ほ"
		Case "ボ"
			StrConvHiragana = StrConvHiragana & "ぼ"
		Case "ポ"
			StrConvHiragana = StrConvHiragana & "ぽ"
		Case "マ"
			StrConvHiragana = StrConvHiragana & "ま"
		Case "ミ"
			StrConvHiragana = StrConvHiragana & "み"
		Case "ム"
			StrConvHiragana = StrConvHiragana & "む"
		Case "メ"
			StrConvHiragana = StrConvHiragana & "め"
		Case "モ"
			StrConvHiragana = StrConvHiragana & "も"
		Case "ャ"
			StrConvHiragana = StrConvHiragana & "ゃ"
		Case "ヤ"
			StrConvHiragana = StrConvHiragana & "や"
		Case "ュ"
			StrConvHiragana = StrConvHiragana & "ゅ"
		Case "ユ"
			StrConvHiragana = StrConvHiragana & "ゆ"
		Case "ョ"
			StrConvHiragana = StrConvHiragana & "ょ"
		Case "ヨ"
			StrConvHiragana = StrConvHiragana & "よ"
		Case "ラ"
			StrConvHiragana = StrConvHiragana & "ら"
		Case "リ"
			StrConvHiragana = StrConvHiragana & "り"
		Case "ル"
			StrConvHiragana = StrConvHiragana & "る"
		Case "レ"
			StrConvHiragana = StrConvHiragana & "れ"
		Case "ロ"
			StrConvHiragana = StrConvHiragana & "ろ"
		Case "ヮ"
			StrConvHiragana = StrConvHiragana & "ゎ"
		Case "ワ"
			StrConvHiragana = StrConvHiragana & "わ"
		Case "ヰ"
			StrConvHiragana = StrConvHiragana & "ゐ"
		Case "ヱ"
			StrConvHiragana = StrConvHiragana & "ゑ"
		Case "ヲ"
			StrConvHiragana = StrConvHiragana & "を"
		Case "ン"
			StrConvHiragana = StrConvHiragana & "ん"
		Case Else
			StrConvHiragana = StrConvHiragana & sChr
		End Select
	Next
End Function

'***************************************************
' StrConvUnicode
'---------------------------------------------------
' 用途 : StrConv(s,vbUnicode) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvUnicode(sInp)
	Dim nCnt
	Dim nLen
	Dim nAsc
	Dim nChr
	nLen = LenB(sInp)
	For nCnt = 1 To nLen
		nAsc = AscB(MidB(sInp, nCnt, 1))
		If (&h81 <= nAsc And nAsc <= &h9F) Or (&hE0 <= nAsc And nAsc <= &hEF) Then
			nChr = nAsc * 256 + AscB(MidB(sInp, nCnt+1, 1))
			StrConvUnicode = StrConvUnicode & Chr(nChr)
			nCnt = nCnt + 1
		Else
			StrConvUnicode = StrConvUnicode & Chr(AscB(MidB(sInp, nCnt, 1)))
		End If
	Next
End Function

'***************************************************
' StrConvFromUnicode
'---------------------------------------------------
' 用途 : StrConv(s,vbFromUnicode) のクローン
' 引数 : 変換する文字列
' 戻値 : 変換された文字列
'***************************************************
Function StrConvFromUnicode(sInp)
	Dim nCnt
	Dim nLen
	Dim nAsc
	nLen = Len(sInp)
	For nCnt = 1 to nLen
		nAsc = Asc(Mid(sInp, nCnt, 1))
		If nAsc And &hFF00 Then
			StrConvFromUnicode = StrConvFromUnicode & ChrB(Int(nAsc / 256) And &hFF)
			StrConvFromUnicode = StrConvFromUnicode & ChrB(nAsc And &hFF)
		Else
			StrConvFromUnicode = StrConvFromUnicode & ChrB(nAsc)
		End If
	Next
End Function

'***************************************************
' StrConv が使用する定数郡
'***************************************************
' Enum VbStrConv
Const vbUpperCase=1
Const vbLowerCase=2
Const vbProperCase=3
Const vbWide=4
Const vbNarrow=8
Const vbKatakana=16
Const vbHiragana=32
Const vbUnicode = 64
Const vbFromUnicode = 128

'***************************************************
' StrConv
'---------------------------------------------------
' 引数 : 変換する文字列,変換処理
' 戻値 : 変換された文字列
'***************************************************
Function StrConv(sInp,eCnv)
	StrConv = sInp
	' Cnv に対して処理を振り分け
	If eCnv And vbUpperCase Then
		StrConv = StrConvUpperCase(StrConv)
	End If
	If eCnv And vbLowerCase Then
		StrConv = StrConvLowerCase(StrConv)
	End If
	If eCnv = vbProperCase Then
		StrConv = StrConvProperCase(StrConv)
	End If
	If eCnv And vbWide Then
		StrConv = StrConvWide(StrConv)
	End If
	If eCnv And vbNarrow Then
		StrConv = StrConvNarrow(StrConv)
	End If
	If eCnv And vbKatakana Then
		StrConv = StrConvKatakana(StrConv)
	End If
	If eCnv And vbHiragana Then
		StrConv = StrConvHiragana(StrConv)
	End If
	If eCnv And vbUnicode Then
		StrConv = StrConvUnicode(StrConv)
	End If
	If eCnv And vbFromUnicode Then
		StrConv = StrConvFromUnicode(StrConv)
	End If
End Function

Function GetQty(strB,strMode)
	dim	nLen
	dim	nCnt
	dim	nChr
	dim	strRet

	strRet = ""
	nLen = Len(strB)
	For nCnt = 1 to nLen
		nChr = Mid(strB, nCnt, 1)
		if nChr = "0" then
			nChr = strMode
		else
		end if
		strRet = strRet & nChr
	Next
	GetQty = strRet
End Function

Function GetSaisu(strSaisu)
	dim	dblSaisu
	dim	lngSaisu
	
	lngSaisu = 0
	dblSaisu	= cdbl(Rtrim(strSaisu))
	if dblSaisu > 0 then
		lngSaisu = round(dblSaisu,0)
		if lngSaisu = 0 then
			lngSaisu = 1
		end if
	end if
	GetSaisu = lngSaisu
End Function

