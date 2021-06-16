<%
' include�p
function GetShimeDt20(byval dt,byval intShime)
	dim	intYear
	dim	intMonth
	dim	intDay

	intYear		= Year(dt)
	intMonth	= Month(dt)
	intDay		= Day(dt)

	select case intShime
	case 21	' 20���̊J�n���擾 ex.20100821
		if intDay <= 21 then
			intMonth = intMonth - 1
			if intMonth < 1 then
				intMonth = 12
				intYear	= intYear - 1
			end if
		end if
		intDay = 21
	case 20	' 20���̏I�����擾 ex.20100920
		if intDay > 21 then
			intMonth = intMonth + 1
			if intMonth > 12 then
				intMonth = 1
				intYear	= intYear + 1
			end if
		end if
		intDay = 20
	end select
	GetShimeDt20 = intYear & right("0" & intMonth,2) & right("0" & intDay,2)
end function

function makeWhere(byval strWhere,byval strField,byval strValue1,byval strValue2)
	dim	strAnd
	dim	strNot
	dim	strCmp

	if len(strValue1) > 0 then
		if len(strWhere) > 0 then
			strAnd = " and "
		else
			strAnd = " where "
		end if
		select case strField
		case "z.ZAIKO_SHOGO_FLG"
			' �݌ɏƍ� 0:�Ώ�/1:�ΏۊO
			if strValue1 = "0" then
				strWhere = strWhere & strAnd & " (z.Soko_No+z.Retu+z.Ren+z.Dan) not in (select distinct (Soko_No+Retu+Ren+Dan) from Tana where ZAIKO_SHOGO_FLG = '1')"
			else
				strWhere = strWhere & strAnd & " (z.Soko_No+z.Retu+z.Ren+z.Dan) in (select distinct (Soko_No+Retu+Ren+Dan) from Tana where ZAIKO_SHOGO_FLG = '1')"
			end if
			strValue1 = ""
		case "z.GOODS_ON"
			' 2:���i����
			select case ucase(strValue1)
			case "2"
				strWhere = strWhere & strAnd & " (z.Soko_No in (select distinct Soko_No from Soko where GOODS_ON_F = '0'))"
				strValue1 = ""
			case "2A"
				strWhere = strWhere & strAnd & " (z.hin_gai in (select distinct hin_gai from zaiko where Soko_No in (select distinct Soko_No from Soko where GOODS_ON_F = '0')))"
				strValue1 = ""
			end select
		case "h.�o��"
			if strValue1 = "0" then
				' 9(�o�ɍ�)������
				strWhere = strWhere & strAnd & " h.IDNo not in (select distinct KEY_ID_NO from y_syuka where KAN_KBN = '9')"
			else
				strWhere = strWhere & strAnd & " h.IDNo in (select distinct KEY_ID_NO from y_syuka where KAN_KBN = '" & strValue1 & "')"
			end if
			strValue1 = ""
		case "L_URIKIN"
			strWhere = strWhere & strAnd & " ("
			strWhere = strWhere & "    convert(L_URIKIN1,sql_decimal) = " & strValue1
			strWhere = strWhere & " or convert(L_URIKIN2,sql_decimal) = " & strValue1
			strWhere = strWhere & " or convert(L_URIKIN3,sql_decimal) = " & strValue1
			strWhere = strWhere & " )"
			strValue1 = ""
		case "---ShisanJCode---"
			select case strValue1
			case "00021259"
				strWhere = makeWhere(strWhere,"JCode"		,"00021259","")
			case else
				strWhere = makeWhere(strWhere,"JCode"		,"00036003","")
			end select
			strAnd = " and "
		case """�K�v���v"""
			strWhere = strWhere & strAnd & strField & strValue1
			strValue1 = ""
		case "ItemExist"
			select case strValue1
			case "0"
'				strWhere = strWhere & strAnd & " Pn not in (select distinct hin_gai from item where jgyobu <> 'S' and naigai = '1')"
				strWhere = strWhere & strAnd & " ShisanJCode+Pn not in (select distinct case jgyobu when '6' then '00021184' when '4' then '00023410' when 'D' then '00023510' when '1' then '00023100' when '7' then '00023210' when 'A' then '00025800' when 'R' then '00021259' end+hin_gai from item where jgyobu <> 'S' and naigai = '1')"
			case "1"
'				strWhere = strWhere & strAnd & " Pn     in (select distinct hin_gai from item where jgyobu <> 'S' and naigai = '1')"
'				strWhere = strWhere & strAnd & " ShisanJCode+Pn     in (select distinct case jgyobu when '6' then '00021184' when '4' then '00023410' when 'D' then '00023510' when '1' then '00023100' when '7' then '00023210' when 'A' then '00025800' when 'R' then '00021259' end+hin_gai from item where jgyobu <> 'S' and naigai = '1')"
				select case strValue2
				case "00023100"
					strWhere = strWhere & strAnd & " Pn in (select distinct hin_gai from item where jgyobu = '1' and naigai = '1')"
				case "00025800"
					strWhere = strWhere & strAnd & " Pn in (select distinct hin_gai from item where jgyobu = 'A' and naigai = '1')"
				case "00023210"
					strWhere = strWhere & strAnd & " Pn in (select distinct hin_gai from item where jgyobu = '7' and naigai = '1')"
				case else
					strWhere = strWhere & strAnd & " JCode+ShisanJCode+Pn     in (select distinct case jgyobu when '6' then '0003600300021184' when '4' then '0003600300023410' when 'D' then '0003600300023510' when '1' then '0003600300023100' when '7' then '0003600300023210' when 'A' then '0003600300025800' when 'R' then '0002125900021259' end+hin_gai from item where jgyobu <> 'S' and naigai = '1')"
				end select
			end select
			strValue1 = ""
		case "z.�d����"
			if left(strValue1,1) = "-" then
				strValue1 = right(strValue1,len(strValue1)-1)
				strWhere = strWhere & strAnd & " z.HIN_GAI not in (select distinct hin_gai from item where SHIMUKE_CODE = '" & strValue1 & "')"
			else
				strWhere = strWhere & strAnd & " z.HIN_GAI     in (select distinct hin_gai from item where SHIMUKE_CODE = '" & strValue1 & "')"
			end if
			strValue1 = ""
		case "z.���`��"
			if left(strValue1,1) = "-" then
				strValue1 = right(strValue1,len(strValue1)-1)
				strWhere = strWhere & strAnd & " z.HIN_GAI not in (select distinct hin_gai from item where K_KEITAI = '" & strValue1 & "')"
			else
				strWhere = strWhere & strAnd & " z.HIN_GAI     in (select distinct hin_gai from item where K_KEITAI = '" & strValue1 & "')"
			end if
			strValue1 = ""
		case "i.���������敪"
			if left(strValue1,1) = "-" then
				strValue1 = right(strValue1,len(strValue1)-1)
				strWhere = strWhere & strAnd & strValue2 & " not in (select distinct hin_gai from item where NAI_BUHIN = '" & strValue1 & "')"
			else
				strWhere = strWhere & strAnd & strValue2 & "     in (select distinct hin_gai from item where NAI_BUHIN = '" & strValue1 & "')"
			end if
			strValue1 = ""
		case "i.�C�O�����敪"
			if left(strValue1,1) = "-" then
				strValue1 = right(strValue1,len(strValue1)-1)
				strWhere = strWhere & strAnd & strValue2 & " not in (select distinct hin_gai from item where GAI_BUHIN = '" & strValue1 & "')"
			else
				strWhere = strWhere & strAnd & strValue2 & "     in (select distinct hin_gai from item where GAI_BUHIN = '" & strValue1 & "')"
			end if
			strValue1 = ""
		case "z.������"
			if instr(1,strValue1,"%") > 0 then
				strCmp = "like"
			else
				strCmp = "="
			end if
			if instr(1,strValue1," ") > 0 then
				strValue1 = replace(strValue1," ","' or ko_hin_gai " & strCmp & "'")
			end if
			strWhere = strWhere & strAnd & " z.HIN_GAI in (select distinct HIN_GAI from p_compo_k where DATA_KBN = '1' and SEQNO = '010' and (ko_hin_gai " & strCmp & " '" & strValue1 & "'))"
			strValue1 = ""
		case "z.���Y��"
			select case strValue1
			case "0"	' ���Y���}�X�^�[�o�^�Ȃ�
				strWhere = strWhere & strAnd & " rtrim(z.jgyobu) + '/' + rtrim(z.naigai) + '/' + rtrim(z.hin_gai) + '/' + rtrim(z.gensankoku)"
				strWhere = strWhere & " not in ("
				strWhere = strWhere & " select distinct"
				strWhere = strWhere & " rtrim(jgyobu) + '/' + rtrim(naigai) + '/' + rtrim(hin_gai) + '/' + rtrim(gensankoku)"
				strWhere = strWhere & " from gensan"
				strWhere = strWhere & " )"
			case "1"	' ���Y���}�X�^�[�o�^����
				strWhere = strWhere & strAnd & " rtrim(z.jgyobu) + '/' + rtrim(z.naigai) + '/' + rtrim(z.hin_gai) + '/' + rtrim(z.gensankoku)"
				strWhere = strWhere & " in ("
				strWhere = strWhere & " select distinct"
				strWhere = strWhere & " rtrim(jgyobu) + '/' + rtrim(naigai) + '/' + rtrim(hin_gai) + '/' + rtrim(gensankoku)"
				strWhere = strWhere & " from gensan"
				strWhere = strWhere & " )"
			case else
				strWhere = makeWhere(strWhere,"z.gensanKoku",strValue1,strValue2)
			end select
			strValue1 = ""
		case "u.MAISU"
			strWhere = strWhere & strAnd & " convert(" & strField & ",sql_numeric) " & strValue1 & ""
			strValue1 = ""
		case "i.S_KOUSU_BAIKA"
			' ���i���P��
			select case left(strValue1,1)
			case "0"	' 0:�P�����o�^
				strWhere = strWhere & strAnd & " rtrim(i.S_KOUSU_BAIKA) = ''"
			case "1"	' 1:�P���o�^��
				strWhere = strWhere & strAnd & " rtrim(i.S_KOUSU_BAIKA) <> ''"
			case "2"	' 2:�P��0�ȏ�
				strWhere = strWhere & strAnd & " (convert(i.S_KOUSU_BAIKA,SQL_NUMERIC) > 0 or convert(i.S_SHIZAI_BAIKA,SQL_NUMERIC) > 0)"
			end select
			strValue1 = ""
		case "HIN_CHECK"
			select case left(strValue1,1)
			case "0"	' 0:�i�ԃ`�F�b�N��
				strWhere = strWhere & strAnd & " HIN_CHECK_DATETIME = ''"
			case "1"	' 1:�i�ԃ`�F�b�N��
				strWhere = strWhere & strAnd & " HIN_CHECK_DATETIME <> ''"
			end select
			strValue1 = ""
		case "s.CANCEL_F"
			select case left(strValue1,1)
			case "0"	' 0:�L�����Z������
				strWhere = strWhere & strAnd & " CANCEL_F = ''"
			case "1"	' 1:�L�����Z���̂�
				strWhere = strWhere & strAnd & " CANCEL_F <> ''"
'			case "2"	' 1:�L�����Z���̂�
'				strWhere = strWhere & strAnd & " left(ID_NO,7) in (select distinct left(ID_NO,7) from " & tblStr & " " & replace(whereStr,"s.","") & andStr & "  CANCEL_F <> '')"
			case else
				strWhere = strWhere & strAnd & " CANCEL_F <> '" & strValue1 & "'"
			end select
			strValue1 = ""
		case "s.MUKE_CODE"
			select case strValue1
			case "�ϐ��S��"
				strWhere = strWhere & strAnd & " s.MUKE_CODE in ('712317','7401UH','7868HA','7868HB','7868HC')"
				strValue1 = ""
			case "�ϐ������Ȃ�"
				strWhere = strWhere & strAnd & " s.MUKE_CODE in ('712317','7401UH','7868HA','7868HB','7868HC')"
				strWhere = strWhere & " and (rtrim(s.SEK_KEN_NO) + rtrim(s.SEK_HIn_NO)) not in (select distinct rtrim(KEN_NO)+rtrim(HIN_NO) from y_syuka_tei)"
				strValue1 = ""
			case "�ϐ���������"
				strWhere = strWhere & strAnd & " s.MUKE_CODE in ('712317','7401UH','7868HA','7868HB','7868HC')"
				strWhere = strWhere & " and (rtrim(s.SEK_KEN_NO) + rtrim(s.SEK_HIn_NO)) in (select distinct rtrim(KEN_NO)+rtrim(HIN_NO) from y_syuka_tei)"
				strValue1 = ""
			end select
		case "Y_SYUKA_H"
			select case strValue1
			case "�o�ח\�肠��"
				strWhere = strWhere & " and (rtrim(t.KEN_NO) + rtrim(t.HIN_NO))     in (select distinct rtrim(SEK_KEN_NO)+rtrim(SEK_HIN_NO) from y_syuka_h)"
				strValue1 = ""
			case "�o�ח\��Ȃ�"
				strWhere = strWhere & " and (rtrim(t.KEN_NO) + rtrim(t.HIN_NO)) not in (select distinct rtrim(SEK_KEN_NO)+rtrim(SEK_HIN_NO) from y_syuka_h)"
				strValue1 = ""
			end select
			strValue1 = ""
		end select
	end if

	if len(strValue1) > 0 then
		strNot = ""
		if left(strValue1,1) = "-" then
			strNot = " not "
			strValue1 = right(strValue1,len(strValue1)-1)
		end if
		if len(strValue2) > 0 then
			strCmp = "between"
			strWhere = strWhere & strAnd & " " & strField & " " & strCmp & " '" & strValue1 & "' and '" & strValue2 & "'"
		else
			select case left(strValue1,1)
			case "<","=",">"
				strCmp = left(strValue1,1)
				strValue1 = "'" & right(strValue1,len(strValue1)-1) & "'"
			case else
				strValue1 = "'" & RTrim(strValue1) & "'"
				if instr(1,strValue1,"%") > 0 _
				or instr(1,strValue1,"_") > 0 then
					strCmp = strNot & "like"
				elseif instr(strValue1,",") > 0 then
					strCmp = strNot & "in "
					strValue1 = "(" & replace(strValue1,",","','") & ")"
				else
					if strNot = "" Then
						strCmp = "="
					else
						strCmp = "<>"
					end if
				end if
			end select
		    select case strField
			case "s.JYUSHO"
    			strWhere = strWhere & strAnd & "( " & strField & " " & strCmp & " " & strValue1 & ""
    			strWhere = strWhere & " or s.YUBIN_No " & strCmp & " " & strValue1 & ")"
			case "s.OKURISAKI"
    			strWhere = strWhere & strAnd & "( " & strField & " " & strCmp & " " & strValue1 & ""
    			strWhere = strWhere & " or s.BIKOU " & strCmp & " " & strValue1 & ")"
		    case "d.ChoCode"
    			strWhere = strWhere & strAnd & "( " & "y.KEY_MUKE_CODE" & " " & strCmp & " " & strValue1 & ""
    			strWhere = strWhere & " or " & "y.LK_MUKE_CODE" & " " & strCmp & " " & strValue1 & ""
    			strWhere = strWhere & " or " & "d.ChoCode" & " " & strCmp & " " & strValue1 & ")"
		    case "y.KEY_MUKE_CODE"
    			strWhere = strWhere & strAnd & "( " & strField & " " & strCmp & " " & strValue1 & ""
    			strWhere = strWhere & " or " & "y.LK_MUKE_CODE" & " " & strCmp & " " & strValue1 & ")"
            case else
    			strWhere = strWhere & strAnd & " " & strField & " " & strCmp & " " & strValue1 & ""
            end select
		end if
	end if
	makeWhere = strWhere
end function

function getTD(byval strName,byval intType)
	select case strName
	case "�g�p��"
		getTD = "nowrap align=""right"" id=""Charactor"""
	case "����","��","�o�ד�"
		getTD = "nowrap align=""center"" id=""Charactor"""
	case else
		select case intType
		Case 2 , 3 , 5 , 131	' ���l(Integer)
			getTD = "nowrap id=""Integer"""
		Case 133				' ���t(Date)
			getTD = "nowrap id=""Date"""
		Case 129				' ������(Charactor)
			getTD = "nowrap align=""left"" id=""Charactor"""
		Case else				' ���̑�
			getTD = "nowrap"
		end select
	end select
end function

function getFieldValue(byval strName,byval intType,byval fValue)
	dim	strValue

	select case strName
	case "�g�p��"
			if VarType(fValue) = vbNull then
				fValue = ""
			end if
			fValue = rtrim(fValue)
			if isNumeric(fValue) then
				strValue = formatnumber(fValue,getPoint(strName),,,-1) & "%"
			end if
	case else
		select case intType
		Case 2 , 3 , 5 , 131	' ���l(Integer)
	'		strValue = VarType(fValue)
			strValue = ""
			if VarType(fValue) <> vbNull then
				if cdbl(fValue) > 0 and cdbl(fValue) < 1 then
					strValue = formatnumber(cdbl(fValue),1,,,-1)
				else
					if cdbl(fValue) <> 0 then
						strValue = formatnumber(fValue,getPoint(strName),,,-1)
					end if
				end if
			end if
		Case 133				' ���t(Date)
			strValue = year(fValue) & "/"
			strValue = strValue & rtrim("0" & month(fValue,2))
			strValue = strValue & "/"
			strValue = strValue & rtrim("0" & day(fValue,2))
		Case 129				' ������(Charactor)
			strValue = rtrim(fValue)
		Case else				' ���̑�
			strValue = rtrim(fValue)
		end select
	end select
	getFieldValue = strValue
end function

function getPoint(byval strName)
	dim	intPoint
	intPoint = 0
	if right(strName,2) = "�P��" or right(strName,3) = "(�g)" then
		intPoint = 2
	elseif strName = "�ː�" or strName = "�ː��v" then
		intPoint = 2
	end if
	getPoint = intPoint
end function

function getNum(byval strValue)
	dim	dblValue
	dblValue = 0
	if strValue <> "" then
		dblValue = cdbl(strValue)
	end if
	getNum = dblValue
end function

function getDateLastModified(byVal strFileName)
	dim	strMapPath
	dim	strDateLastModified
	dim	objFs
	dim	objF

	strDateLastModified = ""
	if strFileName <> "" then
		Set objFS = Server.CreateObject("Scripting.FileSystemObject")
		strMapPath = Server.MapPath(strFileName)
		if objFs.FileExists(strMapPath) then
			Set objF		= objFs.GetFile(strMapPath)
			strDateLastModified	= objF.DateLastModified
			strDateLastModified	= left(strDateLastModified,len(strDateLastModified)-3)
			Set objF		= Nothing
		end if
		Set objFS			= Nothing
	end if
	getDateLastModified = strDateLastModified
end function

function getXfLoc(byval strDbName,byval strTableName)
	dim	strSql
	dim	strXfLoc
	dim	objDb
	dim	objRs

	strXfLoc = ""

	Set objDb = Server.CreateObject("ADODB.Connection")
	objDb.Open strDbName

	strSql = "select * from X$File where Xf$Name = '" & strTableName & "'"
	set objRs = objDb.Execute(strSql)

	if objRs.Eof = False then
		strXfLoc = RTrim(objRs.Fields("Xf$Loc"))
	end if

	set objRs = Nothing
	set objDb = Nothing
	getXfLoc = strXfLoc

end function
'--------------------------------------------------------------------
'POST�f�[�^�̎��
'--------------------------------------------------------------------
Function GetRequest(byVal strName,byVal strDefault)
	dim	strV
	strV = Request.QueryString(strName)
	if Right(strName,1) <> "_" then
		select case strName
		case "ptype","tbl"
		case else
			strV = ucase(strV)
		end select
	end if
	if strV = "" then
		if Request.QueryString(strName).Count = 0 then
			strV = strDefault
		end if
	end if
	GetRequest = strV
End Function
'--------------------------------------------------------------------
Function GetRequestOld(byVal strName)
	dim	strV
	strV = ucase(Request.QueryString(strName))
	if strV = "" then
		select case strName
		case "dbName"
			select case Request.ServerVariables("HTTP_HOST")
			case "192.168.6.31"
				strV = "newsdcnar"
			case else
				strV = "newsdc"
			end select
		end select
	end if
	GetRequestOld = strV
End Function
'--------------------------------------------------------------------
Function GetDbName()
	dim	strDbName
	strDbName = GetRequest("dbName","")
	if strDbName = "" then
'		strDbName = lcase(Split(Request.ServerVariables("URL"),"/")(1))
'		strDbName = lcase(Split(Request.ServerVariables("APPL_PHYSICAL_PATH"),"\")(1))
		strDbName = lcase(Split(Request.ServerVariables("PATH_TRANSLATED"),"\")(1))
		strDbName = lcase(Split(Request.ServerVariables("URL"),"/")(1))
'		select case Request.ServerVariables("HTTP_HOST")
'		case "fhd.osk.sdch"
'			strDbName = "fhd"
'		case "hs1"
'			strDbName = "newsdchir"
'		case else
'			select case LCase(Right(Request.ServerVariables("APPL_MD_PATH"),7))
'			case "newsdc7"
'				strDbName = "newsdc7"
'			case "newsdc8"
'				strDbName = "newsdc8"
'			case "newsdc9"
'				strDbName = "newsdc9"
'			case else
'				strDbName = "newsdc"
'			end select
'		end select
	end if
	GetDbName = strDbName
End Function
'--------------------------------------------------------------------
Function GetCenterName()
	GetCenterName = lcase(Request.ServerVariables("HTTP_HOST"))
	dim	strDir1
	strDir1 = lcase(Split(Request.ServerVariables("URL"),"/")(1))
	select case GetCenterName
	case "w0","192.168.0.12"
						GetCenterName = "w0"	
	case "w1","192.168.1.31"
						GetCenterName = "����Pc"	
	case "w2","192.168.2.31"
						GetCenterName = "�܈�Pc"	
	case "w3","192.168.3.31"
						GetCenterName = "����Pc"	
						select case strDir1
						case "newsdcn"
									GetCenterName = "�R���d�r"
						end select
	case "w4","192.168.4.31"
						GetCenterName = "����Dc"	
						select case strDir1
						case "newsdcr"
									GetCenterName = "�①��"
						end select
	case "w5","192.168.5.31"
						GetCenterName = "��㎖"	
						select case strDir1
						case "fhd"
									GetCenterName = "���g"
						end select
	case "w6","192.168.6.31"
						GetCenterName = "�ޗǉc"
						select case strDir1
						case "newsdc8"
									GetCenterName = "���T"
						case "newsdc9"
									GetCenterName = "�ޗǂb"
						case "newsdcy"
									GetCenterName = "�O�m�e�X�g"
						end select
	case "w7","192.168.7.31"
						GetCenterName = "�L���c"	
	case else
	end select
End Function
'--------------------------------------------------------------------
%>
