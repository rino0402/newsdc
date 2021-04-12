Option Explicit
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main()
'-----------------------------------------------------------------------
Private Function Main()
	dim	objPn
	Set objPn = New Pn
	objPn.Run
	Set objPn = nothing
End Function
'-----------------------------------------------------------------------
'Pn
'-----------------------------------------------------------------------
Class Pn
	'-----------------------------------------------------------------------
	'�g�p���@
	'-----------------------------------------------------------------------
	Private Sub Usage()
		Echo "Pn.vbs [option]"
		Echo "Ex."
		Echo "cscript Pn.vbs /db:newsdc4 /item:R /top:100"
		Echo "cscript Pn.vbs /db:newsdc6 /item:1 /table:PnNew /test"
		Echo "Option."
		Echo "   DBName:" & strDBName
		Echo "    Table:" & strTable
		Echo "     Item:" & strItem
		Echo "      Top:" & strTop
		Echo "       Pn:" & strPn
		Echo "    Field:" & strField
		Echo "    InsDt:" & strInsDt
		Echo "     test:" & optTest
	End Sub
	'-----------------------------------------------------------------------
	'�ϐ�
	'-----------------------------------------------------------------------
	Private	objDB
	Private	strDBName
	Private	strTable
	Private	strList
	Private	strItem
	Private	strTop
	Private	strPn
	Private	strField
	Private	strInsDt
	Private	optTest
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName	= GetOption("db"	,"newsdc")
		strTable	= GetOption("table"	,"PnNew")
		strList		= GetOption("list"	,"")
		strItem		= GetOption("item"	,"")
		strTop		= GetOption("top"	,"")
		strPn		= GetOption("pn"	,"")
		strField	= GetOption("field"	,"")
		strInsDt	= GetOption("InsDt"	,"")
		optTest		= False
		set objDB	= nothing
	End Sub
	'-----------------------------------------------------------------------
	'Init() �I�v�V�����`�F�b�N
	'-----------------------------------------------------------------------
    Private Function Init()
		Debug ".Init()"
		dim	strArg
		Init = False
		For Each strArg In WScript.Arguments.UnNamed
			Echo "�I�v�V�����G���[:" & strArg
			Usage
			Exit Function
		Next
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
				strDBName	= GetOption(strArg,strDBName)
			case "table"
				strTable	= GetOption(strArg,strList)
			case "debug"
			case "test"
				optTest		= True
			case "list"
				strList		= GetOption(strArg,strList)
			case "item"
				strItem		= GetOption(strArg,strItem)
			case "pn"
				strPn		= GetOption(strArg,strPn)
			case "field"
				strField	= GetOption(strArg,strField)
			case "insdt"
				strInsDt	= GetOption(strArg,strInsDt)
			case "top"
				strTop		= GetOption(strArg,strTop)
			case "?"
				Usage
				Exit Function
			case else
				Echo "�I�v�V�����G���[:" & strArg
				Usage
				Exit Function
			end select
		Next
		Init = True
	End Function
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Run() ���s����
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		if Init() = True then
			OpenDb
			if strItem <> "" then
				Item
			else
				List0
			end if
			CloseDb
		end if
	End Function
	'-----------------------------------------------------------------------
	'List0()
	'-----------------------------------------------------------------------
    Private Function List0()
		Debug ".List0()"
		AddSql ""
		AddSql "select distinct"
		AddSql " p.JCode"
		AddSql ",p.ShisanJCode"
		AddSql ",j.JGYOBU"
		AddSql ",Count(*) c"
		AddSql2 "from ",strTable & " p"
		AddSql "left outer join JGyobu j on (p.ShisanJCode = j.JCode)"
		AddSql "group by"
		AddSql " p.JCode"
		AddSql ",p.ShisanJCode"
		AddSql ",j.JGYOBU"
		Write strTable,0
		CallSql strSql
		WriteLine ""
		do while objRs.Eof = False
			Line0
			WriteLine ""
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'Line0() 1�s�\��
	'-----------------------------------------------------------------------
    Private Function Line0()
		Debug ".Line0()"
		Write objRs.Fields("JCode")	,9
		Write objRs.Fields("ShisanJCode"),9
		Write objRs.Fields("JGYOBU"),2
		Write objRs.Fields("c"),-7
	End Function
	'-----------------------------------------------------------------------
	'Item() �i�ڃ}�X�^�[�X�V
	'-----------------------------------------------------------------------
	Private	lngCnt			'������
	Private	lngCntUpd		'�X�V����
	Private	lngCntErr		'Err����
	Private	lngCntNew		'New���� ITEM���o�^
    Private Function Item()
		Debug ".Item()"
		AddSql "select "
		if strTop <> "" then
			AddSql "top " & strTop
		end if
		AddSql " p.JCode pJCode"
		AddSql ",p.ShisanJCode pShisanJCode"
		AddSql ",j.JGYOBU jJGYOBU"

		AddSql ",p.Pn pPn"
		AddSql ",i.HIN_GAI iHIN_GAI"

		AddSql ",p.SPn pSPn"
		AddSql ",i.HIN_NAI iHIN_NAI"

		AddSql ",p.PName pPName"
		AddSql ",i.HIN_NAME iHIN_NAME"

		AddSql ",p.PNameEngA pPNameEngA"
		AddSql ",i.L_HIN_NAME_E iL_HIN_NAME_E"

		AddSql ",p.DModel pDModel"
		AddSql ",i.D_MODEL iD_MODEL"

		AddSql ",p.Hinmoku pHinmoku"
		AddSql ",i.HINMOKU iHINMOKU"

		AddSql ",p.UnitKbn pUnitKbn"
		AddSql ",i.UNIT_BUHIN iUNIT_BUHIN"

		AddSql ",p.NaiKbn pNaiKbn"
		AddSql ",i.NAI_BUHIN iNAI_BUHIN"

		AddSql ",p.GaiKbn pGaiKbn"
		AddSql ",i.GAI_BUHIN iGAI_BUHIN"

		AddSql ",p.Loc1 pLoc1"
		AddSql ",i.GLICS1_TANA iGLICS1_TANA"
		AddSql ",p.Loc2 pLoc2"
		AddSql ",i.GLICS2_TANA iGLICS2_TANA"
		AddSql ",p.Loc3 pLoc3"
		AddSql ",i.GLICS3_TANA iGLICS3_TANA"

		AddSql ",p.KKeitai pKKeitai"
		AddSql ",RTrim(i.K_KEITAI) iK_KEITAI"

		AddSql ",convert(p.Tanka2,sql_decimal) pTanka2"
		AddSql ",convert(i.L_URIKIN1,sql_decimal) iL_URIKIN1"
		AddSql ",convert(p.Tanka3,sql_decimal) pTanka3"
		AddSql ",convert(i.L_URIKIN2,sql_decimal) iL_URIKIN2"
		AddSql ",convert(p.Tanka4,sql_decimal) pTanka4"
		AddSql ",convert(i.L_URIKIN3,sql_decimal) iL_URIKIN3"
		AddSql ",convert(p.HyoTan,sql_decimal) pHyoTan"
		AddSql ",convert(i.HYO_TANKA,sql_decimal) iHYO_TANKA"

		AddSql ",p.KobaiTanto pKobaiTanto"
		AddSql ",i.CS_TANTO_CD iCS_TANTO_CD"

		AddSql ",p.NaiModel pNaiModel"
		AddSql ",i.L_KISHU1 iL_KISHU1"

		AddSql ",p.GaiModel pGaiModel"
		AddSql ",i.L_KISHU2 iL_KISHU2"

		AddSql ",p.MadeIn pMadeIn"
		AddSql ",RTrim(i.GENSANKOKU) iGENSANKOKU"
		AddSql ",p.MadeInCode pMadeInCode"
		AddSql ",ifnull(c.CountryName2,p.MadeInCode) cCountryName"
		AddSql ",RTrim(i.TORI_GENSANKOKU) iTORI_GENSANKOKU"

		AddSql ",if(ascii(i.PRT_GENSANKOKU)=0,'',i.PRT_GENSANKOKU) iPRT_GENSANKOKU"
		AddSql ",i.BIKOU20 iBIKOU20"
		AddSql ",i.KANKYO_KBN iKANKYO_KBN"
		AddSql ",i.INS_TANTO iINS_TANTO"
		AddSql ",Left(i.Ins_DateTime,8) + '-' + SubString(i.Ins_DateTime,9,6) iIns_DateTime"
		AddSql ",i.UPD_TANTO iUPD_TANTO"
		AddSql ",Left(i.UPD_DateTime,8) + '-' + SubString(i.UPD_DateTime,9,6) iUPD_DateTime"
		AddSql2 "from ",strTable & " p"
		AddSql "left outer join JGyobu j on (p.ShisanJCode = j.JCode)"
		AddSql "left outer join Item i on (i.JGYOBU = j.JGYOBU and i.NAIGAI='1' and i.HIN_GAI = p.Pn)"
		AddSql "left outer join Country c on (c.CountryCode = p.MadeInCode)"
		AddWhere "jJGYOBU",strItem
		AddWhere "pPn",strPn
		AddWhere "Left(i.Ins_DateTime,8)",strInsDt
		CallSql strSql
		lngCnt		= 0		'������
		lngCntUpd	= 0		'�X�V����
		lngCntErr	= 0		'Err����
		lngCntNew	= 0		'New���� ITEM���o�^
		do while objRs.Eof = False
			lngCnt		= lngCnt + 1		'������
			Write GetField("jJGYOBU"),2
			Write GetField("pPn"),20
			Write GetField("iINS_TANTO"),6
			Write GetField("iINS_DateTime"),16
			Write GetField("iUPD_TANTO"),6
			Write GetField("iUPD_DateTime"),0
'			Write GetField("iHIN_NAME"),0
			WriteLine ""
			if GetField("iHIN_GAI") = "" then
				lngCntNew		= lngCntNew + 1		'New���� ITEM���o�^
			else
				strUpdate = ""
				ItemUpdate
				if strUpdate <> "" then
					lngCntUpd		= lngCntUpd + 1		'�X�V����
					Wscript.StdErr.Write GetField("jJGYOBU") & " "
					Wscript.StdErr.Write GetField("iHIN_GAI") & " "
					Wscript.StdErr.Write GetField("iINS_TANTO") & " "
					Wscript.StdErr.Write GetField("iINS_DateTime") & " "
					Wscript.StdErr.Write GetField("iUPD_TANTO") & " "
					Wscript.StdErr.Write GetField("iUPD_DateTime") & " "
					Wscript.StdErr.Write GetField("iHIN_NAME") & ""
					Wscript.StdErr.WriteLine ""
					Wscript.StdErr.Write strUpdate
				end if
			end if
			objRs.MoveNext
		loop
		dim	strTest
		strTest = ""
		if optTest then
			strTest = " /test"
		end if
		Wscript.StdErr.WriteLine "�������F" & lngCnt	'������
		Wscript.StdErr.WriteLine "  �X�V�F" & lngCntUpd	& strTest '�X�V����
		Wscript.StdErr.WriteLine "���o�^�F" & lngCntNew	'New���� ITEM���o�^
		Wscript.StdErr.WriteLine "   Err�F" & lngCntErr	'Err����
	End Function
	'-----------------------------------------------------------------------
	'GetField() Field�l
	'-----------------------------------------------------------------------
    Private Function GetField(byVal strName)
		GetField = RTrim("" & rmNull(objRs.Fields(strName)) & "")
	End Function
	'-----------------------------------------------------------------------
	'ItemUpdate() Item�X�V
	'-----------------------------------------------------------------------
	Private	strUpdate
    Private Function ItemUpdate()
		Debug ".ItemUpdate()"
		if GetField("iHIN_GAI") = "" then
			exit function
		end if
		dim	strSet
		strSet = ""

		strSet = SetSql(strSet,"      �Γ��i��"	,"HIN_NAI"		,GetField("iHIN_NAI")		,GetField("pSPn")		)
		strSet = SetSql(strSet,"          �i��"	,"HIN_NAME"		,GetField("iHIN_NAME")		,GetField("pPName")		)
		strSet = SetSql(strSet,"      �p��i��"	,"L_HIN_NAME_E"	,GetField("iL_HIN_NAME_E")	,GetField("pPNameEngA")	)
		strSet = SetSql(strSet,"  ��\�@��i��"	,"D_MODEL"		,GetField("iD_MODEL")		,GetField("pDModel")	)
		strSet = SetSql(strSet,"          �i��"	,"HINMOKU"		,GetField("iHINMOKU")		,GetField("pHinmoku")	)
		strSet = SetSql(strSet,"      ���j�b�g"	,"UNIT_BUHIN"	,GetField("iUNIT_BUHIN")	,GetField("pUnitKbn")	)
		strSet = SetSql(strSet,"      ��������"	,"NAI_BUHIN"	,GetField("iNAI_BUHIN")		,GetField("pNaiKbn")	)
		strSet = SetSql(strSet,"      �C�O����"	,"GAI_BUHIN"	,GetField("iGAI_BUHIN")		,GetField("pGaiKbn")	)
		strSet = SetSql(strSet,"           �I1"	,"GLICS1_TANA"	,GetField("iGLICS1_TANA")	,GetField("pLoc1")		)
		strSet = SetSql(strSet,"           �I2"	,"GLICS2_TANA"	,GetField("iGLICS2_TANA")	,GetField("pLoc2")		)
		strSet = SetSql(strSet,"           �I3"	,"GLICS3_TANA"	,GetField("iGLICS3_TANA")	,GetField("pLoc3")		)
		strSet = SetSql(strSet,"      ���`��"	,"K_KEITAI"		,GetField("iK_KEITAI")		,GetField("pKKeitai")	)
		strSet = SetSql(strSet,"        �}���Z"	,"L_URIKIN1"	,GetField("iL_URIKIN1")		,GetField("pTanka2")	)
		strSet = SetSql(strSet,"            ��"	,"L_URIKIN2"	,GetField("iL_URIKIN2")		,GetField("pTanka3")	)
		strSet = SetSql(strSet,"            ��"	,"L_URIKIN3"	,GetField("iL_URIKIN3")		,GetField("pTanka4")	)
		strSet = SetSql(strSet,"      �W���P��"	,"HYO_TANKA"	,GetField("iHYO_TANKA")		,GetField("pHyoTan")	)
		strSet = SetSql(strSet,"     ��\�@��1"	,"L_KISHU1"		,GetField("iL_KISHU1")		,GetField("pNaiModel")	)
		strSet = SetSql(strSet,"     ��\�@��2"	,"L_KISHU2"		,GetField("iL_KISHU2")		,GetField("pGaiModel")	)
		strSet = SetSql(strSet,"�����\�����Y��"	,"GENSANKOKU"		,GetField("iGENSANKOKU")		,GetField("pMadeIn")	)
		strSet = SetSql(strSet,"        ���Y��"	,"TORI_GENSANKOKU"	,GetField("iTORI_GENSANKOKU")	,GetField("cCountryName"))
		strSet = SetSql(strSet,"    ���Y����"	,"PRT_GENSANKOKU"	,GetField("iPRT_GENSANKOKU")	,""						)
		strSet = SetSql(strSet,"        CS�S��"	,"CS_TANTO_CD"	,GetField("iCS_TANTO_CD")	,GetField("pKobaiTanto"))
		strSet = SetSql(strSet,"        ���i�["	,"BIKOU20"		,GetField("iBIKOU20")		,GetBikou20()			)
'		strSet = SetSql(strSet," �dWC:","TORI_SHIIRE_WORK_CTR",RTrim(objRs.Fields("TORI_SHIIRE_WORK_CTR")),RTrim(objRs.Fields("SHIIRE_WORK_CENTER")))
'		strSet = SetSql(strSet," ��:","KANKYO_KBN",RTrim(objRs.Fields("iKANKYO_KBN")),RTrim(objRs.Fields("hKANKYO_KBN")))
'		strSet = SetSql(strSet," ���n:","KANKYO_KBN_ST",RTrim(objRs.Fields("iKANKYO_KBN_ST")),RTrim(objRs.Fields("hKANKYO_KBN_ST")))
'		strSet = SetSql(strSet," ����:","KANKYO_KBN_SURYO",Trim(objRs.Fields("iKANKYO_KBN_SURYO")),Trim(objRs.Fields("hKANKYO_KBN_SURYO")))
		if strSet = "" then
			exit function
		end if
		AddSql ""
		AddSql "update Item"
		AddSql strSet
		AddSql ",UPD_TANTO='" & strTable & "'"
		AddSql ",UPD_DATETIME = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)"
		AddSql "where JGyobu = '" & GetField("jJGYOBU") & "'"
		AddSql "  and NAIGAI = '1'"
		AddSql "  and HIN_GAI = '" & GetField("iHIN_GAI") & "'"
		if optTest = True then
			strUpdate = strUpdate & Replace(strSql,vbCrLf," ") & vbCrLf
			exit function
		end if
		WriteLine "update:" & Execute(strSql)
	End Function
	'-----------------------------------------------------------------------
	'SetSql() 
	'-----------------------------------------------------------------------
    Private Function SetSql(byVal strSet,byVal strTitle,byVal strName,byVal strSrc,byVal strDst)
		Debug ".SetSql()"
		if strField <> "" then
			if ucase(strField) <> ucase(strName) and ucase(strField) <> trim(ucase(strTitle)) then
				SetSql = strSet
				exit function
			end if
		end if
		strSrc = RTrim(strSrc & "")
		strDst = RTrim(strDst & "")
		select case strName
		case "L_HIN_NAME_E"
			select case GetField("jJGYOBU")
			case "A"	' �G�A�R�� �C�O�����敪:1,2 �̏ꍇ�A�i�ڕʖ����Z�b�g
				select case GetField("pGaiKbn")
				case "1","2","3"
				case else
					strDst = strSrc
				end select
			end select
		case "GLICS1_TANA","GLICS2_TANA","GLICS3_TANA"
			select case GetField("jJGYOBU")
			case "A"			' �G�A�R�� Glics�I1 ���Z�b�g���Ȃ�
				select case strName
				case "GLICS1_TANA"
					if strSrc <> "" then
						strDst = strSrc
					end if
				end select
			case "R"			'
				if lcase(strTable) = "pnnew" then
					strDst = strSrc
				end if
			case "4","5","D"	' ���� Glics�I2,�I3���Z�b�g���Ȃ�
				select case strName
				case "GLICS2_TANA","GLICS3_TANA"
					if strSrc <> "" then
						strDst = strSrc
					end if
				end select
			end select
		case "L_URIKIN1","L_URIKIN2","L_URIKIN3"
			if strSrc = "99999999" then
				strDst = strSrc
			end if
		case "NAI_BUHIN"
			select case GetField("jJGYOBU")
			case "R"
				if lcase(strTable) = "pnnew" then
					strDst = strSrc
				end if
			end select
		case "HINMOKU"
			select case GetField("jJGYOBU")
			case "7"
			case "R"
				if lcase(strTable) = "pnnew" then
					strDst = strSrc
				end if
			end select
		case "L_KISHU1"
			select case GetField("jJGYOBU")
			case "1","7","2"
				'���p�[�c���x��Pn�A�g���ځF����@ ��\�@��C�O ���Z�b�g
				'  �@��P�F��\�@�퍑���F�����敪�����F1(�Ώ�),2(�Ő؈ē���),3(�Ő�)
				'          ��\�@��C�O�F�����敪�����F0(�ΏۊO)
				'                        �����敪�C�O�F1(�Ώ�),2(�Ő؈ē���),3(�Ő�)
				'  �@��Q�F��
				if left(strSrc,6) = "���p���ɂ��" then
					strDst = strSrc
				else
					select case GetField("pNaiKbn")
					case "1","2","3"
					case "0"
						select case GetField("pGaiKbn")
						case "1","2","3"
							strDst = GetField("pGaiModel")
						end select
					end select
				end if
			case "R"	'�①�� ITEM�̑�\�@�킪�󔒂̏ꍇ�Z�b�g
				if strSrc <> "" then
					strDst = strSrc
				end if
			case "A"
			case else
				strDst = strSrc
			end select
		case "L_KISHU2"
			select case GetField("jJGYOBU")
			case "1","7","2"
				if left(GetField("iL_KISHU1"),6) = "���p���ɂ��" _
				or left(GetField("iL_KISHU2"),6) = "���p���ɂ��" then
					strDst = strSrc
				else
					strDst = ""
				end if
			case "A"
			case else
				strDst = strSrc
			end select
		case "K_KEITAI"
			select case GetField("jJGYOBU")
			case "2"	'2017.05.29 �H�� ���`�Ԃ��Z�b�g���Ȃ�
				strDst = strSrc
			case else
			end select
		case "GENSANKOKU"
			select case GetField("jJGYOBU")
			case "R"	' �①�� �����\�����Y�����X�V���Ȃ�
				strDst = strSrc
			case "4","5","D","1","2"
'GlicsPn�̌����\�����Y�����󔒂̏ꍇ�́APos�i�ڃ}�X�^�[���X�V���Ȃ��B
'�Ώێ��ƕ��F4 ����/5 BL����/D IH
'--------------------------------------------
'�@�@�@�@�@�@�@�@Pos�@�@Glics
'--------------------------------------------
'�����\�����Y���FJAPAN  �󔒁@�@���X�V���Ȃ� �����܂ł͋󔒂ɕύX
'�����\�����Y���FJAPAN  CHINA �@��CHINA�ɕύX
'�����\�����Y���F��   JAPAN �@��JAPAN�ɕύX
'--------------------------------------------
				if strDst = "" then
					if strSrc <> "" then
						strDst = strSrc
						strUpdate = strUpdate & strTitle & "�F" & strSrc & "(Pn��)" & vbCrLf
					end if
				end if
			end select
		case "TORI_GENSANKOKU"
			select case GetField("jJGYOBU")
			case "A"	' �G�A�R��
				select case GetField("pGaiKbn")
				case "1","2","3"
				case else
					strDst = strSrc
				end select
			end select
		case "PRT_GENSANKOKU"
			select case GetField("jJGYOBU")
			case "1"	'����@ �C�O�����Ώۂ͌��Y���󎚂���
				strDst = "0"
				select case GetField("pGaiKbn")
				case "1","2","3"
					strDst = "1"
				end select
			case "7"	'�N���[�i�[ ���Y���󎚂��Ȃ����f�t�H���g�ɕύX(2017.5.8�`)
				strDst = "0"
			case else
				strDst = strSrc
			end select
		case "BIKOU20","HIN_NAI"
			if strDst = "" then
				strDst = strSrc
			end if
		case "KANKYO_KBN_SURYO"
			if strDst = "0" then
				strDst = strSrc
			end if
		case "INSP_MESSAGE"
			if strSrc = "�P������ ���`�E���d�r����" then
				strDst = strSrc
			end if
		end select
		if strDst = strSrc then
			WriteLine strTitle & "�F" & strSrc
'			WriteLine strTitle & "�F" & strDst
		else
			WriteLine strTitle & "�F" & strSrc
			WriteLine strTitle & "��" & strDst
			strUpdate = strUpdate & strTitle & "�F" & strSrc & vbCrLf
			strUpdate = strUpdate & strTitle & "��" & strDst & vbCrLf
			if strSet = "" then
				strSet = " set "
			else
				strSet = strSet & " ,"
			end if
			strDst = Replace(strDst,"'","''")
			strSet = strSet & strName & " = '" & strDst & "'"
		end if
		SetSql = strSet
	End Function
	'-----------------------------------------------------------------------
	'GetTanto() �S���Җ�
	'-----------------------------------------------------------------------
    Private Function GetTanto(byVal strTantoNm)
		GetTanto = ""
		if inStr(strTantoNm,"����") > 0 then
			GetTanto = "����"
		elseif inStr(strTantoNm,"����") > 0 then
			GetTanto = "����"
		elseif inStr(strTantoNm,"��") > 0 then
			GetTanto = "��"
		elseif inStr(strTantoNm,"���c") > 0 then
			GetTanto = "���c"
		elseif inStr(strTantoNm,"�") > 0 then
			GetTanto = "�"
		elseif inStr(strTantoNm,"����") > 0 then
			GetTanto = "����"
		elseif inStr(strTantoNm,"���R") > 0 then
			GetTanto = "���R"
		elseif inStr(strTantoNm,"�쑺") > 0 then
			GetTanto = "�쑺"
		elseif inStr(strTantoNm,"�c��") > 0 then
			GetTanto = "�c��"
		elseif inStr(strTantoNm,"��t") > 0 then
			GetTanto = "��t"
		elseif inStr(strTantoNm,"���") > 0 then
			GetTanto = "���"
		end if
	End Function
	'-----------------------------------------------------------------------
	'GetBikou20() ���i�[���l
	'-----------------------------------------------------------------------
    Private Function GetBikou20()
		GetBikou20 = ""
		dim	strTantoNm
		strTantoNm = GetTantoNm(RTrim(GetField("pKobaiTanto")))
		GetBikou20 = strTantoNm
		if strTantoNm = "" then
			exit function
		end if
		dim	strBikou20
		strBikou20 = RTrim(GetField("iBIKOU20")) & ""
		if strBikou20 = "" then
			exit function
		end if
		if strBikou20 = GetBikou20 then
			exit function
		end if
		dim	strTantoNmOld
		strTantoNmOld = GetTanto(strBikou20)
		if strTantoNmOld = "" then
			GetBikou20 = strBikou20
			exit function
		end if
		GetBikou20 = Replace(strBikou20,strTantoNmOld,strTantoNm)
	End Function
	'-----------------------------------------------------------------------
	'GetTantoNm() �S���Җ�
	'-----------------------------------------------------------------------
    Private Function GetTantoNm(byVal strTanto)
		GetTantoNm = ""
		select case strTanto
		case "R101"
				GetTantoNm = "����"
		case "R102"
				GetTantoNm = "����"
		case "R103"
				GetTantoNm = "���R"
		case "R104"
				GetTantoNm = "����"
		case "R105"
				GetTantoNm = "�쑺"
		case "R106"
				GetTantoNm = "����"
		end select
	End Function
	'-------------------------------------------------------------------
	'GroupHead() �O���[�v�w�b�_�[
	'	True:�O���[�v�w�b�_�[
	'  Flase:�p���s
	'-------------------------------------------------------------------
	Private	curHead
	Private	newHead
	Private	Function GroupHead(byVal intHead)
		if intHead < 0 then
			curHead = ""
			exit function
		end if
		dim	i
		newHead = ""
		for i = 0 to intHead
			newHead = newHead + objRs.Fields(i)
		next
		if curHead = newHead then
			GroupHead = False
			exit function
		end if
		curHead = newHead
		GroupHead = True
	End Function
	'-----------------------------------------------------------------------
	'ItemInsert() Item�ǉ�
	'-----------------------------------------------------------------------
    Private Function ItemInsert()
		Debug ".ItemInsert()"
		AddSql ""
		AddSql "insert into Item"
		AddSql "("
		AddSql " JGYOBU"				'Char(  1) //���ƕ��敪
		AddSql ",NAIGAI"				'Char(  1) //�����O
		AddSql ",HIN_GAI"				'Char( 20) //�i�ԁi�O���j
		AddSql ",HIN_NAME"				'Char( 40) //�i��
		AddSql ",HIN_NAI"				'Char( 20) //�i�ԁi�����j
		AddSql ",GLICS1_TANA"			'Char( 10) //�O���b�N�X�I�ԂP   2005.05
		AddSql ",GLICS2_TANA"			'Char( 10) //�O���b�N�X�I�ԂQ   2005.05
		AddSql ",GLICS3_TANA"			'Char( 10) //�O���b�N�X�I�ԂR   2005.05
		AddSql ",L_HIN_NAME_E"			'Char( 30) //���i����   �i��
		AddSql ",L_KISHU1"				'Char( 25) //           �@��(1)
		AddSql ",L_KISHU3"				'Char(150) //           �@��(3)(���K�p�@���
		AddSql ",L_URIKIN1"				'Char( 10) //           ���i(1)	//NUMERICSA(10,0)
		AddSql ",L_URIKIN2"				'Char( 10) //           ���i(2)	//NUMERICSA(10,0)
		AddSql ",L_URIKIN3"				'Char( 10) //           ���i(3)	//NUMERICSA(10,0)
		AddSql ",UNIT_BUHIN"			'Char(  1) //�Ưĕ��i�敪       2006.07.28
		AddSql ",NAI_BUHIN"				'Char(  1) //�����������i�敪   2006.07.28
		AddSql ",GAI_BUHIN"				'Char(  1) //�C�O�������i�敪   2006.07.28
		AddSql ",HYO_TANKA"				'Char( 10) //�W���P��   2006.07.28
		AddSql ",KANKYO_KBN"			'Char(  3) //����ދ敪       2010.07.27
		AddSql ",KANKYO_KBN_ST"			'Char(  8) //����ދ敪�K�p�J�n 2010.07.
		AddSql ",KANKYO_KBN_SURYO"		'Char( 10) //����ދ敪����   2010.07.27
		AddSql ",CS_TANTO_CD"			'Char(  8) //CS�S������
		AddSql ",D_MODEL"				'Char(  8) //��\�@��i�ڃR�[�h PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",HINMOKU"				'Char(  8) //�i�ڃR�[�h         PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",K_KEITAI"				'Char( 14) //���`��(14��)     2012.03.13
		AddSql ",INS_TANTO"				'Char(  5) //�ǉ��@�S����
		AddSql ",Ins_DateTime"			'Char( 14) //�ǉ��@����  
		AddSql ",BIKOU20"
		AddSql ",L_PAPER"				' not null   //           ��
		AddSql ",L_PLASTIC"             ' not null   //           �v���X�`�b�N
		AddSql ",L_LABEL"               ' not null   //           �K�p�@������
		AddSql ")"
		AddSql "select top 1"
		AddSql " h.JGyobu"				'//���ƕ��敪
		AddSql ",'1'"					'//�����O
		AddSql ",p.Pn"					'Char( 20) //�i�ԁi�O���j
		AddSql ",p.PnBetsu"				'Char( 40) //�i��
		AddSql ",p.SPn"					'Char( 20) //�i�ԁi�����j
		AddSql ",p.Loc1"				'Char( 10) //�O���b�N�X�I�ԂP   2005.05
		AddSql ",p.Loc2"				'Char( 10) //�O���b�N�X�I�ԂQ   2005.05
		AddSql ",p.Loc3"				'Char( 10) //�O���b�N�X�I�ԂR   2005.05
		AddSql ",RTrim(p.PNameEngA)"	'Char( 30) //���i����   �i��
		AddSql ",p.NaiModel"			'Char( 25) //           �@��(1)
		AddSql ",p.GaiModel"			'Char(150) //           �@��(3)(���K�p�@���
		AddSql ",p.Tanka2"				'Char( 10) //           ���i(1)	//NUMERICSA(10,0)
		AddSql ",p.Tanka3"				'Char( 10) //           ���i(2)	//NUMERICSA(10,0)
		AddSql ",p.Tanka4"				'Char( 10) //           ���i(3)	//NUMERICSA(10,0)
		AddSql ",p.UnitKbn"				'Char(  1) //�Ưĕ��i�敪       2006.07.28
		AddSql ",p.NaiKbn"				'Char(  1) //�����������i�敪   2006.07.28
		AddSql ",p.GaiKbn"				'Char(  1) //�C�O�������i�敪   2006.07.28
		AddSql ",p.HyoTan"				'Char( 10) //�W���P��   2006.07.28
		AddSql ",h.KANKYO_KBN"			'Char(  3) //����ދ敪       2010.07.27
		AddSql ",h.KANKYO_KBN_ST"		'Char(  8) //����ދ敪�K�p�J�n 2010.07.
		AddSql ",h.KANKYO_KBN_SURYO"	'Char( 10) //����ދ敪����   2010.07.27
		AddSql ",p.KobaiTanto"			'Char(  8) //CS�S������
		AddSql ",p.DModel"				'Char(  8) //��\�@��i�ڃR�[�h PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",p.Hinmoku"				'Char(  8) //�i�ڃR�[�h         PN�A�g�ŃZ�b�g 2011.12.28
		AddSql ",p.KKeitai"				'Char( 14) //���`��(14��)     2012.03.13
		AddSql ",'HM500'"				'Char(  5) //�ǉ��@�S����
		AddSql ",left(replace(replace(replace(convert(Now(),sql_char),'-',''),':',''),' ',''),14)"	'Char( 14) //�ǉ��@����  
		AddSql ",case p.KobaiTanto"
		AddSql " when 'R101' then '����'"
		AddSql " when 'R102' then '����'"
		AddSql " when 'R103' then '���R'"
		AddSql " when 'R104' then '����'"
		AddSql " when 'R105' then '�쑺'"
		AddSql " when 'R106' then '����'"
		AddSql " else ''"
		AddSql " end"
		AddSql ",'0'"	'//           ��
		AddSql ",'0'"	'//           �v���X�`�b�N
		AddSql ",'0'"	'//           �K�p�@������
		AddSql "from Pn h"
		AddSql "inner join Pn5 p on (h.Pn = p.Pn)"
		AddWhere "h.Filename",RTrim(objRs.Fields("Filename"))
		AddWhere "h.Pn",RTrim(objRs.Fields("Pn"))
		Write ":" & Execute(strSql) ,0
	End Function
	'-------------------------------------------------------------------
	'Sql���s
	'-------------------------------------------------------------------
	Private Function Execute(byVal strSql)
		Debug ".Execute():" & strSql
		on error resume next
		Call objDb.Execute(strSql)
		Execute = Err.Number
		select case Execute
		case 0
		case -2147467259	'0x80004005 �d���L�[
		case else
			Wscript.StdErr.WriteLine
			Wscript.StdErr.WriteLine Err.Description
			Wscript.StdErr.WriteLine strSql
		end select
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'Sql���s
	'-------------------------------------------------------------------
	Private	objRs
	Private Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
		on error resume next
		set objRs = objDb.Execute(strSql)
		if Err.Number <> 0 then
			Wscript.StdErr.WriteLine "0x" & Hex(Err.Number)
			Wscript.StdErr.WriteLine Err.Description
			Wscript.StdErr.WriteLine strSql
			Wscript.Quit
		end if
		on error goto 0
    End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'AddSql2
	'-------------------------------------------------------------------
	Private	Function AddSql2(byVal str1,byVal str2)
		if Right(str1,1) = "'" then
			'Char
			str2 = Replace(RTrim(str2),"'","''") & "'"
		end if
		AddSql str1 & str2
	End Function
	'-------------------------------------------------------------------
	'Where strSql
	'-------------------------------------------------------------------
	Private	Function AddWhere(byVal strF,byVal strV)
		if strV = "" then
			exit function
		end if
		if inStr(strSql,"where") > 0 then
			AddSql " and "
		else
			AddSql " where "
		end if
		dim	strCmp
		strCmp = "="
		if left(strV,1) = "-" then
			strV = Right(strV,len(strV)-1)
			strCmp = "<>"
		elseif left(strV,1) = "+" then
			strV = Right(strV,len(strV)-1)
			strCmp = ">"
		end if
		if inStr(strV,"%") > 0 then
			if strCmp = "=" then
				strCmp = " like "
			else
				strCmp = " not like "
			end if
		end if
		AddSql strF & " " & strCmp & " '" & strV & "'"
	End Function
	'-------------------------------------------------------------------
	'������ǉ� strSql
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
	'-----------------------------------------------------------------------
	'�f�o�b�O�p /debug
	'-----------------------------------------------------------------------
	Private Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'�I�v�V�����擾
	'-----------------------------------------------------------------------
	Private Function GetOption(byval strName ,byval strDefault)
		dim	strValue

		if strName = "" then
			strValue = ""
			if WScript.Arguments.Named.Exists(strDefault) then
				strValue = strDefault
			end if
		else
			strValue = strDefault
			if WScript.Arguments.Named.Exists(strName) then
				strValue = WScript.Arguments.Named(strName)
			end if
		end if
		GetOption = strValue
	End Function
	'-----------------------------------------------------------------------
	'inNull() Null�����邩�`�F�b�N
	'-----------------------------------------------------------------------
	Private Function inNull(byVal s)
		inNull = ""
		dim	c
		dim	i
		if isNull(s) = True then
			inNull = "(null)"
		end if
		for i = 1 to len(s)
			c = mid(s,i,1)
			if Asc(c) = 0 then
				inNull = inNull & "(null:" & i & "/" & len(s) & ")"
				exit for
			end if
		next
	End Function
	'-----------------------------------------------------------------------
	'rmNull() Null���폜
	'-----------------------------------------------------------------------
	Private Function rmNull(byVal s)
'		rmNull = Replace(s,0,"")
		dim	t
		dim	c
		dim	i
		t = ""
		for i = 1 to len("" & s)
			c = mid(s,i,1)
			if Asc(c) = 0 then
				c = ""
			end if
			t = t & c
		next
		rmNull = t
	End Function
	'-----------------------------------------------------------------------
	'WriteLine
	'-----------------------------------------------------------------------
	Private Sub WriteLine(byVal s)
		Wscript.StdOut.WriteLine rmNull(s) & inNull(s)
	End Sub
	'-----------------------------------------------------------------------
	'Write
	'-----------------------------------------------------------------------
	Private Sub Write(byVal s,byVal i)
		dim	t
		t = rmNull(s)
		if i > 0 then
			t = left(RTrim(t) & space(i),i)
		elseif i < 0 then
			t = right(space(-i) & LTrim(t),-i)
		end if
		Wscript.StdOut.Write t & inNull(s)
	End Sub
	'-----------------------------------------------------------------------
	'Echo
	'-----------------------------------------------------------------------
	Private Sub Echo(byVal s)
		Wscript.Echo s
	End Sub
End Class
