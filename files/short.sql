/*
short.sql
2017.10.10 PN���ʔ��l�� �ȍ~�폜
*/
select
 i.HIN_GAI	"�i��"
,i.HIN_NAI	"�����i��"
,i.HIN_NAME	"�i��"
,GetSupplyNm(i.NAI_BUHIN)
"����
�����敪"
,GetSupplyNm(i.GAI_BUHIN)
"�C�O
�����敪"
,z.qty			"�݌ɐ�"
,round(ifnull(convert(a.AVE_SYUKA,sql_decimal),0),1)
"�R��������
��/�o�א�"
,if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
   ,null()
   ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
 )
"�݌Ɍ���"
,if(if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
	  ,null()
	  ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
	  )<=1
   ,'��'
   ,''
   )
"�P�����ȉ�"
,if(if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
	  ,null()
	  ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
	  )<=(5/30)
   ,'��'
   ,''
   )
"�T���ȉ�"
,pn.NaiDisconYm
"��������
�ŐؔN��"
/*
,h.Biko
"PN���ʔ��l��"
,s.Dt_1
"�[���\���"
,s.Qt_1
"�[���\�萔"
,s.Dt_2
"�[���\���2"
,s.Qt_2
"�[���\�萔2"
,s.Dt_3
"�[���\���3"
,s.Qt_3
"�[���\�萔3"
*/
from item i
left outer join (
	select
//	 top 100
	 HIN_GAI
	,sum(convert(YUKO_Z_QTY,sql_decimal))	qty
	from Zaiko
	where JGYOBU='7'
	  and NAIGAI='1'
	group by HIN_GAI
) z
	on (z.HIN_GAI=i.HIN_GAI)
left outer join ave_syuka a
	on (a.JGYOBU='7' and a.NAIGAI='1' and i.HIN_GAI=a.HIN_GAI)
inner join PnNew pn
	on (pn.JCode='00036003' and pn.ShisanJCode='00023210' and i.HIN_GAI=pn.Pn)
/*
left outer join PnHosoku h
	on (h.ShisanJCode='00023210' and i.HIN_GAI=h.Pn)
left outer join SaDelvSum s
	on (left(i.HIN_NAI,8)=s.Pn)
*/
where i.JGYOBU='7' and i.NAIGAI='1'
and ("�݌ɐ�" > 0 or ifnull(convert(a.AVE_SYUKA,sql_decimal),0) > 0)
order by
 "�T���ȉ�" desc
,"�P�����ȉ�" desc
,if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0,1,0)
,"�݌Ɍ���"
,"�R��������
��/�o�א�" desc
,"�݌ɐ�" desc
