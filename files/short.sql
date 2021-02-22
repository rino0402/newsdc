/*
short.sql
2017.10.10 PN‹¤’Ê”õl—“ ˆÈ~íœ
*/
select
 i.HIN_GAI	"•i”Ô"
,i.HIN_NAI	"“à•”•i”Ô"
,i.HIN_NAME	"•i–¼"
,GetSupplyNm(i.NAI_BUHIN)
"‘“à
‹Ÿ‹‹‹æ•ª"
,GetSupplyNm(i.GAI_BUHIN)
"ŠCŠO
‹Ÿ‹‹‹æ•ª"
,z.qty			"İŒÉ”"
,round(ifnull(convert(a.AVE_SYUKA,sql_decimal),0),1)
"‚Rƒ–Œ•½‹Ï
Œ/o‰×”"
,if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
   ,null()
   ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
 )
"İŒÉŒ”"
,if(if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
	  ,null()
	  ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
	  )<=1
   ,'›'
   ,''
   )
"‚Pƒ–ŒˆÈ‰º"
,if(if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
	  ,null()
	  ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
	  )<=(5/30)
   ,'›'
   ,''
   )
"‚T“úˆÈ‰º"
,pn.NaiDisconYm
"‘“à‹Ÿ‹‹
‘ÅØ”NŒ"
/*
,h.Biko
"PN‹¤’Ê”õl—“"
,s.Dt_1
"”[“ü—\’è“ú"
,s.Qt_1
"”[“ü—\’è”"
,s.Dt_2
"”[“ü—\’è“ú2"
,s.Qt_2
"”[“ü—\’è”2"
,s.Dt_3
"”[“ü—\’è“ú3"
,s.Qt_3
"”[“ü—\’è”3"
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
and ("İŒÉ”" > 0 or ifnull(convert(a.AVE_SYUKA,sql_decimal),0) > 0)
order by
 "‚T“úˆÈ‰º" desc
,"‚Pƒ–ŒˆÈ‰º" desc
,if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0,1,0)
,"İŒÉŒ”"
,"‚Rƒ–Œ•½‹Ï
Œ/o‰×”" desc
,"İŒÉ”" desc
