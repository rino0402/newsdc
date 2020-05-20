/*
select * from y_syuka
where KEY_ID_NO in (select distinct idno from HMTAH015)
and LK_SEQ_NO <> ''

complt.sql
pvddl newsdc complt.sql -server w3
*/
update y_syuka
set LK_SEQ_NO = ''
,UPD_NOW = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)
where KEY_ID_NO in (select distinct idno from HMTAH015) and LK_SEQ_NO <> ''
