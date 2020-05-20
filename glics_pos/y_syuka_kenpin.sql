/*
pvddl newsdc y_syuka_kenpin.sql -log y_syuka_kenpin.log -server w1
*/
update y_syuka
set	KENPIN_TANTO_CODE = '00000'
,	KENPIN_YMD = '00000000'
,	KENPIN_HMS = '000000'
,	UPD_NOW = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)
where JGYOBA <> '00036003'
and KEY_SYUKA_YMD < replace(convert(CURDATE(),sql_char),'-','')
and KAN_KBN = '9'
and KENPIN_TANTO_CODE = ''
#

update y_syuka
set KAN_KBN = '9'
,	JITU_SURYO = SURYO
,	UPD_NOW = left(replace(replace(replace(convert(now(),sql_char),'-',''),':',''),' ',''),14)
where JGYOBA <> '00036003'
and KEY_SYUKA_YMD < replace(convert(CURDATE(),sql_char),'-','')
and KAN_KBN <> '9'
and RTrim(KENPIN_TANTO_CODE) <> ''
#
