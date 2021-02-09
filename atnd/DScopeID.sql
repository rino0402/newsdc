/*
DScopeID.sql
pvddl newsdc DScopeID.sql -server w1
pvddl newsdc DScopeID.sql -server w3
pvddl newsdc DScopeID.sql -server w4
pvddl newsdc DScopeID.sql -server w5
pvddl newsdc ..\atnd\DScopeID.sql -server w6
pvddl newsdc ..\atnd\DScopeID.sql -server w2
pvddl newsdc ..\atnd\DScopeID.sql -server w7
select
*
from DScope d
left outer join Tanto t
 on (d.Name = t.TANTO_NAME)

select
*
from DscopeID d
left outer join TANTO t
	on (d.TANTO_CODE = t.TANTO_CODE)

*/
/*
insert into DScopeID (
	ID
,	TANTO_CODE
)
select
distinct
 d.ID
,t.TANTO_CODE
from DScope d
inner join Tanto t
 on (replace(d.Name,'　','') = t.TANTO_NAME)
where d.ID <> '' and ID not in (select ID from DScopeID)
and not (t.TANTO_CODE = '30005' and t.TANTO_NAME = '八太友希')

insert into DScopeID (ID,TANTO_CODE) value ('20657','20657')
*/

Drop Table DScopeID IN DICTIONARY #

Create Table DScopeID using 'DScopeID.DAT' with replace (
	ID			Char(40) default '' not null	//ID
,	TANTO_CODE	Char(10) default '' not null	//担当者コード
,	PRIMARY KEY (
		ID
	)
) #

