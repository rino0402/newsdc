/*
staff.sql
pvddl newsdc staff.sql -server w3
pvddl newsdc staff.sql -server w1
pvddl newsdc staff.sql -server w5
pvddl newsdc staff.sql -server w7
pvddl newsdc staff.sql -server w2
pvddl newsdc staff.sql -server w6
*/
Drop Table Staff IN DICTIONARY #

Create Table Staff using 'Staff.DAT' with replace (
	StaffNo	Char( 8) default '' not null	//1"]‹Æˆõ”Ô†"
,	Name	Char(40) default '' not null	//2"]‹Æˆõ–¼"
,	Post	Char( 8) default '' not null	//3"Š‘®”Ô†"
,	Shift	Char( 2) default '' not null	//5"ƒVƒtƒg”Ô†"
,	Quit	Char(10) default '' not null	//‘ŞE¨ƒƒ‚‚É•ÏX
,	QuitDt	Date							//‘ŞE“ú
,	PRIMARY KEY (
		StaffNo
	)
) #

insert into Staff (
 StaffNo
,Name
,Post
,Shift
)
select distinct
 right(StaffNo,5)
,Name
,Shift
,Shift
from timepack

