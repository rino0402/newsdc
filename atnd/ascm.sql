/*
ascm.sql
pvddl newsdc ascm.sql -server w6
*/
Drop Table Ascm IN DICTIONARY #

Create Table Ascm using 'Ascm.DAT' with replace (
	StaffNo	Char( 8) default '' not null	//ΠυNo
,	Name	Char(40) default '' not null	//Ό
,	Dt		Date				not null	//ϊt
,	Kubun	Char(20) default '' not null	//ζͺ
,	Awh		Currency default 0	not null	//ΐ­
,	Shift	Char(10) default '' not null	//VtgNo
,	ShiftNm	Char(20) default '' not null	//ΌΜΔΌ
,	BegTm	Time							//oΞ
,	FinTM	Time							//ήΞ
,	Late	Currency default 0	not null	//x
,	Early	Currency default 0	not null	//ή
,	Extra	Currency default 0	not null	//ΚcΖ
,	Night	Currency default 0	not null	//[ιcΖ
,	H1Extra	Currency default 0	not null	//@θxc
,	H1Night	Currency default 0	not null	//@θx[
,	H2Extra	Currency default 0	not null	//θxc
,	H2Night	Currency default 0	not null	//θx[
,	PTO		Currency default 0	not null	//LxΙ
,	Actual	Currency default 0	not null	//ΚΤ
,	Memo	Char(80) default '' not null	//υl
,	PRIMARY KEY (
		StaffNo
	,	Dt
	)
) #
