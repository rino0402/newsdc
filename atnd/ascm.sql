/*
ascm.sql
pvddl newsdc ascm.sql -server w6
*/
Drop Table Ascm IN DICTIONARY #

Create Table Ascm using 'Ascm.DAT' with replace (
	StaffNo	Char( 8) default '' not null	//社員No
,	Name	Char(40) default '' not null	//氏名
,	Dt		Date				not null	//日付
,	Kubun	Char(20) default '' not null	//区分
,	Awh		Currency default 0	not null	//実働
,	Shift	Char(10) default '' not null	//シフトNo
,	ShiftNm	Char(20) default '' not null	//ｼﾌﾄ名
,	BegTm	Time							//出勤
,	FinTM	Time							//退勤
,	Late	Currency default 0	not null	//遅刻
,	Early	Currency default 0	not null	//早退
,	Extra	Currency default 0	not null	//普通残業
,	Night	Currency default 0	not null	//深夜残業
,	H1Extra	Currency default 0	not null	//法定休残
,	H1Night	Currency default 0	not null	//法定休深
,	H2Extra	Currency default 0	not null	//所定休残
,	H2Night	Currency default 0	not null	//所定休深
,	PTO		Currency default 0	not null	//有給休暇
,	Actual	Currency default 0	not null	//普通時間
,	Memo	Char(80) default '' not null	//備考
,	PRIMARY KEY (
		StaffNo
	,	Dt
	)
) #
