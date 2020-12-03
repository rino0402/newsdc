/*
timepack.sql
pvddl newsdc timepack.sql -server w3
py timepack.py --dns newsdc3 "\\wakame\c$\kintai\滋賀 (7)"
pvddl newsdc timepack.sql -server w1
@TPEX
*/
Drop Table TimePack IN DICTIONARY #

Create Table TimePack using 'TimePack.DAT' with replace (
	CardNo	Char(10) default '' not null	//0"カード番号"
,	StaffNo	Char( 8) default '' not null	//1"従業員番号"
,	Name	Char(40) default '' not null	//2"従業員氏名"
,	Post	Char( 8) default '' not null	//3"所属番号"
,	Dt		Date				not null	//4"年/月/日"
,	Shift	Char( 2) default '' not null	//5"シフト番号"
,	Holiday	Char( 2) default '' not null	//6"平日/休日区分"
,	Absence	Char(40) default '' not null	//7"不在理由"
,	BgnTM	Time							//8"出勤打刻"
,	BgnMK	Char( 2) default '' not null	//9"出勤マーク"
,	OutTM	Time							//10"外出打刻"
,	OutMK	Char( 2) default '' not null	//11"外出マーク"
,	BckTM	Time							//12"戻打刻"
,	BckMK	Char( 2) default '' not null	//13"戻マーク"
,	FinTM	Time							//14"退勤打刻"
,	FinMK	Char( 2) default '' not null	//15"退勤マーク"
,	Ex1TM	Time							//16"例外１"
,	Ex1MK	Char( 2) default '' not null	//17"例外マーク"
,	Ex2TM	Time							//18"例外２"
,	Ex2MK	Char( 2) default '' not null	//19"例外２マーク"
,	Actual	Currency default 0	not null	//20"所定内時間"		,20"基準内時間"
,	Extra	Currency default 0	not null	//21"延長時間"			,21"基準外時間"
,	ExtEarly		Currency default 0	not null	//22"早出残業"	,22"深夜時間"
,	Night			Currency default 0	not null	//23"深夜時間"	,23"基準外深夜"
,	ExtNight		Currency default 0	not null	//24"深夜残業"	,24"休１時間"
,	Holiday1		Currency default 0	not null	//25"休１時間"	,25"休１深夜"
,	HolidayNight1	Currency default 0	not null	//26"休１深夜"	,26"休２時間"
,	Holiday2		Currency default 0	not null	//27"休２時間"	,27"休２深夜"
,	HolidayNight2	Currency default 0	not null	//28"休２深夜"	,28"外出時間"
,	LateEarly		Currency default 0	not null	//29"遅早時間"	,29"コメント","","","","","","",""
,	Private			Currency default 0	not null	//30"外出時間"
,	Memo			Char(40) default '' not null	//31"コメント","","","","",""
,	PRIMARY KEY (
		StaffNo
	,	Dt
	)
) #
