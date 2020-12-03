/*
atnd.sql
pvddl newsdc atnd.sql -server w3
pvddl newsdc atnd.sql -server w1
pvddl newsdc atnd.sql -server w4
pvddl newsdc atnd.sql -server w5
*/
Drop Table Atnd IN DICTIONARY #

Create Table Atnd using 'Atnd.DAT' with replace (
	StaffNo	Char( 8) default '' not null	//従業員番号
,	Dt		Date				not	null	//日付
,	Shift	Char( 2) default '' not null	//入力 シフト
,	BegTm	Time							//出勤打刻
,	FinTm	Time							//退勤打刻
,	Late	Currency default 0	not null	//入力 遅刻
,	Early	Currency default 0	not null	//入力 早退
,	PTO		Currency default 0	not null	//入力 有給, 有給休暇: paid holiday, paid time off, PTO
,	Actual	Currency default 0	not null	//所定内
,	Extra	Currency default 0	not null	//残業
,	Night	Currency default 0	not null	//深夜
,	Memo	Char(80) default '' not null	//入力 備考
,	BegTm_i		Time						//入力 出勤
,	FinTm_i		Time						//入力 退勤
,	Actual_i	Currency 					//入力 所定内
,	Extra_i		Currency					//入力 残業
,	Night_i		Currency					//入力 深夜
,	PRIMARY KEY (
		StaffNo
	,	Dt
	)
) #
