/*
atnd_alt.sql
pvddl newsdc atnd_alt.sql -server w1
pvddl newsdc atnd_alt.sql -server w2
pvddl newsdc atnd_alt.sql -server w3
pvddl newsdc atnd_alt.sql -server w4
pvddl newsdc atnd_alt.sql -server w6
pvddl newsdc atnd_alt.sql -server w7
pvddl newsdc atnd_alt.sql -server w5

Alter Table Atnd (
 add	StartTm		Time							//始業 add
,add	FinishTm	Time							//終業 add
,add	StartTm_i	Time							//入力 始業 add
,add	FinishTm_i	Time							//入力 終業 add
)
Alter Table Atnd (
 add	PTO_tm		Currency default 0	not null	//入力 有給時間, 有給休暇: paid holiday, paid time off, PTO
)
*/
Alter Table Atnd (
 add	Dayoff		Currency default 0	not null	//休出 2021.04.12
,add	Dayoff_i	Currency						//入力 休出 2021.04.12
)
