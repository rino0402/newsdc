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
 add	StartTm		Time							//�n�� add
,add	FinishTm	Time							//�I�� add
,add	StartTm_i	Time							//���� �n�� add
,add	FinishTm_i	Time							//���� �I�� add
)
Alter Table Atnd (
 add	PTO_tm		Currency default 0	not null	//���� �L������, �L���x��: paid holiday, paid time off, PTO
)
*/
Alter Table Atnd (
 add	Dayoff		Currency default 0	not null	//�x�o 2021.04.12
,add	Dayoff_i	Currency						//���� �x�o 2021.04.12
)
