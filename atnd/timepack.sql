/*
timepack.sql
pvddl newsdc timepack.sql -server w3
py timepack.py --dns newsdc3 "\\wakame\c$\kintai\���� (7)"
pvddl newsdc timepack.sql -server w1
@TPEX
*/
Drop Table TimePack IN DICTIONARY #

Create Table TimePack using 'TimePack.DAT' with replace (
	CardNo	Char(10) default '' not null	//0"�J�[�h�ԍ�"
,	StaffNo	Char( 8) default '' not null	//1"�]�ƈ��ԍ�"
,	Name	Char(40) default '' not null	//2"�]�ƈ�����"
,	Post	Char( 8) default '' not null	//3"�����ԍ�"
,	Dt		Date				not null	//4"�N/��/��"
,	Shift	Char( 2) default '' not null	//5"�V�t�g�ԍ�"
,	Holiday	Char( 2) default '' not null	//6"����/�x���敪"
,	Absence	Char(40) default '' not null	//7"�s�ݗ��R"
,	BgnTM	Time							//8"�o�Αō�"
,	BgnMK	Char( 2) default '' not null	//9"�o�΃}�[�N"
,	OutTM	Time							//10"�O�o�ō�"
,	OutMK	Char( 2) default '' not null	//11"�O�o�}�[�N"
,	BckTM	Time							//12"�ߑō�"
,	BckMK	Char( 2) default '' not null	//13"�߃}�[�N"
,	FinTM	Time							//14"�ދΑō�"
,	FinMK	Char( 2) default '' not null	//15"�ދ΃}�[�N"
,	Ex1TM	Time							//16"��O�P"
,	Ex1MK	Char( 2) default '' not null	//17"��O�}�[�N"
,	Ex2TM	Time							//18"��O�Q"
,	Ex2MK	Char( 2) default '' not null	//19"��O�Q�}�[�N"
,	Actual	Currency default 0	not null	//20"���������"		,20"�������"
,	Extra	Currency default 0	not null	//21"��������"			,21"��O����"
,	ExtEarly		Currency default 0	not null	//22"���o�c��"	,22"�[�鎞��"
,	Night			Currency default 0	not null	//23"�[�鎞��"	,23"��O�[��"
,	ExtNight		Currency default 0	not null	//24"�[��c��"	,24"�x�P����"
,	Holiday1		Currency default 0	not null	//25"�x�P����"	,25"�x�P�[��"
,	HolidayNight1	Currency default 0	not null	//26"�x�P�[��"	,26"�x�Q����"
,	Holiday2		Currency default 0	not null	//27"�x�Q����"	,27"�x�Q�[��"
,	HolidayNight2	Currency default 0	not null	//28"�x�Q�[��"	,28"�O�o����"
,	LateEarly		Currency default 0	not null	//29"�x������"	,29"�R�����g","","","","","","",""
,	Private			Currency default 0	not null	//30"�O�o����"
,	Memo			Char(40) default '' not null	//31"�R�����g","","","","",""
,	PRIMARY KEY (
		StaffNo
	,	Dt
	)
) #
