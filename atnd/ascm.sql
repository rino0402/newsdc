/*
ascm.sql
pvddl newsdc ascm.sql -server w6
*/
Drop Table Ascm IN DICTIONARY #

Create Table Ascm using 'Ascm.DAT' with replace (
	StaffNo	Char( 8) default '' not null	//�Ј�No
,	Name	Char(40) default '' not null	//����
,	Dt		Date				not null	//���t
,	Kubun	Char(20) default '' not null	//�敪
,	Awh		Currency default 0	not null	//����
,	Shift	Char(10) default '' not null	//�V�t�gNo
,	ShiftNm	Char(20) default '' not null	//��Ė�
,	BegTm	Time							//�o��
,	FinTM	Time							//�ދ�
,	Late	Currency default 0	not null	//�x��
,	Early	Currency default 0	not null	//����
,	Extra	Currency default 0	not null	//���ʎc��
,	Night	Currency default 0	not null	//�[��c��
,	H1Extra	Currency default 0	not null	//�@��x�c
,	H1Night	Currency default 0	not null	//�@��x�[
,	H2Extra	Currency default 0	not null	//����x�c
,	H2Night	Currency default 0	not null	//����x�[
,	PTO		Currency default 0	not null	//�L���x��
,	Actual	Currency default 0	not null	//���ʎ���
,	Memo	Char(80) default '' not null	//���l
,	PRIMARY KEY (
		StaffNo
	,	Dt
	)
) #
