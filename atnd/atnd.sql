/*
atnd.sql
pvddl newsdc atnd.sql -server w3
pvddl newsdc atnd.sql -server w1
pvddl newsdc atnd.sql -server w4
pvddl newsdc atnd.sql -server w5
*/
Drop Table Atnd IN DICTIONARY #

Create Table Atnd using 'Atnd.DAT' with replace (
	StaffNo	Char( 8) default '' not null	//�]�ƈ��ԍ�
,	Dt		Date				not	null	//���t
,	Shift	Char( 2) default '' not null	//���� �V�t�g
,	BegTm	Time							//�o�Αō�
,	FinTm	Time							//�ދΑō�
,	Late	Currency default 0	not null	//���� �x��
,	Early	Currency default 0	not null	//���� ����
,	PTO		Currency default 0	not null	//���� �L��, �L���x��: paid holiday, paid time off, PTO
,	Actual	Currency default 0	not null	//�����
,	Extra	Currency default 0	not null	//�c��
,	Night	Currency default 0	not null	//�[��
,	Memo	Char(80) default '' not null	//���� ���l
,	BegTm_i		Time						//���� �o��
,	FinTm_i		Time						//���� �ދ�
,	Actual_i	Currency 					//���� �����
,	Extra_i		Currency					//���� �c��
,	Night_i		Currency					//���� �[��
,	PRIMARY KEY (
		StaffNo
	,	Dt
	)
) #
