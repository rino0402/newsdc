@echo off
rem �h���C��Glics
rem 2013.11.07 ng�t�@�C����ng�t�H���_�ɕۑ�����悤�ɕύX
rem 2014.02.28 ���O�o�͑Ή�
rem 2016.03.07 g�h���C�u�ɕύX
@if exist in_ok\%1 goto _End
call d:\log\batlog �� %0 %*

@echo ��getinz %*
@if exist %2 goto _Convert
	@echo ���׎w���f�[�^�ϊ�
	copy g:\ftpsend\%3\%2
	@if not exist %2 goto _Error
	goto _Convert

:_Convert
@for %%i in (d:\newsdc\hostfile\shiji_out_?.txt) do copy nul %%i > nul && echo %%i
@for %%i in (d:\newsdc\hostfile\shiji_in_?.txt) do copy nul %%i > nul && echo %%i
@for %%i in (d:\newsdc\hostfile\shiji_out_?.dat) do copy nul %%i > nul && echo %%i
@for %%i in (d:\newsdc\hostfile\shiji_in_?.dat) do copy nul %%i > nul && echo %%i
rem @del/q getinz.lzh
rem @tool\lha32 a getinz %2
rem @dir %2 > getinz.txt
rem @if %~z2 NEQ 0 @tool\blatj getinz.txt -attach getinz.lzh -s "[getinz %4] %2" -t system@kk-sdc.co.jp -c m.yoshizawa@adhoc.iplus.to
copy nul  d:\newsdc\hostfile\shiji_out_%4.txt
copy %2   d:\newsdc\hostfile\shiji_in_%4.txt
tool\convcrlf d:\newsdc\hostfile\shiji_in_%4.txt
sort /r d:\newsdc\hostfile\shiji_in_%4.txt > d:\newsdc\hostfile\shiji_in_%4.dat
rem copy  d:\newsdc\hostfile\shiji_in_%4.txt d:\newsdc\hostfile\shiji_in_%4.dat
rem sort < d:\newsdc\hostfile\shiji_in_%4.txt > d:\newsdc\hostfile\shiji_in_%4.dat

@echo ���׃f�[�^�ϊ������v���O����
@if exist d:\newsdc\FILES\NG_FILE.TXT del d:\newsdc\FILES\NG_FILE.TXT
tool\lha32 a in_save\%2.lzh d:\newsdc\hostfile\shiji_in_%4.dat
@d:\newsdc\exe\f102010
xcopy/y/d %2 in_save\
set SZ=%~z2
del %2
@if exist d:\newsdc\FILES\NG_FILE.TXT goto _Error

@type beeps.txt
xcopy/y/d %1 in_ok\

@if %SZ% equ 0 @goto _End_Log

@if exist delete-gensan.sql pvddl newsdc delete-gensan.sql -stoponfail
@if exist y-nyuka-set-9.sql pvddl newsdc y-nyuka-set-9.sql -stoponfail
@d:\newsdc\exe\F102090
@d:\newsdc\exe\F103000

@if not "%3" == "ono" @goto _InspMessageEnd
@echo �� item ���i���b�Z�[�W�X�V�F���`�E���d�r����
if exist item-insp-message.log del item-insp-message.log
cscript     item-insp-message.vbs
for %%i in (item-insp-message.log) do if %%~zi neq 0 (
	echo ������������������������������������������������ > mail.txt
	echo ���i��MST ���i���b�Z�[�W�X�V�F���`�E���d�r���ځ� >> mail.txt
	echo ���@�@���}�F���i���w�}���ɓo�^���Ă��������B�@�� >> mail.txt
	echo ������������������������������������������������ >> mail.txt
	type item-insp-message.log >>mail.txt
	tool\blatj mail.txt -attach item-insp-message.log -s "�i��MST ���i���b�Z�[�W�X�V�F���`�E���d�r����" -t %ML% -c system@kk-sdc.co.jp
)
@echo �� item ���i���b�Z�[�W�X�V�F���p���i
if exist item_insp_message.log del item_insp_message.log
cscript     item_insp_message.vbs /update
for %%i in (item_insp_message.log) do if %%~zi neq 0 (
	echo ���������������������������������������������� > mail.txt
	echo ���i��MST ���i���b�Z�[�W�X�V�F���p���i�ł��B�� >> mail.txt
	echo ���������������������������������������������� >> mail.txt
	type item_insp_message.log >>mail.txt
	tool\blatj mail.txt -attach item_insp_message.log -s "�i��MST ���i���b�Z�[�W�X�V�F���p���i�ł��B" -t %ML% -c system@kk-sdc.co.jp
)
:_InspMessageEnd
@echo �� ���Y���}�X�^�[ �X�V���ɓ��׃f�[�^�̓o�^�����Z�b�g
cscript y_nyuka_gensan.vbs /update
@goto _End_Log

@:_Error
@type beeps.txt
@type beeps.txt
@type beeps.txt
@type beeps.txt
@type beeps.txt
copy d:\newsdc\FILES\NG_FILE.TXT ng\%2.ng
@echo ���׎w�� %2 �ŃG���[���������܂����B
@echo 1����ɍĎ��s���܂��B

:_End_Log
call d:\log\batlog �� %0 %*
:_End
@del %1
@for %%i in (in_save\%2) do @echo ��getinz %2 %3 %4 %%~zi
