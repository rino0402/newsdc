@echo off
rem ����
rem 2013.11.07 ng�t�@�C����ng�t�H���_�ɕۑ�����悤�ɕύX
rem 2014.02.28 ���O�o�͑Ή�
@if exist in_ok\%1 goto _End
call d:\log\batlog �� %0 %*
@echo ��getins %*
@if exist %2 goto _Convert
	@echo ���׎w���f�[�^�ϊ�
	copy H:\ftpsend\%3\%2
	@if not exist %2 goto _Error
	goto _Convert
:_Convert
copy nul  d:\newsdc\hostfile\shiji_out_%4.txt
copy %2   d:\newsdc\hostfile\shiji_in_%4.txt
tool\convcrlf d:\newsdc\hostfile\shiji_in_%4.txt
sort < d:\newsdc\hostfile\shiji_in_%4.txt > d:\newsdc\hostfile\shiji_in_%4.dat

@echo �ϊ������v���O����
@if exist d:\newsdc\FILES\NG_FILE.TXT del d:\newsdc\FILES\NG_FILE.TXT
d:\newsdc\exe\f102010
xcopy/y/d %2 in_save\
del %2
@if exist d:\newsdc\FILES\NG_FILE.TXT goto _Error

type beeps.txt
xcopy/y/d %1 in_ok\
call d:\log\batlog �� %0 %*
goto _End

:_Error
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
copy/y d:\newsdc\FILES\NG_FILE.TXT ng\%2.ng
@echo ���׎w�� %2 �ŃG���[���������܂����B
@echo 1����ɍĎ��s���܂��B

:_End
@del %1
@for %%i in (in_save\%2) do @echo ��getins %2 %3 %4 %%~zi
