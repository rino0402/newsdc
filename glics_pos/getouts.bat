@echo off
rem 2013.11.07 ng�t�@�C����ng�t�H���_�ɕۑ�����悤�ɕύX
rem 2014.02.28 ���O�o�͑Ή�

@if exist out_ok\%1 goto _End
@echo ��getouts %*
call d:\log\batlog �� %0 %*
@if exist %2 goto _Convert
	copy H:\ftpsend\%3\%2 > nul
:_Convert
copy %2  d:\newsdc\hostfile\shiji_out_%4.dat > nul
tool\convcrlf d:\newsdc\hostfile\shiji_out_%4.dat
copy nul d:\newsdc\hostfile\shiji_in_%4.dat > nul

@echo �ϊ������v���O����
@if exist d:\newsdc\FILES\NG_FILE.TXT del d:\newsdc\FILES\NG_FILE.TXT
d:\newsdc\exe\f102010
xcopy/y/d %~n1 out_save\
@del %2
@if exist d:\newsdc\FILES\NG_FILE.TXT goto _Error

@type beeps.txt
xcopy/y/d %1 out_ok\
@echo ����getouts �o�׎w���ϊ� ok %2 ����
call d:\log\batlog �� %0 %*
goto _End

:_Error
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
copy d:\newsdc\FILES\NG_FILE.TXT ng\%2.ng
@echo ����getouts �o�׎w���ϊ� �G���[ %2 ����
@echo 1����ɍĎ��s���܂��B

:_End
@del %1
@for %%i in (out_save\%2) do @echo ��getouts %2 %3 %4 %%~zi
