@echo off
setlocal
rem	2016.07.29 �ȑf�� WebDrive���g�p���Ȃ�
rem 2016.10.05 �ȑf��
rem 2017.03.09 ActiveGift�Ή�
pushd %~dp0
echo.%0 %* �o�׊����f�[�^���M %DATE:/=%
echo.%0 %* > complt.log

call d:\log\batlog �� %0 %*
for /f "tokens=1,2 delims=:" %%i in ( 'time/t' ) do set HH=%%i
set HMEM790=hmem790r%1.dat

echo.�o�׎��эđ��M�f�[�^�`�F�b�N
if exist getouty.done (
	dir getouty.done | findstr /i getouty.done >> complt.log
	pvddl newsdc complt.sql
	call d:\log\batlog �� %0 %* %HMEM790%
	del getouty.done
)

echo.f120090:�o�׎��уf�[�^�o��
del d:\newsdc\hostfile\syuka.txt
call d:\log\batlog �� %0 %* f120090
d:\newsdc\exe\f120090
call d:\log\batlog �� %0 %* f120090

echo.�t�@�C�����M %HMEM790%
copy/y d:\newsdc\hostfile\syuka.txt g:\active\%HMEM790%
echo.�t�@�C�����M %HMEM790%.ok
copy/y nul  g:\active\%HMEM790%.ok

rem �t�@�C���̍X�V������DTTM�ɃZ�b�g
for %%i in ( g:\active\%HMEM790% )  do set DTTM=%%~ti
for /f "tokens=1,2,3,4,5 delims=/: " %%i in ( "%DTTM%" ) do set DTTM=%%i%%j%%k-%%l%%m
copy/y g:\active\%HMEM790%		complt\%HMEM790%.%DTTM%
copy/y g:\active\%HMEM790%.ok	complt\%HMEM790%.%DTTM%.ok
rem ���M�����߁[��
dir g:\active\%HMEM790%.* | findstr/i %HMEM790% >> complt.log
rem tool\blatj mail.txt -s "%0 %*" -t log@kk-sdc.co.jp
call d:\newsdc\tool\slack "��Active�o�׊����f�[�^���M" %cd%\complt.log
call y_syuka_check.bat
call d:\log\batlog �� %0 %* %HMEM790%.%DTTM%
:_End
popd
endlocal
