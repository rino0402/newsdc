@echo off
setlocal
rem 2011.02.10 copy�Ɏ��s�����ꍇ�I������悤�ɕύX
rem 2013.04.05 ConvCrlf �� more �ł���悤�ɕύX ���R�}���h�v�����v�g���c��΍�
rem 2014.02.27 �o�b�`���s���O(batlog)
rem 2014.04.22 ����p�t�@�C���o��(getcomps.done)
rem 2017.03.08 ActiveGift�Ή�
set absPath=%1
set relPath=%~nx1
set Bu=%2
if exist complt\%relPath% goto _End
call d:\log\batlog �� %0 %*
echo.��getcomps %*
xcopy/d/y %absPath% complt\
if %~z1 NEQ 0 (
	type beeps.txt
	echo.%relPath% >> complt\%relPath%
	tool\blatj complt\%relPath% -s "��Active�o�׎��уG���[" -t %ML% -c system@kk-sdc.co.jp
	type beeps.txt
)
call d:\newsdc\tool\slack "��Active�o�׎��уG���[" %cd%\complt\%relPath%
copy/y complt\%relPath getcomps.done
call d:\log\batlog �� %0 %*
:_End
for %%i in (complt\%relPath%) do echo.��getcomps %* %%~zi %ML%
endlocal
