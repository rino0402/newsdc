@echo off
setlocal
cd/d %~dp0
echo.%0 2016.10.03
echo.��d�N���`�F�b�N
tasklist /V /FI "IMAGENAME eq cmd.exe" | findstr "Glics�A�g"
if %ERRORLEVEL% == 0 goto _End
if exist y_syuka_check.send del y_syuka_check.send
title Glics�A�g
rem 2008�� TZUTIL ���Ȃ��̂ŃG���[���b�Z�[�W���\��
TZUTIL /s "Tokyo Standard Time_dstoff" > nul 2>&1
if "%ComputerName%" == "W1" (
	set ML=onopc@kk-sdc.co.jp
)
if "%ComputerName%" == "W2" (
	set ML=fukuroipc@kk-sdc.co.jp
)
if "%ComputerName%" == "W3" (
	set ML=shigapc@kk-sdc.co.jp
)
if "%ComputerName%" == "W4" (
	set ML=shigadc@kk-sdc.co.jp
)
if "%ComputerName%" == "W6" (
	set ML=kabu.mo@kk-sdc.co.jp
)
set DT=%DATE:/=%
set PrevDate=%DT%
:_Top
	set TM=%TIME: =0%
	set TM=%TM:~0,5%
	title Glics�A�g %DT% %TM% %ComputerName%
	echo.Glics�A�g %DT% %TM% %ComputerName%

	cscript //Nologo date.vbs
	set TmpDt=%errorlevel%
	if not "%TmpDt%" == "%PrevDate%" (
		echo.Glics�A�g�J�n %date% %time:~0,8% > mail.txt
		echo.�O�� %TmpDt% >> mail.txt
		echo.���� %PrevDate% >> mail.txt
		echo.%0 %* >> mail.txt
		set ComputerName >> mail.txt
		set ML >> mail.txt
		call d:\newsdc\tool\slack "Glics�A�g�J�n" %cd%\mail.txt

		echo.�O������ %PrevDate% %TmpDt%
		call getall   %TmpDt%
		set PrevDate=%TmpDt%
	)
	call getall %DT%
	echo.getall %DT%.%ERRORLEVEL%
	if ERRORLEVEL 1 goto _Top

	set WT=60
	set TM=%TIME: =0%
	set TM=%TM:~0,5%
	if %TM% gtr 13:00 set WT=30
	if %TM% gtr 18:00 set WT=300
	if %TM% gtr 22:00 goto _End
	echo.%WT%�b�ҋ@��...�ʏ�60�b 18:00�ȍ~30�� %DT% %TM%
rem	start /min /wait timeout /t %WT%
	timeout /t %WT%
if %DT% == %DATE:/=% goto _Top
echo.%DT% %TM% �{���̏����͏I��
:_End
title �A�g�I�� %date% %time%
rem debug�v
echo.Glics�A�g�I�� %date% %time:~0,8% > mail.txt
echo.%0 %* >> mail.txt
call d:\newsdc\tool\slack "Glics�A�g�I��" %cd%\mail.txt
call ItemInsp.bat
endlocal
exit
