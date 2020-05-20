@echo off
setlocal
cd/d %~dp0
echo.%0 2016.10.03
echo.二重起動チェック
tasklist /V /FI "IMAGENAME eq cmd.exe" | findstr "Glics連携"
if %ERRORLEVEL% == 0 goto _End
if exist y_syuka_check.send del y_syuka_check.send
title Glics連携
rem 2008で TZUTIL がないのでエラーメッセージを非表示
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
	title Glics連携 %DT% %TM% %ComputerName%
	echo.Glics連携 %DT% %TM% %ComputerName%

	cscript //Nologo date.vbs
	set TmpDt=%errorlevel%
	if not "%TmpDt%" == "%PrevDate%" (
		echo.Glics連携開始 %date% %time:~0,8% > mail.txt
		echo.前日 %TmpDt% >> mail.txt
		echo.当日 %PrevDate% >> mail.txt
		echo.%0 %* >> mail.txt
		set ComputerName >> mail.txt
		set ML >> mail.txt
		call d:\newsdc\tool\slack "Glics連携開始" %cd%\mail.txt

		echo.前日処理 %PrevDate% %TmpDt%
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
	echo.%WT%秒待機中...通常60秒 18:00以降30分 %DT% %TM%
rem	start /min /wait timeout /t %WT%
	timeout /t %WT%
if %DT% == %DATE:/=% goto _Top
echo.%DT% %TM% 本日の処理は終了
:_End
title 連携終了 %date% %time%
rem debug要
echo.Glics連携終了 %date% %time:~0,8% > mail.txt
echo.%0 %* >> mail.txt
call d:\newsdc\tool\slack "Glics連携終了" %cd%\mail.txt
call ItemInsp.bat
endlocal
exit
