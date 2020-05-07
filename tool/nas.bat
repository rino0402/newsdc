@echo off
setlocal
pushd %~dp0
rem nasバックアップ
rem nas.bat \\nas02\backup\hs1 \\nas01\backup\hs1
title %date% %time:~0,8% %0 %*
echo.%date% %time:~0,8% %0 %* △> nas.tim
rem if not exist \\nas01\backup net use \\nas01\backup /user:administrator 123daaa
rem if not exist \\nas02\backup net use \\nas02\backup /user:sdch\administrator 123daaa
set	XD=/XD "$RECYCLE.BIN" "System Volume Information" temp "Application Data"
set	XF=/XF pagefile.sys Thumbs.db
set	XJ=/XJ
VER | find "Version 5.2." > nul
IF not errorlevel 1 (
	echo.Server2003
rem	set	XJ=/XJD /XJF
)
if "%3" == "" (
	set	MIR=/mir
) else (
	set	MIR=%3
)
set OPT=%MIR% /XO %XJ% /r:0 /w:0 %XD% %XF% /TS /TEE /NDL /FFT /256

set tm=%time: =0%
set tm=%tm::=%
set tm=%tm:.=%
set log=nas.log.%tm%
robocopy %1 %2 %OPT% /LOG:%log%
nkf32 -Lw %log% | findstr /V /C:"%%" > nas.tmp
rem nkf32 -Lw nas.log | findstr /V /C:"%%" | tail -40l > nas.tmp
echo.%date% %time:~0,8% %0 %* ▽>> nas.tim
type nas.tmp >  nas.txt
type nas.tim >> nas.txt
call slack "%0 %*" %cd%\nas.txt
popd
endlocal
