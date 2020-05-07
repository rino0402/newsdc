@echo off
rem nas02バックアップ
rem nas02.bat e:\backup w3
rem nas02.bat e:\backup w6
rem nas02.bat e:\backup w7
echo.%date% %time:~0,8% %0 %* > d:\log\nas02.txt
if not exist \\nas02\backup net use \\nas02\backup /user:sdch\administrator 123daaa
set	XD=/XD "$RECYCLE.BIN" "System Volume Information" temp
set	XF=/XF pagefile.sys
if /i "%Computername%" == "w7" (
	set	XJ=
) else (
	set	XJ=/XJD /XJF
)
set OPT=/mir /XO %XJ% /r:0 /w:0 %XD% %XF% /TS /TEE /NDL /FFT
robocopy %1 \\nas02\backup\%2 %OPT% /LOG:d:\log\nas02.log
echo.%date% %time:~0,8% %0 %* >> d:\log\nas02.txt
call d:\newsdc\tool\slack "backup nas02" d:\log\nas02.txt
