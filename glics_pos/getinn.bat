@echo off
setlocal
rem ¤Ê
rem 2014.02.28 OoÍÎ
rem 2017.03.08 ActiveGiftÎ
rem getinn HMTAH500SEC.dat.20170308-115509.OK HMTAH500SEC.dat.20170308-115509 7 
rem getinn A:\HMTAH500SEC.dat.20170308-115509 7
set ret=0
set absPath=%1
set relPath=%~nx1
set Bu=%2
if exist in_save\%relPath% goto _End
call d:\log\batlog ¢ %0 %*
color 9F
echo ¡getinn %*
dir %absPath% | findstr /i %relPath% > getinn.txt
echo.%DATE%  %TIME:~0,8% ¢ >> getinn.txt

echo.¡¡ü×w¦f[^Ï·
echo.xcopy/d/y %absPath% ...%time%
xcopy/d/y %absPath%
echo.xcopy/d/y %absPath% ...%time%®¹
if not "%ERRORLEVEL%"  == "0" (
	call d:\log\batlog ¥ %0 %* xcopy:%ERRORLEVEL%
	call d:\newsdc\tool\slack "¥ %0 %* xcopy:%ERRORLEVEL%"
	del %relPath%
	GOTO _End
)

copy %relPath% d:\newsdc\hostfile\new_shiji_in_%Bu%.dat > nul
tool\convcrlf d:\newsdc\hostfile\new_shiji_in_%Bu%.dat
copy nul d:\newsdc\hostfile\new_shiji_out_%Bu%.dat > nul
if "%Bu%" == "ono" (
	copy d:\newsdc\hostfile\new_shiji_in_%Bu%.dat d:\newsdc\hostfile\new_shiji_in_4.dat
	copy nul d:\newsdc\hostfile\new_shiji_in_5.dat
	copy nul d:\newsdc\hostfile\new_shiji_in_d.dat
	copy nul d:\newsdc\hostfile\new_shiji_out_4.dat
	copy nul d:\newsdc\hostfile\new_shiji_out_5.dat
	copy nul d:\newsdc\hostfile\new_shiji_out_d.dat
)
:_ExeConv
echo.¡¡Ï·
if exist d:\newsdc\FILES\NG_FILE.TXT del d:\newsdc\FILES\NG_FILE.TXT
d:\newsdc\exe\F102015
if exist d:\newsdc\FILES\NG_FILE.TXT goto _Error

if exist y_nyuka.vbs (
	echo.¡¡getinn ÇiÔi`FbN ¡¡
	cscript y_nyuka.vbs
	for %%i in (y_nyuka.txt) do if %%~zi neq 0 (
		echo.¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡ > mail.txt
rem          ¡ÇiÔi  àÅØXg:HMTAH500SCS.dat.20110106-000001¡
		echo.¡ÇiÔi  àÅØXg:%2¡ >> mail.txt
		echo.¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡ >> mail.txt
		type y_nyuka.txt >>mail.txt
		tool\blatj mail.txt -s "ÇiÔi àÅØXg:%2" -t %ML% -c system@kk-sdc.co.jp
		echo.%0 %* >> mail.txt
		call d:\newsdc\tool\slack "¡ÇiÔi àÅØXg" %cd%\mail.txt
		copy/y mail.txt in_save\%relPath%.txt
	)
)
goto _End_Log

:_Error
	type beep1.txt
	echo.¡¡getinn üoÉw¦Ï· G[ %relPath% ¡¡
	echo.1ªãÉÄsµÜ·B
	call d:\newsdc\tool\slack "¡¡getinn üoÉw¦Ï· G[ %cd%\%relPath% ¡¡"
	goto _End
:_End_Log
rem  -------------------------------
cscript//nologo d:\NewSdc\files\glicspos.vbs %relPath%
cscript//nologo d:\NewSdc\files\hmem500.vbs /table:hmtah500 %relPath% >> getinn.txt

type beeps.txt
echo.¡¡getinn üoÉw¦Ï· ok %relPath% ¡¡
call d:\log\batlog ¤ %0 %*
echo.%DATE%  %TIME:~0,8% ¤ >> getinn.txt
echo.%0 %* >> getinn.txt
call d:\newsdc\tool\slack "¡ActiveUÖf[^" %cd%\getinn.txt
move/y %relPath% in_save\
echo.¡¡t»fBXvCÊm
xcopy/d/y getinn.txt d:\newsdc\files\notice\
xcopy/d/y getinn.txt \\hs1\it\pos\newsdc\files\notice\
set ret=1
:_End
color
for %%i in (in_save\%relPath%) do echo.¡getinn  %* %%~zi
endlocal && set ret=%ret%
exit/b %ret%
