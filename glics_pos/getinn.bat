@echo off
setlocal
rem 共通
rem 2014.02.28 ログ出力対応
rem 2017.03.08 ActiveGift対応
rem getinn HMTAH500SEC.dat.20170308-115509.OK HMTAH500SEC.dat.20170308-115509 7 
rem getinn A:\HMTAH500SEC.dat.20170308-115509 7
set ret=0
set absPath=%1
set relPath=%~nx1
set Bu=%2
if exist in_save\%relPath% goto _End
call d:\log\batlog △ %0 %*
color 9F
echo ■getinn %*
dir %absPath% | findstr /i %relPath% > getinn.txt
echo.%DATE%  %TIME:~0,8% △ >> getinn.txt

echo.■■入荷指示データ変換
echo.xcopy/d/y %absPath% ...%time%
xcopy/d/y %absPath%
echo.xcopy/d/y %absPath% ...%time%完了
if not "%ERRORLEVEL%"  == "0" (
	call d:\log\batlog ▼ %0 %* xcopy:%ERRORLEVEL%
	call d:\newsdc\tool\slack "▼ %0 %* xcopy:%ERRORLEVEL%"
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
echo.■■変換処理
if exist d:\newsdc\FILES\NG_FILE.TXT del d:\newsdc\FILES\NG_FILE.TXT
d:\newsdc\exe\F102015
if exist d:\newsdc\FILES\NG_FILE.TXT goto _Error

if exist y_nyuka.vbs (
	echo.■■getinn 良品返品チェック ■■
	cscript y_nyuka.vbs
	for %%i in (y_nyuka.txt) do if %%~zi neq 0 (
		echo.■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ > mail.txt
rem          ■良品返品  国内供給打切リスト:HMTAH500SCS.dat.20110106-000001■
		echo.■良品返品  国内供給打切リスト:%2■ >> mail.txt
		echo.■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ >> mail.txt
		type y_nyuka.txt >>mail.txt
		tool\blatj mail.txt -s "良品返品 国内供給打切リスト:%2" -t %ML% -c system@kk-sdc.co.jp
		echo.%0 %* >> mail.txt
		call d:\newsdc\tool\slack "■良品返品 国内供給打切リスト" %cd%\mail.txt
		copy/y mail.txt in_save\%relPath%.txt
	)
)
goto _End_Log

:_Error
	type beep1.txt
	echo.■■getinn 入出庫指示変換 エラー %relPath% ■■
	echo.1分後に再試行します。
	call d:\newsdc\tool\slack "■■getinn 入出庫指示変換 エラー %cd%\%relPath% ■■"
	goto _End
:_End_Log
rem  -------------------------------
cscript//nologo d:\NewSdc\files\glicspos.vbs %relPath%
cscript//nologo d:\NewSdc\files\hmem500.vbs /table:hmtah500 %relPath% >> getinn.txt

type beeps.txt
echo.■■getinn 入出庫指示変換 ok %relPath% ■■
call d:\log\batlog ▽ %0 %*
echo.%DATE%  %TIME:~0,8% ▽ >> getinn.txt
echo.%0 %* >> getinn.txt
call d:\newsdc\tool\slack "■Active振替データ" %cd%\getinn.txt
move/y %relPath% in_save\
echo.■■液晶ディスプレイ通知
xcopy/d/y getinn.txt d:\newsdc\files\notice\
xcopy/d/y getinn.txt \\hs1\it\pos\newsdc\files\notice\
set ret=1
:_End
color
for %%i in (in_save\%relPath%) do echo.■getinn  %* %%~zi
endlocal && set ret=%ret%
exit/b %ret%
