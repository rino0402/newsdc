@echo off
rem Glics出荷データ連携
rem 2016.06.22 w6レンジ対応
rem 2017.05.16 ログ出力メール log@kk-sdc.co.jp
rem 2017.07.03 ログ出力slack
set ret=0
set fName=%1
set fSave=out_save\%1
set Bu=%2
if exist %fSave% goto _End
call d:\log\batlog △ %0 %*
echo.%0 %* > getoutg.txt
dir g:\gift\recv\%fName% | findstr /i %fName% >> getoutg.txt
echo.%DATE%  %TIME:~0,8% △ >> getoutg.txt
color 9F

echo.■getoutg %*
echo.■■Glics出荷 データ連携
set NewSdc=%3
if "%NewSdc%" == ""	set NewSdc=newsdc

echo.xcopy/d/y g:\gift\recv\%fName% out_save\ ...%time%
xcopy/d/y g:\gift\recv\%fName% out_save\
echo.xcopy/d/y g:\gift\recv\%fName% out_save\ ...%time%完了
for %%i in (d:\%NewSdc%\hostfile\shiji_out_?.dat) do copy nul %%i > nul && echo.%%i
for %%i in (d:\%NewSdc%\hostfile\shiji_in_?.dat) do copy nul %%i > nul && echo.%%i
copy %fSave%	d:\%NewSdc%\hostfile\shiji_out_%Bu%.dat
copy nul		d:\%NewSdc%\hostfile\shiji_in_%Bu%.dat
tool\convcrlf d:\%NewSdc%\hostfile\shiji_out_%Bu%.dat

echo.■■Glics出荷 データ変換処理
if exist d:\%NewSdc%\FILES\NG_FILE.TXT del d:\%NewSdc%\FILES\NG_FILE.TXT
tool\lha32 a %fSave%.lzh d:\%NewSdc%\hostfile\shiji_*_?.dat
d:\%NewSdc%\exe\f102010
if exist d:\%NewSdc%\FILES\NG_FILE.TXT goto _Error
type beeps.txt

rem if "%Bu%" == "R" (
rem  	echo.冷蔵庫の出荷予定を出庫済にセット
rem  	pvddl %NewSdc% y_syuka_rf.sql
rem )
rem  -------------------------------
cscript//nologo d:\%NewSdc%\files\glicspos.vbs /db:%NewSdc% out_save\%fName%
cscript//nologo d:\%NewSdc%\files\hmem700.vbs /db:%NewSdc% %fName% >> getoutg.txt
rem  -------------------------------
goto _End_Log

:_Error
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
copy d:\%NewSdc%\FILES\NG_FILE.TXT %fSave%.ng
echo.■■Glics出荷 データ変換 %fName% でエラーが発生しました。
echo.■■再試行します。
tool\blatj d:\%NewSdc%\FILES\NG_FILE.TXT -s "Error:%0 %*" -t system@kk-sdc.co.jp
call d:\newsdc\tool\slack "■Glics出荷 `Error` %0 %*" d:\%NewSdc%\FILES\NG_FILE.TXT
del %fSave% 
goto _End

:_End_Log
echo.%DATE%  %TIME:~0,8% ▽ >> getoutg.txt
call d:\newsdc\tool\slack "■Glics出荷 %NewSdc%" %cd%\getoutg.txt
echo.■■液晶ディスプレイ通知
xcopy/d/y getoutg.txt d:\%NewSdc%\files\notice\
xcopy/d/y getoutg.txt \\hs1\it\pos\newsdc\files\notice\

rem  -------------------------------(★)
if /i "%Bu%" == "6" (
	d:\newsdc\tool\blatj getoutg.txt -s "■Glics出荷 %NewSdc% %Bu%" -t sdc.nara.e5@gmail.com -c system@kk-sdc.co.jp
)
rem -------
call d:\log\batlog ▽ %0 %*
set ret=1
:_End
color
for %%i in (%fSave%) do (
	echo.■getoutg %1 %2 %%~zi
)
exit/b %ret%
