@echo off
rem 2013.11.07 ngファイルをngフォルダに保存するように変更
rem 2014.02.28 ログ出力対応
rem 2015.02.04 冷蔵庫の出荷予定を出庫済にセット
rem 2015.11.25 冷蔵庫の出荷予定を出庫済にセット CS本部のみ

if exist out_ok\%1 goto _End
echo ■getoutz(本番用) %*
call d:\log\batlog △ %0 %*

if exist %2 goto _Convert
	copy g:\ftpsend\%3\%2 > nul

:_Convert
for %%i in (d:\newsdc\hostfile\shiji_out_?.dat) do copy nul %%i
for %%i in (d:\newsdc\hostfile\shiji_in_?.dat) do copy nul %%i

copy %2  d:\newsdc\hostfile\shiji_out_%4.dat > nul
tool\convcrlf d:\newsdc\hostfile\shiji_out_%4.dat
copy nul d:\newsdc\hostfile\shiji_in_%4.dat > nul

echo 変換処理プログラム
if exist d:\newsdc\FILES\NG_FILE.TXT del d:\newsdc\FILES\NG_FILE.TXT
tool\lha32 a out_save\%2.lzh d:\newsdc\hostfile\shiji_out_%4.dat
d:\newsdc\exe\f102010
xcopy/y/d %~n1 out_save\
del %2
if exist d:\newsdc\FILES\NG_FILE.TXT goto _Error

type beeps.txt
xcopy/y/d %1 out_ok\
if "%4" == "R" (
 	echo 冷蔵庫の出荷予定を出庫済にセット
 	pvddl newsdc y_syuka_rf.sql
)
echo ■■getoutz 出荷指示変換 ok %2 ■■
call d:\log\batlog ▽ %0 %*
goto _End

:_Error
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
type beeps.txt
copy d:\newsdc\FILES\NG_FILE.TXT ng\%2.ng
echo ■■getoutz 出荷指示変換 エラー %2 ■■
echo 1分後に再試行します。

:_End
del %1
for %%i in (out_save\%2) do echo ■getoutz %2 %3 %4 %%~zi
