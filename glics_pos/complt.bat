@echo off
setlocal
rem	2016.07.29 簡素化 WebDriveを使用しない
rem 2016.10.05 簡素化
rem 2017.03.09 ActiveGift対応
pushd %~dp0
echo.%0 %* 出荷完了データ送信 %DATE:/=%
echo.%0 %* > complt.log

call d:\log\batlog △ %0 %*
for /f "tokens=1,2 delims=:" %%i in ( 'time/t' ) do set HH=%%i
set HMEM790=hmem790r%1.dat

echo.出荷実績再送信データチェック
if exist getouty.done (
	dir getouty.done | findstr /i getouty.done >> complt.log
	pvddl newsdc complt.sql
	call d:\log\batlog □ %0 %* %HMEM790%
	del getouty.done
)

echo.f120090:出荷実績データ出力
del d:\newsdc\hostfile\syuka.txt
call d:\log\batlog ┬ %0 %* f120090
d:\newsdc\exe\f120090
call d:\log\batlog ┴ %0 %* f120090

echo.ファイル送信 %HMEM790%
copy/y d:\newsdc\hostfile\syuka.txt g:\active\%HMEM790%
echo.ファイル送信 %HMEM790%.ok
copy/y nul  g:\active\%HMEM790%.ok

rem ファイルの更新日時をDTTMにセット
for %%i in ( g:\active\%HMEM790% )  do set DTTM=%%~ti
for /f "tokens=1,2,3,4,5 delims=/: " %%i in ( "%DTTM%" ) do set DTTM=%%i%%j%%k-%%l%%m
copy/y g:\active\%HMEM790%		complt\%HMEM790%.%DTTM%
copy/y g:\active\%HMEM790%.ok	complt\%HMEM790%.%DTTM%.ok
rem 送信完了めーる
dir g:\active\%HMEM790%.* | findstr/i %HMEM790% >> complt.log
rem tool\blatj mail.txt -s "%0 %*" -t log@kk-sdc.co.jp
call d:\newsdc\tool\slack "■Active出荷完了データ送信" %cd%\complt.log
call y_syuka_check.bat
call d:\log\batlog ▽ %0 %* %HMEM790%.%DTTM%
:_End
popd
endlocal
