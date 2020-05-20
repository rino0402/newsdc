@echo off
setlocal
pushd %~dp0
rem 2008.02.08 complt.batより分離
rem 2011.02.07 バックコマンドを変更 copy → xcopy
rem 2011.02.07 作業ログ集計処理 を追加
rem 2013.11.12 作業ログ集計処理 の処理順を最初から最後に変更
rem 2013.12.26 filesバックアップ を最後に変更
rem	           welcat.txtローテーション を追加
rem 2014.02.27 バッチ実行ログ(batlog)
call d:\log\batlog △ %0 %*
set	NewSdc=%1
if "%NewSdc%"=="" set	NewSdc=newsdc
set NewSdc
echo.%DATE% %TIME:~0,8% %NewSdc% △ > posend.txt

tasklist /FI "IMAGENAME eq F110010.exe" | findstr /i F110010.exe
if "%ERRORLEVEL%" == "0" (
	echo.■スキャナ制御起動中...
	echo.%DATE% %TIME:~0,8% F110010 スキャナ制御起動中 >> posend.txt
)

echo.■Glics連携チェック■
echo.%DATE% %TIME:~0,8% Glics連携チェック >> posend.txt
pvddl %NewSdc% complt.sql

if /i "%NewSdc%" neq "newsdcn" (
	echo.■y_syuka 出庫済を検品済にセット
	echo.%DATE% %TIME:~0,8% y_syuka 出庫済を検品済にセット >> posend.txt
	pvddl %NewSdc% y_syuka_kenpin.sql
)

echo.■F110070:出荷予定削除
echo.%DATE% %TIME:~0,8% F110070:出荷予定削除 >> posend.txt
d:\%NewSdc%\exe\f110070.exe

echo.■F110030:不要データ削除
echo.%DATE% %TIME:~0,8% F110030:不要データ削除 >> posend.txt
d:\%NewSdc%\exe\F110030.exe

echo ■作業時間セット■
rem call D:\newsdc\FILES\sagyolog.bat

rem echo ■welcat.txtローテーション■
rem call \\hs1\it\bin\rotate d:\newsdc\files\welcat\welcat.txt

echo.■POS filesバックアップ■
echo.%DATE% %TIME:~0,8% filesバックアップ >> posend.txt
xcopy/d/y d:\%NewSdc%\files\*.* d:\%NewSdc%\backup\files\

rem echo %date% %time% >> posend.log
rem echo ■在庫集計処理:F109010■ 2007.05.22 処理停止(zaiko.batのみ実行)
rem \\w1\newsdc\exe\F109010.exe
echo.%DATE% %TIME:~0,8% %NewSdc% ▽ >> posend.txt
call d:\log\batlog ▽ %0 %*
call d:\newsdc\tool\slack "%0 %*" %cd%\posend.txt
popd
endlocal
rem call d:\log\batlog ● %0 %* shutdown
rem shutdown /r /f
