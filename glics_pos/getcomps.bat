@echo off
setlocal
rem 2011.02.10 copyに失敗した場合終了するように変更
rem 2013.04.05 ConvCrlf を more でするように変更 ※コマンドプロンプトが残る対策
rem 2014.02.27 バッチ実行ログ(batlog)
rem 2014.04.22 制御用ファイル出力(getcomps.done)
rem 2017.03.08 ActiveGift対応
set absPath=%1
set relPath=%~nx1
set Bu=%2
if exist complt\%relPath% goto _End
call d:\log\batlog △ %0 %*
echo.■getcomps %*
xcopy/d/y %absPath% complt\
if %~z1 NEQ 0 (
	type beeps.txt
	echo.%relPath% >> complt\%relPath%
	tool\blatj complt\%relPath% -s "■Active出荷実績エラー" -t %ML% -c system@kk-sdc.co.jp
	type beeps.txt
)
call d:\newsdc\tool\slack "■Active出荷実績エラー" %cd%\complt\%relPath%
copy/y complt\%relPath getcomps.done
call d:\log\batlog ▽ %0 %*
:_End
for %%i in (complt\%relPath%) do echo.■getcomps %* %%~zi %ML%
endlocal
