@echo off
rem 共通 2008.11.28 Active対応(メール送信機能追加)
rem 共通 2008.12.26 メール送信機能停止
rem 共通 2011.02.04 y_syuka の 取引区分=19 振替入庫 を削除
rem 共通 2012.03.26 BLBU対応
rem 共通 2013.11.05 出荷予定データ変換コントロール対応
rem 2013.11.07 ngファイルをngフォルダに保存するように変更
rem 2014.02.28 ログ出力対応
rem 2015.07.28 韓国リサイクルマーク個別表示確認
rem 2017.03.08 ActiveGift対応
set ret=0
set absPath=%1
set relPath=%~nx1
set Bu=%2
if exist out_save\%relPath% goto _End
echo.■getoutn %*
call d:\log\batlog △ %0 %*
color 9F
rem call slack "%relPath%:△"
echo.■■Active出荷
echo.xcopy/d/y %absPath% ...%time%
xcopy/d/y %absPath%
echo.xcopy/d/y %absPath% ...%time%完了
if not "%ERRORLEVEL%"  == "0" (
	call d:\log\batlog ▼ %0 %* xcopy:%ERRORLEVEL%
	call d:\newsdc\tool\slack "■getoutn %relPath%:▼xcopy:%ERRORLEVEL%"
	del %relPath%
	GOTO _End
)
set fSize=0
for %%i in ( %relPath% ) do set fSize=%%~zi
if %fSize% == 0 (
	call d:\log\batlog ▼ %0 %* fSize:%fSize%
	call d:\newsdc\tool\slack "■getoutn %relPath%:▼fSize:%fSize%"
	del %relPath%
	GOTO _End
)

for %%i in (d:\newsdc\hostfile\new_shiji_out_?.dat) do copy nul %%i
for %%i in (d:\newsdc\hostfile\new_shiji_in_?.dat) do copy nul %%i
copy %relPath% d:\newsdc\hostfile\new_shiji_out_%Bu%.dat > nul
tool\convcrlf d:\newsdc\hostfile\new_shiji_out_%Bu%.dat
copy nul d:\newsdc\hostfile\new_shiji_in_%Bu%.dat > nul
if "%Bu%" == "ono" (
	findstr 0002139700021397 d:\newsdc\hostfile\new_shiji_out_%Bu%.dat > d:\newsdc\hostfile\new_shiji_out_5.dat
	findstr 0002310000023100 d:\newsdc\hostfile\new_shiji_out_%Bu%.dat > d:\newsdc\hostfile\new_shiji_out_1.dat
	findstr 0002341000023410 d:\newsdc\hostfile\new_shiji_out_%Bu%.dat > d:\newsdc\hostfile\new_shiji_out_4.dat
	findstr 0002351000023510 d:\newsdc\hostfile\new_shiji_out_%Bu%.dat > d:\newsdc\hostfile\new_shiji_out_d.dat
	copy nul d:\newsdc\hostfile\new_shiji_in_5.dat
	copy nul d:\newsdc\hostfile\new_shiji_in_1.dat
	copy nul d:\newsdc\hostfile\new_shiji_in_4.dat
	copy nul d:\newsdc\hostfile\new_shiji_in_d.dat
)
:_ExeConv
tool\lha32 a out_save\%relPath%.lzh d:\newsdc\hostfile\new_shiji_*
if exist d:\newsdc\FILES\NG_FILE.TXT del d:\newsdc\FILES\NG_FILE.TXT
d:\newsdc\exe\F102015
if exist d:\newsdc\FILES\NG_FILE.TXT copy d:\newsdc\FILES\NG_FILE.TXT out_save\%relPath%.ng & goto _Error

type beeps.txt
echo.■■getoutn 出荷指示変換 ok %relPath% ■■
echo.■■getoutn 出荷指示変換 ok %relPath% ■■ > mail.txt
set DT=%DATE:/=%
for /f "tokens=1,2 delims=:" %%i in ( 'time/t' ) do set TM=%%i%%j
cscript getoutn.vbs newsdc %DT%%TM%
type getoutn.txt >> mail.txt
rem echo.>>mail.txt
echo.%0 %*>>mail.txt
call d:\newsdc\tool\slack "■Active出荷データ" %cd%\getoutn.txt
rem tool\blatj mail.txt -s "出荷指示変換:%2" -t %ML% -c system@kk-sdc.co.jp
if exist y_syuka-delete-19.sql pvddl newsdc y_syuka-delete-19.sql
echo.■■出荷予定データ変換コントロールファイル出力
copy/y mail.txt getoutn.ok
echo.■■液晶ディスプレイ通知
xcopy/d/y getoutn.ok d:\newsdc\files\notice\
xcopy/d/y getoutn.ok \\hs1\it\pos\newsdc\files\notice\
move/y %relPath% out_save\
echo.■■韓国リサイクルマーク個別表示確認
call HMTH011 out_save\%relPath%
call d:\log\batlog ▽ %0 %*
rem call slack "%relPath%:▽"
if exist d:\newsdc\files\ySize.sql (
	pvddl newsdc d:\newsdc\files\ySize.sql > ySize.log
	call d:\newsdc\tool\slack "ySize.log" %cd%\ySize.log
)
set ret=1
goto _End

:_Error
	type beeps.txt
	type beeps.txt
	type beeps.txt
	type beeps.txt
	type beeps.txt
	echo.■■getoutn 出荷指示変換 エラー %relPath% ■■
	echo.1分後に再試行します。
	tool\blatj out_save\%relPath%.ng% -s "出荷指示変換ERROR:%relPath%" -t %ML% -c system@kk-sdc.co.jp
	echo.>>d:\newsdc\FILES\NG_FILE.TXT
	echo.%0 %*>>d:\newsdc\FILES\NG_FILE.TXT
	call d:\newsdc\tool\slack "■出荷指示変換 エラー" d:\newsdc\FILES\NG_FILE.TXT
:_End
color
for %%i in (out_save\%relPath%) do echo ■getoutn %* %%~zi
exit/b %ret%
