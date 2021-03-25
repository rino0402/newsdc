@echo off
rem Glics振替データ連携
rem 2016.03.08 gift対応版
rem 2016.06.22 w6レンジ対応
rem 2017.05.16 ログ出力メール log@kk-sdc.co.jp
rem 2017.07.03 ログ出力slack
rem 2017.07.26 処理速度改善中...
rem 2019.06.06 伝票No.10桁
set ret=0
set fName=%1
set Bu=%2
if exist in_save\%fName% goto _End
call d:\log\batlog △ %0 %*
echo.%0 %* > geting.txt
dir g:\gift\recv\%fName% | findstr /i %fName% >> geting.txt
echo.%DATE%  %TIME:~0,8% △ >> geting.txt
color 9F

echo.■geting %*
echo.■■Glics振替 データ連携
set NewSdc=%3
if "%NewSdc%" == ""	set NewSdc=newsdc

echo.xcopy/d/y g:\gift\recv\%fName% in_save\ ...%time%
rem xcopy/d/y g:\gift\recv\%fName% in_save\
copy/y g:\gift\recv\%fName% in_save\
echo.xcopy/d/y g:\gift\recv\%fName% in_save\ ...%time%完了
for %%i in (d:\%NewSdc%\hostfile\shiji_out_?.txt) do copy nul %%i > nul && echo.%%i
for %%i in (d:\%NewSdc%\hostfile\shiji_in_?.txt ) do copy nul %%i > nul && echo.%%i
for %%i in (d:\%NewSdc%\hostfile\shiji_out_?.dat) do copy nul %%i > nul && echo.%%i
for %%i in (d:\%NewSdc%\hostfile\shiji_in_?.dat ) do copy nul %%i > nul && echo.%%i
copy nul  				d:\%NewSdc%\hostfile\shiji_out_%Bu%.txt
rem copy in_save\%fName%	d:\%NewSdc%\hostfile\shiji_in_%Bu%.txt
rem tool\convcrlf d:\%NewSdc%\hostfile\shiji_in_%Bu%.txt
echo.■■271→265
python d:\%NewSdc%\files\hmem500_271.py in_save\%fName% > d:\%NewSdc%\hostfile\shiji_in_%Bu%.265
sort /r d:\%NewSdc%\hostfile\shiji_in_%Bu%.265 > d:\%NewSdc%\hostfile\shiji_in_%Bu%.dat
rem  -------------------------------
echo.■■Glics振替 データ変換処理
if exist d:\%NewSdc%\FILES\NG_FILE.TXT del d:\%NewSdc%\FILES\NG_FILE.TXT
tool\lha32 a in_save\%fName%.lzh d:\%NewSdc%\hostfile\shiji_*_?.dat
d:\%NewSdc%\exe\f102010
if exist d:\%NewSdc%\FILES\NG_FILE.TXT goto _Error
type beeps.txt

rem  -------------------------------
rem echo.%DATE% %TIME:~0,8% 原産国削除 >> geting.txt
if /i "%Computername%" == "w4" (
	echo.■■原産国削除
	pvddl %NewSdc% delete-gensan.sql -stoponfail
)
rem  -------------------------------(★)
if /i "%Computername%" == "w1" (
	echo.■■y_nyukaフラグセット
	rem echo.%DATE% %TIME:~0,8% y_nyukaフラグセット >> geting.txt
	rem if exist y-nyuka-set-9.sql pvddl %NewSdc% y-nyuka-set-9.sql -stoponfail
	rem  -------------------------------
	echo.■■複数原産国部品入庫管理リスト
	d:\%NewSdc%\exe\F102090
	rem  -------------------------------
	echo.■■入庫／棚番チェックリスト
rem	d:\%NewSdc%\exe\F103000
)
rem  -------------------------------
if 0 == 1 (
	rem 処理が遅いので別で実行
	echo.■■item 検品メッセージ更新：リチウム電池搭載
	echo.%DATE% %TIME:~0,8% ■■item 検品メッセージ更新 >> geting.txt
	if exist item-insp-message.log del item-insp-message.log
	cscript     item-insp-message.vbs
	for %%i in (item-insp-message.log) do if %%~zi neq 0 (
		echo.■■■■■■■■■■■■■■■■■■■■■■■■ > mail.txt
		echo.■品目MST 検品メッセージ更新：リチウム電池搭載■ >> mail.txt
		echo.■　　至急：商品化指図書に登録してください。　■ >> mail.txt
		echo.■■■■■■■■■■■■■■■■■■■■■■■■ >> mail.txt
		type item-insp-message.log >>mail.txt
		tool\blatj mail.txt -attach item-insp-message.log -s "品目M 検品メッセージ更新：リチウム電池搭載" -t %ML%
		call d:\newsdc\tool\slack "品目M 検品メッセージ更新：リチウム電池搭載" %cd%\mail.txt
	)

	rem  -------------------------------
	echo.■■item 検品メッセージ更新：共用部品
	if exist item_insp_message.log del item_insp_message.log
	cscript     item_insp_message.vbs /update /db:%NewSdc%
	for %%i in (item_insp_message.log) do if %%~zi neq 0 (
		echo.■■■■■■■■■■■■■■■■■■■■■■■ > mail.txt
		echo.■品目MST 検品メッセージ更新：共用部品です。■ >> mail.txt
		echo.■■■■■■■■■■■■■■■■■■■■■■■ >> mail.txt
		type item_insp_message.log >>mail.txt
		tool\blatj mail.txt -attach item_insp_message.log -s "品目M 検品メッセージ更新：共用部品です。" -t %ML%
		call d:\newsdc\tool\slack "品目M 検品メッセージ更新：共用部品です。" %cd%\mail.txt
	)
)

if /i "%Computername%" == "w1" (
	rem  -------------------------------
	echo.■■原産国マスター 更新日に入荷データの登録日をセット
	rem echo.%DATE% %TIME:~0,8% 原産国マスター更新日セット >> geting.txt
	cscript y_nyuka_gensan.vbs /update /db:%NewSdc%
)
rem  -------------------------------
if exist d:\%NewSdc%\files\hmem500_denno.sql (
	echo.■■hmem500 伝票No 10桁
	pvddl %NewSdc% d:\%NewSdc%\files\hmem500_denno.sql > d:\%NewSdc%\files\hmem500_denno.log
	del d:\%NewSdc%\files\hmem500_denno.sql
)
echo.■■hmem500に登録
cscript//nologo d:\%NewSdc%\files\glicspos.vbs /db:%NewSdc% in_save\%fName%
if /i "%Bu%" == "A" (
	echo.SJ入荷予定登録
	cscript//nologo d:\%NewSdc%\files\hmem500.vbs /db:%NewSdc% %fName% /y_nyuka /z:SJ010101
	echo.SJ出荷予定登録＆伝票No10桁更新
	python d:\%NewSdc%\files\hmem500.py --dns %NewSdc% %fName% > hmem500.log
	call d:\newsdc\tool\slack "hmem500.py %NewSdc% %Bu%" %cd%\hmem500.log
)
cscript//nologo d:\%NewSdc%\files\hmem500.vbs /db:%NewSdc% %fName% >> geting.txt
echo.■■Pn連携：新品番のみ
call d:\newsdc\app\pn.bat %Bu% %NewSdc%
goto _End_Log

rem  -------------------------------
:_Error
	type beeps.txt
	type beeps.txt
	type beeps.txt
	type beeps.txt
	type beeps.txt
	copy d:\%NewSdc%\FILES\NG_FILE.TXT in_save\%1.ng
	echo.■■Glics振替 %1 でエラーが発生しました。
	echo.■■再試行します。
	echo.%0 %* >> geting.txt
	type d:\%NewSdc%\FILES\NG_FILE.TXT >> geting.txt
	tool\blatj %cd%\geting.txt -s "■Glics振替 Error: %0 %*" -t system@kk-sdc.co.jp
	call d:\newsdc\tool\slack "■Glics振替 Error %0 %*" %cd%\geting.txt
	del in_save\%fName%
	goto _End

rem  -------------------------------
:_End_Log
rem debug要
rem dir in_save\%fName% d:\%NewSdc%\hostfile\shiji_*_?.dat>mail.txt
rem echo. >>geting.txt
echo.%DATE%  %TIME:~0,8% ▽ >> geting.txt
rem tool\blatj mail.txt -s "%0  %*" -t log@kk-sdc.co.jp
call d:\newsdc\tool\slack "■Glics振替 %NewSdc% %Bu%" %cd%\geting.txt
rem  -------------------------------(★)
if /i "%Bu%" == "6" (
	d:\newsdc\tool\blatj geting.txt -s "■Glics振替 %NewSdc% %Bu%" -t sdc.nara.e5@gmail.com -c system@kk-sdc.co.jp
)
echo.■■液晶ディスプレイ通知
xcopy/d/y geting.txt d:\%NewSdc%\files\notice\
xcopy/d/y geting.txt \\hs1\it\pos\newsdc\files\notice\
rem -------
call d:\log\batlog ▽ %0 %*
set ret=1
:_End
color
for %%i in (in_save\%1) do (
	echo.■geting  %1 %2 %%~zi
)
exit/b %ret%
