@echo off
rem setlocal enabledelayedexpansion
setlocal
pushd %~dp0
rem 共通 2008.11.29 Active対応(出荷予定データ照合機能追加)
rem 共通 2009.01.05 出荷予定照合OK のメール送信解除
rem 共通 2013.11.05 出荷予定データ変換コントロール対応
rem 共通 2013.11.06 出荷予定データ変換コントロール対応
rem 2013.11.07 ngファイルをngフォルダに保存するように変更
rem 2014.02.28 ログ出力対応
rem 2014.04.24 出荷実績連携の完了チェック
rem 2014.04.28 出荷実績連携の完了メール送信
rem 2015.12.24 blajの送信時刻が+10になるのを修正
rem 2016.07.19 進捗区分４以外削除：暫定対処
rem 2016.10.01 P産機対応
rem 2016.10.16 P産機送り先データ登録
rem 2016.10.26 P産機SSX(R-smile)送り状データ作成
rem 2017.03.08 ActiveGift対応
rem
rem getouty HMTAH015SZZ.dat.20161001-141032.OK HMTAH015SZZ.dat.20161001-141032
rem g_syuka.vbs
rem g_syuka_del.sql
rem d:\newsdc\files\glicspos.vbs
rem d:\newsdc\files\HMTAH015.sql
rem outy\HMTAH015SZZ.dat.20161001-141032.OK
rem outy\HMTAH015SZZ.dat.20161001-141032
rem getoutn.ok
rem check_g_syuka.vbs
rem check_g_syuka.txt
rem outy\HMTAH015SZZ.dat.20161001-141032.txt
rem tool\blatj
rem getcomps.done
rem complt.sql
rem d:\newsdc\files\y_syuka_check.vbs
rem y_syuka_check.txt
rem y_syuka_check.end
rem y_syuka_check.send
rem d:\newsdc\files\b2data.vbs
set ret=0
set absPath=%1
set relPath=%~nx1
if exist outy\%relPath% goto _End
echo.■getouty %*
rem call d:\log\batlog △ %0 %*
rem call slack "%relPath%:△"
color 9F
echo.■■Active出荷予定
echo.xcopy/d/y %absPath% ...%time%
xcopy/d/y %absPath%
echo.xcopy/d/y %absPath% ...%time%完了
if not "%ERRORLEVEL%"  == "0" (
	call d:\log\batlog ▼ %0 %* xcopy:%ERRORLEVEL%
	call d:\newsdc\tool\slack "■getouty %relPath%:▼xcopy:%ERRORLEVEL%"
	del %relPath%
	GOTO _End
)
set fSize=0
for %%i in ( %relPath% ) do set fSize=%%~zi
if %fSize% == 0 (
	call d:\log\batlog ▼ %0 %* fSize:%fSize%
	call d:\newsdc\tool\slack "■getouty %relPath%:▼fSize:%fSize%"
	del %relPath%
	GOTO _End
)

if exist getouty.done del getouty.done
echo.■■出荷予定データ変換 HMTAH015_t
cscript//Nologo d:\newsdc\files\glicspos.vbs %relPath%
echo.■■出荷予定データ変換 HMTAH015
cscript//Nologo d:\newsdc\files\HMTAH015.vbs
echo.■■直送データ登録 HMTAH015_c
python d:\newsdc\files\HMTAH015.py>HMTAH015_c.log
call d:\newsdc\tool\slack "HMTAH015_c" %cd%\HMTAH015_c.log
type beeps.txt
move/y %relPath% outy\

echo.■■出荷予定データ照合
cscript//Nologo check_g_syuka.vbs
copy /y check_g_syuka.txt outy\%relPath%.txt
for %%i in ( check_g_syuka.txt ) do set FSize=%%~zi
if not %FSize% == 0 (
rem	echo.>>check_g_syuka.txt
	echo.%0 %* >>check_g_syuka.txt
	tool\blatj check_g_syuka.txt -attach outy\%relPath%.txt -s "■Active出荷予定照合エラー" -t %ML% -c system@kk-sdc.co.jp
	call d:\newsdc\tool\slack "■Active出荷予定照合エラー" %cd%\check_g_syuka.txt
)
echo.■■出荷完了チェック
cscript//Nologo d:\newsdc\files\y_syuka_check.vbs > y_syuka_check.txt
if %ERRORLEVEL% == 0 (
	echo.出荷完了:%ERRORLEVEL% >> y_syuka_check.txt
	echo.%0 %* >> y_syuka_check.txt
	rem 本日の出荷実績連携残＝０
	if not exist y_syuka_check.end (
		tool\blatj y_syuka_check.txt -s "出荷実績連携(完了)" -t %ML% -c system@kk-sdc.co.jp
		call d:\newsdc\tool\slack "■Active出荷実績連携(完了)" %cd%\y_syuka_check.txt
		copy /y y_syuka_check.txt y_syuka_check.end
	) else (
rem		tool\blatj y_syuka_check.txt -s "(済)%relPath%" -t log@kk-sdc.co.jp
		call d:\newsdc\tool\slack "□Active出荷実績連携(完了済)" %cd%\y_syuka_check.txt
	)
) else (
	echo.実績残:%ERRORLEVEL% >> y_syuka_check.txt
	echo.%0 %* >> y_syuka_check.txt
	if exist y_syuka_check.send (
		tool\blatj y_syuka_check.txt -s "出荷実績連携(未完了)" -t %ML% -c system@kk-sdc.co.jp
		call d:\newsdc\tool\slack "■Active出荷実績連携(未完了)" %cd%\y_syuka_check.txt
		del y_syuka_check.send
	) else (
		rem 状況チェック用メール送信(安定稼働すれば解除)
rem		tool\blatj y_syuka_check.txt -s "(未)%relPath%" -t log@kk-sdc.co.jp
		call d:\newsdc\tool\slack "□Active出荷実績連携" %cd%\y_syuka_check.txt
	)
	del y_syuka_check.end
)
echo %0 %* > getouty.done
if exist d:\newsdc\B2\INPUT (
	call d:\newsdc\B2\makecsv.bat
)

xcopy/d/y \\w4\newsdc\files\ACSHORT.DAT d:\newsdc\files\ > acshort.log 2>&1
call d:\newsdc\tool\slack "acshort.log" %cd%\acshort.log

rem call d:\log\batlog ▽ %0 %*
rem call slack "%relPath%:▽"
set ret=1
:_End
color
for %%i in (outy\%relPath%) do echo.■getouty %* %%~zi %ML%
popd
endlocal && set ret=%ret%
exit/b %ret%
