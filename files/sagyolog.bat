@echo off
rem 2014.02.27 バッチ実行ログ(batlog)
rem 2016.11.09 全センター対応
setlocal
pushd %~dp0
call d:\log\batlog △ %0 %*
set Computername
if /i "%Computername%" == "w1" (
	%windir%\syswow64\cscript sagyolog.vbs -stclear > sagyolog.log
)
if /i "%Computername%" == "w2" (
	cscript sagyolog.vbs -stnoclear > sagyolog.log
)
if /i "%Computername%" == "w3" (
	cscript sagyolog.vbs -stnoclear > sagyolog.log
	cscript sagyolog.vbs -stnoclear /db:newsdcn >> sagyolog.log
)
if /i "%Computername%" == "w4" (
	cscript sagyolog.vbs -stnoclear > sagyolog.log
	cscript sagyolog.vbs -stnoclear /db:newsdcr >> sagyolog.log
)
if /i "%Computername%" == "w5" (
	cscript sagyolog.vbs -stnoclear > sagyolog.log
	cscript sagyolog.vbs -stnoclear /db:fhd >> sagyolog.log
)
if /i "%Computername%" == "w6" (
	cscript sagyolog.vbs -stnoclear > sagyolog.log
	cscript sagyolog.vbs -stnoclear /db:newsdc8 >> sagyolog.log
	cscript sagyolog.vbs -stnoclear /db:newsdc9 >> sagyolog.log
)
if /i "%Computername%" == "w7" (
	cscript sagyolog.vbs -stnoclear > sagyolog.log
)
call d:\log\batlog ▽ %0 %*
cscript p_sagyo_log.vbs /dt:%DATE:/=% /list:1 > p_sagyo_log.log
if /i "%Computername%" == "w3" (
	echo.newsdcn >> p_sagyo_log.log
	cscript p_sagyo_log.vbs /dt:%DATE:/=% /list:1 /db:newsdcn >> p_sagyo_log.log
)
if /i "%Computername%" == "w4" (
	echo.newsdcr >> p_sagyo_log.log
	cscript p_sagyo_log.vbs /dt:%DATE:/=% /list:1 /db:newsdcr >> p_sagyo_log.log
)
if /i "%Computername%" == "w5" (
	echo.fhd >> p_sagyo_log.log
	cscript p_sagyo_log.vbs /dt:%DATE:/=% /list:1 /db:fhd >> p_sagyo_log.log
)
if /i "%Computername%" == "w6" (
	echo.newsdc8 >> p_sagyo_log.log
	cscript p_sagyo_log.vbs /dt:%DATE:/=% /list:1 /db:newsdc8 >> p_sagyo_log.log
	echo.newsdc9 >> p_sagyo_log.log
	cscript p_sagyo_log.vbs /dt:%DATE:/=% /list:1 /db:newsdc9 >> p_sagyo_log.log
)
call d:\newsdc\tool\slack "%0 %*" %cd%\p_sagyo_log.log
rem net view > netview.txt
rem call d:\newsdc\tool\nsession
rem type d:\newsdc\tool\nsession.txt >> netview.txt
rem net session >> netview.txt
rem call d:\newsdc\tool\ulist
rem type ulist.txt >> netview.txt
rem call d:\newsdc\tool\slack "net view" %cd%\netview.txt
popd
endlocal
timeout /T 10
