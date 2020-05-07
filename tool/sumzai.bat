@echo off
rem sumzai.bat
rem ピボットテーブル更新
rem sumzai_5.xlsx
setlocal
pushd %~dp0
echo.%date% %time:~0,8% %0 %* > sumzai.log
rem cscript ctrxls.vbs sumzai_5.xlsx /setts:O1 >> sumzai.log
rem dir sumzai_5.xlsx | findstr \/ | findstr /v DIR >> sumzai.log

rem cscript ctrxls.vbs DailyZaiko_w1.xlsx /setts >> sumzai.log
rem dir DailyZaiko_w1.xlsx | findstr \/ | findstr /v DIR >> sumzai.log

rem cscript ctrxls.vbs DailyZaiko.xlsx /setts >> sumzai.log
rem dir DailyZaiko.xlsx | findstr \/ | findstr /v DIR >> sumzai.log

echo.%date% %time:~0,8% %0 %* >> sumzai.log
call slack "%0 %*" %cd%\sumzai.log
popd
endlocal
