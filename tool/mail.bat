@echo off
rem mail.bat メール送信
rem mail.bat [Sub] [File]
setlocal
pushd %~dp0
set Sub=%1
set File=%2
set	Size=0
for %%i in (%File%) do set Size=%%~zi
echo. Sub=%Sub%
echo.File=%File%
echo.Size=%Size%
echo.  ML=%ML%
if "%ML%" == "log@kk-sdc.co.jp" goto _Skip
if %Size% lss 0 goto _Skip

rem blatj zaiko.log -s "在庫データ連携:zaiko %*" -t %ML% -c system@kk-sdc.co.jp
echo.blatj %File% -s %Sub% -t %ML% -c system@kk-sdc.co.jp
if "%ML%" == "" (
	blatj %File% -s %Sub% -t system@kk-sdc.co.jp
) else (
	blatj %File% -s %Sub% -t %ML% -c system@kk-sdc.co.jp
)

:_Skip
popd
endlocal
