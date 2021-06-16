@echo off
rem Glics入荷 エアコン/燃料電池 hmem506szz
rem app\hmem500.bat %1 %2
rem	%1 ファイル名	hmem506szz
rem					hmem50Bszz
rem	%2 日付			20210525
rem 2021.02.27
setlocal ENABLEDELAYEDEXPANSION
pushd %~dp0
set ret=0
echo.■%0 %*
set FILE=%1
if not defined FILE goto :_END
set DT=%2
if not defined DT set DT=%DATE:/=%
set	LOG=hmem500.log
if "%FILE%" == "hmem506szz" set JNAME=エアコン & set JGYOBU=A
if "%FILE%" == "hmem50Bszz" set JNAME=燃料電池 & set JGYOBU=N
if not defined JNAME goto :_END
rem config.ini
rem DSN
for /F "eol=; delims== tokens=1,2" %%x in (config.ini) do set %%x=%%y
rem
for %%i in (g:\glics\%FILE%.dat.%DT%-*.*) do (
	if exist glics\%%~nxi (
		echo.■%0 %%~nxi
	) else (
		echo.■%0 %%~nxi ★
		echo.%DATE%  %TIME:~0,8% △ > %LOG%
		dir %%i | findstr /i %%~nxi >> %LOG%
		xcopy/d/y/z %%i glics\

		py hmem500.py --dsn %DSN% glics\%%~nxi >> %LOG%
		py hmem500.py --dsn %DSN% %%~nxi --y_nyuka >> %LOG%
		py hmem500.py --dsn %DSN% %%~nxi --zaiko >> %LOG%
		py hmem500.py --dsn %DSN% %%~nxi --y_syuka >> %LOG%
		py hmem500.py --dsn %DSN% %%~nxi --item >> %LOG%
		py hmem500.py --dsn %DSN% %%~nxi --list >> %LOG% 2> hmem500.err & if errorlevel 1 call :_List %%~nxi
		echo.%DATE%  %TIME:~0,8% ▽ >> %LOG%
		rem call ..\tool\slack "`■hmem500` %%~nxi" %cd%\%LOG%
		py slack.py %computername% %computername:w=w% "`■hmem500` %%~nxi" %LOG% > nul
		xcopy/d/y %LOG% notice0\
		type %LOG%
		copy /y %LOG% glics\%%~nxi.log
		call pn.bat %JGYOBU% %DT%
		set /a ret+=1
	)
)
:_END
popd
rem endlocal
exit/b !ret!
rem :_List
:_List
echo.-- >> hmem500.err
echo.%1 >> hmem500.err
py slack.py %computername% %computername:w=w% "`■入荷リスト：%JNAME%` %1" hmem500.err
nkf -xw8 -O hmem500.err hmem500.utf8
blatj hmem500.utf8 -utf8 -s "■入荷リスト：%JNAME%" -t %MAIL_TO% -c system@kk-sdc.co.jp -server ns -f %computername%
copy/y hmem500.err ..\files\notice0\hmem506.txt
copy/y hmem500.err ..\files\notice\hmem506.txt
exit/b

xcopy/d/y hmem500.py \\w3\newsdc\app\
xcopy/d/y hmem500.py \\w3\newsdcn\app\
xcopy/d/y hmem500.bat \\w3\newsdc\app\
xcopy/d/y hmem500.bat \\w3\newsdcn\app\

xcopy/d/y hmtah500.py \\w3\newsdc\app\
xcopy/d/y hmem506.bat \\w3\newsdc\app\
