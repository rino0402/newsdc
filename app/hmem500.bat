@echo off
rem app\hmem500.bat
rem 2021.02.27
rem 2021.04.30
setlocal
pushd %~dp0
echo.Å°%0 %*
set BU=%1
if not defined BU goto :_END
set DT=%2
if "%DT%" == "" set DT=%DATE:/=%
for /F "eol=; delims== tokens=1,2" %%x in (config.ini) do (
	set %%x=%%y
)
set FILE=
for %%i in (\\hs2\gift\glics\hmem50%BU%szz.dat.%DT%-*.*) do (
	if exist glics\%%~nxi (
		echo.Å°%0 %* %%~nxi
	) else (
		echo.Å°%0 %* %%~nxi new
		echo.%DATE%  %TIME:~0,8% Å¢ > hmem500.log
		dir %%i | findstr /i %%~nxi >> hmem500.log
		xcopy/d/y/z %%i glics\
		py hmem500.py --dsn %DSN% glics\%%~nxi >> hmem500.log
		py hmem500.py --dsn %DSN% %%~nxi --y_nyuka >> hmem500.log
		py hmem500.py --dsn %DSN% %%~nxi --zaiko >> hmem500.log
		py hmem500.py --dsn %DSN% %%~nxi --y_syuka >> hmem500.log
		py hmem500.py --dsn %DSN% %%~nxi --item >> hmem500.log
		echo.%DATE%  %TIME:~0,8% Å§ >> hmem500.log
		py slack.py %computername% %computername:w=w% "`Å°hmem500` %%~nxi" hmem500.log
		xcopy/d/y hmem500.log notice0\
		copy /y hmem500.log glics\%%~nxi.log
rem		call pn.bat A
	)
)
popd
if defined FILE exit/b 1
exit/b 0

xcopy/d/y hmem500.py \\w3\newsdc\app\
xcopy/d/y hmem500.bat \\w3\newsdc\app\
xcopy/d/y hmtah500.py \\w3\newsdc\app\

xcopy/d/y hmem500.bat \\w3\newsdcn\app\
xcopy/d/y hmem500.py \\w3\newsdcn\app\
