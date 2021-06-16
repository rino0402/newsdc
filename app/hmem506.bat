@echo off
rem hmem506.bat
rem Glics入荷 エアコン hmem506szz
rem 2021.02.27
setlocal ENABLEDELAYEDEXPANSION
pushd %~dp0
set DT=%1
if "%DT%" == "" set DT=%DATE:/=%
for /F "eol=; delims== tokens=1,2" %%x in (config.ini) do (
	set %%x=%%y
)
echo.■%0 %*
set ret=0
for %%i in (g:\glics\hmem506szz.dat.%DT%-*.*) do (
	if exist glics\%%~nxi (
		echo.■%0 %* %%~nxi
	) else (
		echo.■%0 %* %%~nxi new
		echo.%DATE%  %TIME:~0,8% △ > hmem506.log
		dir %%i | findstr /i %%~nxi >> hmem506.log
		xcopy/d/y/z %%i glics\
		echo.■%0 xcopy=%errorlevel%
		py hmem500.py glics\%%~nxi >> hmem506.log
		py hmem500.py %%~nxi --y_nyuka >> hmem506.log
		py hmem500.py %%~nxi --zaiko >> hmem506.log
		py hmem500.py %%~nxi --y_syuka >> hmem506.log
		py hmem500.py %%~nxi --item >> hmem506.log
		py hmem500.py %%~nxi --list >> hmem506.log 2> hmem500.err & if errorlevel 1 call :_List %%~nxi
		echo.%DATE%  %TIME:~0,8% ▽ >> hmem506.log
		call ..\tool\slack "`■hmem506` %%~nxi" %cd%\hmem506.log
		xcopy/d/y hmem506.log notice0\
		type hmem506.log
		copy /y hmem506.log glics\%%~nxi.log
		call pn.bat A %DT%
		set /a ret+=1
	)
)
popd
rem endlocal
exit/b !ret!
rem :_List
:_List
echo.%1 >> hmem500.err
py slack.py %computername% %computername:w=w% "`■入荷リスト：エアコン` %1" hmem500.err
blatj hmem500.err -s "■入荷リスト：エアコン" -t %MAIL_TO% -c system@kk-sdc.co.jp -server ns -f %computername%
copy/y hmem500.err ..\files\notice0\hmem506.txt
copy/y hmem500.err ..\files\notice\hmem506.txt
exit/b

xcopy/d/y hmem500.py \\w3\newsdc\app\
xcopy/d/y hmem500.bat \\w3\newsdc\app\
xcopy/d/y hmtah500.py \\w3\newsdc\app\
xcopy/d/y hmem506.bat \\w3\newsdc\app\
