@echo off
pushd %~dp0
echo.¡net view > svstat.log
net view | findstr \\ >> svstat.log
net view /domain:sdch | findstr \\ >> svstat.log
rem for /f "tokens=1,2*" %%i in ('net view /domain:sdch') do (
rem 	echo %%i	%%j	%%k | findstr \\ >> svstat.log
rem )
echo.¡net session >> svstat.log
for /f "tokens=1,2*" %%i in ('net session') do (
	echo %%i	%%j	%%k | findstr \\ >> svstat.log
)
echo.¡openfiles >> svstat.log
openfiles | findstr /r /c:"^[=0-9]" | sort /+10 >> svstat.log
echo.¡loadpercentage
wmic cpu get loadpercentage >> svstat.log
echo.%date% %time:~0,8% %0 %* >> svstat.log
slack "%0 %*" %cd%\svstat.log
popd
