@echo off
pushd %~dp0
echo.��net view > svstat.log
for /f "skip=3 eol=�R" %%a in ('net view /domain') do (
	echo.%%a >> svstat.log
	net view /domain:%%a | findstr \\ >> svstat.log
)
rem net view | findstr \\ >> svstat.log
rem net view /domain:sdch | findstr \\ >> svstat.log
rem for /f "tokens=1,2*" %%i in ('net view /domain:sdch') do (
rem 	echo %%i	%%j	%%k | findstr \\ >> svstat.log
rem )
echo.��net session >> svstat.log
for /f "tokens=1,2*" %%i in ('net session') do (
	echo.%%i	%%j	%%k | findstr \\ >> svstat.log
)
echo.��openfiles >> svstat.log
openfiles | findstr /r /c:"^[=0-9]" | sort /+10 >> svstat.log
echo.��loadpercentage
wmic cpu get loadpercentage | findstr "." >> svstat.log
echo.��nbtstat -n >> svstat.log
nbtstat -n | findstr "." >> svstat.log
whoami >> svstat.log
echo.%date% %time:~0,8% %0 %* >> svstat.log
slack "%0 %*" %cd%\svstat.log
popd
