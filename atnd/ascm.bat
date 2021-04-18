@echo off
setlocal
pushd %~dp0
title %DATE% %TIME:~0,8% %0 %*

echo.%DATE% %TIME:~0,8% ¢ > ascm.log

if not exist \\hirame\ascm net use \\hirame\ascm /user:sdch\hs2 123daaa! && echo.\\hirame\ascm .ok
py ascm_dat.py \\hirame\ascm\ascm.dat --dns newsdc6 --month >> ascm.log
py atnd_ascm.py --month --dns newsdc6 >> ascm.log

echo.%DATE% %TIME:~0,8% ¤ >> ascm.log
call ..\tool\slack "%0 %*" %cd%\ascm.log
popd
endlocal
exit/b
