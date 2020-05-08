@echo off
echo %CD%
for /F "tokens=3 delims=\" %%a in ("%CD%") do set cdir=%%a
xcopy/d/y %1 \\hs1\newsdc\%cdir%\
xcopy/d/y %1 \\w1\newsdc\%cdir%\
xcopy/d/y %1 \\w2\newsdc\%cdir%\
xcopy/d/y %1 \\w3\newsdc\%cdir%\
xcopy/d/y %1 \\w4\newsdc\%cdir%\
xcopy/d/y %1 \\w5\newsdc\%cdir%\
xcopy/d/y %1 \\w6\newsdc\%cdir%\
xcopy/d/y %1 \\w7\newsdc\%cdir%\
exit/b
