@echo off
for /F "tokens=3 delims=\" %%a in ("%CD%") do set cdir=%%a
echo.%cdir%
call :ftm %1
call :ftm \\hs1\newsdc\%cdir%\%1
call :ftm \\w1\newsdc\%cdir%\%1
call :ftm \\w2\newsdc\%cdir%\%1
call :ftm \\w3\newsdc\%cdir%\%1
call :ftm \\w4\newsdc\%cdir%\%1
call :ftm \\w5\newsdc\%cdir%\%1
call :ftm \\w5\fhd\%cdir%\%1
call :ftm \\w6\newsdc\%cdir%\%1
call :ftm \\w6\newsdc8\%cdir%\%1
call :ftm \\w6\newsdc9\%cdir%\%1
call :ftm \\w7\newsdc\%cdir%\%1
exit/b
:ftm
for %%i in ("%1") do (
	if exist "%%i" (
		echo.%%~ti %%~zi %%~fi
	) else (
rem          2015/11/25 09:48 494 \\w4\glics\glics_pos\y_syuka_rf.sql
		echo.                     %%~fi
	)
)
exit/b
