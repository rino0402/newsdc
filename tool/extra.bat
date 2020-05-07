@echo off
rem extra \\nas01\backup\w1\glics \\w1\glics
call :_extra \\nas02\backup\w1\doc		\\w1\doc		"/D -365"
call :_extra \\nas02\backup\w1\manager	\\w1\manager	"/D -365"
call :_extra \\nas02\backup\w1\y		\\w1\y			"/D -365"
call :_extra \\nas02\backup\w1\newsdc	\\w1\newsdc
call :_extra \\nas02\backup\w1\newsdc0	\\w1\newsdc0
call :_extra \\nas02\backup\w1\glics	\\w1\glics
exit/b
:_extra
	echo.%0 %*
	setlocal
	set Log=%2\extra.log
	echo.%date% %time:~0,8% %0 %* >> %Log%
	pushd %1
	forfiles %~3 /s /c "cmd /c if not exist %2\@relpath (echo @path @fdate @ftime @fsize && if @isdir==TRUE (rmdir @path) else (del @path) && echo.@path @fdate @ftime @fsize>> %Log%)"
	popd
	endlocal
	exit/b
