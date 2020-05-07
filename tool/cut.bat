@echo off
set sz=0
set	ln=100
:Top
	tail -%ln%l %1 > cut.tmp
	for %%i in (cut.tmp) do (
		set sz=%%~zi
	)
	echo.%ln% %sz%
	set/a ln=%ln% - 1
if %sz% GTR 3000 goto :Top
