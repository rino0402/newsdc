@echo off
rem slackへメッセージ送信
setlocal
pushd %~dp0
set	User=%ComputerName%
set	Chnl=w0
if /i "%ComputerName%" == "w1" (
	set	Chnl=w1
)
if /i "%ComputerName%" == "doel" (
	set	Chnl=w1
)
if /i "%ComputerName%" == "w2" (
	set	Chnl=w2
)
if /i "%ComputerName%" == "w3" (
	set	Chnl=w3
)
if /i "%ComputerName%" == "w4" (
	set	Chnl=w4
)
if /i "%ComputerName%" == "w5" (
	set	Chnl=w5
)
if /i "%ComputerName%" == "w6" (
	set	Chnl=w6
)
if /i "%ComputerName%" == "w7" (
	set	Chnl=w7
)
if /i "%ComputerName%" == "hs1" (
	set	Chnl=w0
)
if /i "%ComputerName%" == "hs2" (
	set	Chnl=w0
)
set Text=%1
python -V > nul
if %errorlevel% == 0 (
	if "%2" == "" (
		python slack.py %User% %Chnl% %Text% nul
	) else (
		python slack.py %User% %Chnl% %Text% %2
	)
) else (
	call :_Cut %2
	echo.cscript >> slack.tmp
	cscript slack.vbs %User% %Chnl% %Text% slack.tmp
)
popd
endlocal
exit/b
:_Cut
	setlocal
	set sz=0
	set	ln=100
	if "%1" == "" (
		copy nul slack.tmp
	) else (
		tail -%ln%l %1 > slack.tmp
	)
	for %%i in (slack.tmp) do (
		set sz=%%~zi
	)
	echo.%ln% %sz%
	set/a ln=%ln% - 1
	if %sz% GTR 3000 goto :_Cut
	endlocal
	exit/b
