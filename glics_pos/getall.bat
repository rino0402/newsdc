@echo off
setlocal ENABLEDELAYEDEXPANSION
set ret=0
echo.��getall %* %Computername%
rem 2016.11.17 gift��M�t�@�C�����`���ύX
rem 2017.03.08 ActiveGift�Ή�����
rem 00023100�Fhmem501szz.dat 1 ����@
rem 00023210�Fhmem502szz.dat 7 �|���@
rem 00023510�Fhmem504szz.dat D IH
rem 00023410�Fhmem505szz.dat 4 ����
rem 00025800�Fhmem506szz.dat A �G�A�R��
rem 00021259�Fhmem507szz.dat R �①��
rem 00021397�Fhmem508szz.dat 5 BL����
rem 00021184�Fhmem509szz.dat 6 �����W
rem 00021529�Fhmem50Aszz.dat 2 �H��
rem 00023100�Fhmem701szz.dat 1 ����@
rem 00023210�Fhmem702szz.dat 7 �|���@
rem 00023510�Fhmem704szz.dat D IH
rem 00023410�Fhmem705szz.dat 4 ����
rem 00025800�Fhmem706szz.dat A �G�A�R��
rem 00021259�Fhmem707szz.dat R �①��
rem 00021397�Fhmem708szz.dat 5 BL����
rem 00021184�Fhmem709szz.dat 6 �����W
rem 00021529�Fhmem70Aszz.dat 2 �H��

set DT=%1
if "%DT%" == "" set DT=%DATE:/=%
set DT
set ML
goto %Computername%

rem --------------------------------------
rem ���ꕨ��
rem --------------------------------------
:W4
rem tool\blatj -install ns w4
if not exist B:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.���Đڑ�
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w4 /y
	net use B: \\hs1\glics\gift\recv 123daaa /USER:w4 /y
	net use G: \\hs1\glics 123daaa /USER:w4 /y
)
rem Glics���o�� 00025800 �G�A�R��
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem506szz.dat.%DT%-*.*.) do (
	call geting %%~nxi A
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics���o�� 00021259 �①��
if not exist g:\gift\recv.ok goto _End
for %%i in (b:\hmem507szz.dat.%DT%-*.*.) do (
rem	call geting %%~nxi R
	call hmem500 newsdcr %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00025800 �G�A�R��
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem706szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi A
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00021259 �①��
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem707szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi R newsdcr
	set /a ret=%ret% + %ERRORLEVEL%
)
for %%i in (b:\hmem707szz.dat.%DT%-*.*.) do (
	call hmem700 newsdcr %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
endlocal && set ret=%ret%
exit/b %ret%
rem --------------------------------------
rem �ޗ�
rem --------------------------------------
rem w6 �����W newsdc8
:W6
if not defined ML (
	set ML=sdc.nara.e5@gmail.com
)
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.���Đڑ�
	net use * /del /y
	net use G: \\hs1\glics
	net use A: \\hs1\glics\gift\acgps
rem	net use G: \\hs1\glics 123daaa /USER:w6 /y
rem	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w6 /y
)
rem Glics���o�� 00021184 �����W
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem509szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 6 newsdc8
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics���o�� 00021529 �H��
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem50Aszz.dat.%DT%-*.*.) do (
	call geting %%~nxi 2 newsdc8
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00021184 �����W
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem709szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 6 newsdc8
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00021529 �H��
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem70Aszz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 2 newsdc8
)
call d:\newsdc9\files\pop3w9.bat
endlocal && set ret=%ret%
exit/b %ret%
rem --------------------------------------
rem �܈�
rem --------------------------------------
:W2
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.���Đڑ�
	net use * /del /y
	net use G: \\hs1\glics 123daaa /USER:w2 /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w2 /y
)
rem Active���ь���
for %%i in (A:\hmtac770ahu.dat.%DT%-*) do (
	call getcomps %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics���� 00023100 ����@
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem501szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 1
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00023100 ����@
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem701szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 1
	set /a ret=%ret% + %ERRORLEVEL%
)
endlocal && set ret=%ret%
exit/b %ret%
rem --------------------------------------
rem ����p
rem --------------------------------------
:W3
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.���Đڑ�
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w3 /y
	net use G: \\hs1\glics 123daaa /USER:w3 /y
)
rem Active����
for %%i in (A:\hmtah500sec.dat.%DT%-*) do (
	call getinn %%i 7
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active�o��
for %%i in (A:\hmtah011sec.dat.%DT%-*) do (
	call getoutn %%i 7
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active�o�ח\��
for %%i in (A:\hmtah015szz.dat.%DT%-*) do (
	call getouty %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active���ь���
for %%i in (A:\hmtac770aec.dat.%DT%-*) do (
	call getcomps %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics���� 00023210 �|���@
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem502szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 7
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00023210 �|���@
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem702szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 7
	set /a ret=%ret% + %ERRORLEVEL%
)
endlocal && set ret=%ret%
exit/b %ret%
rem --------------------------------------
rem ����
rem --------------------------------------
:W1
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.���Đڑ�
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w1 /y
	net use G: \\hs1\glics 123daaa /USER:w1 /y
)
rem Active����
for %%i in (A:\hmtah500scs.dat.%DT%-*) do (
	call getinn %%i ono
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active�o��
set bin=-1
for %%i in (A:\hmtah011scs.dat.%DT%-*) do (
	set/a bin = !bin! + 1
	call getoutn %%i ono !bin!
	set /a ret=%ret% + %ERRORLEVEL%

)
rem Active�o�ח\��
set biny=-1
for %%i in (A:\hmtah015szz.dat.%DT%-*) do (
	set/a biny = !biny! + 1
	call getouty %%i !biny!
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active���ь���
for %%i in (A:\hmtac770acs.dat.%DT%-*) do (
	call getcomps %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics���� 00023510 IH
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem504szz.dat.%DT%-*.*.) do (
	call geting %%~nxi D
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics���� 00023410 ����
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem505szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 4
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics���� 00021397 BL����
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem508szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 5
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00023510 IH
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem704szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi D
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00023410 ����
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem705szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 4
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics�o�� 00021397 BL����
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem708szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 5
	set /a ret=%ret% + %ERRORLEVEL%
)
:_End
endlocal && set ret=%ret%
exit/b %ret%
:W11
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.���Đڑ�
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps /y
	net use G: \\hs1\glics /y
)
set ret=%ERRORLEVEL%
endlocal && set ret=%ret%
exit/b %ret%
