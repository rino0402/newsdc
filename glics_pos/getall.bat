@echo off
setlocal ENABLEDELAYEDEXPANSION
set ret=0
echo.■getall %* %Computername%
rem 2016.11.17 gift受信ファイル名形式変更
rem 2017.03.08 ActiveGift対応準備
rem 00023100：hmem501szz.dat 1 洗濯機
rem 00023210：hmem502szz.dat 7 掃除機
rem 00023510：hmem504szz.dat D IH
rem 00023410：hmem505szz.dat 4 炊飯
rem 00025800：hmem506szz.dat A エアコン
rem 00021259：hmem507szz.dat R 冷蔵庫
rem 00021397：hmem508szz.dat 5 BL調理
rem 00021184：hmem509szz.dat 6 レンジ
rem 00021529：hmem50Aszz.dat 2 食洗
rem 00023100：hmem701szz.dat 1 洗濯機
rem 00023210：hmem702szz.dat 7 掃除機
rem 00023510：hmem704szz.dat D IH
rem 00023410：hmem705szz.dat 4 炊飯
rem 00025800：hmem706szz.dat A エアコン
rem 00021259：hmem707szz.dat R 冷蔵庫
rem 00021397：hmem708szz.dat 5 BL調理
rem 00021184：hmem709szz.dat 6 レンジ
rem 00021529：hmem70Aszz.dat 2 食洗

set DT=%1
if "%DT%" == "" set DT=%DATE:/=%
set DT
set ML
goto %Computername%

rem --------------------------------------
rem 滋賀物流
rem --------------------------------------
:W4
rem tool\blatj -install ns w4
if not exist B:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.■再接続
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w4 /y
	net use B: \\hs1\glics\gift\recv 123daaa /USER:w4 /y
	net use G: \\hs1\glics 123daaa /USER:w4 /y
)
rem Glics入出庫 00025800 エアコン
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem506szz.dat.%DT%-*.*.) do (
	call geting %%~nxi A
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics入出庫 00021259 冷蔵庫
if not exist g:\gift\recv.ok goto _End
for %%i in (b:\hmem507szz.dat.%DT%-*.*.) do (
rem	call geting %%~nxi R
	call hmem500 newsdcr %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00025800 エアコン
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem706szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi A
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00021259 冷蔵庫
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
rem 奈良
rem --------------------------------------
rem w6 レンジ newsdc8
:W6
if not defined ML (
	set ML=sdc.nara.e5@gmail.com
)
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.■再接続
	net use * /del /y
	net use G: \\hs1\glics
	net use A: \\hs1\glics\gift\acgps
rem	net use G: \\hs1\glics 123daaa /USER:w6 /y
rem	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w6 /y
)
rem Glics入出庫 00021184 レンジ
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem509szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 6 newsdc8
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics入出庫 00021529 食洗
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem50Aszz.dat.%DT%-*.*.) do (
	call geting %%~nxi 2 newsdc8
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00021184 レンジ
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem709szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 6 newsdc8
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00021529 食洗
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem70Aszz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 2 newsdc8
)
call d:\newsdc9\files\pop3w9.bat
endlocal && set ret=%ret%
exit/b %ret%
rem --------------------------------------
rem 袋井
rem --------------------------------------
:W2
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.■再接続
	net use * /del /y
	net use G: \\hs1\glics 123daaa /USER:w2 /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w2 /y
)
rem Active実績結果
for %%i in (A:\hmtac770ahu.dat.%DT%-*) do (
	call getcomps %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics入荷 00023100 洗濯機
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem501szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 1
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00023100 洗濯機
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem701szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 1
	set /a ret=%ret% + %ERRORLEVEL%
)
endlocal && set ret=%ret%
exit/b %ret%
rem --------------------------------------
rem 滋賀p
rem --------------------------------------
:W3
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.■再接続
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w3 /y
	net use G: \\hs1\glics 123daaa /USER:w3 /y
)
rem Active入荷
for %%i in (A:\hmtah500sec.dat.%DT%-*) do (
	call getinn %%i 7
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active出荷
for %%i in (A:\hmtah011sec.dat.%DT%-*) do (
	call getoutn %%i 7
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active出荷予定
for %%i in (A:\hmtah015szz.dat.%DT%-*) do (
	call getouty %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active実績結果
for %%i in (A:\hmtac770aec.dat.%DT%-*) do (
	call getcomps %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics入荷 00023210 掃除機
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem502szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 7
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00023210 掃除機
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem702szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 7
	set /a ret=%ret% + %ERRORLEVEL%
)
endlocal && set ret=%ret%
exit/b %ret%
rem --------------------------------------
rem 小野
rem --------------------------------------
:W1
if not exist A:\nul (
	net use * /del /y
)
if not exist G:\nul (
	echo.■再接続
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps 123daaa /USER:w1 /y
	net use G: \\hs1\glics 123daaa /USER:w1 /y
)
rem Active入荷
for %%i in (A:\hmtah500scs.dat.%DT%-*) do (
	call getinn %%i ono
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active出荷
set bin=-1
for %%i in (A:\hmtah011scs.dat.%DT%-*) do (
	set/a bin = !bin! + 1
	call getoutn %%i ono !bin!
	set /a ret=%ret% + %ERRORLEVEL%

)
rem Active出荷予定
set biny=-1
for %%i in (A:\hmtah015szz.dat.%DT%-*) do (
	set/a biny = !biny! + 1
	call getouty %%i !biny!
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Active実績結果
for %%i in (A:\hmtac770acs.dat.%DT%-*) do (
	call getcomps %%i
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics入荷 00023510 IH
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem504szz.dat.%DT%-*.*.) do (
	call geting %%~nxi D
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics入荷 00023410 炊飯
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem505szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 4
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics入荷 00021397 BL調理
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem508szz.dat.%DT%-*.*.) do (
	call geting %%~nxi 5
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00023510 IH
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem704szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi D
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00023410 炊飯
if not exist g:\gift\recv.ok goto _End
for %%i in (g:\gift\recv\hmem705szz.dat.%DT%-*.*.) do (
	call getoutg %%~nxi 4
	set /a ret=%ret% + %ERRORLEVEL%
)
rem Glics出荷 00021397 BL調理
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
	echo.■再接続
	net use * /del /y
	net use A: \\hs1\glics\gift\acgps /y
	net use G: \\hs1\glics /y
)
set ret=%ERRORLEVEL%
endlocal && set ret=%ret%
exit/b %ret%
