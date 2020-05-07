@echo off
pushd %~dp0
rem if not exist \\w2\d$ net use \\w2\d$ /user:sdc2\hs1 123daaa!
rem if not exist \\w3\d$ net use \\w3\d$ /user:sdc3\hs1 123daaa!
rem if not exist \\w5\c$ net use \\w5\c$ /user:hs1 123daaa!
rem if not exist \\w6\d$ net use \\w6\d$ /user:administrator 123daaa!
echo.%DATE%  %TIME:~0,8% △ > LogP.log
del u_ex.log
call :LogMake hs1
call :LogMake w1
call :LogMake w2
call :LogMake w3
call :LogMake w4
call :LogMake w5
call :LogMake w6
call :LogMake w7
echo.%DATE%  %TIME:~0,8% ▽ >> LogP.log
rem echo.http://hs1.kk-sdc.co.jp/it/pos/newsdc/tool/List.html >> LogP.log
echo.http://hs2/logparser/List.html >> LogP.log
call ..\backup\slack "%0 %*" %cd%\LogP.log
net view /domain:sdch > netview.txt
net session >> netview.txt
rem call ulist
rem type ulist.txt >> netview.txt
call ..\backup\slack "net view" %cd%\netview.txt
rem ping vpn2 > ping.txt
rem call slack "ping vpn2" %cd%\ping.txt
popd
exit/b
rem ----------------------------------------
rem IISアクセスログ収集
rem ----------------------------------------
:LogMake
setlocal
set w=%1
for /f "delims=" %%a in ('dir \\%w%\d$\log\W3SVC\*.log /O-D /b /s /a-d') do (
	set LogPath=%%a
	set LogFile=%%~nxa
	goto next
)
for /f "delims=" %%a in ('dir \\%w%\d$\log\W3SVC1\*.log /O-D /b /s /a-d') do (
	set LogPath=%%a
	set LogFile=%%~nxa
	goto next
)
for /f "delims=" %%a in ('dir \\%w%\c$\inetpub\logs\LogFiles\W3SVC\u_ex*.log /O-D /b /s /a-d') do (
	set LogPath=%%a
	set LogFile=%%~nxa
	goto next
)
for /f "delims=" %%a in ('dir \\%w%\c$\inetpub\logs\LogFiles\W3SVC1\u_ex*.log /O-D /b /s /a-d') do (
	set LogPath=%%a
	set LogFile=%%~nxa
	goto next
)
:next

echo.%LogPath% >> LogP.log
dir %LogPath% | findstr/i %LogFile% >> LogP.log
xcopy/y %LogPath% Log\
type %LogPath%>>u_ex.log

set LogParser="c:\Program Files\Log Parser 2.2\LogParser.exe"
set LogParser="C:\Program Files (x86)\Log Parser 2.2\LogParser.exe"

%LogParser% -i:IISW3C -o:TPL -tpl:PageView.tpl ^
"SELECT TOP 10 cs-uri-stem as uri,c-ip as ip,COUNT(*) as view INTO View%w%.html FROM Log\%LogFile% GROUP BY uri,ip ORDER BY view DESC"

%LogParser% -i:IISW3C -o:TPL -tpl:PageList.tpl ^
"SELECT date , time as utc_time, to_localtime(time) as time ,s-ip, c-ip ,cs-method ,cs-uri-stem ,cs-uri-query^
 ,s-port ^
 ,cs-username ^
 ,c-ip ^
 ,cs(User-Agent) ^
 ,cs(Referer) ^
 ,sc-status ^
 ,sc-substatus ^
 ,sc-win32-status ^
 ,time-taken ^
 INTO List%w%.html^
 FROM Log\%LogFile%^
 order by date , time desc"

%LogParser%^
 -i:IISW3C -o:TPL -tpl:PageList.tpl ^
"SELECT^
 date^
 ,time as utc_time^
 ,to_localtime(time) as time^
 ,s-ip^
 ,s-sitename^
 ,s-computername^
 ,c-ip^
 ,cs-host^
 ,cs-method^
 ,cs-uri-stem^
 ,cs-uri-query^
 ,s-port ^
 ,cs-username ^
 ,c-ip ^
 ,cs(User-Agent) ^
 ,cs(Referer) ^
 ,sc-status ^
 ,sc-substatus ^
 ,sc-win32-status ^
 ,time-taken ^
 INTO List.html^
 FROM u_ex.log^
 WHERE STRLEN(rtrim(cs-uri-query)) ^> 0^
 OR cs-uri-stem like '%.xlsx'^
 order by to_localtime(time) desc"

%LogParser%^
 -i:IISW3C -o:TPL -tpl:PageView.tpl ^
"SELECT cs-uri-stem as uri,s-ip,c-ip as ip,cs-username,COUNT(*) as view^
 INTO View.html^
 FROM u_ex.log^
 WHERE STRLEN(rtrim(cs-uri-query)) ^> 0^
 OR cs-uri-stem like '%.xlsx'^
 GROUP BY uri,s-ip,ip,cs-username^
 ORDER BY view DESC"

endlocal
exit/b
