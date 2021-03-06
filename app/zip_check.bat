@echo off
rem zip_check.bat
setlocal
pushd %~dp0
echo.‘is2 XΦΤ`FbN
if "%1" == "" (
	py zip_check.py 2> zip_check.err
) else (
	py zip_check.py --dt %1 --table del_syuka_h 2> zip_check.err
)
if errorlevel 1 (
	echo.%~f0 %* >> zip_check.err
	call ..\tool\slack "‘XΦΤG[" %cd%\zip_check.err
	..\tool\blatj zip_check.err -s "‘XΦΤG[" -t osakapc@kk-sdc.co.jp -c system@kk-sdc.co.jp -server ns -f is2
rem	..\tool\blatj zip_check.err -s "‘XΦΤG[" -t kubo@kk-sdc.co.jp -server ns -f is2
) else (
	call ..\tool\slack "‘is2 zip_check.err" %cd%\zip_check.err
)
type zip_check.err
popd
endlocal
exit/b

dir \\w5\newsdc\istar2\zip_check.* zip_check.*
xcopy/u/d zip_check.* \\w5\newsdc\istar2
py -m pip install jusho

git commit -am "zip_check 0.01"
git push origin master
