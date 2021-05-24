@echo off
rem zip_check.bat
setlocal
pushd %~dp0
echo.■is2 郵便番号チェック
if "%1" == "" (
	py zip_check.py 2> zip_check.err
) else (
	py zip_check.py --dt %1 --table del_syuka_h 2> zip_check.err
)
if errorlevel 1 (
	echo.%~f0 %* >> zip_check.err
	call ..\tool\slack "■is2 郵便番号エラー" %cd%\zip_check.err
rem	..\tool\blatj zip_check.err -s "■is2 郵便番号エラー" -t osakapc@kk-sdc.co.jp -c system@kk-sdc.co.jp -server ns -f is2
	..\tool\blatj zip_check.err -s "■is2 郵便番号エラー" -t kubo@kk-sdc.co.jp -server ns -f is2
) else (
	call ..\tool\slack "■is2 zip_check.err" %cd%\zip_check.err
)
type zip_check.err
popd
endlocal
exit/b

dir \\w5\newsdc\istar2\zip_check.*
xcopy/d zip_check.* \\w5\newsdc\istar2
py -m pip install jusho
