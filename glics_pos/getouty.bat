@echo off
rem setlocal enabledelayedexpansion
setlocal
pushd %~dp0
rem ���� 2008.11.29 Active�Ή�(�o�ח\��f�[�^�ƍ��@�\�ǉ�)
rem ���� 2009.01.05 �o�ח\��ƍ�OK �̃��[�����M����
rem ���� 2013.11.05 �o�ח\��f�[�^�ϊ��R���g���[���Ή�
rem ���� 2013.11.06 �o�ח\��f�[�^�ϊ��R���g���[���Ή�
rem 2013.11.07 ng�t�@�C����ng�t�H���_�ɕۑ�����悤�ɕύX
rem 2014.02.28 ���O�o�͑Ή�
rem 2014.04.24 �o�׎��јA�g�̊����`�F�b�N
rem 2014.04.28 �o�׎��јA�g�̊������[�����M
rem 2015.12.24 blaj�̑��M������+10�ɂȂ�̂��C��
rem 2016.07.19 �i���敪�S�ȊO�폜�F�b��Ώ�
rem 2016.10.01 P�Y�@�Ή�
rem 2016.10.16 P�Y�@�����f�[�^�o�^
rem 2016.10.26 P�Y�@SSX(R-smile)�����f�[�^�쐬
rem 2017.03.08 ActiveGift�Ή�
rem
rem getouty HMTAH015SZZ.dat.20161001-141032.OK HMTAH015SZZ.dat.20161001-141032
rem g_syuka.vbs
rem g_syuka_del.sql
rem d:\newsdc\files\glicspos.vbs
rem d:\newsdc\files\HMTAH015.sql
rem outy\HMTAH015SZZ.dat.20161001-141032.OK
rem outy\HMTAH015SZZ.dat.20161001-141032
rem getoutn.ok
rem check_g_syuka.vbs
rem check_g_syuka.txt
rem outy\HMTAH015SZZ.dat.20161001-141032.txt
rem tool\blatj
rem getcomps.done
rem complt.sql
rem d:\newsdc\files\y_syuka_check.vbs
rem y_syuka_check.txt
rem y_syuka_check.end
rem y_syuka_check.send
rem d:\newsdc\files\b2data.vbs
set ret=0
set absPath=%1
set relPath=%~nx1
if exist outy\%relPath% goto _End
echo.��getouty %*
rem call d:\log\batlog �� %0 %*
rem call slack "%relPath%:��"
color 9F
echo.����Active�o�ח\��
echo.xcopy/d/y %absPath% ...%time%
xcopy/d/y %absPath%
echo.xcopy/d/y %absPath% ...%time%����
if not "%ERRORLEVEL%"  == "0" (
	call d:\log\batlog �� %0 %* xcopy:%ERRORLEVEL%
	call d:\newsdc\tool\slack "��getouty %relPath%:��xcopy:%ERRORLEVEL%"
	del %relPath%
	GOTO _End
)
set fSize=0
for %%i in ( %relPath% ) do set fSize=%%~zi
if %fSize% == 0 (
	call d:\log\batlog �� %0 %* fSize:%fSize%
	call d:\newsdc\tool\slack "��getouty %relPath%:��fSize:%fSize%"
	del %relPath%
	GOTO _End
)

if exist getouty.done del getouty.done
echo.�����o�ח\��f�[�^�ϊ� HMTAH015_t
cscript//Nologo d:\newsdc\files\glicspos.vbs %relPath%
echo.�����o�ח\��f�[�^�ϊ� HMTAH015
cscript//Nologo d:\newsdc\files\HMTAH015.vbs
echo.���������f�[�^�o�^ HMTAH015_c
python d:\newsdc\files\HMTAH015.py>HMTAH015_c.log
call d:\newsdc\tool\slack "HMTAH015_c" %cd%\HMTAH015_c.log
type beeps.txt
move/y %relPath% outy\

echo.�����o�ח\��f�[�^�ƍ�
cscript//Nologo check_g_syuka.vbs
copy /y check_g_syuka.txt outy\%relPath%.txt
for %%i in ( check_g_syuka.txt ) do set FSize=%%~zi
if not %FSize% == 0 (
rem	echo.>>check_g_syuka.txt
	echo.%0 %* >>check_g_syuka.txt
	tool\blatj check_g_syuka.txt -attach outy\%relPath%.txt -s "��Active�o�ח\��ƍ��G���[" -t %ML% -c system@kk-sdc.co.jp
	call d:\newsdc\tool\slack "��Active�o�ח\��ƍ��G���[" %cd%\check_g_syuka.txt
)
echo.�����o�׊����`�F�b�N
cscript//Nologo d:\newsdc\files\y_syuka_check.vbs > y_syuka_check.txt
if %ERRORLEVEL% == 0 (
	echo.�o�׊���:%ERRORLEVEL% >> y_syuka_check.txt
	echo.%0 %* >> y_syuka_check.txt
	rem �{���̏o�׎��јA�g�c���O
	if not exist y_syuka_check.end (
		tool\blatj y_syuka_check.txt -s "�o�׎��јA�g(����)" -t %ML% -c system@kk-sdc.co.jp
		call d:\newsdc\tool\slack "��Active�o�׎��јA�g(����)" %cd%\y_syuka_check.txt
		copy /y y_syuka_check.txt y_syuka_check.end
	) else (
rem		tool\blatj y_syuka_check.txt -s "(��)%relPath%" -t log@kk-sdc.co.jp
		call d:\newsdc\tool\slack "��Active�o�׎��јA�g(������)" %cd%\y_syuka_check.txt
	)
) else (
	echo.���юc:%ERRORLEVEL% >> y_syuka_check.txt
	echo.%0 %* >> y_syuka_check.txt
	if exist y_syuka_check.send (
		tool\blatj y_syuka_check.txt -s "�o�׎��јA�g(������)" -t %ML% -c system@kk-sdc.co.jp
		call d:\newsdc\tool\slack "��Active�o�׎��јA�g(������)" %cd%\y_syuka_check.txt
		del y_syuka_check.send
	) else (
		rem �󋵃`�F�b�N�p���[�����M(����ғ�����Ή���)
rem		tool\blatj y_syuka_check.txt -s "(��)%relPath%" -t log@kk-sdc.co.jp
		call d:\newsdc\tool\slack "��Active�o�׎��јA�g" %cd%\y_syuka_check.txt
	)
	del y_syuka_check.end
)
echo %0 %* > getouty.done
if exist d:\newsdc\B2\INPUT (
	call d:\newsdc\B2\makecsv.bat
)

xcopy/d/y \\w4\newsdc\files\ACSHORT.DAT d:\newsdc\files\ > acshort.log 2>&1
call d:\newsdc\tool\slack "acshort.log" %cd%\acshort.log

rem call d:\log\batlog �� %0 %*
rem call slack "%relPath%:��"
set ret=1
:_End
color
for %%i in (outy\%relPath%) do echo.��getouty %* %%~zi %ML%
popd
endlocal && set ret=%ret%
exit/b %ret%
