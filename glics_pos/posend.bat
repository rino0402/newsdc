@echo off
setlocal
pushd %~dp0
rem 2008.02.08 complt.bat��蕪��
rem 2011.02.07 �o�b�N�R�}���h��ύX copy �� xcopy
rem 2011.02.07 ��ƃ��O�W�v���� ��ǉ�
rem 2013.11.12 ��ƃ��O�W�v���� �̏��������ŏ�����Ō�ɕύX
rem 2013.12.26 files�o�b�N�A�b�v ���Ō�ɕύX
rem	           welcat.txt���[�e�[�V���� ��ǉ�
rem 2014.02.27 �o�b�`���s���O(batlog)
call d:\log\batlog �� %0 %*
set	NewSdc=%1
if "%NewSdc%"=="" set	NewSdc=newsdc
set NewSdc
echo.%DATE% %TIME:~0,8% %NewSdc% �� > posend.txt

tasklist /FI "IMAGENAME eq F110010.exe" | findstr /i F110010.exe
if "%ERRORLEVEL%" == "0" (
	echo.���X�L���i����N����...
	echo.%DATE% %TIME:~0,8% F110010 �X�L���i����N���� >> posend.txt
)

echo.��Glics�A�g�`�F�b�N��
echo.%DATE% %TIME:~0,8% Glics�A�g�`�F�b�N >> posend.txt
pvddl %NewSdc% complt.sql

if /i "%NewSdc%" neq "newsdcn" (
	echo.��y_syuka �o�ɍς����i�ςɃZ�b�g
	echo.%DATE% %TIME:~0,8% y_syuka �o�ɍς����i�ςɃZ�b�g >> posend.txt
	pvddl %NewSdc% y_syuka_kenpin.sql
)

echo.��F110070:�o�ח\��폜
echo.%DATE% %TIME:~0,8% F110070:�o�ח\��폜 >> posend.txt
d:\%NewSdc%\exe\f110070.exe

echo.��F110030:�s�v�f�[�^�폜
echo.%DATE% %TIME:~0,8% F110030:�s�v�f�[�^�폜 >> posend.txt
d:\%NewSdc%\exe\F110030.exe

echo ����Ǝ��ԃZ�b�g��
rem call D:\newsdc\FILES\sagyolog.bat

rem echo ��welcat.txt���[�e�[�V������
rem call \\hs1\it\bin\rotate d:\newsdc\files\welcat\welcat.txt

echo.��POS files�o�b�N�A�b�v��
echo.%DATE% %TIME:~0,8% files�o�b�N�A�b�v >> posend.txt
xcopy/d/y d:\%NewSdc%\files\*.* d:\%NewSdc%\backup\files\

rem echo %date% %time% >> posend.log
rem echo ���݌ɏW�v����:F109010�� 2007.05.22 ������~(zaiko.bat�̂ݎ��s)
rem \\w1\newsdc\exe\F109010.exe
echo.%DATE% %TIME:~0,8% %NewSdc% �� >> posend.txt
call d:\log\batlog �� %0 %*
call d:\newsdc\tool\slack "%0 %*" %cd%\posend.txt
popd
endlocal
rem call d:\log\batlog �� %0 %* shutdown
rem shutdown /r /f
