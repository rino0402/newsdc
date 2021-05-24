@echo off
rem app\HegZaiko.bat
rem 2021.05.12
setlocal
pushd %~dp0
rem pvddl newsdc sumzai02.sql -server w5
w5\f109010 /NoDialog
py HegZaiko.py "\\w5\newsdc\backup\usb\在庫データ(品番別)_0ZG_S2.xlsx" > HegZaiko_s2.log
py HegZaiko.py "\\w5\newsdc\backup\usb\在庫データ(品番別)_0ZG_S3.xlsx" > HegZaiko_s3.log
py HegZaiko.py --shift
py HegZaiko.py --sumzai
\\w5\newsdc\exe\F109030.EXE /NoDialog
call :_TZF "\\w5\newsdc\backup\usb\在庫データ(品番別)_0ZG_S2.xlsx" > HegZaiko.log
call :_TZF "\\w5\newsdc\backup\usb\在庫データ(品番別)_0ZG_S3.xlsx" >> HegZaiko.log
call :_TZF "\\w5\newsdc\work\SUMZAI_B.CSV" >> HegZaiko.log
blatj HegZaiko.log -s "HEG_在庫集計 完了" -t osakapc@kk-sdc.co.jp -c system@kk-sdc.co.jp -server ns -f %computername%
rem blatj HegZaiko.log -s "HEG_在庫集計 完了"  -t kubo@kk-sdc.co.jp -server ns -f %computername%
py slack.py %computername% %computername:w=w% "%~f0 %*" HegZaiko.log
py slack.py %computername% %computername:w=w% HegZaiko_s2.log HegZaiko_s2.log
py slack.py %computername% %computername:w=w% HegZaiko_s3.log HegZaiko_s3.log

popd
endlocal
exit/b
rem --------------------------------
:_TZF
echo.%~tzf1
exit/b
rem --------------------------------
xcopy/d/u hegzaiko.* \\w5\newsdc\app\
rem ---------------------------------------------------------
rem memo
rem ---------------------------------------------------------
在庫データ一括処理
Call ZaikoSyukei
    strCommand = "\\w5\newsdc\EXE\F109010.EXE /NoDialog"

Call ZaikoSub("S2")
    If strKind = "S2" Then
        strFilename = ThisWorkbook.Path & "\在庫データ(品番別)_0ZG_S2.xlsx"
        strRange = "G12"
        strCommand = "\\w5\newsdc\EXE\F109028.EXE"
    Else    ' S3
        strFilename = ThisWorkbook.Path & "\在庫データ(品番別)_0ZG_S3.xlsx"
        strRange = "G13"
        strCommand = "\\w5\newsdc\EXE\F109029.EXE"
    End If
    Call ZaikoConvert(strFilename, strCommand, strRange)
	    strZaiko = "\\w5\newsdc\hostfile\HS_ZAIKO_B.txt"
Call ZaikoSub("S3")

Call ZaikoCheck
    strCommand = "\\w5\newsdc\EXE\F109030.EXE /NoDialog"

Call ZaikoDaily
    strCommand = "\\w5\newsdc\EXE\dzaiko.bat"
	pvddl newsdc \\w5\newsdc\exe\makedz.sql  -server w5
	pvddl newsdc \\w5\newsdc\exe\makedzs.sql -server w5
rem ---------------------------------------------------------
