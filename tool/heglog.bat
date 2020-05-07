setlocal
set SV=%1
if "%SV%" == "" set SV=hs1
cscript tail-f.VBS \\%SV%\heg\y_syuka_heg.log
endlocal
