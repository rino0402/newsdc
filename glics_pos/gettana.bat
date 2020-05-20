rem @echo off
@for /f "tokens=1,2,3 delims=/, " %%i in ( 'date /t' ) do set DT=%%i%%j%%k
@echo ¡GLICS’I”ÔˆêŠ‡“o˜^Œ‹‰Ê(00023410/hdrba750b.csv)óM
@echo ¡GLICS’I”ÔˆêŠ‡“o˜^Œ‹‰Ê(00023510/hdraa750b.csv)óM
@echo ¡GLICS’I”ÔˆêŠ‡“o˜^Œ‹‰Ê(00023100/hdqa750b.csv )óM
del hdrba750b.csv
"c:\Program Files\ffftp\FFFTP" -q -d -f SDCPOS:SDCPOS@a0o106/apl1/euc/euc/00023410/hdrba750b.csv
copy hdrba750b.csv tana\hdrba750b-%DT%.csv

del hdraa750b.csv
"c:\Program Files\ffftp\FFFTP" -q -d -f SDCPOS:SDCPOS@a0o106/apl1/euc/euc/00023510/hdraa750b.csv
copy hdraa750b.csv tana\hdraa750b-%DT%.csv

del hdqa750b.csv
"c:\Program Files\ffftp\FFFTP" -q -d -f SDCPOS:SDCPOS@a0o106/apl1/euc/euc/00023100/hdqa750b.csv
copy hdqa750b.csv  tana\hdqa750b.csv-%DT%.csv
