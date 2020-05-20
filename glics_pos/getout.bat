@echo ¡getout %1 ¬–ìPCyBUGlicsˆÚs‘Î‰ž 1/7–{”ÔØ‘Öz

@:_hmra71bo
		@if exist G:\ftpsend\ono\hmra71bo.dat.%1-*.OK @copy G:\ftpsend\ono\hmra71bo.dat.%1-*.OK > nul
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@for %%i in (hmra71bo.*.OK)   do @call getouts %%i %%~ni ono D
@:_hmra71bo_end

@:_hmrb71bo
		@if exist G:\ftpsend\ono\hmrb71bo.dat.%1-*.OK @copy G:\ftpsend\ono\hmrb71bo.dat.%1-*.OK > nul
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@for %%i in (hmrb71bo.*.OK)   do @call getouts %%i %%~ni ono 4
@:_hmrb71bo_end

@:_hmqa71bo
		@if exist G:\ftpsend\ono\hmqa71bo.dat.%1-*.OK @copy G:\ftpsend\ono\hmqa71bo.dat.%1-*.OK > nul
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@for %%i in (hmqa71bo.*.OK)   do @call getouts %%i %%~ni ono 1
@:_hmqa71bo_end

@:_hmem011scs
		@if exist H:\ftpsend\HMTAH011SCS.dat.%1-*.OK @copy H:\ftpsend\HMTAH011SCS.dat.%1-*.OK > nul
		@for %%i in (HMTAH011SCS*.OK) do @call getoutn %%i %%~ni ono
@:_hmem011scs_end

@:_hmem011shu
		@if exist H:\ftpsend\HMTAH011SHU.dat.%1-*.OK @copy H:\ftpsend\HMTAH011SHU.dat.%1-*.OK > nul
		@for %%i in (HMTAH011SHU.*.OK) do @call getoutn %%i %%~ni ono
@:_hmem011shu_end

@:_hmem015szz
		@if exist H:\ftpsend\HMTAH015SZZ.dat.%1-*.OK @copy H:\ftpsend\HMTAH015SZZ.dat.%1-*.OK > nul
		@for %%i in (HMTAH015SZZ.*.OK) do @call getouty %%i %%~ni
@:_hmem015szz_end

@:_hmem703szz
		@if exist G:\ftpsend\ono\hmem703szz.dat.%1-*.OK @copy G:\ftpsend\ono\hmem703szz.dat.%1-*.OK > nul
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@for %%i in (hmem703szz.*.OK)   do @call getoutz %%i %%~ni ono 1
@:_hmem703szz_end

@:_hmem704szz
		@if exist G:\ftpsend\ono\hmem704szz.dat.%1-*.OK @copy G:\ftpsend\ono\hmem704szz.dat.%1-*.OK > nul
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@for %%i in (hmem704szz.*.OK)   do @call getoutz %%i %%~ni ono D
@:_hmem704szz_end

@:_hmem705szz
		@if exist G:\ftpsend\ono\hmem705szz.dat.%1-*.OK @copy G:\ftpsend\ono\hmem705szz.dat.%1-*.OK > nul
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@for %%i in (hmem705szz.*.OK)   do @call getoutz %%i %%~ni ono 4
@:_hmem705szz_end
