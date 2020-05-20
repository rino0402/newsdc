@rem echo ¡getin %1 ¬–ìPCyBUGlicsˆÚs‘Î‰ž 1/7–{”ÔØ‘Öz
@echo ¡getin %1 ¬–ìPCyActive‘Î‰ž 20081124z
@:_hmraa50oi
	@if not exist G:\ftpsend\ono\hmraa50oi.dat.%1-*.OK goto _hmraa50oi_end
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@rem dir G:\ftpsend\ono\hmraa50oi.dat.%1-* | findstr :
		@copy G:\ftpsend\ono\hmraa50oi.dat.%1-*.OK > nul
		@for %%i in (hmraa50oi.*.OK) do @call getins %%i %%~ni ono D
@:_hmraa50oi_end

@:_hmrba50oi
	@if not exist G:\ftpsend\ono\hmrba50oi.dat.%1-*.OK goto _hmrba50oi_end
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@rem dir G:\ftpsend\ono\hmrba50oi.dat.%1-* | findstr :
		@copy G:\ftpsend\ono\hmrba50oi.dat.%1-*.OK > nul
		@for %%i in (hmrba50oi.*.OK) do @call getins %%i %%~ni ono 4
@:_hmrba50oi_end

@:_hmqa50oi
	@if not exist G:\ftpsend\ono\hmqa50oi.dat.%1-*.OK goto _hmqa50oi_end
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@rem dir G:\ftpsend\ono\hmqa50oi.dat.%1-* | findstr :
		@copy G:\ftpsend\ono\hmqa50oi.dat.%1-*.OK > nul
		@for %%i in (hmqa50oi.*.OK) do @call getins %%i %%~ni ono 1
@:_hmqa50oi_end

@:_hmem500scs
	@if not exist H:\ftpsend\HMTAH500SCS.dat.%1-*.OK goto _hmem500scs_end
		@copy H:\ftpsend\HMTAH500SCS.dat.%1-*.OK > nul
		@for %%i in (HMTAH500SCS.*.OK) do @call getinn %%i %%~ni ono
@:_hmem500scs_end

@:_hmem500shu
	@if not exist H:\ftpsend\hmem500shu.dat.%1-*.OK goto _hmem500shu_end
		@copy H:\ftpsend\hmem500shu.dat.%1-*.OK > nul
		@for %%i in (hmem500shu.*.OK) do @call getinn %%i %%~ni ono
@:_hmem500shu_end

@:_hmem503szz
	@if not exist G:\ftpsend\ono\hmem503szz.dat.%1-*.OK goto _hmem503szz_end
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@copy G:\ftpsend\ono\hmem503szz.dat.%1-*.OK > nul
		@for %%i in (hmem503szz.*.OK) do @call getinz %%i %%~ni ono 1
@:_hmem503szz_end

@:_hmem504szz
	@if not exist G:\ftpsend\ono\hmem504szz.dat.%1-*.OK goto _hmem504szz_end
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@copy G:\ftpsend\ono\hmem504szz.dat.%1-*.OK > nul
		@for %%i in (hmem504szz.*.OK) do @call getinz %%i %%~ni ono D
@:_hmem504szz_end

@:_hmem505szz
	@if not exist G:\ftpsend\ono\hmem505szz.dat.%1-*.OK goto _hmem505szz_end
		@copy nul d:\newsdc\hostfile\shiji_out_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_out_d.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_1.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_4.dat > nul
		@copy nul d:\newsdc\hostfile\shiji_in_d.dat > nul
		@copy G:\ftpsend\ono\hmem505szz.dat.%1-*.OK > nul
		@for %%i in (hmem505szz.*.OK) do @call getinz %%i %%~ni ono 4
@:_hmem505szz_end
