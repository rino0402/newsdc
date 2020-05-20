@echo ¡getcomp %*
@rem dir  H:\ftpsend\hmec770a%2.dat.%1-* | findstr :

@:_hmec770a
@if not exist H:\ftpsend\hmtac770a%2.dat.%1-*.OK goto _hmec770a_end
	@copy  H:\ftpsend\hmtac770a%2.dat.%1-*.OK > nul
	@for %%i in (hmtac770a%2.*.OK) do @call getcomps %%i %%~ni
@:_hmec770a_end
