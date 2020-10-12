setlocal
pushd %~dp0
if "%1" == "" (
	xcopy/d/y s_error.html \\w1\newsdc\files\notice0\
	xcopy/d/y s_error.html \\w1\newsdc\files\notice\
	xcopy/d/y s_error.html \\w1\newsdc\files\notice2\
	xcopy/d/y s_error.html \\w1\newsdc\files\notice3\
	xcopy/d/y s_error.html \\w1\newsdc\files\notice4\
rem	xcopy/d/y s_error.html \\w1\newsdc\files\notice5\
	xcopy/d/y s_error.html \\w1\newsdc\files\notice6\

	xcopy/d/y s_error.html \\w2\newsdc\files\notice\

	xcopy/d/y s_error.html \\w3\newsdc\files\notice\
	xcopy/d/y s_error.html \\w3\newsdc\files\notice2\
	xcopy/d/y s_error.html \\w3\newsdc\files\notice3\
	xcopy/d/y s_error.html \\w3\newsdc\files\notice4\

	xcopy/d/y s_error.html \\w3\newsdcn\files\notice\

	xcopy/d/y s_error.html \\w4\newsdc\files\notice\
	xcopy/d/y s_error.html \\w4\newsdcr\files\notice\

	xcopy/d/y s_error.html \\w5\newsdc\files\notice\

	xcopy/d/y s_error.html \\w6\newsdc\files\notice\

	xcopy/d/y s_error.html \\w7\newsdc\files\notice\
	xcopy/d/y s_error.html \\w7\newsdc\files\notice2\
	xcopy/d/y s_error.html \\w7\newsdc\files\notice3\
) else (
	xcopy/d/y s_error.html %1\
)
popd
endlocal
exit/b
