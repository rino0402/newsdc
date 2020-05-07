@echo off
VER | find "Version 5.2." > nul
IF not errorlevel 1 echo.Server2003
