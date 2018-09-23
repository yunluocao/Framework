@echo off
set CURRDIR=%CD%
cd %~dp0
set INSTALLDIR=%CD%

IF not [%1]==[] GOTO other  
start csc_server_install.exe %*
goto end

:other
csc_server_install.exe %*

:end
cd %CURRDIR%
