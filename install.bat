@echo off
rem if exist "%appdata%\Microsoft\Excel\XLSTART\fpsgmacros.xla" del "%appdata%\Microsoft\Excel\XLSTART\fpsgmacros.xla"
rem if exist "%appdata%\Microsoft\Excel\XLSTART\fpsgmacros.xlam" del "%appdata%\Microsoft\Excel\XLSTART\fpsgmacros.xlam"
@echo on
if exist "%PROGRAMFILES(X86)%\Microsoft Office\Office15\Library" copy "\\glasfile\software\excelmacros\fpsgMacrosMaster.xlam" "%PROGRAMFILES(X86)%\Microsoft Office\Office15\Library\fpsgMacros.xlam" /Y
if exist "%PROGRAMFILES(X86)%\Microsoft Office\Office15\Library" copy "\\glasfile\software\excelmacros\jacobsmaster.xlam" "%PROGRAMFILES(X86)%\Microsoft Office\Office15\Library\jacobs.xlam" /Y
if exist "%PROGRAMFILES%\Microsoft Office\Office15\Library" copy "\\glasfile\software\excelmacros\fpsgMacrosMaster.xlam" "%PROGRAMFILES%\Microsoft Office\Office15\Library\fpsgMacros.xlam" /Y
if exist "%PROGRAMFILES%\Microsoft Office\Office15\Library" copy "\\glasfile\software\excelmacros\jacobsmaster.xlam" "%PROGRAMFILES%\Microsoft Office\Office15\Library\jacobs.xlam" /Y

pause