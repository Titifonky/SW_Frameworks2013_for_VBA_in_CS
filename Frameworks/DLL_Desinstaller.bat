echo off

mode con cols=150 lines=60

set Titre=Suppression des DLLs
echo.
echo ================================================================
echo =
echo =                      %Titre%
echo =
echo ================================================================

Title %Titre%

@setlocal enableextensions
@cd /d "%~dp0"

rem cd %SYSTEMDRIVE%\Framework_SW

echo.
echo Dossier courant :
echo    %cd%

setlocal enableDelayedExpansion
for %%i in (*.dll) do (
set NomDLL=%%~ni.dll
set NomTLB=%%~ni.tlb
)

for /r %%i in (%NomTLB%) do (
set FichierTLB=%%i
)

for /r %%i in (%NomDLL%) do (
set FichierDLL=%%i
)

set DossierCourant=%cd%
set DossierNET="%WINDIR%\Microsoft.NET\Framework\v4.0.30319"
set DossierNET64="%WINDIR%\Microsoft.NET\Framework64\v4.0.30319"

If exist %DossierNET% (
cd /d %DossierNET%
)
If exist %DossierNET64% (
cd /d %DossierNET64%
)

echo.
echo.
echo %Titre%
echo --------------------------------------------------------
echo Nom de la DLL :
echo    %NomDLL%
echo Dossier courant :
echo    %cd%
if %DossierCourant%==%cd% (
echo.
echo Le framework .NET v4.0.30319 n'est pas installe sur la machine
echo Il est necessaire au fonctionnement de la dll
echo.
echo FIN
pause
exit
)



echo.
echo --------------------------------------------------------
RegAsm.exe "%FichierDLL%" /codebase /tlb:"%FichierTLB%" /unregister
echo --------------------------------------------------------
echo.
echo FIN
Pause