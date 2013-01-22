echo off

set Titre=Inscription des DLLs

Title %Titre%

cd %SYSTEMDRIVE%\Framework_SW

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

set DossierNET="%WINDIR%\Microsoft.NET\Framework\v4.0.30319"
set DossierNET64="%WINDIR%\Microsoft.NET\Framework64\v4.0.30319"

If exist %DossierNET% cd %DossierNET%
If exist %DossierNET64% cd %DossierNET64%

echo.
echo.
echo %Titre%
echo --------------------------------------------------------
echo Nom de la DLL :
echo    %NomDLL%
echo Dossier courant :
echo    %cd%

echo.
echo --------------------------------------------------------
RegAsm.exe "%FichierDLL%" /codebase /tlb:"%FichierTLB%"
echo --------------------------------------------------------
Pause