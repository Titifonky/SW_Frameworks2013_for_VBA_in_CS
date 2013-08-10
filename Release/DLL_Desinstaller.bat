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

echo.
echo Dossier courant :
echo    %cd%

setlocal enableDelayedExpansion
for %%i in (*.dll) do (
set FichierDLL=%%~fi
set FichierTLB=%%~dpni.tlb
)

<<<<<<< HEAD
for %%i in (%NomTLB%) do (
set FichierTLB=%%~si
)

for %%i in (%NomDLL%) do (
set FichierDLL=%%~si
)

=======
>>>>>>> b53be8ba9234f98978e3808d465d384100e015d9
if not exist "%FichierTLB%" (
echo.
echo Pas de fichier .tlb pour la desinscription de la dll
echo.
pause
exit
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
echo Dossier du framework .NET
echo    %cd%

echo.
echo.
echo %Titre%
echo --------------------------------------------------------
echo Nom de la DLL a supprimer :
echo    %FichierDLL%

if "%DossierCourant%"=="%cd%" (
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

del "%FichierTLB%"

echo FIN
Pause