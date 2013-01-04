echo off

Title Inscription des DLLs

setlocal enableDelayedExpansion
for %%i in (*.csproj) do (
set NomProjet=%%i
set NomDLL=%%~ni.dll
set NomTLB=%%~ni.tlb
)

echo --------------------------------------------------------
echo Nom du Projet : %NomProjet%
echo Nom de la DLL : %NomDLL%
echo Nom de la TLB : %NomTLB%
cd bin
cd Release

for /r %%i in (%NomTLB%) do (
set FichierTLB=%%i
)

echo Chemin de la TLB : %FichierTLB%
echo --------------------------------------------------------
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe %NomDLL% /codebase /tlb:"%FichierTLB%"
echo --------------------------------------------------------
Pause