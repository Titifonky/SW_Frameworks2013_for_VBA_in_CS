; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define Application "Framework VBA pour Solidworks 2012"
#define Developpeur "Etienne Canuel"
#define Version "1-1"
#define FichierFW "Framework_SW2012_V" 

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{B8E2812F-D5F8-4043-9F22-86996372C462}
AppName={#Application}
AppVersion={#Version}
;AppVerName={#Application} {#Version}
AppPublisher={#Developpeur}
DefaultDirName={pf}\{#Application}
DisableDirPage=yes
DefaultGroupName={#Application}
DisableProgramGroupPage=yes
OutputDir=Setup
OutputBaseFilename={#FichierFW}{#Version}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Dirs]
Name: "{app}"; Permissions: users-modify

[Files]

Source: "{#FichierFW}{#Version}.dll"; DestDir: "{app}"; Flags: ignoreversion; Permissions: users-modify
Source: "DLL_Desinstaller.bat"; DestDir: "{app}"; Flags: ignoreversion; Permissions: users-modify
Source: "DLL_Installer.bat"; DestDir: "{app}"; Flags: ignoreversion; Permissions: users-modify
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Run]
Filename: "{dotnet40}\RegAsm.exe"; Parameters: {#FichierFW}{#Version}.dll /codebase /tlb:{#FichierFW}{#Version}.tlb; WorkingDir: {app}; StatusMsg: "Installation de la dll..."; Flags: runminimized runascurrentuser 

[Icons]
Name: "{group}\Dossier d'installation"; Filename: "{app}\"; IconIndex: 0
Name: "{group}\D�sinstaller le framework"; Filename: "{uninstallexe}"; IconIndex: 1

[UninstallRun]
Filename: "{dotnet40}\RegAsm.exe"; Parameters: {#FichierFW}{#Version}.dll /codebase /tlb:{#FichierFW}{#Version}.tlb /unregister; WorkingDir: {app}; StatusMsg: "Suppression de la dll..."; Flags: runminimized

[UninstallDelete]
Type: files; Name: "{app}\{#FichierFW}{#Version}.tlb"
