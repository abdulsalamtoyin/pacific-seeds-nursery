; Pacific Seeds — Nursery Fieldbook installer (Inno Setup)
;
; Builds a Windows installer that places the .exe under Program Files,
; creates Start Menu + Desktop shortcuts, and registers an uninstaller.
;
; Requirements:
;   1. Build the .exe first via build.bat (produces dist\PacificSeeds.exe)
;   2. Install Inno Setup: https://jrsoftware.org/isinfo.php
;   3. Open this .iss in Inno Setup Compiler and click Compile.
;      → Output\PacificSeedsSetup.exe is the installable distributable.

#define MyAppName      "Pacific Seeds Nursery Fieldbook"
#define MyAppVersion   "1.0.0"
#define MyAppPublisher "Pacific Seeds"
#define MyAppURL       "https://www.pacificseeds.com.au/"
#define MyAppExeName   "PacificSeeds.exe"

[Setup]
AppId={{A3F1E7C2-9D4B-4F1A-B8E2-7C3D5A9E1F2B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\PacificSeeds
DefaultGroupName=Pacific Seeds
OutputDir=..\dist\installer
OutputBaseFilename=PacificSeedsSetup-{#MyAppVersion}
; SetupIconFile=..\pwa\ps-logo.ico   ← uncomment after exporting a real .ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DisableProgramGroupPage=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\PacificSeeds\__pycache"
