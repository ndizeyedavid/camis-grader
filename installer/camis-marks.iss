; Inno Setup script for CAMIS Marks.
; Built by .github/workflows/release.yml:
;   iscc /DAppVersion=2.0.0 installer\camis-marks.iss
; Expects the PyInstaller output in dist\CAMIS Marks.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

#define AppName "CAMIS Marks"
#define AppExe "CAMIS Marks.exe"

[Setup]
AppId={{6F4C2B1E-8A3D-4E7B-9C21-CA315A4B2E10}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=CAMIS Marks
AppPublisherURL=https://github.com/ndizeyedavid/camis-grader
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=CAMIS-Marks-Setup-{#AppVersion}
SetupIconFile=..\assets\icon.ico
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
Source: "..\dist\CAMIS Marks\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\How to use"; Filename: "{app}\HOW TO USE.txt"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Open {#AppName} now"; Flags: nowait postinstall skipifsilent
