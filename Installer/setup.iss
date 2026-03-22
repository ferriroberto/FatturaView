; Inno Setup script for FatturaView 1.1.0
; Requires Inno Setup 6

#define AppName "FatturaView"
#define AppVersion "1.1.0"
#define AppPublisher "Roberto Ferri"
#define AppExeName "FatturaView.exe"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={pf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=no
OutputBaseFilename=FatturaView_Setup_{#AppVersion}
Compression=lzma
SolidCompression=yes

; License file included and shown during install
LicenseFile=..\LICENSE

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Files]
Source: "..\Installer\dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "*.AppImage"
Source: "..\Resources\*"; DestDir: "{app}\Resources"; Flags: ignoreversion recursesubdirs createallsubdirs

[Tasks]
Name: "desktopicon"; Description: "Crea un'icona sul desktop"; GroupDescription: "Icone:"; Flags: unchecked

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{#AppName} - Uninstall"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram, {#AppName}}"; Flags: nowait postinstall skipifsilent
