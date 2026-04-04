; Inno Setup script for FatturaView 1.2.0
; Requires Inno Setup 6

#define AppName "FatturaView"
#define AppVersion "1.2.0"
#define AppPublisher "Roberto Ferri"
#define AppExeName "FatturaView.exe"
#define AppGUID "{{55618536-2590-4AD9-ADA8-DEE56FFBB9EE}"

[Setup]
AppId={#AppGUID}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=no
OutputBaseFilename=FatturaView_Setup_v{#AppVersion}_x64
Compression=lzma2/ultra64
SolidCompression=yes
ChangesAssociations=yes

; Architecture: 64-bit
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible

; License file included and shown during install
LicenseFile=..\LICENSE

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

; [Registry] rimosse associazioni file come richiesto dall'utente

[Files]
Source: "..\x64\Release\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Resources\*"; DestDir: "{app}\Resources"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Redist\vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Tasks]
Name: "desktopicon"; Description: "Crea un'icona sul desktop"; GroupDescription: "Icone aggiuntive:"; Flags: unchecked

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"
Name: "{group}\Disinstalla FatturaView"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
; Installa VC++ Redistributable silenziosamente se non già presente
Filename: "{tmp}\vc_redist.x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "Installazione Visual C++ Redistributable..."; Flags: waituntilterminated skipifdoesntexist; Check: not VCRedistInstalled

Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram, {#AppName}}"; Flags: nowait postinstall skipifsilent

[Code]
function VCRedistInstalled(): Boolean;
var
  Version: String;
begin
  Result := RegQueryStringValue(
    HKLM,
    'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64',
    'Version',
    Version
  );
  if Result then
    Result := (CompareStr(Version, 'v14.30') >= 0);
end;
