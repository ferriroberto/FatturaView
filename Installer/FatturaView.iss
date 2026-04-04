; ===========================================================================
;  FatturaView - Script di installazione Inno Setup
;  Applicazione: Visualizzatore Fatture Elettroniche
;  Autore: Roberto Ferri
;  Versione: 1.3.3
; ===========================================================================

#define AppName        "FatturaView"
#define AppFullName    "FatturaView - Visualizzatore Fatture Elettroniche"
#define AppVersion     "1.3.3"
#define AppPublisher   "Roberto Ferri"
#define AppURL         "https://github.com/ferriroberto/FatturaView"
#define AppExeName     "FatturaView.exe"
#define AppGUID        "{{55618536-2590-4AD9-ADA8-DEE56FFBB9EE}"

; Percorso della build Release (relativo a questa cartella .iss)
#define BuildDir       "..\x64\Release"
#define SourceDir      ".."

[Setup]
AppId={#AppGUID}
AppName={#AppFullName}
AppVersion={#AppVersion}
AppVerName={#AppFullName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
AppCopyright=Copyright (C) 2026 {#AppPublisher}

; Percorso di installazione predefinito
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppFullName}

; Output installer
OutputDir=Output
OutputBaseFilename=FatturaView_Setup_v{#AppVersion}_x64

; Icona dell'installer
SetupIconFile={#SourceDir}\FatturaView.ico

; Richiede Windows 10 o superiore, 64-bit
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible

; Compressione LZMA (massima)
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

; Richiede privilegi amministrativi
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; Wizard
WizardStyle=modern
WizardResizable=no

; Licenza e informazioni (mostrata durante il wizard)
LicenseFile={#SourceDir}\LICENSE

; Abilita la pagina di selezione componenti
DisableProgramGroupPage=no

; Aggiunge voce in "Programmi e funzionalità"
UninstallDisplayIcon={app}\FatturaView.ico
UninstallDisplayName={#AppFullName}
ChangesAssociations=yes

; Lingua predefinita
ShowLanguageDialog=auto

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

; [Registry] rimosse associazioni file come richiesto dall'utente

; ===========================================================================
;  FILE DA INSTALLARE
; ===========================================================================
[Files]
; Eseguibile principale
Source: "{#BuildDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Icona applicazione (usata per i collegamenti)
Source: "{#SourceDir}\FatturaView.ico"; DestDir: "{app}"; Flags: ignoreversion

; Fogli di stile XSL (obbligatori per la visualizzazione)
Source: "{#SourceDir}\Resources\*"; DestDir: "{app}\Resources"; Flags: ignoreversion recursesubdirs createallsubdirs

; Documentazione 
Source: "{#SourceDir}\README.md";             DestDir: "{app}\Doc"; Flags: ignoreversion skipifsourcedoesntexist; Components: docs

; Visual C++ 2022 Redistributable x64
Source: "Redist\vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Components: vcredist

; ===========================================================================
;  ESECUZIONE POST-INSTALLAZIONE
; ===========================================================================
[Run]
; Installa VC++ Redistributable silenziosamente se non già presente
Filename: "{tmp}\vc_redist.x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "Installazione Visual C++ Redistributable..."; Flags: waituntilterminated skipifdoesntexist; Check: not VCRedistInstalled; Components: vcredist

Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram, {#AppName}}"; Flags: nowait postinstall skipifsilent

; Pulisci la cache delle icone di Windows per forzare il refresh
Filename: "ie4uinit.exe"; Parameters: "-show"; Flags: nowait runhidden; StatusMsg: "Aggiornamento icone..."

; ===========================================================================
;  COMPONENTI SELEZIONABILI
; ===========================================================================
[Components]
Name: "main";      Description: "Applicazione principale (obbligatorio)"; Types: full compact custom; Flags: fixed
Name: "vcredist"; Description: "Microsoft Visual C++ Redistributable";  Types: full compact custom; Flags: fixed
Name: "docs";      Description: "Documentazione (README, Guide)";         Types: full

; ===========================================================================
;  CARTELLE DA CREARE
; ===========================================================================
[Dirs]
Name: "{app}\Resources"
Name: "{app}\Doc"

; ===========================================================================
;  ICONE (COLLEGAMENTO START MENU E DESKTOP)
; ===========================================================================
[Icons]
; Start Menu
Name: "{group}\{#AppFullName}";        Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\FatturaView.ico"
Name: "{group}\Disinstalla {#AppName}"; Filename: "{uninstallexe}"

; Desktop (opzionale, l'utente può deselezionarlo)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\FatturaView.ico"; Tasks: desktopicon

; ===========================================================================
;  TASK OPZIONALI (PAGINA WIZARD)
; ===========================================================================
[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

; ===========================================================================
;  CODICE PASCAL - CONTROLLO VC++ REDISTRIBUTABLE
; ===========================================================================
[Code]
// ---------------------------------------------------------------------------
// Verifica se Visual C++ 2022 Redistributable x64 è già installato
// ---------------------------------------------------------------------------
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
  // Verifica versione minima 14.30 (VS2022)
  if Result then
    Result := (CompareStr(Version, 'v14.30') >= 0);
end;

// ---------------------------------------------------------------------------
// Prima dell'installazione: informa se il VC++ Redistributable verrà installato
// ---------------------------------------------------------------------------
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if not VCRedistInstalled() then
  begin
    MsgBox(
      'Visual C++ 2022 Redistributable (x64) non rilevato.' + #13#10
      + #13#10
      + 'Verrà installato automaticamente durante il setup.',
      mbInformation,
      MB_OK
    );
  end;
end;
