# =============================================================================
#  Build-Installer.ps1
#  Compila FatturaView in Release x64 e genera il pacchetto di installazione
#  Uso: .\Installer\Build-Installer.ps1
#  Firma: .\Installer\Build-Installer.ps1 -CertPath "C:\Cert\MyCert.pfx" -CertPassword "pwd"
# =============================================================================

param(
    [string]$Version      = "1.0.0",
    [switch]$SkipBuild,
    # Firma digitale (opzionale) - fornire entrambi per attivare la firma
    [string]$CertPath     = "",
    [string]$CertPassword = "",
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
$RootDir   = Split-Path $PSScriptRoot -Parent
$BuildDir  = Join-Path $RootDir "x64\Release"
$ProjectFile = Join-Path $RootDir "FatturaView.vcxproj"
$IssFile   = Join-Path $PSScriptRoot "FatturaView.iss"
$OutputDir = Join-Path $PSScriptRoot "Output"

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  FatturaView - Build & Installer v$Version" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 1. Trova MSBuild
# ---------------------------------------------------------------------------
$MsBuild = Get-ChildItem "C:\Program Files\Microsoft Visual Studio" -Recurse -Filter "MSBuild.exe" -ErrorAction SilentlyContinue |
           Where-Object { $_.FullName -like "*\amd64\MSBuild.exe" } |
           Select-Object -First 1 -ExpandProperty FullName

if (-not $MsBuild) {
    $MsBuild = Get-ChildItem "C:\Program Files\Microsoft Visual Studio" -Recurse -Filter "MSBuild.exe" -ErrorAction SilentlyContinue |
               Select-Object -First 1 -ExpandProperty FullName
}

if (-not $MsBuild) {
    Write-Error "MSBuild.exe non trovato. Installare Visual Studio 2022."
}
Write-Host "[OK] MSBuild: $MsBuild" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 2. Build Release x64
# ---------------------------------------------------------------------------
if (-not $SkipBuild) {
    Write-Host "`n[1/3] Compilazione Release x64..." -ForegroundColor Yellow
    Push-Location $RootDir
    & $MsBuild $ProjectFile /p:Configuration=Release /p:Platform=x64 /m /nologo /verbosity:minimal
    if ($LASTEXITCODE -ne 0) {
        Pop-Location
        Write-Error "Build fallita (exit code $LASTEXITCODE). Controllare gli errori sopra."
    }
    Pop-Location
    Write-Host "[OK] Build completata." -ForegroundColor Green
} else {
    Write-Host "[--] Build saltata (flag -SkipBuild attivo)." -ForegroundColor Gray
}

# Verifica che l'exe esista
$ExePath = Join-Path $BuildDir "FatturaView.exe"
if (-not (Test-Path $ExePath)) {
    Write-Error "Eseguibile non trovato: $ExePath"
}

# ---------------------------------------------------------------------------
# 2b. Firma digitale dell'eseguibile (opzionale)
# ---------------------------------------------------------------------------
if ($CertPath -and $CertPassword) {
    Write-Host "`n[2b] Firma digitale FatturaView.exe..." -ForegroundColor Yellow

    $SigntoolPaths = @(
        "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe"
    ) + (Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe" -ErrorAction SilentlyContinue |
         Sort-Object -Descending | Select-Object -First 1 -ExpandProperty FullName)

    $Signtool = $SigntoolPaths | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

    if (-not $Signtool) {
        Write-Host "[AVVISO] signtool.exe non trovato. Installare Windows SDK." -ForegroundColor Red
        Write-Host "         Firma saltata - l'installer non sarà firmato." -ForegroundColor Yellow
    } else {
        Write-Host "[OK] signtool: $Signtool" -ForegroundColor Green
        & $Signtool sign /fd sha256 /tr $TimestampUrl /td sha256 /f $CertPath /p $CertPassword $ExePath
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Firma eseguibile fallita (exit code $LASTEXITCODE)."
        }
        Write-Host "[OK] Eseguibile firmato." -ForegroundColor Green

        # Aggiorna anche la direttiva SignTool nel file .iss per firmare l'installer
        $IssContent = Get-Content $IssFile -Raw
        $SignCmd = "sign /fd sha256 /tr $TimestampUrl /td sha256 /f `"$CertPath`" /p `"$CertPassword`" `$f"
        $IssContent = $IssContent -replace '; SignTool=standard .*', "SignTool=standard $SignCmd"
        Set-Content $IssFile $IssContent -Encoding UTF8
        Write-Host "[OK] Firma installer configurata nel .iss." -ForegroundColor Green
    }
} else {
    Write-Host "`n[--] Firma digitale saltata (parametri -CertPath/-CertPassword non forniti)." -ForegroundColor Gray
    Write-Host "     Per firmare: .\Build-Installer.ps1 -CertPath 'C:\Cert\MyCert.pfx' -CertPassword 'pwd'" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# 3. Trova Inno Setup Compiler
# ---------------------------------------------------------------------------
Write-Host "`n[2/3] Ricerca Inno Setup..." -ForegroundColor Yellow

$IsccPaths = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
)
$Iscc = $IsccPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $Iscc) {
    Write-Host "[AVVISO] Inno Setup non trovato." -ForegroundColor Red
    Write-Host "Scaricalo da: https://jrsoftware.org/isdl.php" -ForegroundColor Yellow
    Write-Host "Dopo l'installazione, riesegui questo script." -ForegroundColor Yellow
    exit 1
}
Write-Host "[OK] Inno Setup: $Iscc" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. Crea cartella output
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# ---------------------------------------------------------------------------
# 5. Compila l'installer
# ---------------------------------------------------------------------------
Write-Host "`n[3/3] Generazione installer..." -ForegroundColor Yellow
& $Iscc $IssFile "/DAppVersion=$Version" /O"$OutputDir"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Inno Setup fallito (exit code $LASTEXITCODE)."
}

# ---------------------------------------------------------------------------
# 6. Risultato
# ---------------------------------------------------------------------------
$InstallerFile = Get-ChildItem $OutputDir -Filter "*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Write-Host "`n=================================================" -ForegroundColor Cyan
Write-Host "  Installer creato con successo!" -ForegroundColor Green
Write-Host "  $($InstallerFile.FullName)" -ForegroundColor White
Write-Host "  Dimensione: $([math]::Round($InstallerFile.Length / 1MB, 2)) MB" -ForegroundColor White
Write-Host "=================================================" -ForegroundColor Cyan
