# Script per firmare l'eseguibile FatturaView.exe con certificato digitale
# Questo elimina l'avviso "senza una firma valida" in Foxit PDF Reader

param(
    [Parameter(Mandatory=$false)]
    [string]$ExePath = ".\x64\Debug\FatturaView.exe",
    
    [Parameter(Mandatory=$false)]
    [string]$CertPath = "",
    
    [Parameter(Mandatory=$false)]
    [string]$CertPassword = "",
    
    [Parameter(Mandatory=$false)]
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  FatturaView - Firma Eseguibile" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verifica che l'eseguibile esista
if (-not (Test-Path $ExePath)) {
    Write-Host "ERRORE: File non trovato: $ExePath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Compila il progetto prima di firmare l'eseguibile." -ForegroundColor Yellow
    exit 1
}

Write-Host "Eseguibile trovato: $ExePath" -ForegroundColor Green

# Se non è specificato un certificato, cerca nella configurazione di FatturaView
if ([string]::IsNullOrEmpty($CertPath)) {
    $configPath = "$env:APPDATA\FatturaView\pdfsign.cfg"
    if (Test-Path $configPath) {
        Write-Host "Lettura configurazione da: $configPath" -ForegroundColor Yellow
        
        $config = Get-Content $configPath | ConvertFrom-StringData
        
        if ($config.ContainsKey("certPath")) {
            $CertPath = $config["certPath"]
            Write-Host "Certificato configurato: $CertPath" -ForegroundColor Green
        }
        
        if ($config.ContainsKey("certPassword")) {
            $CertPassword = $config["certPassword"]
        }
    }
}

# Se ancora non abbiamo un certificato, chiedi all'utente
if ([string]::IsNullOrEmpty($CertPath)) {
    Write-Host ""
    Write-Host "Nessun certificato configurato." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Per firmare l'eseguibile hai bisogno di un certificato code-signing (.pfx o .p12)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Opzioni:" -ForegroundColor White
    Write-Host "  1. Configura un certificato in FatturaView (Menu → Impostazioni → Firma PDF)" -ForegroundColor White
    Write-Host "  2. Specifica il percorso: .\SignExecutable.ps1 -CertPath 'C:\Cert\MyCert.pfx' -CertPassword 'pwd'" -ForegroundColor White
    Write-Host "  3. Usa signtool.exe manualmente (Windows SDK)" -ForegroundColor White
    Write-Host ""
    
    # Cerca signtool.exe come alternativa
    $signtoolPaths = @(
        "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe",
        "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe",
        "C:\Program Files (x86)\Microsoft SDKs\Windows\*\bin\signtool.exe"
    )
    
    $signtool = $null
    foreach ($path in $signtoolPaths) {
        $found = Get-ChildItem $path -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) {
            $signtool = $found.FullName
            break
        }
    }
    
    if ($signtool) {
        Write-Host "ALTERNATIVA: Trovato signtool.exe" -ForegroundColor Green
        Write-Host "Percorso: $signtool" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Puoi usare:" -ForegroundColor Cyan
        Write-Host "  & '$signtool' sign /f 'C:\Path\To\Cert.pfx' /p 'password' /t $TimestampServer '$ExePath'" -ForegroundColor Gray
    } else {
        Write-Host "signtool.exe non trovato. Installa Windows SDK:" -ForegroundColor Yellow
        Write-Host "  https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/" -ForegroundColor Gray
    }
    
    Write-Host ""
    exit 1
}

# Verifica che il certificato esista
if (-not (Test-Path $CertPath)) {
    Write-Host "ERRORE: Certificato non trovato: $CertPath" -ForegroundColor Red
    exit 1
}

# Verifica che abbiamo la password
if ([string]::IsNullOrEmpty($CertPassword)) {
    Write-Host ""
    $securePassword = Read-Host "Inserisci la password del certificato" -AsSecureString
    $CertPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
    )
}

Write-Host ""
Write-Host "Firma dell'eseguibile in corso..." -ForegroundColor Cyan

# Cerca signtool.exe
$signtoolPaths = @(
    "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe",
    "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe",
    "C:\Program Files (x86)\Microsoft SDKs\Windows\*\bin\*\signtool.exe"
)

$signtool = $null
foreach ($path in $signtoolPaths) {
    $found = Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) {
        $signtool = $found.FullName
        break
    }
}

if (-not $signtool) {
    Write-Host "ERRORE: signtool.exe non trovato!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Installa Windows SDK da:" -ForegroundColor Yellow
    Write-Host "  https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Oppure usa Set-AuthenticodeSignature di PowerShell (meno affidabile):" -ForegroundColor Yellow
    Write-Host ""
    
    # Fallback: Usa Set-AuthenticodeSignature
    try {
        Write-Host "Tentativo con Set-AuthenticodeSignature..." -ForegroundColor Yellow
        
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertPath, $CertPassword)
        Set-AuthenticodeSignature -FilePath $ExePath -Certificate $cert -TimestampServer $TimestampServer
        
        Write-Host ""
        Write-Host "Eseguibile firmato con successo! (metodo PowerShell)" -ForegroundColor Green
        Write-Host ""
        Write-Host "NOTA: La firma con signtool.exe è più affidabile." -ForegroundColor Yellow
        
        exit 0
    } catch {
        Write-Host "ERRORE: $_" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Trovato signtool: $signtool" -ForegroundColor Green
Write-Host ""

# Firma con signtool
try {
    $arguments = @(
        "sign",
        "/f", "`"$CertPath`"",
        "/p", "`"$CertPassword`"",
        "/t", $TimestampServer,
        "/fd", "SHA256",
        "/v",
        "`"$ExePath`""
    )
    
    Write-Host "Esecuzione: signtool $($arguments -join ' ')" -ForegroundColor Gray
    Write-Host ""
    
    $process = Start-Process -FilePath $signtool -ArgumentList $arguments -NoNewWindow -Wait -PassThru
    
    if ($process.ExitCode -eq 0) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  ESEGUIBILE FIRMATO CON SUCCESSO!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "L'eseguibile $ExePath è ora firmato digitalmente." -ForegroundColor White
        Write-Host ""
        Write-Host "Foxit PDF Reader non mostrerà più l'avviso:" -ForegroundColor Cyan
        Write-Host "  'senza una firma valida'" -ForegroundColor Gray
        Write-Host ""
        
        # Verifica la firma
        $sig = Get-AuthenticodeSignature -FilePath $ExePath
        if ($sig.Status -eq "Valid") {
            Write-Host "Verifica firma: OK" -ForegroundColor Green
            Write-Host "Soggetto: $($sig.SignerCertificate.Subject)" -ForegroundColor White
            Write-Host "Validità: $($sig.SignerCertificate.NotBefore) - $($sig.SignerCertificate.NotAfter)" -ForegroundColor White
        } else {
            Write-Host "Verifica firma: $($sig.Status)" -ForegroundColor Yellow
            Write-Host "La firma è presente ma potrebbe non essere trusted (certificato self-signed?)" -ForegroundColor Yellow
        }
        
        Write-Host ""
        exit 0
    } else {
        Write-Host ""
        Write-Host "ERRORE: signtool ha restituito codice di errore $($process.ExitCode)" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host ""
    Write-Host "ERRORE durante la firma: $_" -ForegroundColor Red
    exit 1
}
