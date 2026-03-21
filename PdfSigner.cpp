#include "PdfSigner.h"
#include <fstream>
#include <shlobj.h>
#include <atlbase.h>
#include <comutil.h>

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comsuppw.lib")

std::wstring PdfSigner::GetConfigFilePath()
{
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_APPDATA, NULL, 0, path)))
    {
        std::wstring configPath = path;
        configPath += L"\\FatturaView";
        CreateDirectoryW(configPath.c_str(), NULL);
        configPath += L"\\pdfsign.cfg";
        return configPath;
    }
    return L"pdfsign.cfg";
}

bool PdfSigner::LoadConfig(PdfSignConfig& config)
{
    std::wstring configFile = GetConfigFilePath();
    std::wifstream file(configFile);
    
    if (!file.is_open())
        return false;

    std::wstring line;
    while (std::getline(file, line))
    {
        size_t pos = line.find(L'=');
        if (pos == std::wstring::npos)
            continue;

        std::wstring key = line.substr(0, pos);
        std::wstring value = line.substr(pos + 1);

        if (key == L"enabled")
            config.enabled = (value == L"1" || value == L"true");
        else if (key == L"certPath")
            config.certPath = value;
        else if (key == L"certPassword")
            config.certPassword = value;
        else if (key == L"signReason")
            config.signReason = value;
        else if (key == L"signLocation")
            config.signLocation = value;
        else if (key == L"signContact")
            config.signContact = value;
    }

    file.close();
    return true;
}

bool PdfSigner::SaveConfig(const PdfSignConfig& config)
{
    std::wstring configFile = GetConfigFilePath();
    std::wofstream file(configFile);
    
    if (!file.is_open())
        return false;

    file << L"enabled=" << (config.enabled ? L"1" : L"0") << L"\n";
    file << L"certPath=" << config.certPath << L"\n";
    file << L"certPassword=" << config.certPassword << L"\n";
    file << L"signReason=" << config.signReason << L"\n";
    file << L"signLocation=" << config.signLocation << L"\n";
    file << L"signContact=" << config.signContact << L"\n";

    file.close();
    return true;
}

bool PdfSigner::CheckPrerequisites(std::wstring& errorMsg)
{
    // Verifica PowerShell
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\PowerShell\\3", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
    {
        errorMsg = L"PowerShell 3.0 o superiore non trovato";
        return false;
    }
    RegCloseKey(hKey);

    return true;
}

std::wstring PdfSigner::EscapePowerShellString(const std::wstring& str)
{
    std::wstring escaped = str;
    size_t pos = 0;
    while ((pos = escaped.find(L'"', pos)) != std::wstring::npos)
    {
        escaped.replace(pos, 1, L"`\"");
        pos += 2;
    }
    return escaped;
}

std::wstring PdfSigner::CreateSigningScript(const std::wstring& pdfPath, const PdfSignConfig& config)
{
    std::wstring script = LR"(
# Script per firmare PDF con iTextSharp
$ErrorActionPreference = "Stop"

try {
    # Carica iTextSharp
    Add-Type -Path "$env:TEMP\itextsharp.dll" -ErrorAction SilentlyContinue
    
    # Se non esiste, scaricalo
    if (-not ([System.Management.Automation.PSTypeName]'iTextSharp.text.pdf.PdfReader').Type) {
        $itextUrl = "https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3"
        $tempZip = "$env:TEMP\itext.zip"
        $extractPath = "$env:TEMP\itext"
        
        Write-Host "Download iTextSharp..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $itextUrl -OutFile $tempZip
        
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
        Expand-Archive -Path $tempZip -DestinationPath $extractPath -Force
        
        Copy-Item "$extractPath\lib\itextsharp.dll" "$env:TEMP\itextsharp.dll" -Force
        Add-Type -Path "$env:TEMP\itextsharp.dll"
    }

    $pdfPath = ")";
    script += EscapePowerShellString(pdfPath);
    script += LR"("
    $certPath = ")";
    script += EscapePowerShellString(config.certPath);
    script += LR"("
    $certPass = ")";
    script += EscapePowerShellString(config.certPassword);
    script += LR"("
    $reason = ")";
    script += EscapePowerShellString(config.signReason);
    script += LR"("
    $location = ")";
    script += EscapePowerShellString(config.signLocation);
    script += LR"("
    $contact = ")";
    script += EscapePowerShellString(config.signContact);
    script += LR"("

    # Carica il certificato
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath, $certPass)

    # Crea il file temporaneo firmato
    $signedPdf = $pdfPath -replace '\.pdf$', '_signed.pdf'

    # Apri il PDF originale
    $reader = New-Object iTextSharp.text.pdf.PdfReader($pdfPath)
    $stream = New-Object System.IO.FileStream($signedPdf, [System.IO.FileMode]::Create)

    # Crea il PdfStamper per la firma
    $stamper = [iTextSharp.text.pdf.PdfStamper]::CreateSignature($reader, $stream, [char]0, $null, $true)

    # Aggiungi metadati al PDF per identificare l'applicazione
    $info = $stamper.MoreInfo
    $info["Producer"] = "FatturaView v1.0.0 - Roberto Ferri"
    $info["Creator"] = "FatturaView - Lettore Fatture Elettroniche PA"
    $info["Author"] = $contact
    if (-not $info.ContainsKey("Title")) {
        $info["Title"] = "Fattura Elettronica Italiana"
    }

    # Configura l'aspetto della firma
    $appearance = $stamper.SignatureAppearance
    $appearance.Reason = $reason
    $appearance.Location = $location
    $appearance.Contact = $contact
    $appearance.CertificationLevel = [iTextSharp.text.pdf.PdfSignatureAppearance+CertificationLevel]::NOT_CERTIFIED
    $appearance.SetVisibleSignature((New-Object iTextSharp.text.Rectangle(0, 0, 0, 0)), 1, "FatturaViewSignature")

    # Crea la firma con catena di certificazione completa
    $pk = [Org.BouncyCastle.Pkcs.Pkcs12Store]::new()
    $pkStream = [System.IO.File]::OpenRead($certPath)
    $pk.Load($pkStream, $certPass.ToCharArray())
    $pkStream.Close()

    $alias = $null
    foreach ($a in $pk.Aliases) {
        if ($pk.IsKeyEntry($a)) {
            $alias = $a
            break
        }
    }

    if ($null -eq $alias) {
        throw "Nessuna chiave privata trovata nel certificato"
    }

    $privateKey = $pk.GetKey($alias).Key
    $chain = $pk.GetCertificateChain($alias)

    # Converti in array di certificati per iTextSharp
    $certChain = @()
    foreach ($c in $chain) {
        $certChain += $c.Certificate
    }

    # Crea il digest esterno per la firma
    $digest = [iTextSharp.text.pdf.security.MakeSignature+CryptoStandard]::CMS

    # Firma il documento con tutti i parametri corretti
    [iTextSharp.text.pdf.security.MakeSignature]::SignDetached(
        $appearance,
        $privateKey,
        $certChain,
        $null,  # CRL (Certificate Revocation List)
        $null,  # OCSP (Online Certificate Status Protocol)
        $null,  # TSA (Time Stamp Authority)
        0,      # Estimated size
        $digest
    )
    
    # Chiudi gli stream
    $stamper.Close()
    $reader.Close()
    $stream.Close()
    
    # Sostituisci il file originale
    Remove-Item $pdfPath -Force
    Rename-Item $signedPdf $pdfPath -Force
    
    Write-Host "PDF firmato con successo"
    exit 0
    
} catch {
    Write-Error "Errore durante la firma: $_"
    exit 1
}
)";

    return script;
}

bool PdfSigner::SignPdf(const std::wstring& pdfPath, const PdfSignConfig& config, std::wstring& errorMsg)
{
    if (!config.enabled)
    {
        return true; // Non è un errore, semplicemente la firma è disabilitata
    }

    // Verifica prerequisiti
    if (!CheckPrerequisites(errorMsg))
    {
        return false;
    }

    // Verifica che il certificato esista
    if (GetFileAttributesW(config.certPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        errorMsg = L"Certificato non trovato: " + config.certPath;
        return false;
    }

    // Verifica che il PDF esista
    if (GetFileAttributesW(pdfPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        errorMsg = L"File PDF non trovato: " + pdfPath;
        return false;
    }

    // Crea lo script PowerShell
    std::wstring script = CreateSigningScript(pdfPath, config);
    
    // Salva lo script in un file temporaneo
    wchar_t tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempPath);
    std::wstring scriptPath = std::wstring(tempPath) + L"sign_pdf.ps1";
    
    std::wofstream scriptFile(scriptPath);
    if (!scriptFile.is_open())
    {
        errorMsg = L"Impossibile creare lo script temporaneo";
        return false;
    }
    scriptFile << script;
    scriptFile.close();

    // Esegui lo script PowerShell
    std::wstring psCommand = L"powershell.exe -ExecutionPolicy Bypass -NoProfile -File \"" + scriptPath + L"\"";
    
    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = { 0 };
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    if (!CreateProcessW(NULL, (LPWSTR)psCommand.c_str(), NULL, NULL, FALSE,
        CREATE_NO_WINDOW, NULL, NULL, &si, &pi))
    {
        errorMsg = L"Impossibile avviare PowerShell";
        DeleteFileW(scriptPath.c_str());
        return false;
    }

    // Attendi il completamento
    WaitForSingleObject(pi.hProcess, 60000); // Timeout 60 secondi

    DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);

    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    // Elimina lo script temporaneo
    DeleteFileW(scriptPath.c_str());

    if (exitCode != 0)
    {
        errorMsg = L"Errore durante la firma del PDF (codice: " + std::to_wstring(exitCode) + L")";
        return false;
    }

    return true;
}
