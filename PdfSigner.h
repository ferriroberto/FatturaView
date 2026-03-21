#pragma once
#include <windows.h>
#include <string>

// Struttura per le configurazioni della firma PDF
struct PdfSignConfig
{
    bool enabled;
    std::wstring certPath;
    std::wstring certPassword;
    std::wstring signReason;
    std::wstring signLocation;
    std::wstring signContact;

    PdfSignConfig()
        : enabled(false)
    {
        signReason = L"Fattura Elettronica Italiana";
        signLocation = L"Italia";
        signContact = L"";
    }
};

class PdfSigner
{
public:
    // Carica la configurazione dal file
    static bool LoadConfig(PdfSignConfig& config);
    
    // Salva la configurazione nel file
    static bool SaveConfig(const PdfSignConfig& config);
    
    // Firma un PDF usando PowerShell e iTextSharp
    static bool SignPdf(const std::wstring& pdfPath, const PdfSignConfig& config, std::wstring& errorMsg);
    
    // Verifica se PowerShell e .NET sono disponibili
    static bool CheckPrerequisites(std::wstring& errorMsg);
    
    // Crea lo script PowerShell per la firma
    static std::wstring CreateSigningScript(const std::wstring& pdfPath, const PdfSignConfig& config);

private:
    static std::wstring GetConfigFilePath();
    static std::wstring EscapePowerShellString(const std::wstring& str);
};
