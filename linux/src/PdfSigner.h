#pragma once
#include <string>

// ---------------------------------------------------------------------------
// Struttura di configurazione firma PDF — identica alla versione Windows
// ---------------------------------------------------------------------------
struct PdfSignConfig
{
    bool        enabled      = false;
    std::string certPath;
    std::string certPassword;
    std::string signReason   = "Fattura Elettronica Italiana";
    std::string signLocation = "Italia";
    std::string signContact;
};

// ---------------------------------------------------------------------------
// PdfSigner (Linux)
// Config: ~/.config/FatturaView/pdfsign.cfg  (stesso formato key=value)
// Firma:  java + iTextSharp  oppure  python3 + pyhanko
// ---------------------------------------------------------------------------
class PdfSigner
{
public:
    static bool LoadConfig(PdfSignConfig& config);
    static bool SaveConfig(const PdfSignConfig& config);

    static bool SignPdf(const std::string& pdfPath,
                        const PdfSignConfig& config,
                        std::string& errorMsg);

    static bool CheckPrerequisites(std::string& errorMsg);

private:
    static std::string GetConfigFilePath();
};
