#pragma once
#include <string>
#include <vector>
#include <filesystem>

// ---------------------------------------------------------------------------
// Struttura dati fattura — identica alla versione Windows ma con std::string
// (UTF-8 invece di UTF-16 wstring)
// ---------------------------------------------------------------------------
struct FatturaInfo
{
    std::string filePath;
    std::string cedenteDenominazione;
    std::string cessionarioDenominazione;
    std::string numero;
    std::string data;
    std::string importoTotale;
    std::string tipoDocumento;

    std::string GetDisplayText() const
    {
        if (numero.empty() && data.empty() && cessionarioDenominazione.empty())
            return std::filesystem::path(filePath).filename().string();

        std::string text = "  ";
        if (!numero.empty())
            text += "N. " + numero;
        if (!data.empty())
        {
            if (!numero.empty()) text += " del ";
            text += data;
        }
        if (!cessionarioDenominazione.empty())
        {
            if (!numero.empty() || !data.empty()) text += " | ";
            text += cessionarioDenominazione;
        }
        if (!importoTotale.empty())
            text += " | \u20AC " + importoTotale;
        return text;
    }
};

// ---------------------------------------------------------------------------
// FatturaParser — sostituisce FatturaElettronicaParser (Windows/MSXML)
// Backend: libxml2 + libxslt + libzip
// ---------------------------------------------------------------------------
class FatturaParser
{
public:
    FatturaParser();
    ~FatturaParser();

    // Estrae un archivio ZIP nella cartella indicata
    bool ExtractZipFile(const std::string& zipPath, const std::string& extractPath);

    // Carica un file XML in memoria
    bool LoadXmlFile(const std::string& xmlPath);

    // Applica una trasformazione XSLT al documento caricato e restituisce l'HTML
    bool ApplyXsltTransform(const std::string& xsltPath, std::string& htmlOutput);

    // Restituisce tutti i file .xml in una cartella (ricorsivo)
    std::vector<std::string> GetXmlFilesFromFolder(const std::string& folderPath);

    // Estrae le informazioni principali da un file XML di fattura
    bool ExtractFatturaInfo(const std::string& xmlPath, FatturaInfo& info);

    // Restituisce le informazioni di tutte le fatture in una cartella
    std::vector<FatturaInfo> GetFattureInfoFromFolder(const std::string& folderPath);

    // Verifica se il file XML è una FatturaPA
    bool IsFatturaPA(const std::string& xmlPath);

private:
    // Pimpl: nasconde i tipi di libxml2 dall'header
    struct XmlImpl;
    XmlImpl* m_impl;

    // Helper: estrae il testo di un nodo XPath su un documento già caricato
    static std::string GetNodeTextFromDoc(void* doc,
                                          const std::string& xpath,
                                          const std::string& nsPrefix = "",
                                          const std::string& nsUri   = "");
};
