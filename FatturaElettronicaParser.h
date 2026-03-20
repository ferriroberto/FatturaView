#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include <msxml6.h>
#include <comdef.h>

#pragma comment(lib, "msxml6.lib")

// Struttura per memorizzare i dati principali di una fattura
struct FatturaInfo
{
    std::wstring filePath;              // Percorso completo del file XML
    std::wstring cedenteDenominazione;  // Nome emittente/fornitore
    std::wstring cessionarioDenominazione; // Nome cliente/committente
    std::wstring numero;                // Numero fattura
    std::wstring data;                  // Data fattura
    std::wstring importoTotale;         // Importo totale
    std::wstring tipoDocumento;         // Tipo documento (TD01, TD04, ecc.)

    // Metodo per formattare la visualizzazione (numero e data prima, poi cessionario)
    std::wstring GetDisplayText() const
    {
        // Se non abbiamo dati, mostra almeno il nome del file
        if (numero.empty() && data.empty() && cessionarioDenominazione.empty())
        {
            size_t pos = filePath.find_last_of(L"\\");
            return (pos != std::wstring::npos) ? filePath.substr(pos + 1) : filePath;
        }

        std::wstring text = L"  ";  // Indentazione per mostrare che è sotto un gruppo

        // Prima: numero fattura
        if (!numero.empty())
            text += L"N. " + numero;

        // Poi: data
        if (!data.empty())
        {
            if (!numero.empty())
                text += L" del ";
            text += data;
        }

        // Poi: cessionario (cliente) - verrà reso in grassetto nel rendering
        if (!cessionarioDenominazione.empty())
        {
            if (!numero.empty() || !data.empty())
                text += L" | ";
            text += cessionarioDenominazione;
        }

        // Infine: importo
        if (!importoTotale.empty())
            text += L" | \x20AC " + importoTotale;  // \x20AC = simbolo Euro

        return text;
    }
};

class FatturaElettronicaParser
{
public:
    FatturaElettronicaParser();
    ~FatturaElettronicaParser();

    bool ExtractZipFile(const std::wstring& zipPath, const std::wstring& extractPath);
    bool LoadXmlFile(const std::wstring& xmlPath);
    bool ApplyXsltTransform(const std::wstring& xsltPath, std::wstring& htmlOutput);
    std::vector<std::wstring> GetXmlFilesFromFolder(const std::wstring& folderPath);

    // Nuovi metodi per estrarre informazioni dalle fatture
    bool ExtractFatturaInfo(const std::wstring& xmlPath, FatturaInfo& info);
    std::vector<FatturaInfo> GetFattureInfoFromFolder(const std::wstring& folderPath);

private:
    IXMLDOMDocument2* m_pXmlDoc;
    IXMLDOMDocument2* m_pXslDoc;

    bool InitializeCOM();
    void CleanupCOM();
    std::wstring GetNodeText(IXMLDOMDocument2* pDoc, const std::wstring& xpath);
};
