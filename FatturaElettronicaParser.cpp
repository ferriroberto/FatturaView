#include "FatturaElettronicaParser.h"
#include <shobjidl.h>
#include <shlwapi.h>
#include <shldisp.h>
#include <atlbase.h>
#include <fstream>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

FatturaElettronicaParser::FatturaElettronicaParser()
    : m_pXmlDoc(nullptr), m_pXslDoc(nullptr)
{
    InitializeCOM();
}

FatturaElettronicaParser::~FatturaElettronicaParser()
{
    CleanupCOM();
}

bool FatturaElettronicaParser::InitializeCOM()
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
    {
        return false;
    }

    hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER,
        __uuidof(IXMLDOMDocument2), (void**)&m_pXmlDoc);
    if (FAILED(hr))
    {
        return false;
    }

    hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER,
        __uuidof(IXMLDOMDocument2), (void**)&m_pXslDoc);
    if (FAILED(hr))
    {
        if (m_pXmlDoc) m_pXmlDoc->Release();
        return false;
    }

    m_pXmlDoc->put_async(VARIANT_FALSE);
    m_pXslDoc->put_async(VARIANT_FALSE);

    return true;
}

void FatturaElettronicaParser::CleanupCOM()
{
    if (m_pXmlDoc)
    {
        m_pXmlDoc->Release();
        m_pXmlDoc = nullptr;
    }
    if (m_pXslDoc)
    {
        m_pXslDoc->Release();
        m_pXslDoc = nullptr;
    }
    CoUninitialize();
}

bool FatturaElettronicaParser::ExtractZipFile(const std::wstring& zipPath, const std::wstring& extractPath)
{
    HRESULT hr = CoInitialize(NULL);
    
    IShellDispatch* pShellDispatch = NULL;
    hr = CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER, 
        IID_IShellDispatch, (void**)&pShellDispatch);
    
    if (FAILED(hr) || !pShellDispatch)
        return false;

    CreateDirectoryW(extractPath.c_str(), NULL);

    VARIANT vZipFile;
    VariantInit(&vZipFile);
    vZipFile.vt = VT_BSTR;
    vZipFile.bstrVal = SysAllocString(zipPath.c_str());

    Folder* pZipFolder = NULL;
    hr = pShellDispatch->NameSpace(vZipFile, &pZipFolder);
    VariantClear(&vZipFile);

    if (FAILED(hr) || !pZipFolder)
    {
        pShellDispatch->Release();
        return false;
    }

    VARIANT vDestFolder;
    VariantInit(&vDestFolder);
    vDestFolder.vt = VT_BSTR;
    vDestFolder.bstrVal = SysAllocString(extractPath.c_str());

    Folder* pDestFolder = NULL;
    hr = pShellDispatch->NameSpace(vDestFolder, &pDestFolder);
    VariantClear(&vDestFolder);

    if (FAILED(hr) || !pDestFolder)
    {
        pZipFolder->Release();
        pShellDispatch->Release();
        return false;
    }

    FolderItems* pItems = NULL;
    hr = pZipFolder->Items(&pItems);
    
    if (SUCCEEDED(hr) && pItems)
    {
        VARIANT vItems;
        VariantInit(&vItems);
        vItems.vt = VT_DISPATCH;
        vItems.pdispVal = pItems;

        VARIANT vOptions;
        VariantInit(&vOptions);
        vOptions.vt = VT_I4;
        vOptions.lVal = 0x14;

        pDestFolder->CopyHere(vItems, vOptions);
        
        Sleep(1000);

        VariantClear(&vItems);
        VariantClear(&vOptions);
        pItems->Release();
    }

    pDestFolder->Release();
    pZipFolder->Release();
    pShellDispatch->Release();

    return true;
}

bool FatturaElettronicaParser::LoadXmlFile(const std::wstring& xmlPath)
{
    if (!m_pXmlDoc)
        return false;

    VARIANT_BOOL status;
    HRESULT hr = m_pXmlDoc->load(_variant_t(xmlPath.c_str()), &status);
    
    return SUCCEEDED(hr) && (status == VARIANT_TRUE);
}

bool FatturaElettronicaParser::ApplyXsltTransform(const std::wstring& xsltPath, std::wstring& htmlOutput)
{
    if (!m_pXmlDoc || !m_pXslDoc)
        return false;

    VARIANT_BOOL status;
    HRESULT hr = m_pXslDoc->load(_variant_t(xsltPath.c_str()), &status);
    
    if (FAILED(hr) || status != VARIANT_TRUE)
        return false;

    BSTR bstrOutput = NULL;
    hr = m_pXmlDoc->transformNode(m_pXslDoc, &bstrOutput);
    
    if (SUCCEEDED(hr) && bstrOutput)
    {
        htmlOutput = bstrOutput;
        SysFreeString(bstrOutput);
        return true;
    }

    return false;
}

bool FatturaElettronicaParser::IsFatturaPA(const std::wstring& xmlPath)
{
    IXMLDOMDocument2* pDoc = NULL;
    HRESULT hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER,
        __uuidof(IXMLDOMDocument2), (void**)&pDoc);
    if (FAILED(hr) || !pDoc)
        return false;

    pDoc->put_async(VARIANT_FALSE);

    VARIANT_BOOL status;
    hr = pDoc->load(_variant_t(xmlPath.c_str()), &status);
    if (FAILED(hr) || status != VARIANT_TRUE)
    {
        pDoc->Release();
        return false;
    }

    IXMLDOMElement* pRoot = NULL;
    hr = pDoc->get_documentElement(&pRoot);
    if (FAILED(hr) || !pRoot)
    {
        pDoc->Release();
        return false;
    }

    BSTR bstrTag = NULL;
    pRoot->get_tagName(&bstrTag);
    std::wstring tag = bstrTag ? std::wstring(bstrTag) : L"";
    SysFreeString(bstrTag);
    pRoot->Release();
    pDoc->Release();

    return tag.find(L"FatturaElettronica") != std::wstring::npos;
}

std::vector<std::wstring> FatturaElettronicaParser::GetXmlFilesFromFolder(const std::wstring& folderPath)
{
    std::vector<std::wstring> xmlFiles;
    WIN32_FIND_DATAW findData;

    std::wstring normalizedPath = folderPath;
    if (!normalizedPath.empty() && normalizedPath.back() == L'\\')
        normalizedPath.pop_back();

    HANDLE hFind = FindFirstFileW((normalizedPath + L"\\*").c_str(), &findData);
    if (hFind != INVALID_HANDLE_VALUE)
    {
        do
        {
            if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
            {
                // Ricorsione nelle sottocartelle (escludi "." e "..")
                if (wcscmp(findData.cFileName, L".") != 0 && wcscmp(findData.cFileName, L"..") != 0)
                {
                    std::wstring subPath = normalizedPath + L"\\" + findData.cFileName;
                    auto subFiles = GetXmlFilesFromFolder(subPath);
                    xmlFiles.insert(xmlFiles.end(), subFiles.begin(), subFiles.end());
                }
            }
            else
            {
                // Accetta file .xml (case-insensitive)
                std::wstring name = findData.cFileName;
                size_t dot = name.find_last_of(L'.');
                if (dot != std::wstring::npos && _wcsicmp(name.c_str() + dot, L".xml") == 0)
                    xmlFiles.push_back(normalizedPath + L"\\" + name);
            }
        } while (FindNextFileW(hFind, &findData));
        FindClose(hFind);
    }

    return xmlFiles;
}

// Metodo helper per estrarre testo da un nodo XML usando XPath
std::wstring FatturaElettronicaParser::GetNodeText(IXMLDOMDocument2* pDoc, const std::wstring& xpath)
{
    if (!pDoc)
        return L"";

    IXMLDOMNode* pNode = NULL;
    BSTR bstrXPath = SysAllocString(xpath.c_str());

    HRESULT hr = pDoc->selectSingleNode(bstrXPath, &pNode);
    SysFreeString(bstrXPath);

    if (SUCCEEDED(hr) && pNode)
    {
        BSTR bstrText = NULL;
        hr = pNode->get_text(&bstrText);
        pNode->Release();

        if (SUCCEEDED(hr) && bstrText)
        {
            std::wstring text = bstrText;
            SysFreeString(bstrText);
            return text;
        }
    }

    return L"";
}

// Estrae le informazioni principali da un file di fattura XML
bool FatturaElettronicaParser::ExtractFatturaInfo(const std::wstring& xmlPath, FatturaInfo& info)
{
    // Crea un documento XML temporaneo per questa operazione
    IXMLDOMDocument2* pTempDoc = NULL;
    HRESULT hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER,
        __uuidof(IXMLDOMDocument2), (void**)&pTempDoc);

    if (FAILED(hr) || !pTempDoc)
        return false;

    pTempDoc->put_async(VARIANT_FALSE);

    // Carica il file XML
    VARIANT_BOOL status;
    hr = pTempDoc->load(_variant_t(xmlPath.c_str()), &status);

    if (FAILED(hr) || status != VARIANT_TRUE)
    {
        pTempDoc->Release();
        return false;
    }

    // Verifica che sia una FatturaPA: il root element deve contenere "FatturaElettronica"
    IXMLDOMElement* pRoot = NULL;
    hr = pTempDoc->get_documentElement(&pRoot);
    if (FAILED(hr) || !pRoot)
    {
        pTempDoc->Release();
        return false;
    }
    BSTR bstrRootTag = NULL;
    pRoot->get_tagName(&bstrRootTag);
    std::wstring rootTag = bstrRootTag ? std::wstring(bstrRootTag) : L"";
    SysFreeString(bstrRootTag);
    pRoot->Release();

    if (rootTag.find(L"FatturaElettronica") == std::wstring::npos)
    {
        pTempDoc->Release();
        return false; // File XML non e' una FatturaPA - ignoralo silenziosamente
    }

    // Estrai le informazioni usando XPath
    info.filePath = xmlPath;

    // Prova prima con namespace FatturaPA v1.2
    pTempDoc->setProperty(_bstr_t(L"SelectionNamespaces"), 
        _variant_t(L"xmlns:p='http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2'"));

    // Cedente/Prestatore (Emittente)
    info.cedenteDenominazione = GetNodeText(pTempDoc, 
        L"//p:FatturaElettronicaHeader/p:CedentePrestatore/p:DatiAnagrafici/p:Anagrafica/p:Denominazione");

    // Se non troviamo nulla, prova senza namespace prefix
    if (info.cedenteDenominazione.empty())
    {
        // Reset del namespace
        pTempDoc->setProperty(_bstr_t(L"SelectionNamespaces"), _variant_t(L""));

        info.cedenteDenominazione = GetNodeText(pTempDoc, 
            L"//FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione");

        // Se non c'è Denominazione, prova con Nome e Cognome
        if (info.cedenteDenominazione.empty())
        {
            std::wstring nome = GetNodeText(pTempDoc, 
                L"//FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Nome");
            std::wstring cognome = GetNodeText(pTempDoc, 
                L"//FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Cognome");

            if (!nome.empty() || !cognome.empty())
                info.cedenteDenominazione = cognome + L" " + nome;
        }

        // Cessionario/Committente (Cliente)
        info.cessionarioDenominazione = GetNodeText(pTempDoc, 
            L"//FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione");

        if (info.cessionarioDenominazione.empty())
        {
            std::wstring nome = GetNodeText(pTempDoc, 
                L"//FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Nome");
            std::wstring cognome = GetNodeText(pTempDoc, 
                L"//FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Cognome");

            if (!nome.empty() || !cognome.empty())
                info.cessionarioDenominazione = cognome + L" " + nome;
        }

        // Numero fattura
        info.numero = GetNodeText(pTempDoc, 
            L"//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Numero");

        // Data fattura
        info.data = GetNodeText(pTempDoc, 
            L"//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data");

        // Importo totale
        info.importoTotale = GetNodeText(pTempDoc, 
            L"//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/ImportoTotaleDocumento");

        // Tipo documento
        info.tipoDocumento = GetNodeText(pTempDoc, 
            L"//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/TipoDocumento");
    }
    else
    {
        // Con namespace - continua normalmente
        // Se non c'è Denominazione, prova con Nome e Cognome
        if (info.cedenteDenominazione.empty())
        {
            std::wstring nome = GetNodeText(pTempDoc, 
                L"//p:FatturaElettronicaHeader/p:CedentePrestatore/p:DatiAnagrafici/p:Anagrafica/p:Nome");
            std::wstring cognome = GetNodeText(pTempDoc, 
                L"//p:FatturaElettronicaHeader/p:CedentePrestatore/p:DatiAnagrafici/p:Anagrafica/p:Cognome");

            if (!nome.empty() || !cognome.empty())
                info.cedenteDenominazione = cognome + L" " + nome;
        }

        // Cessionario/Committente (Cliente)
        info.cessionarioDenominazione = GetNodeText(pTempDoc, 
            L"//p:FatturaElettronicaHeader/p:CessionarioCommittente/p:DatiAnagrafici/p:Anagrafica/p:Denominazione");

        if (info.cessionarioDenominazione.empty())
        {
            std::wstring nome = GetNodeText(pTempDoc, 
                L"//p:FatturaElettronicaHeader/p:CessionarioCommittente/p:DatiAnagrafici/p:Anagrafica/p:Nome");
            std::wstring cognome = GetNodeText(pTempDoc, 
                L"//p:FatturaElettronicaHeader/p:CessionarioCommittente/p:DatiAnagrafici/p:Anagrafica/p:Cognome");

            if (!nome.empty() || !cognome.empty())
                info.cessionarioDenominazione = cognome + L" " + nome;
        }

        // Numero fattura
        info.numero = GetNodeText(pTempDoc, 
            L"//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:Numero");

        // Data fattura
        info.data = GetNodeText(pTempDoc, 
            L"//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:Data");

        // Importo totale
        info.importoTotale = GetNodeText(pTempDoc, 
            L"//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:ImportoTotaleDocumento");

        // Tipo documento
        info.tipoDocumento = GetNodeText(pTempDoc, 
            L"//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:TipoDocumento");
    }

    pTempDoc->Release();

    // Se non siamo riusciti a estrarre almeno il cedente, usa il nome del file
    if (info.cedenteDenominazione.empty())
    {
        size_t pos = xmlPath.find_last_of(L"\\");
        info.cedenteDenominazione = (pos != std::wstring::npos) ? xmlPath.substr(pos + 1) : xmlPath;
    }

    return true;
}

// Ottiene le informazioni di tutte le fatture in una cartella
std::vector<FatturaInfo> FatturaElettronicaParser::GetFattureInfoFromFolder(const std::wstring& folderPath)
{
    std::vector<FatturaInfo> fattureInfo;
    std::vector<std::wstring> xmlFiles = GetXmlFilesFromFolder(folderPath);

    for (const auto& xmlFile : xmlFiles)
    {
        FatturaInfo info;
        if (ExtractFatturaInfo(xmlFile, info))
        {
            fattureInfo.push_back(info);
        }
    }

    return fattureInfo;
}
