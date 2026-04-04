#include "FatturaElettronicaParser.h"
#include <shobjidl.h>
#include <shlwapi.h>
#include <shldisp.h>
#include <atlbase.h>
#include <fstream>
#include <wincrypt.h>
#include <algorithm>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "crypt32.lib")

// Funzione helper per decodificare un file P7M e salvare il contenuto XML in un file temporaneo.
// Restituisce il percorso del file temporaneo, o una stringa vuota in caso di errore.
std::wstring FatturaElettronicaParser::DecodeP7M(const std::wstring& p7mPath)
{
    HCRYPTMSG hMsg = NULL;
    HCERTSTORE hStore = NULL;
    DWORD dwEncoding, dwContentType, dwFormatType;

    // 1. Prova prima con CryptQueryObject, che è la funzione più completa per tipi diversi (DER, PEM, ecc.)
    BOOL success = CryptQueryObject(
        CERT_QUERY_OBJECT_FILE,
        p7mPath.c_str(),
        CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED | CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED_EMBED | CERT_QUERY_CONTENT_FLAG_ALL,
        CERT_QUERY_FORMAT_FLAG_ALL,
        0,
        &dwEncoding,
        &dwContentType,
        &dwFormatType,
        &hStore,
        &hMsg,
        NULL);

    // 2. Se CryptQueryObject fallisce o non restituisce hMsg (accade con certi file Base64 senza header),
    // proviamo una decodifica manuale leggendo il file come BLOB.
    if (!success || !hMsg)
    {
        if (hStore) { CertCloseStore(hStore, 0); hStore = NULL; }
        if (hMsg) { CryptMsgClose(hMsg); hMsg = NULL; }

        std::ifstream p7mIn(p7mPath, std::ios::binary);
        if (!p7mIn) return L"";
        std::vector<BYTE> buffer((std::istreambuf_iterator<char>(p7mIn)), std::istreambuf_iterator<char>());
        p7mIn.close();
        if (buffer.empty()) return L"";

        // Tentativo di conversione se è Base64 (anche senza intestazioni PEM)
        DWORD dwDecodedSize = 0;
        if (CryptStringToBinaryA((LPCSTR)buffer.data(), (DWORD)buffer.size(), CRYPT_STRING_BASE64_ANY, NULL, &dwDecodedSize, NULL, NULL))
        {
            std::vector<BYTE> decoded(dwDecodedSize);
            if (CryptStringToBinaryA((LPCSTR)buffer.data(), (DWORD)buffer.size(), CRYPT_STRING_BASE64_ANY, decoded.data(), &dwDecodedSize, NULL, NULL))
            {
                buffer = std::move(decoded);
            }
        }

        hMsg = CryptMsgOpenToDecode(X509_ASN_ENCODING | PKCS_7_ASN_ENCODING, 0, 0, NULL, NULL, NULL);
        if (hMsg)
        {
            if (!CryptMsgUpdate(hMsg, buffer.data(), (DWORD)buffer.size(), TRUE))
            {
                CryptMsgClose(hMsg);
                hMsg = NULL;
            }
        }
    }

    if (!hMsg) return L"";

    // 3. Ottieni la dimensione del contenuto della fattura (XML originale)
    DWORD cbContent = 0;
    if (!CryptMsgGetParam(hMsg, CMSG_CONTENT_PARAM, 0, NULL, &cbContent) || cbContent == 0)
    {
        CryptMsgClose(hMsg);
        if (hStore) CertCloseStore(hStore, 0);
        return L"";
    }

    // 4. Estrai i dati XML
    std::vector<BYTE> content(cbContent);
    if (!CryptMsgGetParam(hMsg, CMSG_CONTENT_PARAM, 0, content.data(), &cbContent))
    {
        CryptMsgClose(hMsg);
        if (hStore) CertCloseStore(hStore, 0);
        return L"";
    }

    CryptMsgClose(hMsg);
    if (hStore) CertCloseStore(hStore, 0);

    // 4b. Se il contenuto estratto non sembra XML ma Base64, prova a decodificarlo.
    // Accade se la fattura è stata imbustata come stringa Base64 nel P7M.
    bool looksLikeBase64 = true;
    for (size_t i = 0; i < content.size() && i < 100; ++i) {
        if (content[i] == '<') {
            looksLikeBase64 = false;
            break;
        }
    }
    if (looksLikeBase64 && content.size() > 10) {
        DWORD dwSize = 0;
        if (CryptStringToBinaryA((LPCSTR)content.data(), (DWORD)content.size(), CRYPT_STRING_BASE64_ANY, NULL, &dwSize, NULL, NULL)) {
            std::vector<BYTE> decoded(dwSize);
            if (CryptStringToBinaryA((LPCSTR)content.data(), (DWORD)content.size(), CRYPT_STRING_BASE64_ANY, decoded.data(), &dwSize, NULL, NULL)) {
                content = std::move(decoded);
            }
        }
    }

    // 5. Salva il contenuto in un file temporaneo
    wchar_t tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempPath);
    wchar_t tempFileName[MAX_PATH];
    GetTempFileNameW(tempPath, L"FV", 0, tempFileName);

    // Forza l'estensione .xml (alcune configurazioni MSXML6 in Release lo richiedono)
    std::wstring finalTemp = tempFileName;
    finalTemp += L".xml";
    if (!MoveFileW(tempFileName, finalTemp.c_str()))
        finalTemp = tempFileName;

    std::ofstream tempStream(finalTemp, std::ios::binary);
    if (!tempStream) return L"";
    tempStream.write(reinterpret_cast<const char*>(content.data()), content.size());
    tempStream.close();

    return finalTemp;
}

// Funzione helper per caricare un file XML in un documento DOM tramite IStream
// per una gestione più robusta degli encoding.
bool FatturaElettronicaParser::LoadXmlFileIntoDoc(const std::wstring& xmlPath, IXMLDOMDocument2* pDoc)
{
    if (!pDoc) return false;

    std::ifstream xmlStream(xmlPath, std::ios::binary);
    if (!xmlStream) {
        return false;
    }
    std::vector<char> buffer((std::istreambuf_iterator<char>(xmlStream)), std::istreambuf_iterator<char>());
    xmlStream.close();

    if (buffer.empty()) return false;

    // --- SANITIZZAZIONE XML ---
    // 1. Rimuovi byte nulli
    buffer.erase(std::remove(buffer.begin(), buffer.end(), '\0'), buffer.end());

    // 2. Rimuovi caratteri di controllo illegali (tranne TAB, LF, CR)
    for (auto& c : buffer) {
        unsigned char uc = (unsigned char)c;
        if (uc < 0x20 && uc != 0x09 && uc != 0x0A && uc != 0x0D) {
            c = ' ';
        }
    }

    // 3. Fix per rogue '<' segnalati (es. '< ' o '<' seguito da caratteri spazzatura non ASCII)
    // In XML i tag validi iniziano con lettere, /, ?, !, _, :
    for (size_t i = 0; i + 1 < buffer.size(); ++i) {
        if (buffer[i] == '<') {
            unsigned char next = (unsigned char)buffer[i + 1];
            // Se non è l'inizio di un tag valido
            if (next != '/' && next != '?' && next != '!' && next != '_' && next != ':' &&
                !((next >= 'A' && next <= 'Z') || (next >= 'a' && next <= 'z'))) {
                buffer[i] = ' '; // Neutralizza il '<' orfano
            }
        }
    }

    IStream* pStream = SHCreateMemStream((const BYTE*)buffer.data(), static_cast<UINT>(buffer.size()));
    if (!pStream) {
        return false;
    }

    VARIANT vStream;
    VariantInit(&vStream);
    vStream.vt = VT_UNKNOWN;
    vStream.punkVal = pStream;

    VARIANT_BOOL status;
    HRESULT hr = pDoc->load(vStream, &status);

    VariantClear(&vStream); // This will call pStream->Release()

    return SUCCEEDED(hr) && (status == VARIANT_TRUE);
}

FatturaElettronicaParser::FatturaElettronicaParser()
    : m_pXmlDoc(nullptr), m_pXslDoc(nullptr)
{
    InitializeCOM();
}

bool FatturaElettronicaParser::HasAttachments(const std::wstring& xmlPath)
{
    IXMLDOMDocument2* pDoc = NULL;
    HRESULT hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER,
        __uuidof(IXMLDOMDocument2), (void**)&pDoc);
    if (FAILED(hr) || !pDoc)
        return false;

    pDoc->put_async(VARIANT_FALSE);
    pDoc->setProperty(_bstr_t(L"SelectionLanguage"), _variant_t(L"XPath"));
    pDoc->setProperty(_bstr_t(L"SelectionNamespaces"), _variant_t(L""));

    if (!LoadXmlFileIntoDoc(xmlPath, pDoc))
    {
        pDoc->Release();
        return false;
    }

    IXMLDOMNodeList* pList = NULL;
    BSTR bstrQ = SysAllocString(L"//*[local-name()='Allegati']");
    hr = pDoc->selectNodes(bstrQ, &pList);
    SysFreeString(bstrQ);

    bool has = false;
    if (SUCCEEDED(hr) && pList)
    {
        long count = 0;
        pList->get_length(&count);
        has = (count > 0);
        pList->Release();
    }

    pDoc->Release();
    return has;
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
    if (!m_lastTempXmlPath.empty())
    {
        _wremove(m_lastTempXmlPath.c_str());
        m_lastTempXmlPath.clear();
    }
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
    IShellDispatch* pShellDispatch = NULL;
    HRESULT hr = CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER, 
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

        // Attendi il completamento dell'estrazione con un timeout.
        // La logica precedente riutilizzava la variabile 'pItems' in modo errato,
        // causando memory leak e comportamento non definito.
        // La nuova logica usa una variabile separata per gli elementi di destinazione
        // e confronta il numero di file estratti con il numero di file sorgente.
        long sourceItemCount = 0;
        pItems->get_Count(&sourceItemCount);

        pDestFolder->CopyHere(vItems, vOptions);

        const int timeoutSeconds = 30; // Timeout di 30 secondi
        for (int i = 0; i < timeoutSeconds * 2; i++) // Controlla ogni 500ms
        {
            FolderItems* pDestItems = NULL;
            long destItemCount = 0;

            if (SUCCEEDED(pDestFolder->Items(&pDestItems)) && pDestItems)
            {
                pDestItems->get_Count(&destItemCount);
                pDestItems->Release();
            }

            // Si assume che l'estrazione sia completa quando il numero di elementi
            // nella destinazione è uguale o superiore a quello della sorgente.
            if (sourceItemCount > 0 && destItemCount >= sourceItemCount)
            {
                Sleep(500); // Attesa extra per completare la scrittura dei file
                break;
            }

            // Se la sorgente è vuota, basta attendere che la cartella di destinazione non sia vuota
            // (potrebbe essere creata una cartella vuota).
            if (sourceItemCount == 0 && destItemCount > 0)
            {
                break;
            }

            Sleep(500);
        }

        VariantClear(&vItems); // Rilascia pItems, che era contenuto nel variant
        VariantClear(&vOptions);
        // pItems->Release() non è necessario qui perché VariantClear lo ha già fatto.
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

    // Se un file temporaneo era aperto, cancellalo
    if (!m_lastTempXmlPath.empty()) {
        _wremove(m_lastTempXmlPath.c_str());
        m_lastTempXmlPath.clear();
    }

    std::wstring pathToLoad = xmlPath;

    // Se è un file P7M, decodificalo
    size_t dot = xmlPath.find_last_of(L'.');
    if (dot != std::wstring::npos && _wcsicmp(xmlPath.c_str() + dot, L".p7m") == 0)
    {
        m_lastTempXmlPath = DecodeP7M(xmlPath);
        if (m_lastTempXmlPath.empty())
            return false;

        pathToLoad = m_lastTempXmlPath;
    }

    return LoadXmlFileIntoDoc(pathToLoad, m_pXmlDoc);
}

bool FatturaElettronicaParser::ApplyXsltTransform(const std::wstring& xsltPath, std::wstring& htmlOutput)
{
    if (!m_pXmlDoc || !m_pXslDoc)
        return false;

    if (!LoadXmlFileIntoDoc(xsltPath, m_pXslDoc))
        return false;

    BSTR bstrOutput = NULL;
    HRESULT hr = m_pXmlDoc->transformNode(m_pXslDoc, &bstrOutput);

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

    if (!LoadXmlFileIntoDoc(xmlPath, pDoc))
    {
        pDoc->Release();
        return false;
    }

    // Usa una query XPath indipendente dal namespace per trovare l'elemento radice
    pDoc->setProperty(_bstr_t(L"SelectionLanguage"), _variant_t(L"XPath"));

    IXMLDOMNode* pNode = NULL;
    BSTR bstrXPath = SysAllocString(L"/*[local-name()='FatturaElettronica']");
    hr = pDoc->selectSingleNode(bstrXPath, &pNode);
    SysFreeString(bstrXPath);

    bool isValid = (SUCCEEDED(hr) && pNode != NULL);
    if (pNode) pNode->Release();

    pDoc->Release();

    return isValid;
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
                // Accetta file .xml e .p7m (case-insensitive)
                std::wstring name = findData.cFileName;
                size_t dot = name.find_last_of(L'.');
                if (dot != std::wstring::npos)
                {
                    const wchar_t* ext = name.c_str() + dot;
                    if (_wcsicmp(ext, L".xml") == 0 || _wcsicmp(ext, L".p7m") == 0)
                        xmlFiles.push_back(normalizedPath + L"\\" + name);
                }
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

    if (!LoadXmlFileIntoDoc(xmlPath, pTempDoc))
    {
        pTempDoc->Release();
        return false;
    }

    // Verifica che sia una FatturaPA: il root element deve essere "FatturaElettronica"
    IXMLDOMElement* pRoot = NULL;
    hr = pTempDoc->get_documentElement(&pRoot);
    if (FAILED(hr) || !pRoot)
    {
        pTempDoc->Release();
        return false;
    }
    BSTR bstrBaseName = NULL;
    hr = pRoot->get_baseName(&bstrBaseName);
    std::wstring baseName = bstrBaseName ? std::wstring(bstrBaseName) : L"";
    SysFreeString(bstrBaseName);
    pRoot->Release();

    if (baseName != L"FatturaElettronica")
    {
        pTempDoc->Release();
        return false; // File XML non e' una FatturaPA - ignoralo silenziosamente
    }

    // Estrai le informazioni usando XPath agnostico rispetto al namespace (usando local-name())
    info.filePath = xmlPath;

    // Imposta XPath come linguaggio di selezione
    pTempDoc->setProperty(_bstr_t(L"SelectionLanguage"), _variant_t(L"XPath"));
    pTempDoc->setProperty(_bstr_t(L"SelectionNamespaces"), _variant_t(L""));

    // Helper lambda per abbreviare la query con local-name
    auto getInfo = [&](const std::wstring& path) {
        return GetNodeText(pTempDoc, path);
    };

    // Cedente/Prestatore (Emittente)
    info.cedenteDenominazione = getInfo(L"//*[local-name()='CedentePrestatore']//*[local-name()='Denominazione']");

    // Se non c'è Denominazione, prova con Nome e Cognome
    if (info.cedenteDenominazione.empty())
    {
        std::wstring nome = getInfo(L"//*[local-name()='CedentePrestatore']//*[local-name()='Nome']");
        std::wstring cognome = getInfo(L"//*[local-name()='CedentePrestatore']//*[local-name()='Cognome']");

        if (!nome.empty() || !cognome.empty())
        {
            info.cedenteDenominazione = cognome + (cognome.empty() ? L"" : L" ") + nome;
        }
    }

    // Cessionario/Committente (Cliente)
    info.cessionarioDenominazione = getInfo(L"//*[local-name()='CessionarioCommittente']//*[local-name()='Denominazione']");

    if (info.cessionarioDenominazione.empty())
    {
        std::wstring nome = getInfo(L"//*[local-name()='CessionarioCommittente']//*[local-name()='Nome']");
        std::wstring cognome = getInfo(L"//*[local-name()='CessionarioCommittente']//*[local-name()='Cognome']");

        if (!nome.empty() || !cognome.empty())
        {
            info.cessionarioDenominazione = cognome + (cognome.empty() ? L"" : L" ") + nome;
        }
    }

    // Numero fattura
    info.numero = getInfo(L"//*[local-name()='DatiGeneraliDocumento']/*[local-name()='Numero']");

    // Data fattura
    info.data = getInfo(L"//*[local-name()='DatiGeneraliDocumento']/*[local-name()='Data']");

    // Importo totale
    info.importoTotale = getInfo(L"//*[local-name()='DatiGeneraliDocumento']/*[local-name()='ImportoTotaleDocumento']");

    // Tipo documento
    info.tipoDocumento = getInfo(L"//*[local-name()='DatiGeneraliDocumento']/*[local-name()='TipoDocumento']");

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
    std::vector<std::wstring> files = GetXmlFilesFromFolder(folderPath);
    std::vector<std::wstring> tempFilesToDelete;

    for (const auto& filePath : files)
    {
        std::wstring xmlPathToProcess = filePath;

        // Se è un file P7M, decodificalo in un file XML temporaneo
        size_t dot = filePath.find_last_of(L'.');
        if (dot != std::wstring::npos && _wcsicmp(filePath.c_str() + dot, L".p7m") == 0)
        {
            xmlPathToProcess = DecodeP7M(filePath);
            if (xmlPathToProcess.empty())
                continue; // Errore di decodifica, salta il file

            tempFilesToDelete.push_back(xmlPathToProcess);
        }

        FatturaInfo info;
        if (ExtractFatturaInfo(xmlPathToProcess, info))
        {
            info.filePath = filePath; // Salva il percorso del file originale (.xml o .p7m)
            info.hasAllegati = HasAttachments(xmlPathToProcess);
            fattureInfo.push_back(info);
        }
    }

    // Pulisci i file temporanei creati
    for (const auto& tempFile : tempFilesToDelete)
    {
        _wremove(tempFile.c_str());
    }

    return fattureInfo;
}

std::vector<AllegatoInfo> FatturaElettronicaParser::ExtractAttachments(
    const std::wstring& filePath, const std::wstring& outputFolder)
{
    std::wstring xmlPathToProcess = filePath;
    std::wstring tempFileToDelete;
    bool isTempFile = false;

    // Se è un file P7M, decodificalo prima
    size_t dot = filePath.find_last_of(L'.');
    if (dot != std::wstring::npos && _wcsicmp(filePath.c_str() + dot, L".p7m") == 0)
    {
        xmlPathToProcess = DecodeP7M(filePath);
        if (xmlPathToProcess.empty())
            return {}; // Errore di decodifica

        isTempFile = true;
        tempFileToDelete = xmlPathToProcess;
    }

    std::vector<AllegatoInfo> result;

    IXMLDOMDocument2* pDoc = NULL;
    HRESULT hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER,
        __uuidof(IXMLDOMDocument2), (void**)&pDoc);
    if (FAILED(hr) || !pDoc)
    {
        if (isTempFile) _wremove(tempFileToDelete.c_str());
        return result;
    }

    pDoc->put_async(VARIANT_FALSE);
    pDoc->setProperty(_bstr_t(L"SelectionLanguage"), _variant_t(L"XPath"));
    pDoc->setProperty(_bstr_t(L"SelectionNamespaces"), _variant_t(L""));

    if (!LoadXmlFileIntoDoc(xmlPathToProcess, pDoc))
    {
        pDoc->Release();
        if (isTempFile) _wremove(tempFileToDelete.c_str());
        return result;
    }

    IXMLDOMNodeList* pList = NULL;
    BSTR bstrQ = SysAllocString(L"//*[local-name()='Allegati']");
    hr = pDoc->selectNodes(bstrQ, &pList);
    SysFreeString(bstrQ);

    if (FAILED(hr) || !pList)
    {
        pDoc->Release();
        if (isTempFile) _wremove(tempFileToDelete.c_str());
        return result;
    }

    long count = 0;
    pList->get_length(&count);
    if (count == 0)
    {
        pList->Release();
        pDoc->Release();
        if (isTempFile) _wremove(tempFileToDelete.c_str());
        return result;
    }

    // Crea la cartella di destinazione
    std::wstring folder = outputFolder;
    if (!folder.empty() && folder.back() != L'\\')
        folder += L'\\';
    CreateDirectoryW(folder.c_str(), NULL);

    auto getChildText = [](IXMLDOMNode* pParent, const wchar_t* localName) -> std::wstring
    {
        std::wstring xp = std::wstring(L"*[local-name()='") + localName + L"']";
        BSTR bstr = SysAllocString(xp.c_str());
        IXMLDOMNode* pChild = NULL;
        std::wstring text;
        if (SUCCEEDED(pParent->selectSingleNode(bstr, &pChild)) && pChild)
        {
            BSTR bstrVal = NULL;
            if (SUCCEEDED(pChild->get_text(&bstrVal)) && bstrVal)
            {
                text = bstrVal;
                SysFreeString(bstrVal);
            }
            pChild->Release();
        }
        SysFreeString(bstr);
        return text;
    };

    for (long i = 0; i < count; i++)
    {
        IXMLDOMNode* pNode = NULL;
        if (FAILED(pList->get_item(i, &pNode)) || !pNode)
            continue;

        std::wstring nome    = getChildText(pNode, L"NomeAttachment");
        std::wstring desc    = getChildText(pNode, L"DescrizioneAttachment");
        std::wstring formato = getChildText(pNode, L"FormatoAttachment");
        std::wstring b64     = getChildText(pNode, L"Attachment");
        pNode->Release();

        if (nome.empty() || b64.empty())
            continue;

        // Rimuovi whitespace dal base64
        std::wstring cleanB64;
        cleanB64.reserve(b64.size());
        for (wchar_t c : b64)
            if (c != L' ' && c != L'\t' && c != L'\r' && c != L'\n')
                cleanB64 += c;

        // Decodifica base64 con CryptStringToBinaryW
        DWORD cbBinary = 0;
        if (!CryptStringToBinaryW(cleanB64.c_str(), (DWORD)cleanB64.size(),
            CRYPT_STRING_BASE64, NULL, &cbBinary, NULL, NULL) || cbBinary == 0)
            continue;

        std::vector<BYTE> bin(cbBinary);
        if (!CryptStringToBinaryW(cleanB64.c_str(), (DWORD)cleanB64.size(),
            CRYPT_STRING_BASE64, bin.data(), &cbBinary, NULL, NULL))
            continue;

        // Sanifica il nome del file (rimuovi caratteri illegali)
        std::wstring safeName = nome;
        for (wchar_t& c : safeName)
            if (c == L'/' || c == L'\\' || c == L':' || c == L'*' ||
                c == L'?' || c == L'"' || c == L'<' || c == L'>' || c == L'|')
                c = L'_';

        std::wstring outFilePath = folder + safeName;
        std::ofstream ofs(outFilePath, std::ios::binary);
        if (!ofs.is_open())
            continue;

        ofs.write(reinterpret_cast<const char*>(bin.data()), cbBinary);
        ofs.close();

        AllegatoInfo info;
        info.nomeAttachment = nome;
        info.descrizione    = desc;
        info.formato        = formato;
        info.filePath       = outFilePath;
        result.push_back(info);
    }

    pList->Release();
    pDoc->Release();

    // Pulisci il file temporaneo se è stato creato
    if (isTempFile)
    {
        _wremove(tempFileToDelete.c_str());
    }

    return result;
}
