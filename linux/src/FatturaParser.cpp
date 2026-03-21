#include "FatturaParser.h"

#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxml/xpath.h>
#include <libxml/xpathInternals.h>
#include <libxslt/xslt.h>
#include <libxslt/xsltInternals.h>
#include <libxslt/transform.h>
#include <libxslt/xsltutils.h>

#include <zip.h>

#include <filesystem>
#include <fstream>
#include <algorithm>
#include <cstring>

namespace fs = std::filesystem;

// ---------------------------------------------------------------------------
// Implementazione interna (nascosta dall'header)
// ---------------------------------------------------------------------------
struct FatturaParser::XmlImpl
{
    xmlDocPtr         xmlDoc    = nullptr;
    xsltStylesheetPtr xsltSheet = nullptr;
};

// ---------------------------------------------------------------------------
FatturaParser::FatturaParser()
    : m_impl(new XmlImpl())
{
    xmlInitParser();
    xmlSubstituteEntitiesDefault(1);
}

FatturaParser::~FatturaParser()
{
    if (m_impl->xsltSheet) { xsltFreeStylesheet(m_impl->xsltSheet); }
    if (m_impl->xmlDoc)    { xmlFreeDoc(m_impl->xmlDoc); }
    xsltCleanupGlobals();
    xmlCleanupParser();
    delete m_impl;
}

// ---------------------------------------------------------------------------
// Estrazione ZIP tramite libzip
// ---------------------------------------------------------------------------
bool FatturaParser::ExtractZipFile(const std::string& zipPath, const std::string& extractPath)
{
    int err = 0;
    zip_t* z = zip_open(zipPath.c_str(), ZIP_RDONLY, &err);
    if (!z) return false;

    fs::create_directories(extractPath);

    const zip_int64_t numEntries = zip_get_num_entries(z, 0);
    for (zip_int64_t i = 0; i < numEntries; ++i)
    {
        struct zip_stat st;
        if (zip_stat_index(z, i, 0, &st) != 0) continue;

        std::string entryName = st.name;
        std::string destPath  = extractPath + "/" + entryName;

        // Crea la directory padre se necessaria
        fs::path destFs(destPath);
        fs::create_directories(destFs.parent_path());

        // Salta le voci che sono directory
        if (!entryName.empty() && entryName.back() == '/') continue;

        zip_file_t* zf = zip_fopen_index(z, i, 0);
        if (!zf) continue;

        std::ofstream out(destPath, std::ios::binary);
        if (out.is_open())
        {
            char buf[8192];
            zip_int64_t n;
            while ((n = zip_fread(zf, buf, sizeof(buf))) > 0)
                out.write(buf, n);
        }
        zip_fclose(zf);
    }

    zip_close(z);
    return true;
}

// ---------------------------------------------------------------------------
// Caricamento XML
// ---------------------------------------------------------------------------
bool FatturaParser::LoadXmlFile(const std::string& xmlPath)
{
    if (m_impl->xmlDoc)
    {
        xmlFreeDoc(m_impl->xmlDoc);
        m_impl->xmlDoc = nullptr;
    }
    m_impl->xmlDoc = xmlParseFile(xmlPath.c_str());
    return m_impl->xmlDoc != nullptr;
}

// ---------------------------------------------------------------------------
// Trasformazione XSLT
// ---------------------------------------------------------------------------
bool FatturaParser::ApplyXsltTransform(const std::string& xsltPath, std::string& htmlOutput)
{
    if (!m_impl->xmlDoc) return false;

    // Ricarica il foglio di stile solo se cambiato
    if (m_impl->xsltSheet)
    {
        xsltFreeStylesheet(m_impl->xsltSheet);
        m_impl->xsltSheet = nullptr;
    }

    xmlDocPtr xsltDoc = xmlParseFile(xsltPath.c_str());
    if (!xsltDoc) return false;

    m_impl->xsltSheet = xsltParseStylesheetDoc(xsltDoc);
    if (!m_impl->xsltSheet) return false;

    xmlDocPtr result = xsltApplyStylesheet(m_impl->xsltSheet, m_impl->xmlDoc, nullptr);
    if (!result) return false;

    xmlChar* output  = nullptr;
    int      outLen  = 0;
    xsltSaveResultToString(&output, &outLen, result, m_impl->xsltSheet);
    xmlFreeDoc(result);

    if (output)
    {
        htmlOutput = reinterpret_cast<char*>(output);
        xmlFree(output);
        return true;
    }
    return false;
}

// ---------------------------------------------------------------------------
// Scansione cartella per file .xml (ricorsiva)
// ---------------------------------------------------------------------------
std::vector<std::string> FatturaParser::GetXmlFilesFromFolder(const std::string& folderPath)
{
    std::vector<std::string> result;
    if (!fs::exists(folderPath)) return result;

    for (const auto& entry : fs::recursive_directory_iterator(folderPath))
    {
        if (!entry.is_regular_file()) continue;
        std::string ext = entry.path().extension().string();
        std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
        if (ext == ".xml")
            result.push_back(entry.path().string());
    }
    return result;
}

// ---------------------------------------------------------------------------
// Helper: testo di un nodo XPath su documento già caricato
// ---------------------------------------------------------------------------
std::string FatturaParser::GetNodeTextFromDoc(void*              docPtr,
                                               const std::string& xpath,
                                               const std::string& nsPrefix,
                                               const std::string& nsUri)
{
    auto* doc = static_cast<xmlDocPtr>(docPtr);
    if (!doc) return {};

    xmlXPathContextPtr ctx = xmlXPathNewContext(doc);
    if (!ctx) return {};

    if (!nsPrefix.empty() && !nsUri.empty())
        xmlXPathRegisterNs(ctx,
            reinterpret_cast<const xmlChar*>(nsPrefix.c_str()),
            reinterpret_cast<const xmlChar*>(nsUri.c_str()));

    xmlXPathObjectPtr obj = xmlXPathEvalExpression(
        reinterpret_cast<const xmlChar*>(xpath.c_str()), ctx);

    std::string text;
    if (obj && obj->nodesetval && obj->nodesetval->nodeNr > 0)
    {
        xmlChar* content = xmlNodeGetContent(obj->nodesetval->nodeTab[0]);
        if (content)
        {
            text = reinterpret_cast<char*>(content);
            xmlFree(content);
        }
    }

    if (obj) xmlXPathFreeObject(obj);
    xmlXPathFreeContext(ctx);
    return text;
}

// ---------------------------------------------------------------------------
// Verifica FatturaPA
// ---------------------------------------------------------------------------
bool FatturaParser::IsFatturaPA(const std::string& xmlPath)
{
    xmlDocPtr doc = xmlParseFile(xmlPath.c_str());
    if (!doc) return false;

    xmlNodePtr root = xmlDocGetRootElement(doc);
    std::string tag = root ? reinterpret_cast<const char*>(root->name) : "";
    xmlFreeDoc(doc);

    return tag.find("FatturaElettronica") != std::string::npos;
}

// ---------------------------------------------------------------------------
// Estrazione informazioni dalla fattura (replica logica Windows)
// ---------------------------------------------------------------------------
bool FatturaParser::ExtractFatturaInfo(const std::string& xmlPath, FatturaInfo& info)
{
    xmlDocPtr doc = xmlParseFile(xmlPath.c_str());
    if (!doc) return false;

    // Controlla root element
    xmlNodePtr root = xmlDocGetRootElement(doc);
    if (!root)
    {
        xmlFreeDoc(doc);
        return false;
    }
    std::string rootTag = reinterpret_cast<const char*>(root->name);
    if (rootTag.find("FatturaElettronica") == std::string::npos)
    {
        xmlFreeDoc(doc);
        return false;
    }

    info.filePath = xmlPath;

    const std::string NS_PRE = "p";
    const std::string NS_URI = "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2";

    // Prova prima con namespace FatturaPA v1.2
    info.cedenteDenominazione = GetNodeTextFromDoc(doc,
        "//p:FatturaElettronicaHeader/p:CedentePrestatore/p:DatiAnagrafici/p:Anagrafica/p:Denominazione",
        NS_PRE, NS_URI);

    if (info.cedenteDenominazione.empty())
    {
        // Senza namespace
        info.cedenteDenominazione = GetNodeTextFromDoc(doc,
            "//FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione");

        if (info.cedenteDenominazione.empty())
        {
            std::string nome    = GetNodeTextFromDoc(doc, "//FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Nome");
            std::string cognome = GetNodeTextFromDoc(doc, "//FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Cognome");
            if (!nome.empty() || !cognome.empty())
                info.cedenteDenominazione = cognome + " " + nome;
        }

        info.cessionarioDenominazione = GetNodeTextFromDoc(doc,
            "//FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione");
        if (info.cessionarioDenominazione.empty())
        {
            std::string nome    = GetNodeTextFromDoc(doc, "//FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Nome");
            std::string cognome = GetNodeTextFromDoc(doc, "//FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Cognome");
            if (!nome.empty() || !cognome.empty())
                info.cessionarioDenominazione = cognome + " " + nome;
        }

        info.numero        = GetNodeTextFromDoc(doc, "//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Numero");
        info.data          = GetNodeTextFromDoc(doc, "//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data");
        info.importoTotale = GetNodeTextFromDoc(doc, "//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/ImportoTotaleDocumento");
        info.tipoDocumento = GetNodeTextFromDoc(doc, "//FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/TipoDocumento");
    }
    else
    {
        // Con namespace
        info.cessionarioDenominazione = GetNodeTextFromDoc(doc,
            "//p:FatturaElettronicaHeader/p:CessionarioCommittente/p:DatiAnagrafici/p:Anagrafica/p:Denominazione",
            NS_PRE, NS_URI);
        if (info.cessionarioDenominazione.empty())
        {
            std::string nome    = GetNodeTextFromDoc(doc, "//p:FatturaElettronicaHeader/p:CessionarioCommittente/p:DatiAnagrafici/p:Anagrafica/p:Nome", NS_PRE, NS_URI);
            std::string cognome = GetNodeTextFromDoc(doc, "//p:FatturaElettronicaHeader/p:CessionarioCommittente/p:DatiAnagrafici/p:Anagrafica/p:Cognome", NS_PRE, NS_URI);
            if (!nome.empty() || !cognome.empty())
                info.cessionarioDenominazione = cognome + " " + nome;
        }

        info.numero        = GetNodeTextFromDoc(doc, "//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:Numero",                    NS_PRE, NS_URI);
        info.data          = GetNodeTextFromDoc(doc, "//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:Data",                      NS_PRE, NS_URI);
        info.importoTotale = GetNodeTextFromDoc(doc, "//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:ImportoTotaleDocumento",    NS_PRE, NS_URI);
        info.tipoDocumento = GetNodeTextFromDoc(doc, "//p:FatturaElettronicaBody/p:DatiGenerali/p:DatiGeneraliDocumento/p:TipoDocumento",             NS_PRE, NS_URI);
    }

    xmlFreeDoc(doc);

    if (info.cedenteDenominazione.empty())
        info.cedenteDenominazione = fs::path(xmlPath).filename().string();

    return true;
}

// ---------------------------------------------------------------------------
std::vector<FatturaInfo> FatturaParser::GetFattureInfoFromFolder(const std::string& folderPath)
{
    std::vector<FatturaInfo> result;
    for (const auto& xmlFile : GetXmlFilesFromFolder(folderPath))
    {
        FatturaInfo info;
        if (ExtractFatturaInfo(xmlFile, info))
            result.push_back(std::move(info));
    }
    return result;
}
