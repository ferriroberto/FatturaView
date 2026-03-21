#include "FatturaViewer.h"
#include <fstream>
#include <shlwapi.h>
#include <comdef.h>
#include <ole2.h>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

// Struttura per memorizzare il controllo WebBrowser
struct BrowserData
{
    IWebBrowser2* pWebBrowser;
    IOleObject* pOleObject;
    IOleInPlaceObject* pInPlaceObject;
    IOleClientSite* pClientSite;
};

// Window procedure per il contenitore del browser
static LRESULT CALLBACK BrowserContainerProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_SIZE:
        {
            BrowserData* pData = (BrowserData*)GetWindowLongPtr(hWnd, GWLP_USERDATA);
            if (pData && pData->pInPlaceObject)
            {
                RECT rect;
                GetClientRect(hWnd, &rect);
                pData->pInPlaceObject->SetObjectRects(&rect, &rect);
            }
        }
        return 0;
    case WM_DESTROY:
        {
            BrowserData* pData = (BrowserData*)GetWindowLongPtr(hWnd, GWLP_USERDATA);
            if (pData)
            {
                if (pData->pInPlaceObject)
                {
                    pData->pInPlaceObject->Release();
                    pData->pInPlaceObject = NULL;
                }
                if (pData->pOleObject)
                {
                    pData->pOleObject->Close(OLECLOSE_NOSAVE);
                    pData->pOleObject->Release();
                    pData->pOleObject = NULL;
                }
                if (pData->pClientSite)
                {
                    pData->pClientSite->Release();
                    pData->pClientSite = NULL;
                }
                if (pData->pWebBrowser)
                {
                    pData->pWebBrowser->Release();
                    pData->pWebBrowser = NULL;
                }
                delete pData;
                SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
            }
        }
        break;
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}

// Implementazione semplice di IOleClientSite
class SimpleOleClientSite : public IOleClientSite, public IOleInPlaceSite
{
private:
    LONG m_refCount;
    HWND m_hWnd;

public:
    SimpleOleClientSite(HWND hWnd) : m_refCount(1), m_hWnd(hWnd) {}

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID riid, void** ppv)
    {
        if (riid == IID_IUnknown || riid == IID_IOleClientSite)
        {
            *ppv = static_cast<IOleClientSite*>(this);
        }
        else if (riid == IID_IOleInPlaceSite)
        {
            *ppv = static_cast<IOleInPlaceSite*>(this);
        }
        else
        {
            *ppv = NULL;
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }

    STDMETHOD_(ULONG, AddRef)() { return ++m_refCount; }
    STDMETHOD_(ULONG, Release)()
    {
        if (--m_refCount == 0)
        {
            delete this;
            return 0;
        }
        return m_refCount;
    }

    // IOleClientSite
    STDMETHOD(SaveObject)() { return E_NOTIMPL; }
    STDMETHOD(GetMoniker)(DWORD, DWORD, IMoniker**) { return E_NOTIMPL; }
    STDMETHOD(GetContainer)(IOleContainer**) { return E_NOINTERFACE; }
    STDMETHOD(ShowObject)() { return S_OK; }
    STDMETHOD(OnShowWindow)(BOOL) { return S_OK; }
    STDMETHOD(RequestNewObjectLayout)() { return E_NOTIMPL; }

    // IOleWindow (base of IOleInPlaceSite)
    STDMETHOD(GetWindow)(HWND* phwnd)
    {
        *phwnd = m_hWnd;
        return S_OK;
    }
    STDMETHOD(ContextSensitiveHelp)(BOOL) { return E_NOTIMPL; }

    // IOleInPlaceSite
    STDMETHOD(CanInPlaceActivate)() { return S_OK; }
    STDMETHOD(OnInPlaceActivate)() { return S_OK; }
    STDMETHOD(OnUIActivate)() { return S_OK; }
    STDMETHOD(GetWindowContext)(IOleInPlaceFrame**, IOleInPlaceUIWindow**, LPRECT, LPRECT, LPOLEINPLACEFRAMEINFO)
    {
        return E_NOTIMPL;
    }
    STDMETHOD(Scroll)(SIZE) { return E_NOTIMPL; }
    STDMETHOD(OnUIDeactivate)(BOOL) { return S_OK; }
    STDMETHOD(OnInPlaceDeactivate)() { return S_OK; }
    STDMETHOD(DiscardUndoState)() { return E_NOTIMPL; }
    STDMETHOD(DeactivateAndUndo)() { return E_NOTIMPL; }
    STDMETHOD(OnPosRectChange)(LPCRECT) { return S_OK; }
};

HWND FatturaViewer::CreateBrowserControl(HWND hParent, HINSTANCE hInst, int x, int y, int width, int height)
{
    // Registra la classe di finestra per il contenitore del browser
    static bool classRegistered = false;
    if (!classRegistered)
    {
        WNDCLASSEXW wc = { sizeof(WNDCLASSEXW) };
        wc.lpfnWndProc = BrowserContainerProc;
        wc.hInstance = hInst;
        wc.lpszClassName = L"BrowserContainerClass";
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        RegisterClassExW(&wc);
        classRegistered = true;
    }

    // Crea una finestra contenitore
    HWND hContainer = CreateWindowExW(
        0,
        L"BrowserContainerClass",
        L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
        x, y, width, height,
        hParent,
        NULL,
        hInst,
        NULL
    );

    if (!hContainer)
    {
        return NULL;
    }

    CoInitialize(NULL);

    BrowserData* pData = new BrowserData();
    ZeroMemory(pData, sizeof(BrowserData));

    // Crea il controllo WebBrowser
    HRESULT hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_INPROC_SERVER,
        IID_IWebBrowser2, (void**)&pData->pWebBrowser);

    if (FAILED(hr))
    {
        delete pData;
        DestroyWindow(hContainer);
        return NULL;
    }

    // Ottieni IOleObject
    hr = pData->pWebBrowser->QueryInterface(IID_IOleObject, (void**)&pData->pOleObject);
    if (FAILED(hr))
    {
        pData->pWebBrowser->Release();
        delete pData;
        DestroyWindow(hContainer);
        return NULL;
    }

    // Crea il client site
    pData->pClientSite = new SimpleOleClientSite(hContainer);

    // Imposta il client site
    pData->pOleObject->SetClientSite(pData->pClientSite);

    // Attiva il controllo
    RECT rect;
    GetClientRect(hContainer, &rect);
    hr = pData->pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, pData->pClientSite, 0, hContainer, &rect);

    // Ottieni IOleInPlaceObject
    hr = pData->pWebBrowser->QueryInterface(IID_IOleInPlaceObject, (void**)&pData->pInPlaceObject);
    if (SUCCEEDED(hr))
    {
        pData->pInPlaceObject->SetObjectRects(&rect, &rect);
    }

    // Configura il browser
    pData->pWebBrowser->put_Silent(VARIANT_TRUE);

    SetWindowLongPtr(hContainer, GWLP_USERDATA, (LONG_PTR)pData);

    return hContainer;
}

void FatturaViewer::NavigateToHTML(HWND hBrowser, const std::wstring& htmlFilePath)
{
    if (!hBrowser)
        return;

    BrowserData* pData = (BrowserData*)GetWindowLongPtr(hBrowser, GWLP_USERDATA);
    if (!pData || !pData->pWebBrowser)
        return;

    // Converti il percorso file in formato URL file://
    std::wstring fileUrl = L"file:///" + htmlFilePath;

    // Sostituisci i backslash con forward slash per URL
    for (size_t i = 0; i < fileUrl.length(); i++)
    {
        if (fileUrl[i] == L'\\')
            fileUrl[i] = L'/';
    }

    VARIANT vEmpty;
    VariantInit(&vEmpty);

    BSTR bstrURL = SysAllocString(fileUrl.c_str());
    pData->pWebBrowser->Navigate(bstrURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
    SysFreeString(bstrURL);

    VariantClear(&vEmpty);
}

void FatturaViewer::SaveHTMLToFile(const std::wstring& htmlContent, const std::wstring& filePath)
{
    // Apri il file in modalità binaria per scrivere UTF-8 correttamente
    std::ofstream outFile(filePath, std::ios::binary);
    if (outFile.is_open())
    {
        // Scrivi il BOM UTF-8 (opzionale ma aiuta alcuni browser)
        const char bom[] = { (char)0xEF, (char)0xBB, (char)0xBF };
        outFile.write(bom, 3);

        // Converti wstring in UTF-8 string
        int size_needed = WideCharToMultiByte(CP_UTF8, 0, htmlContent.c_str(), (int)htmlContent.length(), NULL, 0, NULL, NULL);
        std::string utf8Content(size_needed, 0);
        WideCharToMultiByte(CP_UTF8, 0, htmlContent.c_str(), (int)htmlContent.length(), &utf8Content[0], size_needed, NULL, NULL);

        // Verifica se l'HTML ha già il tag meta charset
        std::string htmlStr = utf8Content;
        bool hasCharset = (htmlStr.find("charset") != std::string::npos);

        if (!hasCharset)
        {
            // Cerca il tag <head> e inserisci il meta charset
            size_t headPos = htmlStr.find("<head>");
            if (headPos != std::string::npos)
            {
                htmlStr.insert(headPos + 6, "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">");
            }
            else
            {
                // Se non c'è <head>, cerca <html> e aggiungi un <head>
                size_t htmlPos = htmlStr.find("<html");
                if (htmlPos != std::string::npos)
                {
                    size_t closePos = htmlStr.find(">", htmlPos);
                    if (closePos != std::string::npos)
                    {
                        htmlStr.insert(closePos + 1, "<head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head>");
                    }
                }
            }
        }

        outFile.write(htmlStr.c_str(), htmlStr.length());
        outFile.close();
    }
}

void FatturaViewer::CreateWelcomePage(const std::wstring& filePath)
{
    std::wstring welcomeHTML = LR"(<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FatturaView</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        html, body {
            width: 100%;
            height: 100%;
            overflow: hidden;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            display: flex;
            justify-content: center;
            align-items: center;
            color: #fff;
        }
        .container {
            text-align: center;
            padding: 40px;
            background: rgba(255, 255, 255, 0.1);
            border-radius: 20px;
            backdrop-filter: blur(10px);
            box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
            max-width: 90%;
            max-height: 90%;
            overflow-y: auto;
        }
        h1 {
            font-size: 2.5em;
            margin-bottom: 20px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        .icon {
            font-size: 5em;
            margin-bottom: 20px;
        }
        .instructions {
            font-size: 1.2em;
            line-height: 1.8;
            margin-top: 30px;
        }
        .step {
            background: rgba(255, 255, 255, 0.2);
            padding: 15px;
            margin: 10px 0;
            border-radius: 10px;
            text-align: left;
        }
        .step-number {
            display: inline-block;
            width: 30px;
            height: 30px;
            background: #fff;
            color: #667eea;
            border-radius: 50%;
            text-align: center;
            line-height: 30px;
            font-weight: bold;
            margin-right: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="icon">&#128196;</div>
        <h1>FatturaView</h1>
        <p style="font-size: 1.1em; opacity: 0.9;">Benvenuto nel lettore di fatture XML italiane</p>

        <div class="instructions">
            <div class="step">
                <span class="step-number">1</span>
                <strong>Apri un archivio ZIP</strong><br>
                Menu - File - Apri archivio ZIP...
            </div>
            <div class="step">
                <span class="step-number">2</span>
                <strong>Seleziona una fattura</strong><br>
                Doppio click sulla fattura nella lista a sinistra
            </div>
            <div class="step">
                <span class="step-number">3</span>
                <strong>Scegli il foglio di stile</strong><br>
                Seleziona il file XSLT (Ministero o Assosoftware)
            </div>
            <div class="step">
                <span class="step-number">4</span>
                <strong>Visualizza la fattura</strong><br>
                La fattura apparira' qui in formato HTML
            </div>
        </div>
    </div>
</body>
</html>)";

    SaveHTMLToFile(welcomeHTML, filePath);
}

std::wstring FatturaViewer::GetWelcomePageHTML()
{
    std::wstring html;
    html = L"<!DOCTYPE html><html lang=\"it\"><head>";
    html += L"<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">";
    html += L"<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">";
    html += L"<title>FatturaView</title><style>";
    html += L"*{margin:0;padding:0;box-sizing:border-box;}";
    html += L"html,body{width:100%;height:100%;overflow:auto;}";
    html += L"body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;";
    html += L"background:linear-gradient(145deg,#0078d4 0%,#106ebe 50%,#005a9e 100%);";
    html += L"padding:40px 20px;}";
    html += L".container{text-align:center;padding:60px 50px;background:rgba(255,255,255,0.98);";
    html += L"border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.2);max-width:900px;margin:0 auto;}";
    html += L".icon{margin-bottom:22px;display:flex;justify-content:center;"
        L"filter:drop-shadow(0 10px 28px rgba(0,80,160,0.55));"
        L"animation:floatIcon 3.2s ease-in-out infinite;}"
        L"@keyframes floatIcon{0%,100%{transform:translateY(0px)}"
        L"50%{transform:translateY(-12px)}}";
    html += L"h1{font-size:2.8em;margin-bottom:12px;color:#0078d4;font-weight:700;}";
    html += L".subtitle{font-size:1.25em;color:#605e5c;margin-bottom:8px;font-weight:500;}";
    html += L".version-info{font-size:1em;color:#8a8886;margin-top:10px;padding:8px 20px;";
    html += L"background:#f3f2f1;border-radius:20px;display:inline-block;}";
    html += L".badge{background:#0078d4;color:#fff;padding:5px 15px;";
    html += L"border-radius:15px;font-size:0.85em;margin-left:8px;font-weight:600;}";
    html += L".instructions{text-align:left;margin-top:40px;background:#f0f7ff;";
    html += L"padding:30px;border-radius:16px;box-shadow:0 4px 20px rgba(0,120,212,0.1);}";
    html += L".instructions-title{color:#0078d4;font-size:1.5em;font-weight:700;margin-bottom:25px;";
    html += L"text-align:center;}";
    html += L".steps-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px;}";
    html += L".step{background:#fff;padding:20px;border-radius:12px;";
    html += L"box-shadow:0 4px 15px rgba(0,0,0,0.08);transition:all 0.3s ease;";
    html += L"border-left:4px solid #0078d4;}";
    html += L".step:hover{transform:translateY(-5px);box-shadow:0 8px 25px rgba(0,120,212,0.2);}";
    html += L".step-number{display:inline-block;width:40px;height:40px;";
    html += L"background:#0078d4;color:#fff;";
    html += L"border-radius:50%;text-align:center;line-height:40px;font-weight:700;font-size:1.2em;";
    html += L"margin-bottom:12px;}";
    html += L".step-title{color:#0078d4;font-size:1.1em;font-weight:700;margin-bottom:6px;}";
    html += L".step-desc{color:#605e5c;font-size:0.95em;line-height:1.5;}";
    html += L".footer{margin-top:35px;padding-top:25px;border-top:2px solid #e1dfdd;";
    html += L"color:#8a8886;font-size:0.95em;}";
    html += L".features{display:flex;justify-content:center;gap:25px;margin-top:25px;flex-wrap:wrap;}";
    html += L".feature-badge{background:#fff;padding:12px 24px;border-radius:25px;";
    html += L"box-shadow:0 2px 10px rgba(0,0,0,0.08);font-size:0.9em;color:#323130;font-weight:600;}";
    html += L"</style></head><body><div class=\"container\">";
    html += L"<div class=\"icon\">";
    html += L"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\" width=\"128\" height=\"128\">";
    html += L"<defs>";
    html += L"<linearGradient id=\"fvbg\" x1=\"0\" y1=\"1\" x2=\"1\" y2=\"0\">";
    html += L"<stop offset=\"0%\" stop-color=\"#002d6e\"/><stop offset=\"100%\" stop-color=\"#0078d4\"/>"
        L"</linearGradient>";
    html += L"</defs>";
    html += L"<rect x=\"1\" y=\"1\" width=\"98\" height=\"98\" rx=\"20\" fill=\"url(#fvbg)\"/>";
    html += L"<ellipse cx=\"36\" cy=\"24\" rx=\"32\" ry=\"18\" fill=\"rgba(255,255,255,0.11)\"/>";
    html += L"<rect x=\"22\" y=\"17\" width=\"46\" height=\"57\" rx=\"5\" fill=\"rgba(255,255,255,0.93)\"/>";
    html += L"<polygon points=\"57,17 68,17 68,28 57,28\" fill=\"#b8d4ee\"/>";
    html += L"<line x1=\"57\" y1=\"17\" x2=\"57\" y2=\"28\" stroke=\"#90b6d6\" stroke-width=\"0.9\"/>";
    html += L"<rect x=\"30\" y=\"37\" width=\"24\" height=\"3.5\" rx=\"1.8\" fill=\"#0078d4\"/>";
    html += L"<rect x=\"30\" y=\"44\" width=\"18\" height=\"3\" rx=\"1.5\" fill=\"#85b8d8\"/>";
    html += L"<rect x=\"30\" y=\"51\" width=\"20\" height=\"3\" rx=\"1.5\" fill=\"#85b8d8\"/>";
    html += L"<circle cx=\"62\" cy=\"63\" r=\"13\" fill=\"rgba(0,0,0,0.15)\" transform=\"translate(1,2)\"/>";
    html += L"<circle cx=\"62\" cy=\"63\" r=\"13\" fill=\"#0078d4\"/>";
    html += L"<circle cx=\"62\" cy=\"63\" r=\"11\" fill=\"#005fa3\"/>";
    html += L"<text x=\"62\" y=\"67.5\" font-family=\"Arial,sans-serif\" font-size=\"15\" "
        L"font-weight=\"bold\" fill=\"white\" text-anchor=\"middle\">\u20AC</text>";
    html += L"</svg>";
    html += L"</div>";
    html += L"<h1>FatturaView</h1>";
    html += L"<p class=\"subtitle\">&#128196; Gestione Professionale FatturaPA XML</p>";
    html += L"<p class=\"version-info\">Roberto Ferri <span class=\"badge\">v1.0.0</span></p>";
    html += L"<div class=\"features\">";
    html += L"<div class=\"feature-badge\">&#128194; ZIP e XML</div>";
    html += L"<div class=\"feature-badge\">&#128424; Stampa Massiva</div>";
    html += L"<div class=\"feature-badge\">&#127912; Stili Multipli</div>";
    html += L"</div>";
    html += L"<div class=\"instructions\">";
    html += L"<div class=\"instructions-title\">&#128161; Come Iniziare</div>";
    html += L"<div class=\"steps-grid\">";
    html += L"<div class=\"step\"><div class=\"step-number\">1</div>";
    html += L"<div class=\"step-title\">Apri Fatture</div>";
    html += L"<div class=\"step-desc\">Clicca su 'Apri' nella toolbar o menu File (ZIP o XML)</div></div>";
    html += L"<div class=\"step\"><div class=\"step-number\">2</div>";
    html += L"<div class=\"step-title\">Seleziona Fatture</div>";
    html += L"<div class=\"step-desc\">Click singolo o Ctrl+Click per selezione multipla</div></div>";
    html += L"<div class=\"step\"><div class=\"step-number\">3</div>";
    html += L"<div class=\"step-title\">Visualizza</div>";
    html += L"<div class=\"step-desc\">Scegli foglio stile Ministero o Assosoftware</div></div>";
    html += L"<div class=\"step\"><div class=\"step-number\">4</div>";
    html += L"<div class=\"step-title\">Stampa</div>";
    html += L"<div class=\"step-desc\">Stampa singola, multipla o tutte le fatture</div></div>";
    html += L"</div></div>";
    html += L"<div class=\"footer\">";
    html += L"<p><strong>&copy; 2026 Roberto Ferri</strong></p>";
    html += L"<p style=\"margin-top:8px;font-size:0.85em;\">FatturaView - Lettore professionale per Fatture Elettroniche PA</p>";
    html += L"</div></div></body></html>";
    return html;
}

void FatturaViewer::NavigateToString(HWND hBrowser, const std::wstring& htmlContent)
{
    if (!hBrowser)
        return;

    BrowserData* pData = (BrowserData*)GetWindowLongPtr(hBrowser, GWLP_USERDATA);
    if (!pData || !pData->pWebBrowser)
        return;

    // Usa about:blank e poi scrive il contenuto
    VARIANT vEmpty;
    VariantInit(&vEmpty);

    BSTR bstrURL = SysAllocString(L"about:blank");
    pData->pWebBrowser->Navigate(bstrURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
    SysFreeString(bstrURL);

    VariantClear(&vEmpty);

    // Attendi che il browser sia pronto (controlla ReadyState)
    READYSTATE readyState;
    int maxRetries = 100; // Max 5 secondi (50ms * 100)
    int retries = 0;
    do
    {
        Sleep(50);
        HRESULT hr = pData->pWebBrowser->get_ReadyState(&readyState);
        if (FAILED(hr))
            break;
        retries++;
    } while (readyState != READYSTATE_COMPLETE && retries < maxRetries);

    // Se il browser non è pronto dopo il timeout, prova comunque
    if (retries >= maxRetries)
    {
        Sleep(500); // Attendi ancora mezzo secondo
    }

    // Ottieni il document
    IDispatch* pDisp = NULL;
    HRESULT hr = pData->pWebBrowser->get_Document(&pDisp);
    if (SUCCEEDED(hr) && pDisp)
    {
        IHTMLDocument2* pDoc = NULL;
        hr = pDisp->QueryInterface(IID_IHTMLDocument2, (void**)&pDoc);
        if (SUCCEEDED(hr) && pDoc)
        {
            // Crea un SAFEARRAY per il contenuto HTML
            SAFEARRAY* pSA = SafeArrayCreateVector(VT_VARIANT, 0, 1);
            if (pSA)
            {
                VARIANT* pVar;
                SafeArrayAccessData(pSA, (void**)&pVar);
                pVar->vt = VT_BSTR;
                pVar->bstrVal = SysAllocString(htmlContent.c_str());
                SafeArrayUnaccessData(pSA);

                // Scrivi il contenuto nel document
                pDoc->write(pSA);
                pDoc->close();

                SafeArrayDestroy(pSA);
            }
            pDoc->Release();
        }
        pDisp->Release();
    }

    // Forza un refresh del browser
    pData->pWebBrowser->Refresh();
}

void FatturaViewer::CleanupBrowser(HWND hBrowser)
{
    if (hBrowser)
    {
        BrowserData* pData = (BrowserData*)GetWindowLongPtr(hBrowser, GWLP_USERDATA);
        if (pData)
        {
            if (pData->pInPlaceObject)
            {
                pData->pInPlaceObject->Release();
            }
            if (pData->pOleObject)
            {
                pData->pOleObject->Close(OLECLOSE_NOSAVE);
                pData->pOleObject->Release();
            }
            if (pData->pClientSite)
            {
                pData->pClientSite->Release();
            }
            if (pData->pWebBrowser)
            {
                pData->pWebBrowser->Release();
            }
            delete pData;
        }
        DestroyWindow(hBrowser);
    }
}

void FatturaViewer::PrintDocument(HWND hBrowser)
{
    if (!hBrowser)
        return;

    BrowserData* pData = (BrowserData*)GetWindowLongPtr(hBrowser, GWLP_USERDATA);
    if (!pData || !pData->pWebBrowser)
        return;

    // Mostra la finestra di dialogo di stampa standard
    VARIANT vEmpty;
    VariantInit(&vEmpty);
    pData->pWebBrowser->ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT, &vEmpty, &vEmpty);
    VariantClear(&vEmpty);
}

void FatturaViewer::PrintDocumentSilent(HWND hBrowser)
{
    if (!hBrowser)
        return;

    BrowserData* pData = (BrowserData*)GetWindowLongPtr(hBrowser, GWLP_USERDATA);
    if (!pData || !pData->pWebBrowser)
        return;

    // Stampa senza mostrare la finestra di dialogo (usa stampante predefinita)
    VARIANT vEmpty;
    VariantInit(&vEmpty);
    pData->pWebBrowser->ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, &vEmpty, &vEmpty);
    VariantClear(&vEmpty);
}

void FatturaViewer::SetZoom(HWND hBrowser, int zoomPercent)
{
    if (!hBrowser)
        return;

    BrowserData* pData = (BrowserData*)GetWindowLongPtr(hBrowser, GWLP_USERDATA);
    if (!pData || !pData->pWebBrowser)
        return;

#ifndef OLECMDID_OPTICAL_ZOOM
#define OLECMDID_OPTICAL_ZOOM 63
#endif

    VARIANT vZoom;
    VariantInit(&vZoom);
    vZoom.vt = VT_I4;
    vZoom.lVal = zoomPercent;
    pData->pWebBrowser->ExecWB((OLECMDID)OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, &vZoom, NULL);
    VariantClear(&vZoom);
}
