#include "FatturaViewer.h"
#include "Resource.h"
#include <fstream>
#include <shlwapi.h>
#include <comdef.h>
#include <ole2.h>
#include <exdisp.h>
#include <shellapi.h>
#include <wincrypt.h>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "crypt32.lib")

// Struttura per memorizzare il controllo WebBrowser
struct BrowserData
{
    IWebBrowser2*     pWebBrowser;
    IOleObject*       pOleObject;
    IOleInPlaceObject* pInPlaceObject;
    IOleClientSite*   pClientSite;
    IConnectionPoint* pEventCP;
    DWORD             dwEventCookie;
};

// Event sink per intercettare la navigazione e aprire i link http nel browser predefinito
class BrowserEventSink : public IDispatch
{
    LONG m_refCount = 1;
    HWND m_hParent;
public:
    explicit BrowserEventSink(HWND hWnd) : m_hParent(hWnd) {}

    STDMETHOD(QueryInterface)(REFIID riid, void** ppv)
    {
        if (riid == IID_IUnknown || riid == IID_IDispatch || riid == DIID_DWebBrowserEvents2)
        {
            *ppv = this; AddRef(); return S_OK;
        }
        *ppv = nullptr; return E_NOINTERFACE;
    }
    STDMETHOD_(ULONG, AddRef)()  { return InterlockedIncrement(&m_refCount); }
    STDMETHOD_(ULONG, Release)()
    {
        ULONG r = InterlockedDecrement(&m_refCount);
        if (r == 0) delete this;
        return r;
    }
    STDMETHOD(GetTypeInfoCount)(UINT*) { return E_NOTIMPL; }
    STDMETHOD(GetTypeInfo)(UINT, LCID, ITypeInfo**) { return E_NOTIMPL; }
    STDMETHOD(GetIDsOfNames)(REFIID, LPOLESTR*, UINT, LCID, DISPID*) { return E_NOTIMPL; }
    STDMETHOD(Invoke)(DISPID dispid, REFIID, LCID, WORD, DISPPARAMS* pParams,
                      VARIANT*, EXCEPINFO*, UINT*)
    {
        // DISPID 250 = BeforeNavigate2
        // Args (right-to-left): pDisp[6], URL[5], Flags[4], Target[3], PostData[2], Headers[1], Cancel[0]
        if (dispid == 250 && pParams && pParams->cArgs >= 7)
        {
            // URL può arrivare come VT_BSTR oppure come VT_BYREF|VT_VARIANT che wrappa un VT_BSTR
            VARIANT* pURL = &pParams->rgvarg[5];
            BSTR bstrURL = nullptr;
            if (pURL->vt == VT_BSTR)
                bstrURL = pURL->bstrVal;
            else if (pURL->vt == (VT_BYREF | VT_VARIANT) && pURL->pvarVal
                     && pURL->pvarVal->vt == VT_BSTR)
                bstrURL = pURL->pvarVal->bstrVal;

            if (bstrURL)
            {
                std::wstring url(bstrURL);
                bool isHttp       = (url.rfind(L"http", 0) == 0);
                bool isAttachment = (url.rfind(L"fvatt://", 0) == 0);

                if (isHttp || isAttachment)
                {
                    VARIANT* pCancel = &pParams->rgvarg[0];
                    if (pCancel->vt == (VT_BOOL | VT_BYREF) && pCancel->pboolVal)
                        *pCancel->pboolVal = VARIANT_TRUE;

                    if (isAttachment)
                    {
                        // fvatt:///C:/path/to/file  ->  C:\path\to\file
                        std::wstring path = url.substr(8); // salta "fvatt://"
                        while (!path.empty() && path[0] == L'/')
                            path = path.substr(1);
                        std::wstring decoded;
                        for (size_t j = 0; j < path.size(); j++)
                        {
                            if (path[j] == L'%' && j + 2 < path.size())
                            {
                                wchar_t hex[3] = { path[j + 1], path[j + 2], L'\0' };
                                decoded += (wchar_t)wcstol(hex, nullptr, 16);
                                j += 2;
                            }
                            else if (path[j] == L'/')
                                decoded += L'\\';
                            else
                                decoded += path[j];
                        }
                        ShellExecuteW(m_hParent, L"open", decoded.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
                    }
                    else
                    {
                        ShellExecuteW(m_hParent, L"open", url.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
                    }
                }
            }
        }
        return S_OK;
    }
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
                if (pData->pEventCP)
                {
                    pData->pEventCP->Unadvise(pData->dwEventCookie);
                    pData->pEventCP->Release();
                    pData->pEventCP = NULL;
                }
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

    // Collega il sink eventi per intercettare i link http
    IConnectionPointContainer* pCPC = nullptr;
    if (SUCCEEDED(pData->pWebBrowser->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC)))
    {
        if (SUCCEEDED(pCPC->FindConnectionPoint(DIID_DWebBrowserEvents2, &pData->pEventCP)))
        {
            BrowserEventSink* pSink = new BrowserEventSink(hContainer);
            pData->pEventCP->Advise(pSink, &pData->dwEventCookie);
            pSink->Release();
        }
        pCPC->Release();
    }

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
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
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
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4);
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
            background: rgba(59, 130, 246, 0.2);
            padding: 15px;
            margin: 10px 0;
            border-radius: 10px;
            text-align: left;
        }
        .step-number {
            display: inline-block;
            width: 30px;
            height: 30px;
            background: #3b82f6;
            color: #fff;
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
    // Carica l'immagine g6.png dalla risorsa embedded e convertila in base64 data URI
    std::wstring imgDataUri;
    HRSRC hRes = FindResourceW(GetModuleHandle(NULL), MAKEINTRESOURCEW(IDR_WELCOME_IMAGE), L"PNG");
    if (hRes)
    {
        HGLOBAL hData = LoadResource(GetModuleHandle(NULL), hRes);
        DWORD dwSize = SizeofResource(GetModuleHandle(NULL), hRes);
        if (hData && dwSize > 0)
        {
            const BYTE* pData = (const BYTE*)LockResource(hData);
            if (pData)
            {
                DWORD dwB64Size = 0;
                CryptBinaryToStringW(pData, dwSize, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, NULL, &dwB64Size);
                if (dwB64Size > 0)
                {
                    std::wstring b64(dwB64Size, L'\0');
                    CryptBinaryToStringW(pData, dwSize, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, &b64[0], &dwB64Size);
                    b64.resize(dwB64Size);
                    imgDataUri = L"data:image/png;base64," + b64;
                }
            }
        }
    }

    std::wstring html;
    html = L"<!DOCTYPE html><html><head>";
    html += L"<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">";
    html += L"<meta charset=\"UTF-8\">";
    html += L"<title>FatturaView</title><style>";
    html += L"*{margin:0;padding:0;box-sizing:border-box}";
    html += L"html,body{width:100%;height:100%;overflow:hidden;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif}";
    html += L"body{background:linear-gradient(135deg,#1e293b 0%,#0f172a 100%);position:relative}";
    html += L"body::before{content:'';position:absolute;inset:0;background:";
    html += L"radial-gradient(circle at 10% 20%,rgba(59,130,246,0.15) 0%,transparent 50%),";
    html += L"radial-gradient(circle at 90% 80%,rgba(37,99,235,0.1) 0%,transparent 50%);";
    html += L"animation:pulse 8s ease-in-out infinite;z-index:0}";
    html += L"@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.7}}";
    html += L"@keyframes float{0%,100%{transform:translateY(0) rotate(0deg)}50%{transform:translateY(-10px) rotate(2deg)}}";
    html += L"@keyframes slideUp{from{opacity:0;transform:translateY(30px)}to{opacity:1;transform:translateY(0)}}";
    html += L".wrp{width:100%;height:100%;border:none;border-spacing:0;position:relative;z-index:1}";
    html += L".c{text-align:center;padding:36px 32px;display:inline-block;";
    html += L"background:#ffffff;border-radius:24px;max-width:520px;min-width:420px;";
    html += L"box-shadow:0 25px 50px -12px rgba(0,0,0,0.4);";
    html += L"animation:slideUp 0.6s cubic-bezier(0.34,1.56,0.64,1)}";
    html += L".logo{max-width:280px;max-height:120px;margin:0 auto 20px;display:block;animation:float 4s ease-in-out infinite}";
    html += L".hd{margin-bottom:24px}";
    html += L"h1{font-size:2.2em;margin-bottom:8px;font-weight:bold;color:#1d4ed8;letter-spacing:-0.03em}";
    html += L".s{font-size:0.95em;color:#64748b;font-weight:500;margin-bottom:4px}";
    html += L".bg{margin-top:8px;text-align:center}";
    html += L".b{display:inline-block;padding:4px 12px;border-radius:20px;font-size:0.75em;";
    html += L"font-weight:bold;margin:0 3px;text-transform:uppercase}";
    html += L".b1{background:#3b82f6;color:#ffffff}";
    html += L".b2{background:#16a34a;color:#ffffff}";
    html += L".b3{background:#d97706;color:#ffffff}";
    html += L".st{background:#f8fafc;padding:14px 16px;margin:12px 0;border-radius:16px;";
    html += L"text-align:left;display:block;transition:all 0.3s ease;border:2px solid #e2e8f0;position:relative}";
    html += L".st:hover{transform:translateX(6px);box-shadow:0 8px 16px rgba(0,0,0,0.1);border-color:#93c5fd}";
    html += L".n{width:32px;height:32px;background:#3b82f6;color:#ffffff;border-radius:8px;";
    html += L"display:inline-block;text-align:center;line-height:32px;font-weight:bold;";
    html += L"margin-right:12px;vertical-align:middle}";
    html += L".st-inner{display:inline-block;vertical-align:middle;width:calc(100% - 50px)}";
    html += L".st b{display:block;margin-bottom:2px;font-size:1.05em;color:#1e293b}";
    html += L".st span.txt{line-height:1.4;color:#64748b;font-size:0.92em}";
    html += L".ft{margin-top:14px;font-size:0.82em;color:#94a3b8}";
    html += L".ft a{color:#3b82f6;text-decoration:none;font-weight:500}";
    html += L".ft a:hover{text-decoration:underline}";
    html += L"</style></head><body>";
    html += L"<table class=\"wrp\"><tr><td align=\"center\" valign=\"middle\">";
    html += L"<div class=\"c\">";

    // Usa l'immagine g6.png come logo al posto dell'SVG
    if (!imgDataUri.empty())
    {
        html += L"<img class=\"logo\" src=\"" + imgDataUri + L"\" alt=\"FatturaView\">";
    }
    else
    {
        // Fallback: titolo testuale se la risorsa non è disponibile
        html += L"<div style=\"font-size:3em;margin-bottom:20px;\">&#128196;</div>";
    }

    html += L"<div class=\"hd\">";
    html += L"<h1>FatturaView</h1>";
    html += L"<p class=\"s\">Visualizzatore Professionale Fatture Elettroniche PA</p>";
    html += L"<div class=\"bg\">";
    html += L"<span class=\"b b1\">XML/ZIP</span>";
    html += L"<span class=\"b b2\">XSLT</span>";
    html += L"<span class=\"b b3\">PDF</span>";
    html += L"</div></div>";
    html += L"<div class=\"st\"><div class=\"n\">1</div><div class=\"st-inner\"><b>Apri Archivio</b><span class=\"txt\">File - Apri archivio ZIP</span></div></div>";
    html += L"<div class=\"st\"><div class=\"n\">2</div><div class=\"st-inner\"><b>Seleziona Fattura</b><span class=\"txt\">Doppio click nella lista</span></div></div>";
    html += L"<div class=\"st\"><div class=\"n\">3</div><div class=\"st-inner\"><b>Applica Stile</b><span class=\"txt\">Ministero o Assosoftware</span></div></div>";
    html += L"<div class=\"st\"><div class=\"n\">4</div><div class=\"st-inner\"><b>Stampa PDF</b><span class=\"txt\">File - Stampa fatture</span></div></div>";
    html += L"<p class=\"ft\"><a href=\"https://www.fatturaview.it\">www.fatturaview.it</a></p>";
    html += L"</div></td></tr></table></body></html>";
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

    // Attendi che il browser sia pronto (controlla ReadyState) senza bloccare il thread UI
    READYSTATE readyState = READYSTATE_UNINITIALIZED;
    int maxRetries = 200; // Max 2 secondi (10ms * 200)
    int retries = 0;
    do
    {
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        Sleep(10);
        HRESULT hr = pData->pWebBrowser->get_ReadyState(&readyState);
        if (FAILED(hr))
            break;
        retries++;
    } while (readyState != READYSTATE_COMPLETE && retries < maxRetries);

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

void FatturaViewer::ShowNotification(HWND hBrowser, const std::wstring& message, const std::wstring& type)
{
    if (!hBrowser)
        return;

    BrowserData* pData = (BrowserData*)GetWindowLongPtr(hBrowser, GWLP_USERDATA);
    if (!pData || !pData->pWebBrowser)
        return;

    // Ottieni il document
    IDispatch* pDisp = NULL;
    HRESULT hr = pData->pWebBrowser->get_Document(&pDisp);
    if (SUCCEEDED(hr) && pDisp)
    {
        IHTMLDocument2* pDoc = NULL;
        hr = pDisp->QueryInterface(IID_IHTMLDocument2, (void**)&pDoc);
        if (SUCCEEDED(hr) && pDoc)
        {
            // Determina il colore e l'icona in base al tipo
            std::wstring bgColor = L"#0078d4";
            std::wstring icon = L"&#8505;"; // i cerchiata

            if (type == L"success") {
                bgColor = L"#10b981";
                icon = L"&#10004;"; // checkmark
            } else if (type == L"error") {
                bgColor = L"#ef4444";
                icon = L"&#10060;"; // X
            } else if (type == L"warning") {
                bgColor = L"#f59e0b";
                icon = L"&#9888;"; // warning
            }

            // Crea lo script JavaScript per mostrare la notifica
            std::wstring script = L"(function() {";
            script += L"if (!document.getElementById('fv-notification-container')) {";
            script += L"var container = document.createElement('div');";
            script += L"container.id = 'fv-notification-container';";
            script += L"container.style.cssText = 'position:fixed;top:20px;right:20px;z-index:999999;';";
            script += L"document.body.appendChild(container);";
            script += L"}";
            script += L"var notif = document.createElement('div');";
            script += L"notif.innerHTML = '<span style=\"font-size:1.2em;margin-right:10px;\">" + icon + L"</span>" + message + L"';";
            script += L"notif.style.cssText = 'background:" + bgColor + L";color:white;padding:16px 24px;";
            script += L"border-radius:8px;margin-bottom:10px;box-shadow:0 4px 12px rgba(0,0,0,0.3);";
            script += L"font-family:Segoe UI,Arial;font-size:14px;min-width:280px;max-width:400px;";
            script += L"animation:slideIn 0.3s ease-out;display:flex;align-items:center;';";
            script += L"document.getElementById('fv-notification-container').appendChild(notif);";
            script += L"if (!document.getElementById('fv-notification-style')) {";
            script += L"var style = document.createElement('style');";
            script += L"style.id = 'fv-notification-style';";
            script += L"style.textContent = '@keyframes slideIn { from { transform:translateX(400px); opacity:0; } to { transform:translateX(0); opacity:1; } }";
            script += L"@keyframes slideOut { from { transform:translateX(0); opacity:1; } to { transform:translateX(400px); opacity:0; } }';";
            script += L"document.head.appendChild(style);";
            script += L"}";
            script += L"setTimeout(function() {";
            script += L"notif.style.animation = 'slideOut 0.3s ease-in';";
            script += L"setTimeout(function() { notif.remove(); }, 300);";
            script += L"}, 3000);";
            script += L"})();";

            // Esegui lo script
            IHTMLWindow2* pWindow = NULL;
            hr = pDoc->get_parentWindow(&pWindow);
            if (SUCCEEDED(hr) && pWindow)
            {
                VARIANT vResult;
                VariantInit(&vResult);

                BSTR bstrScript = SysAllocString(script.c_str());
                BSTR bstrLanguage = SysAllocString(L"JavaScript");

                pWindow->execScript(bstrScript, bstrLanguage, &vResult);

                SysFreeString(bstrScript);
                SysFreeString(bstrLanguage);
                VariantClear(&vResult);

                pWindow->Release();
            }

            pDoc->Release();
        }
        pDisp->Release();
    }
}
