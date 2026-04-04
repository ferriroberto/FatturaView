#include "AutoUpdater.h"
#include <winhttp.h>
#include <shellapi.h>
#include <commctrl.h>
#include <sstream>
#include <thread>
#include <vector>
#include <string>
#include <fstream>

#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "comctl32.lib")

// Variabile globale per conservare le info dell'aggiornamento trovato in background
static UpdateInfo g_pendingUpdate;

std::string AutoUpdater::HttpsGet(const std::wstring& host, const std::wstring& path, int& httpStatus)
{
    std::string result;
    httpStatus = 0;

    HINTERNET hSession = WinHttpOpen(L"FatturaView-AutoUpdater/1.0",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return result;

    // Abilita TLS 1.2 + TLS 1.3
    DWORD dwSecureProtocols = WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 | WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_3;
    WinHttpSetOption(hSession, WINHTTP_OPTION_SECURE_PROTOCOLS,
        &dwSecureProtocols, sizeof(dwSecureProtocols));

    // Timeout ragionevoli
    DWORD dwTimeout = 10000;
    WinHttpSetOption(hSession, WINHTTP_OPTION_CONNECT_TIMEOUT, &dwTimeout, sizeof(dwTimeout));
    WinHttpSetOption(hSession, WINHTTP_OPTION_SEND_TIMEOUT, &dwTimeout, sizeof(dwTimeout));
    WinHttpSetOption(hSession, WINHTTP_OPTION_RECEIVE_TIMEOUT, &dwTimeout, sizeof(dwTimeout));

    HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(),
        INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return result; }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(),
        NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
        WINHTTP_FLAG_SECURE);
    if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return result; }

    // Header richiesti dall'API GitHub
    LPCWSTR headers = L"Accept: application/vnd.github.v3+json\r\nUser-Agent: FatturaView-AutoUpdater/1.0\r\n";

    if (!WinHttpSendRequest(hRequest, headers, (DWORD)-1L,
        WINHTTP_NO_REQUEST_DATA, 0, 0, 0))
    {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return result;
    }

    if (!WinHttpReceiveResponse(hRequest, NULL))
    {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return result;
    }

    DWORD dwStatusCode = 0;
    DWORD dwSize = sizeof(dwStatusCode);
    WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
        WINHTTP_HEADER_NAME_BY_INDEX, &dwStatusCode, &dwSize, WINHTTP_NO_HEADER_INDEX);

    httpStatus = (int)dwStatusCode;

    if (dwStatusCode != 200)
    {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return result;
    }

    DWORD dwBytesAvailable = 0;
    do
    {
        dwBytesAvailable = 0;
        if (!WinHttpQueryDataAvailable(hRequest, &dwBytesAvailable))
            break;

        if (dwBytesAvailable == 0)
            break;

        std::vector<char> buffer(dwBytesAvailable);
        DWORD dwBytesRead = 0;
        if (WinHttpReadData(hRequest, buffer.data(), dwBytesAvailable, &dwBytesRead))
        {
            result.append(buffer.data(), dwBytesRead);
        }
    } while (dwBytesAvailable > 0);

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    return result;
}

bool AutoUpdater::IsNewerVersion(const std::wstring& local, const std::wstring& remote)
{
    auto parseVersion = [](const std::wstring& v, int out[3])
    {
        out[0] = out[1] = out[2] = 0;
        std::wstring s = v;
        if (!s.empty() && (s[0] == L'v' || s[0] == L'V'))
            s = s.substr(1);

        int idx = 0;
        std::wstringstream ss(s);
        std::wstring token;
        while (std::getline(ss, token, L'.') && idx < 3)
        {
            try { out[idx] = std::stoi(token); } catch (...) {}
            idx++;
        }
    };

    int loc[3], rem[3];
    parseVersion(local, loc);
    parseVersion(remote, rem);

    for (int i = 0; i < 3; i++)
    {
        if (rem[i] > loc[i]) return true;
        if (rem[i] < loc[i]) return false;
    }
    return false;
}

std::wstring AutoUpdater::JsonGetString(const std::string& json, const std::string& key)
{
    std::string search = "\"" + key + "\"";
    size_t pos = json.find(search);
    if (pos == std::string::npos) return L"";

    pos = json.find(':', pos + search.length());
    if (pos == std::string::npos) return L"";

    pos = json.find('"', pos + 1);
    if (pos == std::string::npos) return L"";
    pos++;

    std::string value;
    while (pos < json.size() && json[pos] != '"')
    {
        if (json[pos] == '\\' && pos + 1 < json.size())
        {
            pos++;
            if (json[pos] == 'n') value += '\n';
            else if (json[pos] == 'r') value += '\r';
            else if (json[pos] == 't') value += '\t';
            else if (json[pos] == '"') value += '"';
            else if (json[pos] == '\\') value += '\\';
            else value += json[pos];
        }
        else
        {
            value += json[pos];
        }
        pos++;
    }

    int wlen = MultiByteToWideChar(CP_UTF8, 0, value.c_str(), (int)value.size(), NULL, 0);
    if (wlen <= 0) return L"";
    std::wstring wvalue(wlen, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value.c_str(), (int)value.size(), &wvalue[0], wlen);
    return wvalue;
}

std::wstring AutoUpdater::JsonGetInstallerAssetUrl(const std::string& json)
{
    std::string search = "\"browser_download_url\"";
    size_t pos = 0;

    while ((pos = json.find(search, pos)) != std::string::npos)
    {
        size_t colon = json.find(':', pos + search.length());
        if (colon == std::string::npos) break;

        size_t q1 = json.find('"', colon + 1);
        if (q1 == std::string::npos) break;
        q1++;

        size_t q2 = json.find('"', q1);
        if (q2 == std::string::npos) break;

        std::string url = json.substr(q1, q2 - q1);

        if (url.find(".exe") != std::string::npos ||
            url.find(".msi") != std::string::npos)
        {
            int wlen = MultiByteToWideChar(CP_UTF8, 0, url.c_str(), (int)url.size(), NULL, 0);
            if (wlen > 0)
            {
                std::wstring wurl(wlen, L'\0');
                MultiByteToWideChar(CP_UTF8, 0, url.c_str(), (int)url.size(), &wurl[0], wlen);
                return wurl;
            }
        }

        pos = q2;
    }

    return L"";
}

bool AutoUpdater::CheckForUpdates(const std::wstring& currentVersion, UpdateInfo& info)
{
    info = {};

    std::wstring path = std::wstring(L"/repos/") + GITHUB_OWNER + L"/" + GITHUB_REPO + L"/releases/latest";
    int httpStatus = 0;
    std::string json = HttpsGet(L"api.github.com", path, httpStatus);

    // 404 = nessuna release pubblicata: connessione OK ma niente da aggiornare
    if (httpStatus == 404)
    {
        info.updateAvailable = false;
        return true;
    }

    // Connessione fallita o altro errore HTTP
    if (json.empty() || httpStatus != 200)
        return false;

    std::wstring tagName = JsonGetString(json, "tag_name");
    if (tagName.empty())
    {
        info.updateAvailable = false;
        return true;
    }

    info.latestVersion = tagName;
    info.releaseNotes = JsonGetString(json, "body");
    info.downloadUrl = JsonGetString(json, "html_url");
    info.installerUrl = JsonGetInstallerAssetUrl(json);
    info.updateAvailable = IsNewerVersion(currentVersion, tagName);

    return true;
}

void AutoUpdater::ShowUpdateDialog(HWND hParent, const UpdateInfo& info, const std::wstring& currentVersion)
{
    std::wstring msg = L"Una nuova versione di FatturaView e' disponibile!\n\n";
    msg += L"Versione corrente:  " + currentVersion + L"\n";
    msg += L"Nuova versione:     " + info.latestVersion + L"\n";

    if (!info.releaseNotes.empty())
    {
        std::wstring notes = info.releaseNotes;
        if (notes.length() > 300)
            notes = notes.substr(0, 300) + L"...";
        msg += L"\nNote di rilascio:\n" + notes + L"\n";
    }

    if (!info.installerUrl.empty())
        msg += L"\nVuoi scaricare e installare l'aggiornamento ora?";
    else
        msg += L"\nVuoi aprire la pagina di download?";

    int result = MessageBoxW(hParent, msg.c_str(), L"Aggiornamento Disponibile",
        MB_YESNO | MB_ICONINFORMATION);

    if (result == IDYES)
    {
        if (!info.installerUrl.empty())
        {
            DownloadAndInstall(hParent, info);
        }
        else
        {
            std::wstring url = info.downloadUrl.empty()
                ? std::wstring(L"https://github.com/") + GITHUB_OWNER + L"/" + GITHUB_REPO + L"/releases/latest"
                : info.downloadUrl;
            ShellExecuteW(hParent, L"open", url.c_str(), NULL, NULL, SW_SHOWNORMAL);
        }
    }
}

void AutoUpdater::ShowNoUpdateDialog(HWND hParent, const std::wstring& currentVersion)
{
    std::wstring msg = L"FatturaView e' aggiornato!\n\nVersione corrente: " + currentVersion;
    MessageBoxW(hParent, msg.c_str(), L"Nessun Aggiornamento", MB_OK | MB_ICONINFORMATION);
}

void AutoUpdater::CheckInBackground(HWND hNotifyWnd, const std::wstring& currentVersion)
{
    std::wstring version = currentVersion;
    HWND hwnd = hNotifyWnd;

    std::thread([version, hwnd]()
    {
        UpdateInfo info;
        if (CheckForUpdates(version, info) && info.updateAvailable)
        {
            g_pendingUpdate = info;
            PostMessage(hwnd, WM_APP_UPDATE_AVAILABLE, 0, (LPARAM)&g_pendingUpdate);
        }
    }).detach();
}

// --- Download con progress bar ---

bool AutoUpdater::HttpsDownloadFile(const std::wstring& url, const std::wstring& destPath,
    HWND hProgressBar, HWND hStatusLabel)
{
    // Parsa l'URL per estrarre host e path
    URL_COMPONENTS urlComp = {};
    urlComp.dwStructSize = sizeof(urlComp);
    wchar_t hostBuf[256] = {};
    wchar_t pathBuf[2048] = {};
    urlComp.lpszHostName = hostBuf;
    urlComp.dwHostNameLength = 256;
    urlComp.lpszUrlPath = pathBuf;
    urlComp.dwUrlPathLength = 2048;

    if (!WinHttpCrackUrl(url.c_str(), (DWORD)url.size(), 0, &urlComp))
        return false;

    HINTERNET hSession = WinHttpOpen(L"FatturaView-AutoUpdater/1.0",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return false;

    DWORD dwSecureProtocols = WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 | WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_3;
    WinHttpSetOption(hSession, WINHTTP_OPTION_SECURE_PROTOCOLS, &dwSecureProtocols, sizeof(dwSecureProtocols));

    DWORD dwTimeout = 30000;
    WinHttpSetOption(hSession, WINHTTP_OPTION_CONNECT_TIMEOUT, &dwTimeout, sizeof(dwTimeout));
    WinHttpSetOption(hSession, WINHTTP_OPTION_RECEIVE_TIMEOUT, &dwTimeout, sizeof(dwTimeout));

    INTERNET_PORT port = (urlComp.nScheme == INTERNET_SCHEME_HTTPS) ? INTERNET_DEFAULT_HTTPS_PORT : INTERNET_DEFAULT_HTTP_PORT;
    HINTERNET hConnect = WinHttpConnect(hSession, hostBuf, port, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

    // Combina path + extra info (query string)
    std::wstring fullPath = pathBuf;
    if (urlComp.lpszExtraInfo && urlComp.dwExtraInfoLength > 0)
        fullPath += urlComp.lpszExtraInfo;

    DWORD dwFlags = (urlComp.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", fullPath.c_str(),
        NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, dwFlags);
    if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

    LPCWSTR headers = L"User-Agent: FatturaView-AutoUpdater/1.0\r\n";
    if (!WinHttpSendRequest(hRequest, headers, (DWORD)-1L, WINHTTP_NO_REQUEST_DATA, 0, 0, 0))
    {
        WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession);
        return false;
    }

    if (!WinHttpReceiveResponse(hRequest, NULL))
    {
        WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession);
        return false;
    }

    // Controlla se c'e' un redirect (GitHub usa redirect per gli asset)
    DWORD dwStatusCode = 0;
    DWORD dwSize = sizeof(dwStatusCode);
    WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
        WINHTTP_HEADER_NAME_BY_INDEX, &dwStatusCode, &dwSize, WINHTTP_NO_HEADER_INDEX);

    if (dwStatusCode == 301 || dwStatusCode == 302)
    {
        // Segui il redirect
        DWORD dwRedirSize = 0;
        WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_LOCATION, WINHTTP_HEADER_NAME_BY_INDEX,
            WINHTTP_NO_OUTPUT_BUFFER, &dwRedirSize, WINHTTP_NO_HEADER_INDEX);

        if (GetLastError() == ERROR_INSUFFICIENT_BUFFER && dwRedirSize > 0)
        {
            std::wstring redirectUrl(dwRedirSize / sizeof(wchar_t), L'\0');
            WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_LOCATION, WINHTTP_HEADER_NAME_BY_INDEX,
                &redirectUrl[0], &dwRedirSize, WINHTTP_NO_HEADER_INDEX);
            // Rimuovi null terminatori extra
            while (!redirectUrl.empty() && redirectUrl.back() == L'\0')
                redirectUrl.pop_back();

            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);

            return HttpsDownloadFile(redirectUrl, destPath, hProgressBar, hStatusLabel);
        }
    }

    if (dwStatusCode != 200)
    {
        WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession);
        return false;
    }

    // Ottieni la dimensione totale del file
    DWORD dwContentLength = 0;
    dwSize = sizeof(dwContentLength);
    bool hasContentLength = WinHttpQueryHeaders(hRequest,
        WINHTTP_QUERY_CONTENT_LENGTH | WINHTTP_QUERY_FLAG_NUMBER,
        WINHTTP_HEADER_NAME_BY_INDEX, &dwContentLength, &dwSize, WINHTTP_NO_HEADER_INDEX) != FALSE;

    if (hasContentLength && hProgressBar)
    {
        SendMessage(hProgressBar, PBM_SETRANGE32, 0, (LPARAM)dwContentLength);
        SendMessage(hProgressBar, PBM_SETPOS, 0, 0);
    }

    // Scarica e scrivi su disco
    std::ofstream outFile(destPath, std::ios::binary);
    if (!outFile.is_open())
    {
        WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession);
        return false;
    }

    DWORD dwTotalRead = 0;
    DWORD dwBytesAvailable = 0;

    do
    {
        dwBytesAvailable = 0;
        if (!WinHttpQueryDataAvailable(hRequest, &dwBytesAvailable))
            break;

        if (dwBytesAvailable == 0)
            break;

        std::vector<char> buffer(dwBytesAvailable);
        DWORD dwBytesRead = 0;
        if (WinHttpReadData(hRequest, buffer.data(), dwBytesAvailable, &dwBytesRead))
        {
            outFile.write(buffer.data(), dwBytesRead);
            dwTotalRead += dwBytesRead;

            if (hProgressBar)
                SendMessage(hProgressBar, PBM_SETPOS, (WPARAM)dwTotalRead, 0);

            if (hStatusLabel)
            {
                wchar_t statusText[128];
                if (hasContentLength && dwContentLength > 0)
                {
                    int pct = (int)(((__int64)dwTotalRead * 100) / dwContentLength);
                    swprintf_s(statusText, L"Download in corso... %d%%  (%u KB / %u KB)",
                        pct, dwTotalRead / 1024, dwContentLength / 1024);
                }
                else
                {
                    swprintf_s(statusText, L"Download in corso... %u KB", dwTotalRead / 1024);
                }
                SetWindowTextW(hStatusLabel, statusText);
            }
        }
    } while (dwBytesAvailable > 0);

    outFile.close();
    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    return (dwTotalRead > 0);
}

INT_PTR CALLBACK AutoUpdater::DownloadDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static DownloadContext* pCtx = nullptr;

    switch (msg)
    {
    case WM_INITDIALOG:
    {
        pCtx = (DownloadContext*)lParam;
        pCtx->hDlg = hDlg;

        // Centra la finestra sul parent
        HWND hParent = GetParent(hDlg);
        if (hParent)
        {
            RECT rcParent, rcDlg;
            GetWindowRect(hParent, &rcParent);
            GetWindowRect(hDlg, &rcDlg);
            int x = rcParent.left + ((rcParent.right - rcParent.left) - (rcDlg.right - rcDlg.left)) / 2;
            int y = rcParent.top + ((rcParent.bottom - rcParent.top) - (rcDlg.bottom - rcDlg.top)) / 2;
            SetWindowPos(hDlg, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
        }

        // Avvia il download in un thread separato
        DownloadContext* ctx = pCtx;
        std::thread([ctx]()
        {
            ctx->success = HttpsDownloadFile(ctx->url, ctx->destPath,
                ctx->hProgressBar, ctx->hStatusLabel);

            PostMessage(ctx->hDlg, WM_APP + 200, 0, 0);
        }).detach();

        return TRUE;
    }

    case WM_APP + 200: // Download completato
    {
        if (pCtx && pCtx->success)
        {
            SetWindowTextW(pCtx->hStatusLabel, L"Download completato! Avvio installazione...");
            EndDialog(hDlg, IDOK);
        }
        else
        {
            SetWindowTextW(pCtx->hStatusLabel, L"Errore durante il download.");
            MessageBoxW(hDlg, L"Impossibile scaricare l'aggiornamento.\nRiprova piu' tardi.",
                L"Errore Download", MB_OK | MB_ICONERROR);
            EndDialog(hDlg, IDCANCEL);
        }
        return TRUE;
    }

    case WM_COMMAND:
        if (LOWORD(wParam) == IDCANCEL)
        {
            if (pCtx) pCtx->cancelled = true;
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }
        break;

    case WM_CLOSE:
        if (pCtx) pCtx->cancelled = true;
        EndDialog(hDlg, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

void AutoUpdater::DownloadAndInstall(HWND hParent, const UpdateInfo& info)
{
    if (info.installerUrl.empty())
        return;

    // Prepara il percorso temporaneo per il download
    wchar_t tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempPath);
    std::wstring destPath = std::wstring(tempPath) + L"FatturaView_Setup.exe";

    // Elimina un eventuale download precedente
    DeleteFileW(destPath.c_str());

    // Crea il dialog di download a runtime (senza risorse .rc)
    // Layout: Label status + Progress bar + pulsante Annulla
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtr(hParent, GWLP_HINSTANCE);

    // Template del dialog in memoria
    struct DialogTemplate
    {
        DLGTEMPLATE tmpl;
        WORD menu, classId, title;
    };

    // Usiamo un approccio piu' semplice: creiamo la finestra manualmente
    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"#32770", L"Download Aggiornamento",
        WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU,
        0, 0, 420, 160, hParent, NULL, hInst, NULL);

    if (!hDlg) return;

    // Centra sul parent
    RECT rcParent, rcDlg;
    GetWindowRect(hParent, &rcParent);
    GetWindowRect(hDlg, &rcDlg);
    int x = rcParent.left + ((rcParent.right - rcParent.left) - (rcDlg.right - rcDlg.left)) / 2;
    int y = rcParent.top + ((rcParent.bottom - rcParent.top) - (rcDlg.bottom - rcDlg.top)) / 2;
    SetWindowPos(hDlg, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

    // Font di sistema
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    // Label di stato
    HWND hStatusLabel = CreateWindowExW(0, L"STATIC", L"Download in corso...",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        15, 15, 380, 20, hDlg, NULL, hInst, NULL);
    SendMessage(hStatusLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Progress bar
    INITCOMMONCONTROLSEX icex = { sizeof(icex), ICC_PROGRESS_CLASS };
    InitCommonControlsEx(&icex);

    HWND hProgressBar = CreateWindowExW(0, PROGRESS_CLASSW, NULL,
        WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
        15, 45, 380, 25, hDlg, NULL, hInst, NULL);
    SendMessage(hProgressBar, PBM_SETRANGE32, 0, 100);
    SendMessage(hProgressBar, PBM_SETPOS, 0, 0);

    // Label percentuale
    HWND hPctLabel = CreateWindowExW(0, L"STATIC", L"",
        WS_CHILD | WS_VISIBLE | SS_CENTER,
        15, 75, 380, 18, hDlg, NULL, hInst, NULL);
    SendMessage(hPctLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Pulsante Annulla
    HWND hCancelBtn = CreateWindowExW(0, L"BUTTON", L"Annulla",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        165, 100, 90, 28, hDlg, (HMENU)IDCANCEL, hInst, NULL);
    SendMessage(hCancelBtn, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Disabilita la finestra padre
    EnableWindow(hParent, FALSE);
    ShowWindow(hDlg, SW_SHOW);
    UpdateWindow(hDlg);

    // Context per il download
    DownloadContext ctx = {};
    ctx.hDlg = hDlg;
    ctx.hProgressBar = hProgressBar;
    ctx.hStatusLabel = hStatusLabel;
    ctx.url = info.installerUrl;
    ctx.destPath = destPath;
    ctx.success = false;
    ctx.cancelled = false;

    // Avvia il download in un thread separato
    bool downloadDone = false;
    std::thread downloadThread([&ctx, &downloadDone]()
    {
        ctx.success = HttpsDownloadFile(ctx.url, ctx.destPath,
            ctx.hProgressBar, ctx.hStatusLabel);
        downloadDone = true;
        PostMessage(ctx.hDlg, WM_APP + 200, 0, 0);
    });
    downloadThread.detach();

    // Message loop per la finestra di download
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        if (msg.message == WM_APP + 200 && msg.hwnd == hDlg)
            break;

        // Gestisci click su Annulla
        if (msg.message == WM_COMMAND && msg.hwnd == hDlg && LOWORD(msg.wParam) == IDCANCEL)
        {
            ctx.cancelled = true;
            break;
        }

        // Gestisci click sul pulsante Annulla (figlio)
        if (msg.message == WM_COMMAND && LOWORD(msg.wParam) == IDCANCEL)
        {
            ctx.cancelled = true;
            break;
        }

        TranslateMessage(&msg);
        DispatchMessage(&msg);

        if (downloadDone)
            break;
    }

    // Riabilita la finestra padre e chiudi il dialog
    EnableWindow(hParent, TRUE);
    DestroyWindow(hDlg);
    SetForegroundWindow(hParent);

    if (ctx.cancelled)
    {
        DeleteFileW(destPath.c_str());
        return;
    }

    if (ctx.success)
    {
        // Avvia l'installer e chiudi l'app
        MessageBoxW(hParent,
            L"Download completato!\n\n"
            L"L'installer verra' avviato e FatturaView si chiudera'.\n"
            L"Al termine dell'installazione, riavvia FatturaView.",
            L"Aggiornamento Pronto", MB_OK | MB_ICONINFORMATION);

        ShellExecuteW(NULL, L"open", destPath.c_str(), NULL, NULL, SW_SHOWNORMAL);

        // Chiudi l'applicazione
        PostMessage(hParent, WM_CLOSE, 0, 0);
    }
    else
    {
        DeleteFileW(destPath.c_str());
        MessageBoxW(hParent,
            L"Impossibile scaricare l'aggiornamento.\nRiprova piu' tardi.",
            L"Errore Download", MB_OK | MB_ICONERROR);
    }
}
