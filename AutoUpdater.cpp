#include "AutoUpdater.h"
#include <winhttp.h>
#include <shellapi.h>
#include <sstream>
#include <thread>
#include <vector>
#include <string>

#pragma comment(lib, "winhttp.lib")

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

    msg += L"\nVuoi aprire la pagina di download?";

    int result = MessageBoxW(hParent, msg.c_str(), L"Aggiornamento Disponibile",
        MB_YESNO | MB_ICONINFORMATION);

    if (result == IDYES)
    {
        std::wstring url = info.installerUrl.empty() ? info.downloadUrl : info.installerUrl;
        if (url.empty())
            url = std::wstring(L"https://github.com/") + GITHUB_OWNER + L"/" + GITHUB_REPO + L"/releases/latest";

        ShellExecuteW(hParent, L"open", url.c_str(), NULL, NULL, SW_SHOWNORMAL);
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
