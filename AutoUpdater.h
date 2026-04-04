#pragma once

#include <string>
#include <windows.h>

// Informazioni su una release disponibile
struct UpdateInfo
{
    std::wstring latestVersion;   // Es. "1.3.0"
    std::wstring downloadUrl;     // URL della release page
    std::wstring releaseNotes;    // Note di rilascio (body della release)
    std::wstring installerUrl;    // URL diretto del setup .exe (se presente)
    bool updateAvailable = false;
};

class AutoUpdater
{
public:
    // Proprietario GitHub e nome repo
    static constexpr const wchar_t* GITHUB_OWNER = L"ferriroberto";
    static constexpr const wchar_t* GITHUB_REPO  = L"FatturaView";

    // Controlla se c'e' un aggiornamento disponibile (sincrono, chiamare da thread)
    static bool CheckForUpdates(const std::wstring& currentVersion, UpdateInfo& info);

    // Mostra un dialog con le info dell'aggiornamento e chiede all'utente se vuole scaricare
    static void ShowUpdateDialog(HWND hParent, const UpdateInfo& info, const std::wstring& currentVersion);

    // Mostra un dialog "nessun aggiornamento disponibile"
    static void ShowNoUpdateDialog(HWND hParent, const std::wstring& currentVersion);

    // Controlla aggiornamenti in background (thread separato) e notifica la finestra
    // Invia WM_APP_UPDATE_AVAILABLE alla finestra se trova un aggiornamento
    static void CheckInBackground(HWND hNotifyWnd, const std::wstring& currentVersion);

    // Messaggio personalizzato per notifica aggiornamento disponibile
    static constexpr UINT WM_APP_UPDATE_AVAILABLE = WM_APP + 100;

private:
    // Esegue una richiesta HTTPS GET e restituisce il body come stringa
    // httpStatus riceve il codice HTTP (200, 404, ecc.) o 0 se la connessione fallisce
    static std::string HttpsGet(const std::wstring& host, const std::wstring& path, int& httpStatus);

    // Confronta due versioni "x.y.z", restituisce true se remote > local
    static bool IsNewerVersion(const std::wstring& local, const std::wstring& remote);

    // Estrae un valore stringa da un JSON minimale (senza libreria JSON)
    static std::wstring JsonGetString(const std::string& json, const std::string& key);

    // Cerca l'URL del primo asset .exe nella release
    static std::wstring JsonGetInstallerAssetUrl(const std::string& json);
};
