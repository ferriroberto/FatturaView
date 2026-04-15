#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include <exdisp.h>
#include <mshtml.h>
#include <mshtmhst.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "uuid.lib")

class FatturaViewer
{
public:
    static HWND CreateBrowserControl(HWND hParent, HINSTANCE hInst, int x, int y, int width, int height);
    static void NavigateToHTML(HWND hBrowser, const std::wstring& htmlFilePath);
    static void NavigateToString(HWND hBrowser, const std::wstring& htmlContent);
    static void SaveHTMLToFile(const std::wstring& htmlContent, const std::wstring& filePath);
    static void CreateWelcomePage(const std::wstring& filePath);
    static std::wstring GetWelcomePageHTML();
    static void ShowWelcomePage(HWND hBrowser);
    static void CleanupBrowser(HWND hBrowser);
    static void PrintDocument(HWND hBrowser);
    static void PrintDocumentSilent(HWND hBrowser);
    static void SetZoom(HWND hBrowser, int zoomPercent);
    static void ShowNotification(HWND hBrowser, const std::wstring& message, const std::wstring& type = L"info");
};
