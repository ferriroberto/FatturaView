// xmlRead.cpp : Definisce il punto di ingresso dell'applicazione.
//

#include "framework.h"
#include "xmlRead.h"
#include "FatturaElettronicaParser.h"
#include "FatturaViewer.h"
#include <commdlg.h>
#include <shellapi.h>
#include <vector>
#include <string>
#include <algorithm>
#include <commctrl.h>

#pragma comment(lib, "comctl32.lib")

#define MAX_LOADSTRING 100
#define TOOLBAR_HEIGHT 50

// Variabili globali:
HINSTANCE hInst;                                // istanza corrente
WCHAR szTitle[MAX_LOADSTRING];                  // Testo della barra del titolo
WCHAR szWindowClass[MAX_LOADSTRING];            // nome della classe della finestra principale
std::vector<std::wstring> g_xmlFiles;           // Lista dei file XML estratti
std::vector<FatturaInfo> g_fattureInfo;         // Informazioni dettagliate delle fatture
std::vector<int> g_listBoxToFatturaIndex;       // Mappa indice ListBox -> indice g_fattureInfo (-1 = intestazione)
std::wstring g_currentXmlPath;                  // Percorso del file XML corrente
std::wstring g_extractPath;                     // Percorso di estrazione
std::wstring g_lastXsltPath;                    // Foglio di stile corrente
std::wstring g_welcomePath;                     // Percorso file welcome
std::wstring g_appPath;                         // Percorso della cartella dell'applicazione
FatturaElettronicaParser* g_pParser = nullptr;  // Parser fatture
HWND g_hListBox = NULL;                         // ListBox per le fatture
HWND g_hMainWnd = NULL;                         // Finestra principale
HWND g_hBrowserWnd = NULL;                      // Finestra browser per visualizzazione
HWND g_hToolbar = NULL;                         // Toolbar
HWND g_hStatusBar = NULL;                       // Status bar
HFONT g_hFontNormal = NULL;                     // Font normale per ListBox
HFONT g_hFontBold = NULL;                       // Font grassetto per intestazioni

#define APP_VERSION L"1.0.0"
#define APP_AUTHOR L"Roberto Ferri"

// Dichiarazioni con prototipo di funzioni incluse in questo modulo di codice:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    StylesheetSelector(HWND, UINT, WPARAM, LPARAM);
void                OnOpenZip(HWND hWnd);
void                OnApplyStylesheet(HWND hWnd);
void                OnPrintCurrent(HWND hWnd);
void                OnPrintAll(HWND hWnd);
void                OnPrintSelected(HWND hWnd);
void                LoadFattureList(HWND hWnd);
std::wstring        GetTempExtractionPath();
std::wstring        OpenFileDialog(HWND hWnd, const WCHAR* filter, const WCHAR* title);
HWND                CreateToolbar(HWND hParent, HINSTANCE hInst);
std::wstring        GetApplicationPath();
void                LoadStylesheetPreference();
void                SaveStylesheetPreference(const std::wstring& path);
std::vector<std::wstring> GetStylesheetsFromResources();
std::wstring        GetResourcesPath();

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Inizializza il parser
    g_pParser = new FatturaElettronicaParser();
    g_extractPath = GetTempExtractionPath();
    g_appPath = GetApplicationPath();

    // Carica la preferenza del foglio di stile
    LoadStylesheetPreference();

    // Inizializza Common Controls per la toolbar
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_BAR_CLASSES;
    InitCommonControlsEx(&icex);

    // Inizializzare le stringhe globali
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_XMLREAD, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Eseguire l'inizializzazione dall'applicazione:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_XMLREAD));

    MSG msg;

    // Ciclo di messaggi principale:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // Cleanup
    if (g_pParser)
    {
        delete g_pParser;
        g_pParser = nullptr;
    }

    return (int) msg.wParam;
}



//
//  FUNZIONE: MyRegisterClass()
//
//  SCOPO: Registra la classe di finestre.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_XMLREAD));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = CreateSolidBrush(RGB(243, 243, 243)); // Sfondo grigio chiaro moderno
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_XMLREAD);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNZIONE: InitInstance(HINSTANCE, int)
//
//   SCOPO: Salva l'handle di istanza e crea la finestra principale
//
//   COMMENTI:
//
//        In questa funzione l'handle di istanza viene salvato in una variabile globale e
//        viene creata e visualizzata la finestra principale del programma.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Archivia l'handle di istanza nella variabile globale

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, 1200, 700, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   g_hMainWnd = hWnd;

   // Crea font per la ListBox (owner-drawn)
   g_hFontNormal = CreateFont(
      17, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
      CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
      L"Segoe UI");

   g_hFontBold = CreateFont(
      18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
      CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
      L"Segoe UI");

   // Crea la toolbar
   g_hToolbar = CreateToolbar(hWnd, hInstance);

   // Crea la status bar
   g_hStatusBar = CreateWindowEx(0, STATUSCLASSNAME, NULL,
       WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
       0, 0, 0, 0,
       hWnd, (HMENU)IDC_STATUSBAR, hInstance, NULL);

   if (g_hStatusBar)
   {
       // Configura le parti della status bar (3 sezioni)
       int parts[] = { 300, 500, -1 };
       SendMessage(g_hStatusBar, SB_SETPARTS, 3, (LPARAM)parts);

       // Testo nelle sezioni
       SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Pronto");
       std::wstring authorInfo = std::wstring(L"Autore: ") + APP_AUTHOR;
       SendMessage(g_hStatusBar, SB_SETTEXT, 1, (LPARAM)authorInfo.c_str());
       std::wstring versionInfo = std::wstring(L"Versione ") + APP_VERSION;
       SendMessage(g_hStatusBar, SB_SETTEXT, 2, (LPARAM)versionInfo.c_str());
   }

   // Ottieni dimensioni area client
   RECT rcClient;
   GetClientRect(hWnd, &rcClient);

   // Ottieni altezza status bar
   RECT rcStatus;
   GetWindowRect(g_hStatusBar, &rcStatus);
   int statusHeight = rcStatus.bottom - rcStatus.top;

   int leftPanelWidth = 550;
   int margin = 15;
   int topOffset = TOOLBAR_HEIGHT + margin;

   // Crea una ListBox per mostrare le fatture (pannello sinistro) con selezione multipla
   // Usa LBS_OWNERDRAWFIXED per disegnare manualmente gli elementi con colori
   g_hListBox = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", NULL,
      WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY | LBS_EXTENDEDSEL | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
      margin, topOffset, 
      leftPanelWidth, 
      rcClient.bottom - topOffset - margin - statusHeight,
      hWnd, (HMENU)IDC_LISTBOX_FATTURE, hInstance, NULL);

   // Imposta l'altezza degli elementi (più alta per migliore leggibilità)
   SendMessage(g_hListBox, LB_SETITEMHEIGHT, 0, (LPARAM)28);

   // Crea il controllo browser per visualizzazione (pannello destro)
   g_hBrowserWnd = FatturaViewer::CreateBrowserControl(hWnd, hInstance, 
      leftPanelWidth + 2 * margin, topOffset,
      rcClient.right - leftPanelWidth - 3 * margin,
      rcClient.bottom - topOffset - margin - statusHeight);

   // Verifica che il browser sia stato creato
   if (!g_hBrowserWnd)
   {
       MessageBoxW(hWnd, L"Impossibile creare il controllo browser!", L"Errore", MB_OK | MB_ICONERROR);
   }
   else
   {
       // Usa un timer per navigare dopo che la finestra è completamente inizializzata
       SetTimer(hWnd, 1, 500, NULL); // Timer ID 1, 500ms delay (aumentato)
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNZIONE: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  SCOPO: Elabora i messaggi per la finestra principale.
//
//  WM_COMMAND  - elabora il menu dell'applicazione
//  WM_PAINT    - Disegna la finestra principale
//  WM_DESTROY  - inserisce un messaggio di uscita e restituisce un risultato
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_DRAWITEM:
        {
            // Disegna gli elementi della ListBox con colori personalizzati
            LPDRAWITEMSTRUCT pDIS = (LPDRAWITEMSTRUCT)lParam;

            if (pDIS->CtlID == IDC_LISTBOX_FATTURE)
            {
                // Ottieni il testo dell'elemento
                int textLen = (int)SendMessage(pDIS->hwndItem, LB_GETTEXTLEN, pDIS->itemID, 0);
                if (textLen == LB_ERR)
                    return TRUE;

                std::wstring itemText(textLen + 1, L'\0');
                SendMessage(pDIS->hwndItem, LB_GETTEXT, pDIS->itemID, (LPARAM)itemText.data());

                // Determina se è un'intestazione (indice -1 nel mapping o inizia con ►)
                bool isHeader = false;
                bool isEmptyLine = itemText.empty() || itemText == L"";

                if (pDIS->itemID < g_listBoxToFatturaIndex.size())
                {
                    int fatturaIndex = g_listBoxToFatturaIndex[pDIS->itemID];
                    isHeader = (fatturaIndex == -1 && !isEmptyLine);
                }

                // Colori moderni
                COLORREF clrBackground;
                COLORREF clrText;
                HFONT hFont;

                if (isEmptyLine)
                {
                    // Riga vuota di separazione
                    clrBackground = RGB(230, 230, 230);
                    clrText = RGB(0, 0, 0);
                    hFont = g_hFontNormal;
                }
                else if (isHeader)
                {
                    // Intestazione emittente - Blu moderno con testo bianco
                    if (pDIS->itemState & ODS_SELECTED)
                    {
                        clrBackground = RGB(0, 102, 204);  // Blu più scuro se selezionato
                        clrText = RGB(255, 255, 255);
                    }
                    else
                    {
                        clrBackground = RGB(0, 120, 215);  // Blu Windows 11
                        clrText = RGB(255, 255, 255);
                    }
                    hFont = g_hFontBold;
                }
                else
                {
                    // Fattura normale - Sfondo alternato
                    if (pDIS->itemState & ODS_SELECTED)
                    {
                        // Selezionato - Arancione chiaro
                        clrBackground = RGB(255, 224, 192);
                        clrText = RGB(0, 0, 0);
                    }
                    else
                    {
                        // Alternanza grigio molto chiaro / bianco
                        clrBackground = (pDIS->itemID % 2 == 0) ? RGB(255, 255, 255) : RGB(248, 248, 248);
                        clrText = RGB(64, 64, 64);
                    }
                    hFont = g_hFontNormal;
                }

                // Disegna lo sfondo
                HBRUSH hBrush = CreateSolidBrush(clrBackground);
                FillRect(pDIS->hDC, &pDIS->rcItem, hBrush);
                DeleteObject(hBrush);

                // Disegna il focus rectangle
                if (pDIS->itemState & ODS_FOCUS)
                {
                    DrawFocusRect(pDIS->hDC, &pDIS->rcItem);
                }

                // Imposta il colore del testo
                SetTextColor(pDIS->hDC, clrText);
                SetBkMode(pDIS->hDC, TRANSPARENT);

                // Disegna il testo con padding
                RECT rcText = pDIS->rcItem;
                rcText.left += 8;  // Padding sinistro
                rcText.right -= 8; // Padding destro

                // Per le fatture normali (non header), applica grassetto solo al cessionario
                if (!isHeader && !isEmptyLine)
                {
                    // Cerca i due delimitatori "|" per identificare il nome del cessionario
                    size_t firstPipe = itemText.find(L'|');
                    size_t secondPipe = itemText.find(L'|', firstPipe + 1);

                    if (firstPipe != std::wstring::npos && secondPipe != std::wstring::npos)
                    {
                        // Estrai le tre parti del testo
                        std::wstring part1 = itemText.substr(0, firstPipe + 1);  // "  N. 123 del 2024-01-15 |"
                        std::wstring part2 = itemText.substr(firstPipe + 1, secondPipe - firstPipe - 1);  // " Cessionario "
                        std::wstring part3 = itemText.substr(secondPipe);  // "| € 1000.00"

                        // Prepara il rettangolo per il disegno
                        RECT rcPart = rcText;
                        SIZE textSize;

                        // Disegna la prima parte (numero e data) con font normale
                        HFONT hOldFont = (HFONT)SelectObject(pDIS->hDC, g_hFontNormal);
                        DrawText(pDIS->hDC, part1.c_str(), -1, &rcPart, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_CALCRECT);
                        DrawText(pDIS->hDC, part1.c_str(), -1, &rcPart, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
                        GetTextExtentPoint32(pDIS->hDC, part1.c_str(), (int)part1.length(), &textSize);
                        rcPart.left += textSize.cx;

                        // Disegna la seconda parte (cessionario) con font grassetto
                        SelectObject(pDIS->hDC, g_hFontBold);
                        DrawText(pDIS->hDC, part2.c_str(), -1, &rcPart, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_CALCRECT);
                        DrawText(pDIS->hDC, part2.c_str(), -1, &rcPart, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
                        GetTextExtentPoint32(pDIS->hDC, part2.c_str(), (int)part2.length(), &textSize);
                        rcPart.left += textSize.cx;

                        // Disegna la terza parte (importo) con font normale
                        SelectObject(pDIS->hDC, g_hFontNormal);
                        DrawText(pDIS->hDC, part3.c_str(), -1, &rcPart, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

                        SelectObject(pDIS->hDC, hOldFont);
                    }
                    else
                    {
                        // Fallback: disegna normalmente se non trova i delimitatori
                        HFONT hOldFont = (HFONT)SelectObject(pDIS->hDC, hFont);
                        DrawText(pDIS->hDC, itemText.c_str(), -1, &rcText, 
                            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
                        SelectObject(pDIS->hDC, hOldFont);
                    }
                }
                else
                {
                    // Per header e righe vuote, usa il font predefinito
                    HFONT hOldFont = (HFONT)SelectObject(pDIS->hDC, hFont);
                    DrawText(pDIS->hDC, itemText.c_str(), -1, &rcText, 
                        DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
                    SelectObject(pDIS->hDC, hOldFont);
                }

                return TRUE;
            }
        }
        break;
    case WM_TIMER:
        {
            if (wParam == 1) // Timer per il caricamento del welcome
            {
                KillTimer(hWnd, 1); // Ferma il timer

                // Mostra loader: cursore wait e messaggio status bar
                HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));
                if (g_hStatusBar)
                    SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Caricamento interfaccia...");

                // Naviga al file welcome
                if (g_hBrowserWnd)
                {
                    if (!g_welcomePath.empty() && GetFileAttributesW(g_welcomePath.c_str()) != INVALID_FILE_ATTRIBUTES)
                    {
                        // Prova prima con il file
                        FatturaViewer::NavigateToHTML(g_hBrowserWnd, g_welcomePath);
                    }
                    else
                    {
                        // Se il file non funziona, carica l'HTML direttamente in memoria
                        FatturaViewer::NavigateToString(g_hBrowserWnd, FatturaViewer::GetWelcomePageHTML());
                    }
                }

                // Attendi un momento per il caricamento
                Sleep(500);

                // Ripristina cursore e status bar
                SetCursor(hOldCursor);
                if (g_hStatusBar)
                    SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Pronto");
            }
        }
        break;
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Analizzare le selezioni di menu:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            case IDM_OPEN_ZIP:
                OnOpenZip(hWnd);
                break;
            case IDM_APPLY_MINISTERO:
            case IDM_APPLY_ASSOSOFTWARE:
            case IDM_CHANGE_STYLESHEET:
                OnApplyStylesheet(hWnd);
                break;
            case IDM_PRINT_CURRENT:
                OnPrintCurrent(hWnd);
                break;
            case IDM_PRINT_ALL:
                OnPrintAll(hWnd);
                break;
            case IDM_PRINT_SELECTED:
                OnPrintSelected(hWnd);
                break;
            case IDM_SETTINGS_SELECT_STYLESHEET:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_STYLESHEET_SELECTOR), hWnd, StylesheetSelector);
                break;
            case IDM_TEST_BROWSER:
                {
                    // Test del controllo browser
                    std::wstring testPath = g_extractPath + L"test.html";
                    std::wstring testHtml = L"<html><body><h1>TEST BROWSER</h1><p>Se vedi questo, il browser funziona!</p></body></html>";
                    FatturaViewer::SaveHTMLToFile(testHtml, testPath);

                    if (g_hBrowserWnd)
                    {
                        FatturaViewer::NavigateToHTML(g_hBrowserWnd, testPath);
                    }
                    else
                    {
                        MessageBoxW(hWnd, L"Controllo browser non inizializzato!", L"Test", MB_OK | MB_ICONERROR);
                    }
                }
                break;
            case IDC_LISTBOX_FATTURE:
                if (HIWORD(wParam) == LBN_SELCHANGE)
                {
                    int listBoxIndex = (int)SendMessage(g_hListBox, LB_GETCURSEL, 0, 0);
                    if (listBoxIndex != LB_ERR && listBoxIndex < (int)g_listBoxToFatturaIndex.size())
                    {
                        int fatturaIndex = g_listBoxToFatturaIndex[listBoxIndex];
                        // -1 significa intestazione o riga vuota, ignora
                        if (fatturaIndex >= 0 && fatturaIndex < (int)g_fattureInfo.size())
                        {
                                g_currentXmlPath = g_fattureInfo[fatturaIndex].filePath;

                                // Visualizza automaticamente la fattura con un singolo click
                                OnApplyStylesheet(hWnd);
                            }
                    }
                }
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_SIZE:
        {
            // Ridimensiona i controlli quando la finestra viene ridimensionata
            RECT rcClient;
            GetClientRect(hWnd, &rcClient);

            int leftPanelWidth = 550;
            int margin = 15;
            int topOffset = TOOLBAR_HEIGHT + margin;

            // Ridimensiona la toolbar
            if (g_hToolbar)
            {
                SendMessage(g_hToolbar, TB_AUTOSIZE, 0, 0);
            }

            // Ridimensiona la status bar
            if (g_hStatusBar)
            {
                SendMessage(g_hStatusBar, WM_SIZE, 0, 0);
            }

            // Ottieni altezza status bar
            RECT rcStatus;
            int statusHeight = 0;
            if (g_hStatusBar)
            {
                GetWindowRect(g_hStatusBar, &rcStatus);
                statusHeight = rcStatus.bottom - rcStatus.top;
            }

            if (g_hListBox)
            {
                MoveWindow(g_hListBox, 
                    margin, topOffset, 
                    leftPanelWidth, 
                    rcClient.bottom - topOffset - margin - statusHeight, 
                    TRUE);
            }

            if (g_hBrowserWnd)
            {
                MoveWindow(g_hBrowserWnd,
                    leftPanelWidth + 2 * margin, topOffset,
                    rcClient.right - leftPanelWidth - 3 * margin,
                    rcClient.bottom - topOffset - margin - statusHeight,
                    TRUE);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        // Elimina i font creati
        if (g_hFontNormal)
            DeleteObject(g_hFontNormal);
        if (g_hFontBold)
            DeleteObject(g_hFontBold);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Gestore di messaggi per la finestra Informazioni su.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Gestore di messaggi per la finestra di selezione foglio di stile
INT_PTR CALLBACK StylesheetSelector(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    static std::vector<std::wstring> stylesheets;
    static std::wstring resourcesPath;

    switch (message)
    {
    case WM_INITDIALOG:
        {
            // Trova il percorso corretto della cartella Resources
            resourcesPath = GetResourcesPath();

            // Carica i fogli di stile dalla cartella Resources
            stylesheets = GetStylesheetsFromResources();

            HWND hList = GetDlgItem(hDlg, IDC_STYLESHEET_LIST);

            if (stylesheets.empty())
            {
                MessageBoxW(hDlg, 
                    L"Nessun foglio di stile trovato nella cartella Resources!\n\nAssicurati che i file .xsl siano nella cartella Resources\\",
                    L"Attenzione", MB_OK | MB_ICONWARNING);
                EndDialog(hDlg, IDCANCEL);
                return (INT_PTR)TRUE;
            }

            // Popola la listbox con i nomi dei file
            int selectedIndex = -1;
            for (size_t i = 0; i < stylesheets.size(); i++)
            {
                SendMessageW(hList, LB_ADDSTRING, 0, (LPARAM)stylesheets[i].c_str());

                // Se questo è il foglio corrente, selezionalo
                std::wstring currentName = g_lastXsltPath;
                size_t pos = currentName.find_last_of(L"\\");
                if (pos != std::wstring::npos)
                    currentName = currentName.substr(pos + 1);

                if (stylesheets[i] == currentName)
                    selectedIndex = (int)i;
            }

            // Seleziona il foglio corrente o il primo
            if (selectedIndex >= 0)
                SendMessageW(hList, LB_SETCURSEL, selectedIndex, 0);
            else if (!stylesheets.empty())
                SendMessageW(hList, LB_SETCURSEL, 0, 0);

            // Mostra info sul foglio selezionato
            if (!stylesheets.empty())
            {
                int idx = (selectedIndex >= 0) ? selectedIndex : 0;
                std::wstring info = L"File: " + resourcesPath + stylesheets[idx];
                SetDlgItemTextW(hDlg, IDC_STYLESHEET_PREVIEW, info.c_str());
            }

            return (INT_PTR)TRUE;
        }

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_STYLESHEET_LIST && HIWORD(wParam) == LBN_SELCHANGE)
        {
            // Aggiorna l'anteprima quando cambia la selezione
            HWND hList = GetDlgItem(hDlg, IDC_STYLESHEET_LIST);
            int sel = (int)SendMessageW(hList, LB_GETCURSEL, 0, 0);
            if (sel >= 0 && sel < (int)stylesheets.size())
            {
                std::wstring info = L"File: " + resourcesPath + stylesheets[sel];
                SetDlgItemTextW(hDlg, IDC_STYLESHEET_PREVIEW, info.c_str());
            }
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDOK)
        {
            // Salva la selezione
            HWND hList = GetDlgItem(hDlg, IDC_STYLESHEET_LIST);
            int sel = (int)SendMessageW(hList, LB_GETCURSEL, 0, 0);

            if (sel >= 0 && sel < (int)stylesheets.size())
            {
                std::wstring fullPath = resourcesPath + stylesheets[sel];
                g_lastXsltPath = fullPath;
                SaveStylesheetPreference(fullPath);

                // Aggiorna la status bar nella finestra principale
                if (g_hStatusBar)
                {
                    std::wstring statusMsg = L"⚙ Foglio di stile: " + stylesheets[sel];
                    SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)statusMsg.c_str());
                }
            }

            EndDialog(hDlg, IDOK);
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

std::wstring OpenFileDialog(HWND hWnd, const WCHAR* filter, const WCHAR* title)
{
    WCHAR szFile[MAX_PATH] = { 0 };
    OPENFILENAMEW ofn = { 0 };

    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = filter;
    ofn.lpstrTitle = title;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileNameW(&ofn))
    {
        return std::wstring(szFile);
    }

    return L"";
}

std::wstring GetTempExtractionPath()
{
    WCHAR tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempPath);
    std::wstring extractPath = std::wstring(tempPath) + L"FattureElettroniche\\";
    CreateDirectoryW(extractPath.c_str(), NULL);
    return extractPath;
}

void OnOpenZip(HWND hWnd)
{
    std::wstring zipPath = OpenFileDialog(hWnd, 
        L"File ZIP\0*.zip\0Tutti i file\0*.*\0", 
        L"Seleziona archivio ZIP con fatture elettroniche");

    if (zipPath.empty())
        return;

    if (!g_pParser)
        return;

    // ===== INIZIO LOADER =====

    // Disabilita controlli durante l'estrazione
    EnableWindow(g_hListBox, FALSE);
    EnableWindow(g_hToolbar, FALSE);

    // Cambia cursore in wait (clessidra)
    HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));

    // Aggiorna status bar con animazione
    if (g_hStatusBar)
    {
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"⏳ Estrazione ZIP in corso...");
        UpdateWindow(g_hStatusBar);
    }

    // Aggiorna titolo finestra
    std::wstring oldTitle;
    WCHAR titleBuffer[256];
    GetWindowTextW(hWnd, titleBuffer, 256);
    oldTitle = titleBuffer;
    SetWindowTextW(hWnd, L"⏳ Estrazione in corso - Attendere...");

    // Forza ridisegno
    UpdateWindow(hWnd);

    // Estrai il file ZIP
    bool success = g_pParser->ExtractZipFile(zipPath, g_extractPath);

    if (success)
    {
        // Aggiorna messaggio durante il caricamento della lista
        if (g_hStatusBar)
        {
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"📄 Caricamento fatture...");
            UpdateWindow(g_hStatusBar);
        }

        // Carica la lista dei file XML
        LoadFattureList(hWnd);

        // ===== FINE LOADER - SUCCESSO =====

        // Ripristina titolo
        SetWindowTextW(hWnd, oldTitle.c_str());

        // Ripristina cursore
        SetCursor(hOldCursor);

        // Riabilita controlli
        EnableWindow(g_hListBox, TRUE);
        EnableWindow(g_hToolbar, TRUE);

        // Aggiorna status bar con il numero di fatture
        if (g_hStatusBar)
        {
            std::wstring statusText = L"✓ " + std::to_wstring(g_fattureInfo.size()) + L" fatture caricate con successo";
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)statusText.c_str());
        }

        MessageBoxW(hWnd, L"Fatture estratte con successo!", L"✓ Completato", MB_OK | MB_ICONINFORMATION);
    }
    else
    {
        // ===== FINE LOADER - ERRORE =====

        // Ripristina titolo
        SetWindowTextW(hWnd, oldTitle.c_str());

        // Ripristina cursore
        SetCursor(hOldCursor);

        // Riabilita controlli
        EnableWindow(g_hListBox, TRUE);
        EnableWindow(g_hToolbar, TRUE);

        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"✗ Errore estrazione");

        MessageBoxW(hWnd, L"Errore durante l'estrazione del file ZIP!", L"✗ Errore", MB_OK | MB_ICONERROR);
    }
}

void LoadFattureList(HWND hWnd)
{
    if (!g_pParser)
        return;

    // Svuota la listbox
    SendMessage(g_hListBox, LB_RESETCONTENT, 0, 0);
    g_xmlFiles.clear();
    g_fattureInfo.clear();
    g_listBoxToFatturaIndex.clear();

    // Ottieni i dati delle fatture dalla cartella di estrazione
    g_fattureInfo = g_pParser->GetFattureInfoFromFolder(g_extractPath);

    // Ordina le fatture prima per emittente, poi per numero
    std::sort(g_fattureInfo.begin(), g_fattureInfo.end(), 
        [](const FatturaInfo& a, const FatturaInfo& b) -> bool
        {
            // Prima ordina per emittente
            if (a.cedenteDenominazione != b.cedenteDenominazione)
                return a.cedenteDenominazione < b.cedenteDenominazione;

            // Se stesso emittente, ordina per numero fattura
            auto extractNumber = [](const std::wstring& str) -> int {
                int result = 0;
                for (wchar_t c : str) {
                    if (c >= L'0' && c <= L'9') {
                        result = result * 10 + (c - L'0');
                    }
                }
                return result;
            };

            int numA = extractNumber(a.numero);
            int numB = extractNumber(b.numero);

            if (numA != numB && numA != 0 && numB != 0) {
                return numA < numB;
            }

            return a.numero < b.numero;
        });

    // Mantieni anche la lista dei file per compatibilità
    for (const auto& info : g_fattureInfo)
    {
        g_xmlFiles.push_back(info.filePath);
    }

    // Aggiungi le informazioni formattate con raggruppamento per emittente
    std::wstring currentEmittente = L"";

    for (size_t i = 0; i < g_fattureInfo.size(); i++)
    {
        const auto& info = g_fattureInfo[i];

        // Se cambia l'emittente, aggiungi l'intestazione del gruppo
        if (info.cedenteDenominazione != currentEmittente)
        {
            currentEmittente = info.cedenteDenominazione;

            // Aggiungi una riga vuota di separazione (tranne per il primo gruppo)
            if (!currentEmittente.empty() && SendMessage(g_hListBox, LB_GETCOUNT, 0, 0) > 0)
            {
                SendMessage(g_hListBox, LB_ADDSTRING, 0, (LPARAM)L"");
                g_listBoxToFatturaIndex.push_back(-1);  // -1 = riga vuota
            }

            // Aggiungi l'intestazione dell'emittente
            std::wstring header = L"\x25BA " + currentEmittente;  // ► simbolo
            SendMessage(g_hListBox, LB_ADDSTRING, 0, (LPARAM)header.c_str());
            g_listBoxToFatturaIndex.push_back(-1);  // -1 = intestazione
        }

        // Aggiungi la fattura indentata
        SendMessage(g_hListBox, LB_ADDSTRING, 0, (LPARAM)info.GetDisplayText().c_str());
        g_listBoxToFatturaIndex.push_back((int)i);  // Mappa questo indice ListBox all'indice della fattura
    }
}

void OnApplyStylesheet(HWND hWnd)
{
    if (g_currentXmlPath.empty())
    {
        MessageBoxW(hWnd, L"Seleziona prima una fattura dalla lista!", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    if (!g_pParser)
        return;

    // ===== INIZIO LOADER =====
    HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));
    if (g_hStatusBar)
    {
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"⏳ Caricamento fattura...");
        UpdateWindow(g_hStatusBar);
    }

    // Carica il file XML
    if (!g_pParser->LoadXmlFile(g_currentXmlPath))
    {
        SetCursor(hOldCursor);
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"✗ Errore caricamento XML");
        MessageBoxW(hWnd, L"Errore nel caricamento del file XML!", L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

    // Usa il foglio di stile configurato
    std::wstring xsltPath = g_lastXsltPath;

    // Se non c'è nessun foglio configurato, chiedi di selezionarne uno
    if (xsltPath.empty() || GetFileAttributesW(xsltPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        SetCursor(hOldCursor);

        MessageBoxW(hWnd, 
            L"Nessun foglio di stile configurato!\n\nApri il menu Impostazioni per selezionare un foglio di stile.", 
            L"Configurazione richiesta", MB_OK | MB_ICONINFORMATION);

        // Apri automaticamente il dialog di selezione
        if (DialogBox(hInst, MAKEINTRESOURCE(IDD_STYLESHEET_SELECTOR), hWnd, StylesheetSelector) != IDOK)
        {
            if (g_hStatusBar)
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Operazione annullata");
            return;
        }

        xsltPath = g_lastXsltPath;
        if (xsltPath.empty())
        {
            if (g_hStatusBar)
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Nessun foglio selezionato");
            return;
        }

        hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));
    }

    if (g_hStatusBar)
    {
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"🎨 Applicazione foglio di stile...");
        UpdateWindow(g_hStatusBar);
    }

    // Applica la trasformazione
    std::wstring htmlOutput;
    if (g_pParser->ApplyXsltTransform(xsltPath, htmlOutput))
    {
        if (g_hStatusBar)
        {
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"📄 Rendering HTML...");
            UpdateWindow(g_hStatusBar);
        }

        // Salva l'HTML in un file temporaneo
        std::wstring htmlPath = g_extractPath + L"fattura_visualizzata.html";
        FatturaViewer::SaveHTMLToFile(htmlOutput, htmlPath);

        // Visualizza nel frame destro
        if (g_hBrowserWnd)
        {
            FatturaViewer::NavigateToHTML(g_hBrowserWnd, htmlPath);
        }

        // ===== FINE LOADER - SUCCESSO =====
        SetCursor(hOldCursor);
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"✓ Fattura visualizzata");
    }
    else
    {
        // ===== FINE LOADER - ERRORE =====
        SetCursor(hOldCursor);
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"✗ Errore trasformazione");
        MessageBoxW(hWnd, L"Errore durante l'applicazione del foglio di stile!", L"Errore", MB_OK | MB_ICONERROR);
    }
}

void OnPrintCurrent(HWND hWnd)
{
    if (g_currentXmlPath.empty())
    {
        MessageBoxW(hWnd, L"Seleziona prima una fattura dalla lista!", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    if (!g_pParser || !g_hBrowserWnd)
        return;

    // Verifica che ci sia un foglio di stile selezionato
    if (g_lastXsltPath.empty())
    {
        MessageBoxW(hWnd, L"Seleziona prima un foglio di stile visualizzando una fattura!", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    // Carica il file XML corrente
    if (!g_pParser->LoadXmlFile(g_currentXmlPath))
    {
        MessageBoxW(hWnd, L"Errore nel caricamento del file XML!", L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

    // Applica la trasformazione
    std::wstring htmlOutput;
    if (g_pParser->ApplyXsltTransform(g_lastXsltPath, htmlOutput))
    {
        // Salva l'HTML temporaneo per la stampa
        std::wstring htmlPath = g_extractPath + L"fattura_stampa.html";
        FatturaViewer::SaveHTMLToFile(htmlOutput, htmlPath);
        FatturaViewer::NavigateToHTML(g_hBrowserWnd, htmlPath);

        // Attendi che il documento sia caricato
        Sleep(1000);

        // Stampa con dialogo
        FatturaViewer::PrintDocument(g_hBrowserWnd);
    }
    else
    {
        MessageBoxW(hWnd, L"Errore durante la generazione della fattura per la stampa!", L"Errore", MB_OK | MB_ICONERROR);
    }
}

void OnPrintAll(HWND hWnd)
{
    if (g_fattureInfo.empty())
    {
        MessageBoxW(hWnd, L"Nessuna fattura da stampare! Apri prima un archivio ZIP.", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    if (!g_pParser || !g_hBrowserWnd)
        return;

    // Verifica che ci sia un foglio di stile selezionato
    if (g_lastXsltPath.empty())
    {
        MessageBoxW(hWnd, L"Seleziona prima un foglio di stile visualizzando una fattura!", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    // Chiedi conferma e modalità di stampa
    std::wstring msg = L"Stai per stampare " + std::to_wstring(g_fattureInfo.size()) + L" fatture.\n\n";
    msg += L"Scegli la modalita' di stampa:\n\n";
    msg += L"SI = Scegli la stampante (dialogo per la prima fattura)\n";
    msg += L"NO = Usa stampante predefinita (nessun dialogo)\n";
    msg += L"ANNULLA = Annulla operazione";

    int result = MessageBoxW(hWnd, msg.c_str(), L"Stampa massiva", MB_YESNOCANCEL | MB_ICONQUESTION);

    if (result == IDCANCEL)
        return;

    bool showDialogFirst = (result == IDYES);

    // Mostra un messaggio di avvio
    SetWindowTextW(hWnd, L"Stampa in corso...");

    int successCount = 0;
    int errorCount = 0;

    // Stampa tutte le fatture
    for (size_t i = 0; i < g_fattureInfo.size(); i++)
    {
        // Carica il file XML
        if (!g_pParser->LoadXmlFile(g_fattureInfo[i].filePath))
        {
            errorCount++;
            continue;
        }

        // Applica la trasformazione
        std::wstring htmlOutput;
        if (g_pParser->ApplyXsltTransform(g_lastXsltPath, htmlOutput))
        {
            // Salva l'HTML temporaneo
            std::wstring htmlPath = g_extractPath + L"fattura_stampa_" + std::to_wstring(i) + L".html";
            FatturaViewer::SaveHTMLToFile(htmlOutput, htmlPath);
            FatturaViewer::NavigateToHTML(g_hBrowserWnd, htmlPath);

            // Attendi il caricamento
            Sleep(1500);

            // Stampa: con dialogo solo per la prima se richiesto, poi sempre silent
            if (i == 0 && showDialogFirst)
            {
                FatturaViewer::PrintDocument(g_hBrowserWnd);
                Sleep(1000); // Attendi che l'utente confermi il dialogo
            }
            else
            {
                FatturaViewer::PrintDocumentSilent(g_hBrowserWnd);
                Sleep(500);
            }

            successCount++;
        }
        else
        {
            errorCount++;
        }

        // Aggiorna il titolo della finestra con il progresso
        std::wstring progress = L"Stampa in corso... (" + std::to_wstring(i + 1) + L"/" + std::to_wstring(g_fattureInfo.size()) + L")";
        SetWindowTextW(hWnd, progress.c_str());
    }

    // Ripristina il titolo originale
    SetWindowTextW(hWnd, szTitle);

    // Mostra il risultato
    std::wstring resultMsg = L"Stampa completata!\n\n";
    resultMsg += L"Fatture stampate: " + std::to_wstring(successCount) + L"\n";
    if (errorCount > 0)
    {
        resultMsg += L"Errori: " + std::to_wstring(errorCount);
    }

    MessageBoxW(hWnd, resultMsg.c_str(), L"Stampa massiva", MB_OK | MB_ICONINFORMATION);
}

void OnPrintSelected(HWND hWnd)
{
    if (g_fattureInfo.empty())
    {
        MessageBoxW(hWnd, L"Nessuna fattura disponibile! Apri prima un archivio ZIP.", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    if (!g_pParser || !g_hBrowserWnd)
        return;

    // Verifica che ci sia un foglio di stile selezionato
    if (g_lastXsltPath.empty())
    {
        MessageBoxW(hWnd, L"Seleziona prima un foglio di stile visualizzando una fattura!", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    // Ottieni gli indici delle fatture selezionate
    int selCount = (int)SendMessage(g_hListBox, LB_GETSELCOUNT, 0, 0);
    if (selCount == 0 || selCount == LB_ERR)
    {
        MessageBoxW(hWnd, L"Seleziona almeno una fattura dalla lista!\n\nUsa Ctrl+Click o Shift+Click per selezioni multiple.", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    // Alloca un array per gli indici selezionati (della ListBox)
    std::vector<int> selectedListBoxIndices(selCount);
    SendMessage(g_hListBox, LB_GETSELITEMS, selCount, (LPARAM)selectedListBoxIndices.data());

    // Converti gli indici ListBox in indici fattura (filtrando intestazioni)
    std::vector<int> selectedFatturaIndices;
    for (int listBoxIndex : selectedListBoxIndices)
    {
        if (listBoxIndex >= 0 && listBoxIndex < (int)g_listBoxToFatturaIndex.size())
        {
            int fatturaIndex = g_listBoxToFatturaIndex[listBoxIndex];
            if (fatturaIndex >= 0)  // -1 = intestazione, ignora
            {
                selectedFatturaIndices.push_back(fatturaIndex);
            }
        }
    }

    if (selectedFatturaIndices.empty())
    {
        MessageBoxW(hWnd, L"Nessuna fattura valida selezionata! Hai selezionato solo intestazioni.", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    // Chiedi conferma e modalità di stampa
    std::wstring msg = L"Stai per stampare " + std::to_wstring(selectedFatturaIndices.size()) + L" fatture selezionate.\n\n";
    msg += L"Scegli la modalita' di stampa:\n\n";
    msg += L"SI = Scegli la stampante (dialogo per la prima fattura)\n";
    msg += L"NO = Usa stampante predefinita (nessun dialogo)\n";
    msg += L"ANNULLA = Annulla operazione";

    int result = MessageBoxW(hWnd, msg.c_str(), L"Stampa selezionate", MB_YESNOCANCEL | MB_ICONQUESTION);

    if (result == IDCANCEL)
        return;

    bool showDialogFirst = (result == IDYES);

    // Mostra un messaggio di avvio
    SetWindowTextW(hWnd, L"Stampa in corso...");

    int successCount = 0;
    int errorCount = 0;

    // Stampa le fatture selezionate
    for (size_t i = 0; i < selectedFatturaIndices.size(); i++)
    {
        int fatturaIndex = selectedFatturaIndices[i];
        if (fatturaIndex < 0 || fatturaIndex >= (int)g_fattureInfo.size())
        {
            errorCount++;
            continue;
        }

        // Carica il file XML
        if (!g_pParser->LoadXmlFile(g_fattureInfo[fatturaIndex].filePath))
        {
            errorCount++;
            continue;
        }

        // Applica la trasformazione
        std::wstring htmlOutput;
        if (g_pParser->ApplyXsltTransform(g_lastXsltPath, htmlOutput))
        {
            // Salva l'HTML temporaneo
            std::wstring htmlPath = g_extractPath + L"fattura_stampa_" + std::to_wstring(fatturaIndex) + L".html";
            FatturaViewer::SaveHTMLToFile(htmlOutput, htmlPath);
            FatturaViewer::NavigateToHTML(g_hBrowserWnd, htmlPath);

            // Attendi il caricamento
            Sleep(1500);

            // Stampa: con dialogo solo per la prima se richiesto, poi sempre silent
            if (i == 0 && showDialogFirst)
            {
                FatturaViewer::PrintDocument(g_hBrowserWnd);
                Sleep(1000); // Attendi che l'utente confermi il dialogo
            }
            else
            {
                FatturaViewer::PrintDocumentSilent(g_hBrowserWnd);
                Sleep(500);
            }

            successCount++;
        }
        else
        {
            errorCount++;
        }

        // Aggiorna il titolo della finestra con il progresso
        std::wstring progress = L"Stampa in corso... (" + std::to_wstring(i + 1) + L"/" + std::to_wstring(selectedFatturaIndices.size()) + L")";
        SetWindowTextW(hWnd, progress.c_str());
    }

    // Ripristina il titolo originale
    SetWindowTextW(hWnd, szTitle);

    // Mostra il risultato
    std::wstring resultMsg = L"Stampa completata!\n\n";
    resultMsg += L"Fatture stampate: " + std::to_wstring(successCount) + L"\n";
    if (errorCount > 0)
    {
        resultMsg += L"Errori: " + std::to_wstring(errorCount);
    }

    MessageBoxW(hWnd, resultMsg.c_str(), L"Stampa selezionate", MB_OK | MB_ICONINFORMATION);
}

HWND CreateToolbar(HWND hParent, HINSTANCE hInst)
{
    // Crea la toolbar con stile moderno
    HWND hToolbar = CreateWindowEx(0, TOOLBARCLASSNAME, NULL,
        WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | TBSTYLE_LIST,
        0, 0, 0, 0,
        hParent, (HMENU)IDC_TOOLBAR, hInst, NULL);

    if (!hToolbar)
        return NULL;

    // Imposta la dimensione dei pulsanti più grande per un aspetto moderno
    SendMessage(hToolbar, TB_BUTTONSTRUCTSIZE, (WPARAM)sizeof(TBBUTTON), 0);
    SendMessage(hToolbar, TB_SETBITMAPSIZE, 0, MAKELONG(32, 32));
    SendMessage(hToolbar, TB_SETBUTTONSIZE, 0, MAKELONG(40, 40));

    // Usa icone standard di Windows (versione large per aspetto moderno)
    TBADDBITMAP tbab;
    tbab.hInst = HINST_COMMCTRL;
    tbab.nID = IDB_STD_LARGE_COLOR;
    SendMessage(hToolbar, TB_ADDBITMAP, 0, (LPARAM)&tbab);

    // Definisci i pulsanti della toolbar con testo
    TBBUTTON tbb[] = {
        {STD_FILEOPEN, IDM_OPEN_ZIP, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Apri ZIP"},
        {0, 0, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0},
        {STD_PROPERTIES, IDM_APPLY_MINISTERO, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Ministero"},
        {STD_PROPERTIES, IDM_APPLY_ASSOSOFTWARE, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Assosoftware"},
        {0, 0, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0},
        {STD_PRINT, IDM_PRINT_CURRENT, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Stampa"},
        {STD_PRINT, IDM_PRINT_SELECTED, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Selezionate"},
        {STD_PRINTPRE, IDM_PRINT_ALL, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Tutte"},
    };

    SendMessage(hToolbar, TB_ADDBUTTONS, sizeof(tbb) / sizeof(TBBUTTON), (LPARAM)&tbb);
    SendMessage(hToolbar, TB_AUTOSIZE, 0, 0);

    return hToolbar;
}

// Ottiene il percorso della cartella dell'applicazione
std::wstring GetApplicationPath()
{
    WCHAR exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);

    std::wstring path(exePath);
    size_t lastSlash = path.find_last_of(L"\\/");
    if (lastSlash != std::wstring::npos)
    {
        return path.substr(0, lastSlash + 1);
    }
    return L"";
}

// Carica la preferenza del foglio di stile dal Registry
void LoadStylesheetPreference()
{
    HKEY hKey;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\FatturaElettronicaViewer", 0, KEY_READ, &hKey);

    if (result == ERROR_SUCCESS)
    {
        WCHAR savedPath[MAX_PATH] = { 0 };
        DWORD dataSize = sizeof(savedPath);
        result = RegQueryValueExW(hKey, L"StylesheetPath", NULL, NULL, (LPBYTE)savedPath, &dataSize);

        if (result == ERROR_SUCCESS)
        {
            g_lastXsltPath = savedPath;
        }
        else
        {
            // Default: Assosoftware
            std::wstring resourcesPath = GetResourcesPath();
            g_lastXsltPath = resourcesPath + L"FoglioStileAssoSoftware.xsl";
        }

        RegCloseKey(hKey);
    }
    else
    {
        // Prima esecuzione: usa Assosoftware come predefinito
        std::wstring resourcesPath = GetResourcesPath();
        g_lastXsltPath = resourcesPath + L"FoglioStileAssoSoftware.xsl";
        SaveStylesheetPreference(g_lastXsltPath);
    }
}

// Salva la preferenza del foglio di stile nel Registry
void SaveStylesheetPreference(const std::wstring& path)
{
    HKEY hKey;
    LONG result = RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\FatturaElettronicaViewer", 0, NULL, 
        REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);

    if (result == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, L"StylesheetPath", 0, REG_SZ, 
            (LPBYTE)path.c_str(), 
            (DWORD)((path.length() + 1) * sizeof(wchar_t)));

        RegCloseKey(hKey);
    }
}

// Ottiene la lista dei fogli di stile dalla cartella Resources
std::vector<std::wstring> GetStylesheetsFromResources()
{
    std::vector<std::wstring> stylesheets;

    // Lista di percorsi da cercare (utile per debug e release)
    std::vector<std::wstring> searchPaths = {
        g_appPath + L"Resources\\*.xsl",           // Release: accanto all'exe
        g_appPath + L"..\\..\\Resources\\*.xsl",   // Debug: due livelli sopra (da x64\Debug)
        g_appPath + L"..\\Resources\\*.xsl"        // Debug alternativo: un livello sopra
    };

    for (const auto& searchPath : searchPaths)
    {
        WIN32_FIND_DATAW findData;
        HANDLE hFind = FindFirstFileW(searchPath.c_str(), &findData);

        if (hFind != INVALID_HANDLE_VALUE)
        {
            do
            {
                if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
                {
                    // Verifica che non sia un duplicato
                    std::wstring fileName = findData.cFileName;
                    if (std::find(stylesheets.begin(), stylesheets.end(), fileName) == stylesheets.end())
                    {
                        stylesheets.push_back(fileName);
                    }
                }
            } while (FindNextFileW(hFind, &findData));

            FindClose(hFind);

            // Se abbiamo trovato file, non cercare oltre
            if (!stylesheets.empty())
                break;
        }
    }

    return stylesheets;
}

// Trova il percorso reale della cartella Resources (anche durante il debug)
std::wstring GetResourcesPath()
{
    std::vector<std::wstring> possiblePaths = {
        g_appPath + L"Resources\\",           // Release: accanto all'exe
        g_appPath + L"..\\..\\Resources\\",   // Debug: due livelli sopra (da x64\Debug)
        g_appPath + L"..\\Resources\\"        // Debug alternativo: un livello sopra
    };

    for (const auto& path : possiblePaths)
    {
        DWORD attribs = GetFileAttributesW(path.c_str());
        if (attribs != INVALID_FILE_ATTRIBUTES && (attribs & FILE_ATTRIBUTE_DIRECTORY))
        {
            return path;
        }
    }

    // Default: assume sia accanto all'exe
    return g_appPath + L"Resources\\";
}
