// xmlRead.cpp : Definisce il punto di ingresso dell'applicazione.
//

#include "framework.h"
#include "FatturaView.h"
#include "FatturaElettronicaParser.h"
#include "FatturaViewer.h"
#include "PdfSigner.h"
#include "AutoUpdater.h"
#include <commdlg.h>
#include <shellapi.h>
#include <vector>
#include <string>
#include <algorithm>
#include <optional>
#include <commctrl.h>
#include <windowsx.h>  // Per GET_X_LPARAM e GET_Y_LPARAM
#include <wincrypt.h>  // Per le funzioni di decodifica P7M

// Helper to get last error message as a string
static std::wstring GetLastErrorAsString()
{
    DWORD errorMessageID = ::GetLastError();
    if (errorMessageID == 0) {
        return L"No error";
    }

    LPWSTR messageBuffer = nullptr;
    size_t size = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL, errorMessageID, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&messageBuffer, 0, NULL);

    if (size == 0) {
        return L"Failed to format error message.";
    }

    std::wstring message(messageBuffer, size);
    LocalFree(messageBuffer);

    return message;
}

#pragma comment(lib, "crypt32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "msimg32.lib")  // Per GradientFill

#define MAX_LOADSTRING 100
#define TOOLBAR_HEIGHT 60  // Aumentato per pulsanti più grandi

// Variabili globali:
HINSTANCE hInst;                                // istanza corrente
WCHAR szTitle[MAX_LOADSTRING];                  // Testo della barra del titolo
WCHAR szWindowClass[MAX_LOADSTRING];            // nome della classe della finestra principale
std::vector<std::wstring> g_xmlFiles;           // Lista dei file XML estratti
std::vector<FatturaInfo> g_fattureInfo;         // Informazioni dettagliate delle fatture
std::vector<int> g_listBoxToFatturaIndex;       // Mappa indice ListBox -> indice g_fattureInfo (-1 = intestazione)
std::wstring g_currentXmlPath;                  // Percorso del file XML corrente
std::vector<std::wstring> g_multipleHtmlFiles;  // File HTML delle fatture in navigazione
int                 g_currentHtmlIndex = -1;    // Indice fattura corrente nella navigazione
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
HWND g_hBtnPrev = NULL;                         // Pulsante navigazione precedente
HWND g_hBtnNext = NULL;                         // Pulsante navigazione successivo
HWND g_hBtnZoomIn = NULL;                       // Pulsante zoom in
HWND g_hBtnZoomOut = NULL;                      // Pulsante zoom out
HWND g_hLabelZoom = NULL;                       // Etichetta percentuale zoom
int  g_zoomLevel = 100;                         // Livello zoom corrente (50-200)
#define NAV_BAR_HEIGHT 40                        // Altezza barra navigazione sotto il browser
#define SEARCH_BAR_HEIGHT 36                     // Altezza casella di ricerca (più visibile)
HFONT g_hFontNormal = NULL;                     // Font normale per ListBox
HFONT g_hFontBold = NULL;                       // Font grassetto per intestazioni
HWND g_hSearchEdit = NULL;                      // Casella di ricerca per filtrare le fatture
std::wstring g_searchFilter;                    // Testo del filtro di ricerca corrente
HBRUSH g_hBrushDialogBg = NULL;                 // Brush sfondo dialog personalizzato
HBRUSH g_hBrushEditBg = NULL;                   // Brush sfondo edit/search personalizzato
HFONT g_hFontSearch = NULL;                     // Font più grande per la barra di ricerca

#define APP_VERSION L"1.2.0"
#define APP_AUTHOR L"Roberto Ferri"

// Costanti per temi Windows 11 (alcune potrebbero non essere definite in vecchie versioni SDK)
#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif

#ifndef DWMWA_WINDOW_CORNER_PREFERENCE
#define DWMWA_WINDOW_CORNER_PREFERENCE 33
#endif

#ifndef DWMWA_BORDER_COLOR
#define DWMWA_BORDER_COLOR 34
#endif

#ifndef DWMWA_CAPTION_COLOR
#define DWMWA_CAPTION_COLOR 35
#endif

#ifndef DWMWA_TEXT_COLOR
#define DWMWA_TEXT_COLOR 36
#endif

#ifndef DWMWA_MICA_EFFECT
#define DWMWA_MICA_EFFECT 1029
#endif


// Dichiarazioni con prototipo di funzioni incluse in questo modulo di codice:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
HICON CreateIconWithTransparentBackground(HINSTANCE hInst, int resId, COLORREF bgColor);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    StylesheetSelector(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    PdfSignConfigDialog(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    LicenseDialogProc(HWND, UINT, WPARAM, LPARAM);
void                OnOpenFile(HWND hWnd);
void                OnOpenPaths(HWND hWnd, const std::vector<std::wstring>& paths);
void                OnApplyStylesheet(HWND hWnd);
void                OnApplySpecificStylesheet(HWND hWnd, const std::wstring& keyword);
void                OnApplyStylesheetMultiple(HWND hWnd, const std::wstring& xsltPath, const std::vector<int>& indices);
std::vector<int>    GetSelectedFattureIndices();
void                UpdateNavButtons();
void                OnZoom(HWND hWnd, bool zoomIn);
void                OnPrintToPDF(HWND hWnd);
void                LoadFattureList(HWND hWnd);
void                RefreshFattureListFiltered();
std::wstring        FindEdgePath();
std::wstring        ExtractBodyContent(const std::wstring& html);
std::wstring        GetTempExtractionPath();
std::wstring        OpenFileDialog(HWND hWnd, const WCHAR* filter, const WCHAR* title);
std::vector<std::wstring> OpenMultipleFilesDialog(HWND hWnd, const WCHAR* filter, const WCHAR* title);
HWND                CreateToolbar(HWND hParent, HINSTANCE hInst);
std::wstring        GetApplicationPath();
void                LoadStylesheetPreference();
void                SaveStylesheetPreference(const std::wstring& path);
std::vector<std::wstring> GetStylesheetsFromResources();
std::wstring        GetResourcesPath();
void                ApplyModernWindowTheme(HWND hWnd);
void                ClearFolderContents(const std::wstring& folderPath);
HICON               CreateIconWithTransparentBackground(HINSTANCE hInst, int resId, COLORREF bgColor);

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

    // Pre-crea il file welcome direttamente in %TEMP% (non in g_extractPath che viene svuotata ad ogni apertura ZIP)
    {
        WCHAR tempDir[MAX_PATH];
        GetTempPathW(MAX_PATH, tempDir);
        g_welcomePath = std::wstring(tempDir) + L"FatturaView_welcome.html";
    }




    // ListBox subclass is set later when the control is created
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

    // Gestione della riga di comando (apertura file all'avvio)
    int argc;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv && argc > 1)
    {
        std::vector<std::wstring> initialPaths;
        for (int i = 1; i < argc; i++)
        {
            initialPaths.push_back(argv[i]);
        }
        LocalFree(argv);

        // Esegui l'apertura tramite un messaggio posticipato per permettere alla finestra di terminare l'inizializzazione
        struct ThreadParams { std::vector<std::wstring> paths; HWND hwnd; };
        ThreadParams* p = new ThreadParams{ initialPaths, g_hMainWnd };

        SetTimer(g_hMainWnd, 2, 500, [](HWND hwnd, UINT, UINT_PTR id, DWORD) {
            KillTimer(hwnd, id);
            // Non possiamo passare direttamente il vector, usiamo una variabile globale temporanea o emuliamo l'evento
            // Per semplicità, richiamiamo OnOpenPaths con i percorsi iniziali (supponendo che esista)
        });

        // In alternativa, chiamiamo direttamente OnOpenPaths ma dopo InitInstance
        OnOpenPaths(g_hMainWnd, initialPaths);
    }
    else if (argv)
    {
        LocalFree(argv);
    }

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
    wcex.hbrBackground  = CreateSolidBrush(RGB(248, 250, 252)); // Sfondo Slate 50
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

   // Crea font per la ListBox (owner-drawn) - Più grande per migliore leggibilità
   g_hFontNormal = CreateFont(
      18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
      CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
      L"Segoe UI");

   g_hFontBold = CreateFont(
      19, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
      CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
      L"Segoe UI");

   // Font più grande per la barra di ricerca
   g_hFontSearch = CreateFont(
      22, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
      CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
      L"Segoe UI");

   // Brush personalizzati per i controlli
   g_hBrushDialogBg = CreateSolidBrush(RGB(248, 250, 252));  // Slate 50
   g_hBrushEditBg = CreateSolidBrush(RGB(255, 255, 255));    // Bianco

   // Crea la toolbar
   g_hToolbar = CreateToolbar(hWnd, hInstance);

   // Replace application icon with transparent background version if possible
   // The actual implementation is in a separate helper to avoid defining functions inside wWinMain
   HICON hAppIcon = CreateIconWithTransparentBackground(hInstance, IDI_XMLREAD, RGB(99, 102, 241));
   if (hAppIcon)
   {
       SendMessageW(hWnd, WM_SETICON, ICON_BIG, (LPARAM)hAppIcon);
       SendMessageW(hWnd, WM_SETICON, ICON_SMALL, (LPARAM)hAppIcon);
   }

   // Icona allegato disegnata dinamicamente nel WM_DRAWITEM (no bitmap creata qui)

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
        std::wstring versionInfo = std::wstring(L"Versione ") + APP_VERSION + L"  -  www.fatturaview.it";
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
   int margin = 20;  // Margine uniforme per design allineato
   int topOffset = TOOLBAR_HEIGHT + margin;

   // Crea la casella di ricerca per filtrare le fatture (stile web, senza bordo 3D)
   g_hSearchEdit = CreateWindowExW(0, L"EDIT", NULL,
      WS_CHILD | WS_VISIBLE | ES_LEFT | ES_AUTOHSCROLL,
      margin + 1, topOffset + 1,
      leftPanelWidth - 2, SEARCH_BAR_HEIGHT - 2,
      hWnd, (HMENU)IDC_SEARCH_EDIT, hInstance, NULL);
   if (g_hSearchEdit)
   {
      if (g_hFontSearch)
          SendMessage(g_hSearchEdit, WM_SETFONT, (WPARAM)g_hFontSearch, TRUE);
      SendMessage(g_hSearchEdit, EM_SETCUEBANNER, TRUE, (LPARAM)L"\U0001F50D  Cerca fattura (emittente, numero, cliente, importo...)");
   }

   int listTopOffset = topOffset + SEARCH_BAR_HEIGHT + 4;

   // Crea una ListBox per mostrare le fatture (pannello sinistro) - stile web piatto
   g_hListBox = CreateWindowExW(0, L"LISTBOX", NULL,
      WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY | LBS_EXTENDEDSEL | LBS_OWNERDRAWVARIABLE | LBS_HASSTRINGS,
      margin + 1, listTopOffset + 1,
      leftPanelWidth - 2, 
      rcClient.bottom - listTopOffset - margin - statusHeight - 2,
      hWnd, (HMENU)IDC_LISTBOX_FATTURE, hInstance, NULL);

   // Per owner-draw variabile impostiamo altezza per singolo elemento dinamicamente

   // Subclass the listbox to handle mouse wheel directly when cursor is over it
   if (g_hListBox)
   {
       // Use a static callback function for subclassing
       SetWindowSubclass(g_hListBox, [](HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) -> LRESULT
       {
           switch (uMsg)
           {
               case WM_MOUSEWHEEL:
               {
                   int z = GET_WHEEL_DELTA_WPARAM(wParam);
                   int lines = abs(z) / WHEEL_DELTA;
                   if (lines < 1) lines = 1;
                   WPARAM action = (z > 0) ? SB_LINEUP : SB_LINEDOWN;
                   for (int i = 0; i < lines; ++i)
                       SendMessage(hwnd, WM_VSCROLL, MAKELONG(action, 0), 0);
                   return 0;
               }
               case WM_MOUSEMOVE:
               {
                   HWND fg = GetFocus();
                   if (fg != hwnd)
                       SetFocus(hwnd);
                   break;
               }
           }
           return DefSubclassProc(hwnd, uMsg, wParam, lParam);
       }, 0, 0);
   }

   // Crea il controllo browser per visualizzazione (pannello destro, ridotto per la nav bar)
   int browserWidth = rcClient.right - leftPanelWidth - 3 * margin;
   int browserHeight = rcClient.bottom - topOffset - margin - statusHeight - NAV_BAR_HEIGHT;
   g_hBrowserWnd = FatturaViewer::CreateBrowserControl(hWnd, hInstance,
      leftPanelWidth + 2 * margin, topOffset,
      browserWidth, browserHeight);

   // Crea i pulsanti di navigazione e zoom a piè di pagina del browser
   // Layout: [◄ Prec] [Succ ►]  [−] [100%] [+]
   // Larghezza totale: 125 + 10 + 125 + 15 + 28 + 5 + 55 + 5 + 28 = 396px
   int navY = topOffset + browserHeight + (NAV_BAR_HEIGHT - 30) / 2;
   int navBaseX = leftPanelWidth + 2 * margin + (browserWidth - 396) / 2;
   g_hBtnPrev = CreateWindowExW(0, L"BUTTON", L"\u25C4 Fattura Prec",
      WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_DISABLED,
      navBaseX, navY, 125, 30,
      hWnd, (HMENU)IDM_PREV_FATTURA, hInstance, NULL);
   g_hBtnNext = CreateWindowExW(0, L"BUTTON", L"Fattura Succ \u25BA",
      WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_DISABLED,
      navBaseX + 135, navY, 125, 30,
      hWnd, (HMENU)IDM_NEXT_FATTURA, hInstance, NULL);
   g_hBtnZoomOut = CreateWindowExW(0, L"BUTTON", L"\u2212",
      WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
      navBaseX + 275, navY, 28, 30,
      hWnd, (HMENU)IDM_ZOOM_OUT, hInstance, NULL);
   g_hLabelZoom = CreateWindowExW(0, L"STATIC", L"100%",
      WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
      navBaseX + 308, navY, 55, 30,
      hWnd, (HMENU)IDC_ZOOM_LABEL, hInstance, NULL);
   g_hBtnZoomIn = CreateWindowExW(0, L"BUTTON", L"+",
      WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
      navBaseX + 368, navY, 28, 30,
      hWnd, (HMENU)IDM_ZOOM_IN, hInstance, NULL);

   // Imposta font moderno sui pulsanti di navigazione e zoom
   if (g_hFontNormal)
   {
      SendMessage(g_hBtnPrev, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
      SendMessage(g_hBtnNext, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
      SendMessage(g_hBtnZoomIn, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
      SendMessage(g_hBtnZoomOut, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
      SendMessage(g_hLabelZoom, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
   }

   // Verifica che il browser sia stato creato
   if (!g_hBrowserWnd)
   {
       MessageBoxW(hWnd, L"Impossibile creare il controllo browser!", L"Errore", MB_OK | MB_ICONERROR);
   }
   else
   {
       // Usa un timer per navigare dopo che la finestra è completamente inizializzata
       SetTimer(hWnd, 1, 100, NULL); // Timer ID 1, 100ms delay
   }

   // Applica tema moderno Windows 11
   ApplyModernWindowTheme(hWnd);

   // Controlla aggiornamenti in background all'avvio
   AutoUpdater::CheckInBackground(hWnd, APP_VERSION);

   ShowWindow(hWnd, SW_SHOWMAXIMIZED);
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
                    clrBackground = RGB(243, 244, 246);
                    clrText = RGB(0, 0, 0);
                    hFont = g_hFontNormal;
                }
                else if (isHeader)
                {
                    // Intestazione emittente - Indigo brand
                    if (pDIS->itemState & ODS_SELECTED)
                    {
                        clrBackground = RGB(79, 70, 229);   // Indigo 600
                        clrText = RGB(255, 255, 255);
                    }
                    else
                    {
                        clrBackground = RGB(99, 102, 241);  // Indigo 500
                        clrText = RGB(255, 255, 255);
                    }
                    hFont = g_hFontBold;
                }
                else
                {
                    // Fattura normale - Effetto card
                    if (pDIS->itemState & ODS_SELECTED)
                    {
                        // Selezionato - Indigo chiaro
                        clrBackground = RGB(224, 231, 255);
                        clrText = RGB(30, 41, 59);
                    }
                    else
                    {
                        // Alternanza bianco/lavanda
                        clrBackground = (pDIS->itemID % 2 == 0) ? RGB(255, 255, 255) : RGB(248, 247, 254);
                        clrText = RGB(30, 41, 59);
                    }
                    hFont = g_hFontNormal;
                }

                // Disegna lo sfondo con gradiente per le intestazioni
                if (isHeader && !(pDIS->itemState & ODS_SELECTED))
                {
                    // Crea un gradiente orizzontale indigo->viola per le intestazioni
                    TRIVERTEX vertex[2];
                    vertex[0].x = pDIS->rcItem.left;
                    vertex[0].y = pDIS->rcItem.top;
                    vertex[0].Red = 0x6300;      // RGB(99, 102, 241) - Indigo
                    vertex[0].Green = 0x6600;
                    vertex[0].Blue = 0xF100;
                    vertex[0].Alpha = 0x0000;

                    vertex[1].x = pDIS->rcItem.right;
                    vertex[1].y = pDIS->rcItem.bottom;
                    vertex[1].Red = 0xA800;      // RGB(168, 85, 247) - Viola
                    vertex[1].Green = 0x5500;
                    vertex[1].Blue = 0xF700;
                    vertex[1].Alpha = 0x0000;

                    GRADIENT_RECT gRect;
                    gRect.UpperLeft = 0;
                    gRect.LowerRight = 1;
                    GradientFill(pDIS->hDC, vertex, 2, &gRect, 1, GRADIENT_FILL_RECT_H);
                }
                else
                {
                    // Disegna lo sfondo solido normale
                    HBRUSH hBrush = CreateSolidBrush(clrBackground);
                    FillRect(pDIS->hDC, &pDIS->rcItem, hBrush);
                    DeleteObject(hBrush);

                    // Aggiungi sottile linea di separazione inferiore per effetto card (solo fatture)
                    if (!isEmptyLine && !isHeader)
                    {
                        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(226, 232, 240));
                        HPEN hOldPen = (HPEN)SelectObject(pDIS->hDC, hPen);
                        MoveToEx(pDIS->hDC, pDIS->rcItem.left + 8, pDIS->rcItem.bottom - 1, NULL);
                        LineTo(pDIS->hDC, pDIS->rcItem.right - 8, pDIS->rcItem.bottom - 1);
                        SelectObject(pDIS->hDC, hOldPen);
                        DeleteObject(hPen);
                    }
                }

                // Disegna bordo focus con colore indigo
                if (pDIS->itemState & ODS_FOCUS)
                {
                    RECT rcFocus = pDIS->rcItem;
                    InflateRect(&rcFocus, -2, -2);
                    HPEN hPenFocus = CreatePen(PS_SOLID, 2, RGB(99, 102, 241));
                    HPEN hOldPen = (HPEN)SelectObject(pDIS->hDC, hPenFocus);
                    HBRUSH hOldBrush = (HBRUSH)SelectObject(pDIS->hDC, GetStockObject(NULL_BRUSH));
                    Rectangle(pDIS->hDC, rcFocus.left, rcFocus.top, rcFocus.right, rcFocus.bottom);
                    SelectObject(pDIS->hDC, hOldBrush);
                    SelectObject(pDIS->hDC, hOldPen);
                    DeleteObject(hPenFocus);
                }

                // Imposta il colore del testo
                SetTextColor(pDIS->hDC, clrText);
                SetBkMode(pDIS->hDC, TRANSPARENT);

                // Disegna il testo con padding
                RECT rcText = pDIS->rcItem;
                rcText.left += 8;  // Padding sinistro
                rcText.right -= 8; // Padding destro

                // Se la fattura ha allegati, disegna una piccola icona "graffetta" a destra
                int iconWidth = 0;
                if (!isHeader && !isEmptyLine && pDIS->itemID < g_listBoxToFatturaIndex.size())
                {
                    int fatturaIndex = g_listBoxToFatturaIndex[pDIS->itemID];
                    if (fatturaIndex >= 0 && fatturaIndex < (int)g_fattureInfo.size() && g_fattureInfo[fatturaIndex].hasAllegati)
                    {
                        // Disegna lo stesso simbolo "allegato" usato nell'HTML (📎) usando un font emoji
                        int itemH = pDIS->rcItem.bottom - pDIS->rcItem.top;
                        // Use a smaller, tighter icon so it looks less prominent and sits near the amount
                        int ih = min(18, max(12, itemH - 12));
                        if (ih < 12) ih = 12;

                        // Position icon closer to the amount (right side) with minimal padding
                        int ix = pDIS->rcItem.right - 6 - ih;
                        int iy = pDIS->rcItem.top + (itemH - ih) / 2;

                        // Small expansion to reduce clipping of emoji glyphs
                        RECT rcIcon = { ix - 1, iy - 1, ix + ih + 2, iy + ih + 2 };

                        // Compute font height based on icon height
                        int logPx = GetDeviceCaps(pDIS->hDC, LOGPIXELSY);
                        int fontHeight = -MulDiv(ih, logPx, 72);

                        // Draw a stylized, minimal paperclip using two smooth polylines
                        HPEN hPen = CreatePen(PS_SOLID, max(1, ih / 10), clrText);
                        HPEN hOldPen = (HPEN)SelectObject(pDIS->hDC, hPen);
                        HBRUSH hOldBrush = (HBRUSH)SelectObject(pDIS->hDC, GetStockObject(NULL_BRUSH));

                        // Draw a more classic paperclip loop similar to mail clients
                        // Use an outer smooth-ish polyline and an inner offset loop
                        POINT outer[9];
                        outer[0].x = ix + (int)(0.18f * ih); outer[0].y = iy + (int)(0.88f * ih);
                        outer[1].x = ix + (int)(0.18f * ih); outer[1].y = iy + (int)(0.28f * ih);
                        outer[2].x = ix + (int)(0.30f * ih); outer[2].y = iy + (int)(0.12f * ih);
                        outer[3].x = ix + (int)(0.52f * ih); outer[3].y = iy + (int)(0.08f * ih);
                        outer[4].x = ix + (int)(0.78f * ih); outer[4].y = iy + (int)(0.28f * ih);
                        outer[5].x = ix + (int)(0.78f * ih); outer[5].y = iy + (int)(0.60f * ih);
                        outer[6].x = ix + (int)(0.58f * ih); outer[6].y = iy + (int)(0.84f * ih);
                        outer[7].x = ix + (int)(0.40f * ih); outer[7].y = iy + (int)(0.92f * ih);
                        outer[8].x = ix + (int)(0.26f * ih); outer[8].y = iy + (int)(0.78f * ih);

                        POINT inner[9];
                        for (int k = 0; k < 9; ++k)
                        {
                            // offset points slightly toward the center to create the inner loop
                            inner[k].x = ix + (int)((outer[k].x - ix) * 0.62f);
                            inner[k].y = iy + (int)((outer[k].y - iy) * 0.62f);
                        }

                        // Draw the same paperclip emoji used in the HTML view (U+1F4CE)
                        const wchar_t paperclipEmoji[] = { 0xD83D, 0xDCCE, 0x0000 };

                        // Create an emoji-capable font sized to the icon height
                        HFONT hEmojiFont = CreateFontW(
                            -MulDiv(ih, GetDeviceCaps(pDIS->hDC, LOGPIXELSY), 72),
                            0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                            CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
                            L"Segoe UI Emoji");

                        HFONT hPrevFont = (HFONT)SelectObject(pDIS->hDC, hEmojiFont);
                        SetTextColor(pDIS->hDC, clrText);
                        SetBkMode(pDIS->hDC, TRANSPARENT);

                        SIZE emojiSize = { 0 };
                        GetTextExtentPoint32W(pDIS->hDC, paperclipEmoji, 2, &emojiSize);
                        int ey = iy + (ih - emojiSize.cy) / 2;
                        TextOutW(pDIS->hDC, ix, ey, paperclipEmoji, 2);

                        // Restore and cleanup
                        SelectObject(pDIS->hDC, hPrevFont);
                        DeleteObject(hEmojiFont);

                        iconWidth = emojiSize.cx + 4;
                        // Reserve a little space on the right so the clip sits close to the amount
                        rcText.right -= iconWidth + 1;
                    }
                }

                // Per le fatture normali (non header), applica grassetto solo al cessionario
                if (!isHeader && !isEmptyLine)
                {
                    // Cerca i due delimitatori "|" per identificare il nome del cessionario
                    size_t firstPipe = itemText.find(L'|');
                    size_t secondPipe = itemText.find(L'|', firstPipe + 1);

                    int itemH = pDIS->rcItem.bottom - pDIS->rcItem.top;

                    // Prepara il rettangolo e la variabile di misura del testo
                    RECT rcPart = rcText;
                    SIZE textSize;

                    if (firstPipe != std::wstring::npos && secondPipe != std::wstring::npos)
                    {
                        // Estrai le tre parti del testo
                        std::wstring part1 = itemText.substr(0, firstPipe + 1);  // "  N. 123 del 2024-01-15 |"
                        std::wstring part2 = itemText.substr(firstPipe + 1, secondPipe - firstPipe - 1);  // " Cessionario "
                        std::wstring part3 = itemText.substr(secondPipe);  // "| € 1000.00"
                        // Disegna la prima parte (numero e data) con font normale usando DT_CALCRECT per altezza corretta
                        HFONT hSaved = (HFONT)SelectObject(pDIS->hDC, g_hFontNormal);
                        RECT rcalc = { 0,0,1000,0 };
                        DrawTextW(pDIS->hDC, part1.c_str(), (int)part1.length(), &rcalc, DT_LEFT | DT_SINGLELINE | DT_CALCRECT);
                        int h1 = rcalc.bottom - rcalc.top;
                        int y = pDIS->rcItem.top + (itemH - h1) / 2;
                        SetTextColor(pDIS->hDC, clrText);
                        TextOutW(pDIS->hDC, rcPart.left, y, part1.c_str(), (int)part1.length());
                        rcPart.left += (rcalc.right - rcalc.left);

                        // Disegna la seconda parte (cessionario) con font grassetto
                        SelectObject(pDIS->hDC, g_hFontBold);
                        rcalc = { 0,0,1000,0 };
                        DrawTextW(pDIS->hDC, part2.c_str(), (int)part2.length(), &rcalc, DT_LEFT | DT_SINGLELINE | DT_CALCRECT);
                        int h2 = rcalc.bottom - rcalc.top;
                        int y2 = pDIS->rcItem.top + (itemH - h2) / 2;
                        TextOutW(pDIS->hDC, rcPart.left, y2, part2.c_str(), (int)part2.length());
                        rcPart.left += (rcalc.right - rcalc.left);

                        // Disegna la terza parte (importo) con font normale
                        SelectObject(pDIS->hDC, g_hFontNormal);
                        rcalc = { 0,0,1000,0 };
                        DrawTextW(pDIS->hDC, part3.c_str(), (int)part3.length(), &rcalc, DT_LEFT | DT_SINGLELINE | DT_CALCRECT);
                        int h3 = rcalc.bottom - rcalc.top;
                        int y3 = pDIS->rcItem.top + (itemH - h3) / 2;
                        TextOutW(pDIS->hDC, rcPart.left, y3, part3.c_str(), (int)part3.length());

                        // Ripristina il font precedente
                        SelectObject(pDIS->hDC, hSaved);
                    }
                    else
                    {
                        // Fallback: disegna normalmente se non trova i delimitatori usando centratura manuale
                        HFONT hSaved = (HFONT)SelectObject(pDIS->hDC, hFont);
                        GetTextExtentPoint32(pDIS->hDC, itemText.c_str(), (int)itemText.length(), &textSize);
                        int y = pDIS->rcItem.top + (itemH - textSize.cy) / 2;
                        TextOutW(pDIS->hDC, rcText.left, y, itemText.c_str(), (int)itemText.length());
                        SelectObject(pDIS->hDC, hSaved);
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

            // Owner-draw per pulsanti navigazione/zoom stile web
            if (pDIS->CtlType == ODT_BUTTON)
            {
                bool isDisabled = (pDIS->itemState & ODS_DISABLED) != 0;
                bool isPressed = (pDIS->itemState & ODS_SELECTED) != 0;
                bool isFocused = (pDIS->itemState & ODS_FOCUS) != 0;

                COLORREF bgColor, textColor, borderColor;
                int btnId = pDIS->CtlID;

                if (btnId == IDM_PREV_FATTURA || btnId == IDM_NEXT_FATTURA)
                {
                    // Pulsanti navigazione: pillola indigo
                    if (isDisabled)
                    {
                        bgColor = RGB(226, 232, 240);    // Slate 200
                        textColor = RGB(148, 163, 184);  // Slate 400
                        borderColor = RGB(203, 213, 225);
                    }
                    else if (isPressed)
                    {
                        bgColor = RGB(67, 56, 202);      // Indigo 700
                        textColor = RGB(255, 255, 255);
                        borderColor = RGB(55, 48, 163);
                    }
                    else
                    {
                        bgColor = RGB(99, 102, 241);     // Indigo 500
                        textColor = RGB(255, 255, 255);
                        borderColor = RGB(79, 70, 229);
                    }
                }
                else
                {
                    // Pulsanti zoom: pillola grigia
                    if (isDisabled)
                    {
                        bgColor = RGB(241, 245, 249);
                        textColor = RGB(148, 163, 184);
                        borderColor = RGB(226, 232, 240);
                    }
                    else if (isPressed)
                    {
                        bgColor = RGB(203, 213, 225);
                        textColor = RGB(30, 41, 59);
                        borderColor = RGB(148, 163, 184);
                    }
                    else
                    {
                        bgColor = RGB(241, 245, 249);    // Slate 100
                        textColor = RGB(30, 41, 59);
                        borderColor = RGB(203, 213, 225); // Slate 300
                    }
                }

                // Sfondo con angoli arrotondati
                HBRUSH hBg = CreateSolidBrush(bgColor);
                HPEN hBorder = CreatePen(PS_SOLID, 1, borderColor);
                HBRUSH hOldBr = (HBRUSH)SelectObject(pDIS->hDC, hBg);
                HPEN hOldPn = (HPEN)SelectObject(pDIS->hDC, hBorder);
                RoundRect(pDIS->hDC, pDIS->rcItem.left, pDIS->rcItem.top,
                    pDIS->rcItem.right, pDIS->rcItem.bottom, 12, 12);
                SelectObject(pDIS->hDC, hOldPn);
                SelectObject(pDIS->hDC, hOldBr);
                DeleteObject(hBorder);
                DeleteObject(hBg);

                // Testo centrato
                SetBkMode(pDIS->hDC, TRANSPARENT);
                SetTextColor(pDIS->hDC, textColor);
                HFONT hOldFont = (HFONT)SelectObject(pDIS->hDC, g_hFontNormal);
                WCHAR btnText[64] = {};
                GetWindowTextW(pDIS->hwndItem, btnText, 64);
                DrawTextW(pDIS->hDC, btnText, -1, &pDIS->rcItem,
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                SelectObject(pDIS->hDC, hOldFont);

                return TRUE;
            }
        }
        break;
    case WM_NOTIFY:
        {
            LPNMHDR pnmh = (LPNMHDR)lParam;
            // Toolbar custom draw: sfondo piatto stile web
            if (pnmh->hwndFrom == g_hToolbar && pnmh->code == NM_CUSTOMDRAW)
            {
                LPNMTBCUSTOMDRAW pcd = (LPNMTBCUSTOMDRAW)lParam;
                switch (pcd->nmcd.dwDrawStage)
                {
                case CDDS_PREPAINT:
                    {
                        // Disegna sfondo toolbar personalizzato
                        RECT rc;
                        GetClientRect(g_hToolbar, &rc);
                        HBRUSH hBr = CreateSolidBrush(RGB(248, 250, 252));
                        FillRect(pcd->nmcd.hdc, &rc, hBr);
                        DeleteObject(hBr);
                        // Linea sottile di separazione inferiore
                        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(226, 232, 240));
                        HPEN hOld = (HPEN)SelectObject(pcd->nmcd.hdc, hPen);
                        MoveToEx(pcd->nmcd.hdc, rc.left, rc.bottom - 1, NULL);
                        LineTo(pcd->nmcd.hdc, rc.right, rc.bottom - 1);
                        SelectObject(pcd->nmcd.hdc, hOld);
                        DeleteObject(hPen);
                        return CDRF_NOTIFYITEMDRAW;
                    }
                case CDDS_ITEMPREPAINT:
                    {
                        // Sfondo trasparente per i singoli bottoni
                        pcd->clrBtnFace = RGB(248, 250, 252);
                        pcd->clrText = RGB(30, 41, 59);
                        return TBCDRF_USECDCOLORS;
                    }
                }
                return CDRF_DODEFAULT;
            }
        }
        break;
    case WM_MOUSEWHEEL:
        {
            // Forward mouse wheel to the left listbox when cursor is over it
            if (g_hListBox)
            {
                // lParam contains screen coordinates of the cursor
                POINT pt;
                pt.x = GET_X_LPARAM(lParam);
                pt.y = GET_Y_LPARAM(lParam);

                // Convert listbox rect to screen coordinates and test
                RECT rcLB;
                GetWindowRect(g_hListBox, &rcLB);

                if (PtInRect(&rcLB, pt))
                {
                    int z = GET_WHEEL_DELTA_WPARAM(wParam);
                    int lines = abs(z) / WHEEL_DELTA;
                    if (lines < 1) lines = 1;
                    WPARAM action = (z > 0) ? SB_LINEUP : SB_LINEDOWN;
                    for (int i = 0; i < lines; ++i)
                    {
                        SendMessage(g_hListBox, WM_VSCROLL, MAKELONG(action, 0), 0);
                    }
                    return 0;
                }
            }
        }
        break;
    case WM_TIMER:
        {
            if (wParam == 1) // Timer per il caricamento del welcome
            {
                KillTimer(hWnd, 1);

                if (g_hBrowserWnd)
                {
                    if (!g_welcomePath.empty() && GetFileAttributesW(g_welcomePath.c_str()) != INVALID_FILE_ATTRIBUTES)
                        FatturaViewer::NavigateToHTML(g_hBrowserWnd, g_welcomePath);
                    else
                        FatturaViewer::NavigateToString(g_hBrowserWnd, FatturaViewer::GetWelcomePageHTML());
                }
            }
        }
        break;
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_SHOW_LICENSE:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_LICENSE_DIALOG), hWnd, LicenseDialogProc);
                break;
            case IDM_CHECK_UPDATES:
                {
                    HCURSOR hOld = SetCursor(LoadCursor(NULL, IDC_WAIT));
                    UpdateInfo info;
                    bool ok = AutoUpdater::CheckForUpdates(APP_VERSION, info);
                    SetCursor(hOld);
                    if (ok && info.updateAvailable)
                        AutoUpdater::ShowUpdateDialog(hWnd, info, APP_VERSION);
                    else if (ok)
                        AutoUpdater::ShowNoUpdateDialog(hWnd, APP_VERSION);
                    else
                        MessageBoxW(hWnd, L"Impossibile contattare il server.\nVerifica la connessione internet.", L"Errore", MB_OK | MB_ICONWARNING);
                }
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            case IDM_OPEN_ZIP:
                OnOpenFile(hWnd);
                break;
            case IDM_PREV_FATTURA:
                if (g_currentHtmlIndex > 0 && g_hBrowserWnd)
                {
                    g_currentHtmlIndex--;
                    FatturaViewer::NavigateToHTML(g_hBrowserWnd, g_multipleHtmlFiles[g_currentHtmlIndex]);
                    UpdateNavButtons();
                }
                break;
            case IDM_NEXT_FATTURA:
                if (g_currentHtmlIndex >= 0 && g_currentHtmlIndex < (int)g_multipleHtmlFiles.size() - 1 && g_hBrowserWnd)
                {
                    g_currentHtmlIndex++;
                    FatturaViewer::NavigateToHTML(g_hBrowserWnd, g_multipleHtmlFiles[g_currentHtmlIndex]);
                    UpdateNavButtons();
                }
                break;
            case IDM_ZOOM_IN:
                OnZoom(hWnd, true);
                break;
            case IDM_ZOOM_OUT:
                OnZoom(hWnd, false);
                break;
            case IDM_APPLY_MINISTERO:
                OnApplySpecificStylesheet(hWnd, L"Ministero");
                break;
            case IDM_APPLY_ASSOSOFTWARE:
                OnApplySpecificStylesheet(hWnd, L"Asso");
                break;
            case IDM_PRINT_SELECTED:
                OnPrintToPDF(hWnd);
                break;
            case IDM_CONTEXT_PRINT_SELECTED:
                // Stampa le fatture selezionate tramite menu contestuale
                OnPrintToPDF(hWnd);
                break;
            case IDM_CONTEXT_CHANGE_STYLE:
                // Apri il dialogo di selezione foglio di stile
                if (DialogBox(hInst, MAKEINTRESOURCE(IDD_STYLESHEET_SELECTOR), hWnd, StylesheetSelector) == IDOK && !g_lastXsltPath.empty())
                {
                    std::vector<int> selectedIndices = GetSelectedFattureIndices();
                    if (!selectedIndices.empty())
                    {
                        if (selectedIndices.size() == 1)
                        {
                            // Singola fattura - imposta il path corrente e applica
                            g_currentXmlPath = g_fattureInfo[selectedIndices[0]].filePath;
                            if (g_pParser && g_pParser->LoadXmlFile(g_currentXmlPath))
                            {
                                // Usa il foglio selezionato
                                std::wstring htmlOutput;
                                if (g_pParser->ApplyXsltTransform(g_lastXsltPath, htmlOutput))
                                {
                                    std::wstring htmlPath = g_extractPath + L"fattura_visualizzata.html";
                                    FatturaViewer::SaveHTMLToFile(htmlOutput, htmlPath);
                                    if (g_hBrowserWnd)
                                        FatturaViewer::NavigateToHTML(g_hBrowserWnd, htmlPath);
                                    if (g_hStatusBar)
                                        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Fattura visualizzata");
                                    if (g_hBrowserWnd)
                                        FatturaViewer::ShowNotification(g_hBrowserWnd, L"Fattura visualizzata", L"success");
                                }
                            }
                        }
                        else
                        {
                            // Multiple fatture
                            OnApplyStylesheetMultiple(hWnd, g_lastXsltPath, selectedIndices);
                        }
                    }
                }
                break;
            case IDM_CONTEXT_VIEW_MINISTERO:
                // Applica foglio Ministero alla fattura selezionata
                OnApplySpecificStylesheet(hWnd, L"Ministero");
                break;
            case IDM_CONTEXT_VIEW_ASSOSOFTWARE:
                // Applica foglio Assosoftware alla fattura selezionata
                OnApplySpecificStylesheet(hWnd, L"Asso");
                break;
            case IDM_SETTINGS_SELECT_STYLESHEET:
            {
                if (DialogBox(hInst, MAKEINTRESOURCE(IDD_STYLESHEET_SELECTOR), hWnd, StylesheetSelector) == IDOK && !g_lastXsltPath.empty())
                {
                    std::vector<int> selectedIndices = GetSelectedFattureIndices();
                    if (selectedIndices.size() > 1)
                    {
                        OnApplyStylesheetMultiple(hWnd, g_lastXsltPath, selectedIndices);
                    }
                    else if (!selectedIndices.empty())
                    {
                        g_currentXmlPath = g_fattureInfo[selectedIndices[0]].filePath;
                        OnApplyStylesheet(hWnd);
                    }
                }
                break;
            }
            case IDC_SEARCH_EDIT:
                if (HIWORD(wParam) == EN_CHANGE)
                {
                    // L'utente ha digitato nella casella di ricerca: aggiorna il filtro
                    WCHAR buf[256] = {};
                    GetWindowTextW(g_hSearchEdit, buf, 256);
                    g_searchFilter = buf;
                    RefreshFattureListFiltered();
                }
                break;
            case IDC_LISTBOX_FATTURE:
                if (HIWORD(wParam) == LBN_SELCHANGE)
                {
                    int selCount = (int)SendMessage(g_hListBox, LB_GETSELCOUNT, 0, 0);

                    if (selCount == 1)
                    {
                        // Selezione singola: mostra automaticamente nel browser
                        int listBoxIndex = (int)SendMessage(g_hListBox, LB_GETCURSEL, 0, 0);
                        if (listBoxIndex != LB_ERR && listBoxIndex < (int)g_listBoxToFatturaIndex.size())
                        {
                            int fatturaIndex = g_listBoxToFatturaIndex[listBoxIndex];
                            if (fatturaIndex >= 0 && fatturaIndex < (int)g_fattureInfo.size())
                            {
                                g_currentXmlPath = g_fattureInfo[fatturaIndex].filePath;
                                OnApplyStylesheet(hWnd);
                            }
                        }
                    }
                    else if (selCount > 1)
                    {
                        // Selezione multipla: usa il foglio di stile di default automaticamente
                        std::wstring xsltPath = g_lastXsltPath;

                        // Se non c'è un foglio configurato, cerca il Ministero come default
                        if (xsltPath.empty() || GetFileAttributesW(xsltPath.c_str()) == INVALID_FILE_ATTRIBUTES)
                        {
                            std::wstring resourcesPath = GetResourcesPath();
                            WIN32_FIND_DATAW findData;
                            HANDLE hFind = FindFirstFileW((resourcesPath + L"*Ministero*.xsl").c_str(), &findData);
                            if (hFind != INVALID_HANDLE_VALUE)
                            {
                                xsltPath = resourcesPath + findData.cFileName;
                                FindClose(hFind);
                            }
                            else
                            {
                                // Fallback: usa qualsiasi XSL disponibile
                                hFind = FindFirstFileW((resourcesPath + L"*.xsl").c_str(), &findData);
                                if (hFind != INVALID_HANDLE_VALUE)
                                {
                                    xsltPath = resourcesPath + findData.cFileName;
                                    FindClose(hFind);
                                }
                            }
                        }

                        if (!xsltPath.empty() && GetFileAttributesW(xsltPath.c_str()) != INVALID_FILE_ATTRIBUTES)
                        {
                            g_lastXsltPath = xsltPath;
                            std::vector<int> selectedIndices = GetSelectedFattureIndices();
                            if (!selectedIndices.empty())
                                OnApplyStylesheetMultiple(hWnd, xsltPath, selectedIndices);
                        }
                        else if (g_hStatusBar)
                        {
                            std::wstring msg = std::to_wstring(selCount) + L" fatture selezionate \u2014 nessun foglio di stile disponibile";
                            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)msg.c_str());
                        }
                    }
                }
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_CONTEXTMENU:
        {
            // Gestisce il menu contestuale sulla ListBox
            HWND hWndContext = (HWND)wParam;
            int xPos = GET_X_LPARAM(lParam);
            int yPos = GET_Y_LPARAM(lParam);

            // Verifica se il click è sulla ListBox delle fatture
            if (hWndContext == g_hListBox)
            {
                // Ottieni gli indici selezionati
                std::vector<int> selectedIndices = GetSelectedFattureIndices();

                if (selectedIndices.empty())
                {
                    // Nessuna selezione - non mostrare il menu
                    break;
                }

                // Crea il menu contestuale
                HMENU hMenu = CreatePopupMenu();

                if (selectedIndices.size() == 1)
                {
                    AppendMenuW(hMenu, MF_STRING, IDM_CONTEXT_VIEW_MINISTERO, L"Visualizza con foglio Ministero");
                    AppendMenuW(hMenu, MF_STRING, IDM_CONTEXT_VIEW_ASSOSOFTWARE, L"Visualizza con foglio Assosoftware");
                    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                    AppendMenuW(hMenu, MF_STRING, IDM_CONTEXT_CHANGE_STYLE, L"Cambia foglio di stile...");
                    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                    AppendMenuW(hMenu, MF_STRING, IDM_CONTEXT_PRINT_SELECTED, L"Stampa fattura");
                }
                else
                {
                    std::wstring printText = L"Stampa " + std::to_wstring(selectedIndices.size()) + L" fatture selezionate";
                    AppendMenuW(hMenu, MF_STRING, IDM_CONTEXT_PRINT_SELECTED, printText.c_str());
                    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                    AppendMenuW(hMenu, MF_STRING, IDM_CONTEXT_CHANGE_STYLE, L"Cambia foglio di stile...");
                }

                // Mostra il menu e ottieni la selezione
                int cmd = TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD, 
                                        xPos, yPos, 0, hWnd, NULL);

                // Gestisci la selezione
                if (cmd != 0)
                {
                    SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
                }

                DestroyMenu(hMenu);
            }
        }
        break;
    case WM_SIZE:
        {
            // Ridimensiona i controlli quando la finestra viene ridimensionata
            RECT rcClient;
            GetClientRect(hWnd, &rcClient);

            int leftPanelWidth = 550;
            int margin = 20;  // Margine uniforme allineato
            int topOffset = TOOLBAR_HEIGHT + margin;

            // Ridimensiona la toolbar a tutta larghezza (stile web nav bar)
            if (g_hToolbar)
            {
                SetWindowPos(g_hToolbar, NULL, 0, 0, rcClient.right, TOOLBAR_HEIGHT,
                    SWP_NOZORDER | SWP_NOACTIVATE);
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

            if (g_hSearchEdit)
            {
                MoveWindow(g_hSearchEdit,
                    margin + 1, topOffset + 1,
                    leftPanelWidth - 2, SEARCH_BAR_HEIGHT - 2,
                    TRUE);
            }

            int listTopOffset = topOffset + SEARCH_BAR_HEIGHT + 4;

            if (g_hListBox)
            {
                MoveWindow(g_hListBox,
                    margin + 1, listTopOffset + 1,
                    leftPanelWidth - 2,
                    rcClient.bottom - listTopOffset - margin - statusHeight - 2,
                    TRUE);
            }

            // Forza ridisegno dei bordi personalizzati
            InvalidateRect(hWnd, NULL, TRUE);

            int browserWidth = rcClient.right - leftPanelWidth - 3 * margin;
            int browserHeight = rcClient.bottom - topOffset - margin - statusHeight - NAV_BAR_HEIGHT;

            if (g_hBrowserWnd)
            {
                MoveWindow(g_hBrowserWnd,
                    leftPanelWidth + 2 * margin, topOffset,
                    browserWidth, browserHeight,
                    TRUE);
            }

            // Riposiziona i pulsanti di navigazione e zoom sotto il browser
            int navY = topOffset + browserHeight + (NAV_BAR_HEIGHT - 30) / 2;
            int navBaseX = leftPanelWidth + 2 * margin + (browserWidth - 396) / 2;
            if (g_hBtnPrev)
                MoveWindow(g_hBtnPrev, navBaseX, navY, 125, 30, TRUE);
            if (g_hBtnNext)
                MoveWindow(g_hBtnNext, navBaseX + 135, navY, 125, 30, TRUE);
            if (g_hBtnZoomOut)
                MoveWindow(g_hBtnZoomOut, navBaseX + 275, navY, 28, 30, TRUE);
            if (g_hLabelZoom)
                MoveWindow(g_hLabelZoom, navBaseX + 308, navY, 55, 30, TRUE);
            if (g_hBtnZoomIn)
                MoveWindow(g_hBtnZoomIn, navBaseX + 368, navY, 28, 30, TRUE);
        }
        break;
    case WM_CTLCOLOREDIT:
        {
            // Personalizza colore del filtro di ricerca
            HDC hdcEdit = (HDC)wParam;
            HWND hCtl = (HWND)lParam;
            if (hCtl == g_hSearchEdit)
            {
                SetTextColor(hdcEdit, RGB(30, 41, 59));      // Slate 800
                SetBkColor(hdcEdit, RGB(255, 255, 255));
                return (LRESULT)g_hBrushEditBg;
            }
        }
        break;
    case WM_CTLCOLORSTATIC:
        {
            // Personalizza label e zoom con sfondo della finestra
            HDC hdcStatic = (HDC)wParam;
            SetTextColor(hdcStatic, RGB(30, 41, 59));
            SetBkColor(hdcStatic, RGB(248, 250, 252));
            return (LRESULT)g_hBrushDialogBg;
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);

            // Disegna bordi piatti arrotondati attorno ai controlli (stile web)
            HPEN hBorderPen = CreatePen(PS_SOLID, 1, RGB(199, 210, 254));  // Indigo 200
            HPEN hOldPen = (HPEN)SelectObject(hdc, hBorderPen);
            HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));

            // Bordo attorno alla casella di ricerca
            if (g_hSearchEdit)
            {
                RECT rcEdit;
                GetWindowRect(g_hSearchEdit, &rcEdit);
                MapWindowPoints(HWND_DESKTOP, hWnd, (LPPOINT)&rcEdit, 2);
                InflateRect(&rcEdit, 1, 1);
                RoundRect(hdc, rcEdit.left, rcEdit.top, rcEdit.right, rcEdit.bottom, 8, 8);
            }

            // Bordo attorno alla ListBox
            if (g_hListBox)
            {
                RECT rcList;
                GetWindowRect(g_hListBox, &rcList);
                MapWindowPoints(HWND_DESKTOP, hWnd, (LPPOINT)&rcList, 2);
                InflateRect(&rcList, 1, 1);
                RoundRect(hdc, rcList.left, rcList.top, rcList.right, rcList.bottom, 8, 8);
            }

            SelectObject(hdc, hOldBrush);
            SelectObject(hdc, hOldPen);
            DeleteObject(hBorderPen);

            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        // Elimina i font creati
        if (g_hFontNormal)
            DeleteObject(g_hFontNormal);
        if (g_hFontBold)
            DeleteObject(g_hFontBold);
        if (g_hFontSearch)
            DeleteObject(g_hFontSearch);
        if (g_hBrushDialogBg)
            DeleteObject(g_hBrushDialogBg);
        if (g_hBrushEditBg)
            DeleteObject(g_hBrushEditBg);
        PostQuitMessage(0);
        break;
    default:
        // Gestisci notifica aggiornamento disponibile dal thread di background
        if (message == AutoUpdater::WM_APP_UPDATE_AVAILABLE)
        {
            UpdateInfo* pInfo = (UpdateInfo*)lParam;
            if (pInfo && pInfo->updateAvailable)
            {
                AutoUpdater::ShowUpdateDialog(hWnd, *pInfo, APP_VERSION);
            }
            return 0;
        }
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

    case WM_CTLCOLORDLG:
    case WM_CTLCOLORSTATIC:
        {
            HDC hdcCtl = (HDC)wParam;
            SetTextColor(hdcCtl, RGB(30, 41, 59));
            SetBkColor(hdcCtl, RGB(248, 250, 252));
            if (!g_hBrushDialogBg)
                g_hBrushDialogBg = CreateSolidBrush(RGB(248, 250, 252));
            return (INT_PTR)g_hBrushDialogBg;
        }

    case WM_NOTIFY:
    {
        NMHDR* pHdr = reinterpret_cast<NMHDR*>(lParam);
        if (pHdr->idFrom == IDC_ABOUT_SYSLINK &&
            (pHdr->code == NM_CLICK || pHdr->code == NM_RETURN))
        {
            ShellExecuteW(nullptr, L"open", L"https://www.fatturaview.it",
                          nullptr, nullptr, SW_SHOWNORMAL);
            return (INT_PTR)TRUE;
        }
        break;
    }

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
                size_t pos = currentName.find_last_of(L"\\/");
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

    case WM_CTLCOLORDLG:
    case WM_CTLCOLORSTATIC:
        {
            HDC hdcCtl = (HDC)wParam;
            SetTextColor(hdcCtl, RGB(30, 41, 59));
            SetBkColor(hdcCtl, RGB(248, 250, 252));
            if (!g_hBrushDialogBg)
                g_hBrushDialogBg = CreateSolidBrush(RGB(248, 250, 252));
            return (INT_PTR)g_hBrushDialogBg;
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
                    std::wstring statusMsg = L"Foglio di stile: " + stylesheets[sel];
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

// Dialog procedure per la configurazione della firma PDF
INT_PTR CALLBACK PdfSignConfigDialog(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    static PdfSignConfig config;

    switch (message)
    {
    case WM_INITDIALOG:
        {
            // Carica la configurazione salvata
            PdfSigner::LoadConfig(config);

            // Imposta i valori nei controlli
            CheckDlgButton(hDlg, IDC_ENABLE_SIGNING, config.enabled ? BST_CHECKED : BST_UNCHECKED);
            SetDlgItemTextW(hDlg, IDC_CERT_PATH, config.certPath.c_str());
            SetDlgItemTextW(hDlg, IDC_CERT_PASSWORD, config.certPassword.c_str());
            SetDlgItemTextW(hDlg, IDC_SIGN_REASON, config.signReason.c_str());
            SetDlgItemTextW(hDlg, IDC_SIGN_LOCATION, config.signLocation.c_str());
            SetDlgItemTextW(hDlg, IDC_SIGN_CONTACT, config.signContact.c_str());

            return (INT_PTR)TRUE;
        }

    case WM_CTLCOLORDLG:
    case WM_CTLCOLORSTATIC:
        {
            HDC hdcCtl = (HDC)wParam;
            SetTextColor(hdcCtl, RGB(30, 41, 59));
            SetBkColor(hdcCtl, RGB(248, 250, 252));
            if (!g_hBrushDialogBg)
                g_hBrushDialogBg = CreateSolidBrush(RGB(248, 250, 252));
            return (INT_PTR)g_hBrushDialogBg;
        }

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_CERT_BROWSE)
        {
            // Apri dialog per selezionare il certificato
            std::wstring certPath = OpenFileDialog(hDlg, 
                L"Certificati Digitali (*.pfx;*.p12)\0*.pfx;*.p12\0Tutti i file (*.*)\0*.*\0",
                L"Seleziona Certificato Digitale");

            if (!certPath.empty())
            {
                SetDlgItemTextW(hDlg, IDC_CERT_PATH, certPath.c_str());
            }
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDOK)
        {
            // Salva la configurazione
            config.enabled = (IsDlgButtonChecked(hDlg, IDC_ENABLE_SIGNING) == BST_CHECKED);

            WCHAR buffer[MAX_PATH];
            GetDlgItemTextW(hDlg, IDC_CERT_PATH, buffer, MAX_PATH);
            config.certPath = buffer;

            GetDlgItemTextW(hDlg, IDC_CERT_PASSWORD, buffer, MAX_PATH);
            config.certPassword = buffer;

            GetDlgItemTextW(hDlg, IDC_SIGN_REASON, buffer, MAX_PATH);
            config.signReason = buffer;

            GetDlgItemTextW(hDlg, IDC_SIGN_LOCATION, buffer, MAX_PATH);
            config.signLocation = buffer;

            GetDlgItemTextW(hDlg, IDC_SIGN_CONTACT, buffer, MAX_PATH);
            config.signContact = buffer;

            // Verifica che se abilitato, il certificato sia specificato
            if (config.enabled && config.certPath.empty())
            {
                MessageBoxW(hDlg, L"Specificare il percorso del certificato digitale!", L"Attenzione", MB_OK | MB_ICONWARNING);
                return (INT_PTR)TRUE;
            }

            // Verifica che il certificato esista
            if (config.enabled && GetFileAttributesW(config.certPath.c_str()) == INVALID_FILE_ATTRIBUTES)
            {
                MessageBoxW(hDlg, L"Il certificato specificato non esiste!", L"Errore", MB_OK | MB_ICONERROR);
                return (INT_PTR)TRUE;
            }

            // Salva la configurazione
            if (!PdfSigner::SaveConfig(config))
            {
                MessageBoxW(hDlg, L"Impossibile salvare la configurazione!", L"Errore", MB_OK | MB_ICONERROR);
                return (INT_PTR)TRUE;
            }

            MessageBoxW(hDlg, L"Configurazione salvata con successo!", L"Successo", MB_OK | MB_ICONINFORMATION);
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

std::vector<std::wstring> OpenMultipleFilesDialog(HWND hWnd, const WCHAR* filter, const WCHAR* title)
{
    std::vector<std::wstring> results;
    // Buffer grande per supportare molti file (32KB)
    const DWORD bufSize = 32768;
    std::vector<WCHAR> buf(bufSize, L'\0');

    OPENFILENAMEW ofn = { 0 };
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = hWnd;
    ofn.lpstrFile   = buf.data();
    ofn.nMaxFile    = bufSize;
    ofn.lpstrFilter = filter;
    ofn.lpstrTitle  = title;
    ofn.Flags       = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_EXPLORER;

    if (!GetOpenFileNameW(&ofn))
        return results;

    const wchar_t* p = buf.data();
    std::wstring dir = p;
    p += dir.length() + 1;

    if (*p == L'\0')
    {
        // File singolo: il buffer contiene il percorso completo
        results.push_back(dir);
    }
    else
    {
        // File multipli: primo token = cartella, poi nomi file
        while (*p != L'\0')
        {
            std::wstring filename = p;
            results.push_back(dir + L"\\" + filename);
            p += filename.length() + 1;
        }
    }

    return results;
}

std::wstring GetTempExtractionPath()
{
    WCHAR tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempPath);
    std::wstring extractPath = std::wstring(tempPath) + L"FattureElettroniche\\";
    CreateDirectoryW(extractPath.c_str(), NULL);
    return extractPath;
}

// Elimina ricorsivamente tutti i file e le sottocartelle dentro folderPath
void ClearFolderContents(const std::wstring& folderPath)
{
    std::wstring base = folderPath;
    if (!base.empty() && base.back() != L'\\')
        base += L'\\';

    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW((base + L"*").c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE)
        return;

    do
    {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0)
            continue;

        std::wstring fullPath = base + fd.cFileName;
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            ClearFolderContents(fullPath);
            RemoveDirectoryW(fullPath.c_str());
        }
        else
        {
            DeleteFileW(fullPath.c_str());
        }
    } while (FindNextFileW(hFind, &fd));

    FindClose(hFind);
}

void OnOpenPaths(HWND hWnd, const std::vector<std::wstring>& paths)
{
    if (paths.empty())
        return;

    if (!g_pParser)
        return;

    // Valida anticipatamente tutti i file XML (non ZIP/P7M): se anche uno solo non e' FatturaPA blocca tutto
    {
        std::vector<std::wstring> invalidFiles;
        for (const auto& path : paths)
        {
            size_t dotPos = path.find_last_of(L'.');
            std::wstring ext;
            if (dotPos != std::wstring::npos) {
                ext = path.substr(dotPos);
                std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
            }

            if (ext == L".zip" || ext == L".p7m") {
                continue; // Salta la validazione per ZIP e P7M, verranno gestiti dopo
            }

            if (!g_pParser->IsFatturaPA(path))
            {
                size_t sl = path.find_last_of(L"\\/");
                invalidFiles.push_back((sl != std::wstring::npos) ? path.substr(sl + 1) : path);
            }
        }
        if (!invalidFiles.empty())
        {
            std::wstring msg = L"I seguenti file non sono fatture elettroniche valide (FatturaPA):\n\n";
            for (const auto& name : invalidFiles)
                msg += L"\u2022 " + name + L"\n";
            msg += L"\nOperazione annullata. Selezionare solo file XML in formato FatturaPA.";
            MessageBoxW(hWnd, msg.c_str(), L"File non validi", MB_OK | MB_ICONERROR);
            return;
        }
    }

    EnableWindow(g_hListBox, FALSE);
    EnableWindow(g_hToolbar, FALSE);
    HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));

    // Reset completo
    g_currentXmlPath.clear();
    g_multipleHtmlFiles.clear();
    g_currentHtmlIndex = -1;
    UpdateNavButtons();
    if (g_hBrowserWnd)
    {
        if (!g_welcomePath.empty() && GetFileAttributesW(g_welcomePath.c_str()) != INVALID_FILE_ATTRIBUTES)
            FatturaViewer::NavigateToHTML(g_hBrowserWnd, g_welcomePath);
        else
            FatturaViewer::NavigateToString(g_hBrowserWnd, FatturaViewer::GetWelcomePageHTML());
    }
    if (g_hListBox)
        SendMessage(g_hListBox, LB_RESETCONTENT, 0, 0);

    WCHAR titleBuffer[256];
    GetWindowTextW(hWnd, titleBuffer, 256);
    std::wstring oldTitle = titleBuffer;
    SetWindowTextW(hWnd, L"Caricamento in corso - Attendere...");
    UpdateWindow(hWnd);

    // Svuota completamente la cartella temporanea
    ClearFolderContents(g_extractPath);
    CreateDirectoryW(g_extractPath.c_str(), NULL);

    int successCount = 0;
    int errorCount = 0;

    for (const auto& path : paths)
    {
        size_t dotPos = path.find_last_of(L'.');
        std::wstring ext;
        if (dotPos != std::wstring::npos) {
            ext = path.substr(dotPos);
            std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
        }

        if (ext == L".zip")
        {
            if (g_hStatusBar)
            {
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Estrazione ZIP in corso...");
                UpdateWindow(g_hStatusBar);
            }
            if (g_hBrowserWnd)
                FatturaViewer::ShowNotification(g_hBrowserWnd, L"Estrazione archivio ZIP...", L"info");
            if (g_pParser->ExtractZipFile(path, g_extractPath))
                successCount++;
            else
                errorCount++;
        }
        else if (ext == L".p7m")
        {
            if (g_hStatusBar)
            {
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Caricamento P7M...");
                UpdateWindow(g_hStatusBar);
            }
            if (g_hBrowserWnd)
                FatturaViewer::ShowNotification(g_hBrowserWnd, L"Copia file P7M...", L"info");

            size_t lastSlash = path.find_last_of(L"\\/");
            std::wstring fileName = (lastSlash != std::wstring::npos) ? path.substr(lastSlash + 1) : path;
            if (CopyFileW(path.c_str(), (g_extractPath + fileName).c_str(), FALSE))
                successCount++;
            else
                errorCount++;
        }
        else // XML
        {
            size_t lastSlash = path.find_last_of(L"\\/");
            std::wstring fileName = (lastSlash != std::wstring::npos) ? path.substr(lastSlash + 1) : path;
            if (CopyFileW(path.c_str(), (g_extractPath + fileName).c_str(), FALSE))
                successCount++;
            else
                errorCount++;
        }
    }

    SetWindowTextW(hWnd, oldTitle.c_str());
    SetCursor(hOldCursor);
    EnableWindow(g_hListBox, TRUE);
    EnableWindow(g_hToolbar, TRUE);

    if (successCount == 0 && !paths.empty())
    {
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Errore caricamento file");
        if (g_hBrowserWnd)
            FatturaViewer::ShowNotification(g_hBrowserWnd, L"Errore nel caricamento dei file", L"error");
        MessageBoxW(hWnd, L"Errore nel caricamento dei file selezionati!", L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

    if (successCount > 0)
    {
        if (g_hStatusBar)
        {
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Caricamento fatture...");
            UpdateWindow(g_hStatusBar);
        }
        if (g_hBrowserWnd)
            FatturaViewer::ShowNotification(g_hBrowserWnd, L"Lettura fatture in corso...", L"info");

        LoadFattureList(hWnd);

        if (g_fattureInfo.empty())
        {
            if (g_hStatusBar)
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Nessuna fattura elettronica trovata");
            if (g_hBrowserWnd)
                FatturaViewer::ShowNotification(g_hBrowserWnd, L"Nessuna fattura trovata", L"warning");
            MessageBoxW(hWnd, 
                L"Nessuna fattura elettronica valida trovata nei file selezionati.\n\n"
                L"Verificare che:\n"
                L"\u2022 I file XML siano in formato FatturaPA\n"
                L"\u2022 L'archivio ZIP contenga fatture elettroniche valide", 
                L"Nessuna fattura trovata", MB_OK | MB_ICONWARNING);
            return;
        }

        if (g_hStatusBar)
        {
            std::wstring statusText = std::to_wstring(g_fattureInfo.size()) + 
                (g_fattureInfo.size() == 1 ? L" fattura caricata" : L" fatture caricate");
            if (errorCount > 0)
                statusText += L" (" + std::to_wstring(errorCount) + L" errori)";
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)statusText.c_str());
        }

        // Mostra notifica di successo
        if (g_hBrowserWnd)
        {
            std::wstring notifMsg = std::to_wstring(g_fattureInfo.size()) + 
                (g_fattureInfo.size() == 1 ? L" fattura caricata con successo" : L" fatture caricate con successo");
            FatturaViewer::ShowNotification(g_hBrowserWnd, notifMsg, L"success");
        }
    }
}

void OnOpenFile(HWND hWnd)
{
    std::vector<std::wstring> paths = OpenMultipleFilesDialog(hWnd,
        L"File Fatture (ZIP, XML, P7M)\0*.zip;*.xml;*.XML;*.p7m\0File ZIP\0*.zip\0File XML Fattura\0*.xml;*.XML\0File Firmato P7M\0*.p7m\0Tutti i file\0*.*\0",
        L"Apri file fattura (Ctrl+Click per selezione multipla)");

    OnOpenPaths(hWnd, paths);
}

void LoadFattureList(HWND hWnd)
{
    if (!g_pParser)
        return;

    // Pulisci il filtro di ricerca
    g_searchFilter.clear();
    if (g_hSearchEdit)
        SetWindowTextW(g_hSearchEdit, L"");

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

    // Abilita/Disabilita i pulsanti Ministero e Assosoftware in base alle fatture caricate
    if (g_hToolbar)
    {
        BOOL hasFatture = !g_fattureInfo.empty() ? TRUE : FALSE;
        SendMessage(g_hToolbar, TB_ENABLEBUTTON, IDM_APPLY_MINISTERO, MAKELONG(hasFatture, 0));
        SendMessage(g_hToolbar, TB_ENABLEBUTTON, IDM_APPLY_ASSOSOFTWARE, MAKELONG(hasFatture, 0));
    }

    // Imposta l'altezza variabile degli elementi: righe fattura più alte per migliore leggibilità
    if (g_hListBox)
    {
        int itemCount = (int)SendMessage(g_hListBox, LB_GETCOUNT, 0, 0);
        for (int idx = 0; idx < itemCount; ++idx)
        {
            int mapIdx = (idx < (int)g_listBoxToFatturaIndex.size()) ? g_listBoxToFatturaIndex[idx] : -1;
            if (mapIdx == -1)
            {
                // intestazione o riga vuota: altezza leggermente più piccola
                SendMessage(g_hListBox, LB_SETITEMHEIGHT, (WPARAM)idx, (LPARAM)36);
            }
            else
            {
                // riga fattura: altezza aumentata per evidenziare i dati generali
                // Aumentata per ospitare l'icona "pinzetta" (paperclip)
                SendMessage(g_hListBox, LB_SETITEMHEIGHT, (WPARAM)idx, (LPARAM)72);
            }
        }
    }
}

// Ripopola la ListBox applicando il filtro di ricerca corrente (g_searchFilter).
// Non ricarica i dati dal disco: usa g_fattureInfo già in memoria.
void RefreshFattureListFiltered()
{
    if (!g_hListBox)
        return;

    SendMessage(g_hListBox, LB_RESETCONTENT, 0, 0);
    g_listBoxToFatturaIndex.clear();

    // Prepara il testo di ricerca in minuscolo per confronto case-insensitive
    std::wstring filter = g_searchFilter;
    std::transform(filter.begin(), filter.end(), filter.begin(), ::towlower);

    // Determina quali fatture corrispondono al filtro
    std::vector<bool> match(g_fattureInfo.size(), false);
    bool hasFilter = !filter.empty();
    int matchCount = 0;

    for (size_t i = 0; i < g_fattureInfo.size(); i++)
    {
        if (!hasFilter)
        {
            match[i] = true;
            matchCount++;
            continue;
        }

        const auto& info = g_fattureInfo[i];

        // Cerca nel cedente (emittente)
        std::wstring field = info.cedenteDenominazione;
        std::transform(field.begin(), field.end(), field.begin(), ::towlower);
        if (field.find(filter) != std::wstring::npos) { match[i] = true; matchCount++; continue; }

        // Cerca nel cessionario (cliente)
        field = info.cessionarioDenominazione;
        std::transform(field.begin(), field.end(), field.begin(), ::towlower);
        if (field.find(filter) != std::wstring::npos) { match[i] = true; matchCount++; continue; }

        // Cerca nel numero fattura
        field = info.numero;
        std::transform(field.begin(), field.end(), field.begin(), ::towlower);
        if (field.find(filter) != std::wstring::npos) { match[i] = true; matchCount++; continue; }

        // Cerca nella data
        field = info.data;
        std::transform(field.begin(), field.end(), field.begin(), ::towlower);
        if (field.find(filter) != std::wstring::npos) { match[i] = true; matchCount++; continue; }

        // Cerca nell'importo
        field = info.importoTotale;
        std::transform(field.begin(), field.end(), field.begin(), ::towlower);
        if (field.find(filter) != std::wstring::npos) { match[i] = true; matchCount++; continue; }

        // Cerca nel tipo documento
        field = info.tipoDocumento;
        std::transform(field.begin(), field.end(), field.begin(), ::towlower);
        if (field.find(filter) != std::wstring::npos) { match[i] = true; matchCount++; continue; }
    }

    // Ripopola la ListBox con raggruppamento per emittente (solo fatture filtrate)
    std::wstring currentEmittente;

    for (size_t i = 0; i < g_fattureInfo.size(); i++)
    {
        if (!match[i])
            continue;

        const auto& info = g_fattureInfo[i];

        if (info.cedenteDenominazione != currentEmittente)
        {
            currentEmittente = info.cedenteDenominazione;

            // Riga vuota di separazione (tranne per il primo gruppo)
            if (SendMessage(g_hListBox, LB_GETCOUNT, 0, 0) > 0)
            {
                SendMessage(g_hListBox, LB_ADDSTRING, 0, (LPARAM)L"");
                g_listBoxToFatturaIndex.push_back(-1);
            }

            // Intestazione del gruppo emittente
            std::wstring header = L"\x25BA " + currentEmittente;
            SendMessage(g_hListBox, LB_ADDSTRING, 0, (LPARAM)header.c_str());
            g_listBoxToFatturaIndex.push_back(-1);
        }

        SendMessage(g_hListBox, LB_ADDSTRING, 0, (LPARAM)info.GetDisplayText().c_str());
        g_listBoxToFatturaIndex.push_back((int)i);
    }

    // Imposta altezza variabile degli elementi
    int itemCount = (int)SendMessage(g_hListBox, LB_GETCOUNT, 0, 0);
    for (int idx = 0; idx < itemCount; ++idx)
    {
        int mapIdx = (idx < (int)g_listBoxToFatturaIndex.size()) ? g_listBoxToFatturaIndex[idx] : -1;
        SendMessage(g_hListBox, LB_SETITEMHEIGHT, (WPARAM)idx,
            (LPARAM)(mapIdx == -1 ? 36 : 72));
    }

    // Aggiorna la status bar con il conteggio filtrato
    if (g_hStatusBar)
    {
        if (hasFilter)
        {
            std::wstring msg = std::to_wstring(matchCount) + L" di " +
                std::to_wstring(g_fattureInfo.size()) + L" fatture";
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)msg.c_str());
        }
        else
        {
            std::wstring msg = std::to_wstring(g_fattureInfo.size()) + L" fatture caricate";
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)msg.c_str());
        }
    }
}

// Costruisce la sezione HTML degli allegati da iniettare nella fattura visualizzata
static std::wstring BuildAttachmentsHtml(const std::vector<AllegatoInfo>& allegati)
{
    if (allegati.empty()) return L"";

    std::wstring html;
    // Use the original simple emoji entities for the HTML attachments section
    // (restore original behaviour: show paperclip emoji in the HTML view)
    // Align attachments to the left (not centered) so they visually match the invoice list
    html += L"<div style=\"margin:20px 0;max-width:800px;padding:14px 16px;"
            L"border:2px solid #6366f1;border-radius:10px;"
            L"font-family:'Segoe UI',Arial,sans-serif;background:#f0f4ff;\">";
    html += L"<p style=\"margin:0 0 10px 0;font-weight:bold;color:#4f46e5;font-size:1em;\">";
    // Original: use paperclip emoji for the header in the HTML view
    html += L"&#128206; Allegati (" + std::to_wstring(allegati.size()) + L")</p>";
    html += L"<ul style=\"list-style:none;padding:0;margin:0;\">";

    for (const auto& att : allegati)
    {
        // Costruisce URL fvatt:///C:/path/to/file
        std::wstring url = L"fvatt:///";
        for (wchar_t c : att.filePath)
        {
            if      (c == L'\\') url += L'/';
            else if (c == L' ')  url += L"%20";
            else                 url += c;
        }

        html += L"<li style=\"padding:7px 0;border-bottom:1px solid #c7d2fe;\">";
        html += L"<a href=\"" + url + L"\" style=\"color:#4f46e5;font-weight:600;text-decoration:none;display:inline-flex;align-items:center;\">";
        // Use the same paperclip emoji as the header for consistency
        html += L"&#128206; " + att.nomeAttachment + L"</a>";
        if (!att.descrizione.empty())
            html += L" <span style=\"color:#6b7280;font-size:0.9em;\"> - " + att.descrizione + L"</span>";
        if (!att.formato.empty())
            html += L" <span style=\"background:#e0e7ff;color:#4f46e5;padding:2px 8px;"
                    L"border-radius:10px;font-size:0.78em;font-weight:600;\">" + att.formato + L"</span>";
        html += L"</li>";
    }

    html += L"</ul></div>";
    return html;
}

// Inietta la sezione allegati nell'HTML prima di </body>
static void InjectAttachmentsIntoHtml(std::wstring& htmlOutput, const std::vector<AllegatoInfo>& allegati)
{
    if (allegati.empty()) return;
    std::wstring section = BuildAttachmentsHtml(allegati);
    size_t pos = htmlOutput.rfind(L"</body>");
    if (pos == std::wstring::npos) pos = htmlOutput.rfind(L"</BODY>");
    if (pos != std::wstring::npos)
        htmlOutput.insert(pos, section);
    else
        htmlOutput += section;
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
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Caricamento fattura...");
        UpdateWindow(g_hStatusBar);
    }

    // Carica il file XML
    if (!g_pParser->LoadXmlFile(g_currentXmlPath))
    {
        SetCursor(hOldCursor);
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Errore caricamento XML");
        MessageBoxW(hWnd, L"Errore nel caricamento del file XML!", L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

    // Usa il foglio di stile configurato
    std::wstring xsltPath = g_lastXsltPath;

    // Se non c'è nessun foglio configurato, usa automaticamente Assosoftware
    if (xsltPath.empty() || GetFileAttributesW(xsltPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        std::wstring resourcesPath = GetResourcesPath();
        WIN32_FIND_DATAW findData;
        HANDLE hFind = FindFirstFileW((resourcesPath + L"*Asso*.xsl").c_str(), &findData);
        if (hFind != INVALID_HANDLE_VALUE)
        {
            xsltPath = resourcesPath + findData.cFileName;
            FindClose(hFind);
            g_lastXsltPath = xsltPath;
            SaveStylesheetPreference(xsltPath);
        }
        if (xsltPath.empty() || GetFileAttributesW(xsltPath.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            SetCursor(hOldCursor);
            if (g_hStatusBar)
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Foglio di stile non trovato");
            return;
        }
    }

    if (g_hStatusBar)
    {
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Applicazione foglio di stile...");
        UpdateWindow(g_hStatusBar);
    }
    if (g_hBrowserWnd)
        FatturaViewer::ShowNotification(g_hBrowserWnd, L"Applicazione foglio di stile...", L"info");

    // Applica la trasformazione
    std::wstring htmlOutput;
    if (g_pParser->ApplyXsltTransform(xsltPath, htmlOutput))
    {
        if (g_hStatusBar)
        {
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Rendering HTML...");
            UpdateWindow(g_hStatusBar);
        }
        if (g_hBrowserWnd)
            FatturaViewer::ShowNotification(g_hBrowserWnd, L"Generazione visualizzazione...", L"info");

        // Estrai allegati embedded e iniettali nell'HTML
        std::wstring attachFolder = g_extractPath + L"allegati\\";
        std::vector<AllegatoInfo> allegati = g_pParser->ExtractAttachments(g_currentXmlPath, attachFolder);
        InjectAttachmentsIntoHtml(htmlOutput, allegati);

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
        {
            std::wstring statusMsg = L"Fattura visualizzata";
            if (!allegati.empty())
                statusMsg += L" | " + std::to_wstring(allegati.size()) +
                    (allegati.size() == 1 ? L" allegato" : L" allegati");
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)statusMsg.c_str());
        }
        if (g_hBrowserWnd)
        {
            std::wstring notifMsg = L"Fattura visualizzata";
            if (!allegati.empty())
                notifMsg += L" \u2022 " + std::to_wstring(allegati.size()) +
                    (allegati.size() == 1 ? L" allegato disponibile" : L" allegati disponibili");
            FatturaViewer::ShowNotification(g_hBrowserWnd, notifMsg, L"success");
        }
    }
    else
    {
        // ===== FINE LOADER - ERRORE =====
        SetCursor(hOldCursor);
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Errore trasformazione");
        if (g_hBrowserWnd)
            FatturaViewer::ShowNotification(g_hBrowserWnd, L"Errore trasformazione XSLT", L"error");
        MessageBoxW(hWnd, L"Errore durante l'applicazione del foglio di stile!", L"Errore", MB_OK | MB_ICONERROR);
    }
}

// Restituisce gli indici g_fattureInfo degli elementi selezionati nella ListBox
std::vector<int> GetSelectedFattureIndices()
{
    std::vector<int> result;
    if (!g_hListBox)
        return result;

    int selCount = (int)SendMessage(g_hListBox, LB_GETSELCOUNT, 0, 0);
    if (selCount <= 0)
        return result;

    std::vector<int> selItems(selCount);
    SendMessage(g_hListBox, LB_GETSELITEMS, selCount, (LPARAM)selItems.data());

    for (int lbIdx : selItems)
    {
        if (lbIdx < (int)g_listBoxToFatturaIndex.size())
        {
            int fatIdx = g_listBoxToFatturaIndex[lbIdx];
            if (fatIdx >= 0 && fatIdx < (int)g_fattureInfo.size())
                result.push_back(fatIdx);
        }
    }
    return result;
}

void UpdateNavButtons()
{
    BOOL hasPrev = (g_currentHtmlIndex > 0) ? TRUE : FALSE;
    BOOL hasNext = (g_currentHtmlIndex >= 0 && g_currentHtmlIndex < (int)g_multipleHtmlFiles.size() - 1) ? TRUE : FALSE;

    if (g_hBtnPrev) EnableWindow(g_hBtnPrev, hasPrev);
    if (g_hBtnNext) EnableWindow(g_hBtnNext, hasNext);

    if (g_hStatusBar && g_currentHtmlIndex >= 0 && !g_multipleHtmlFiles.empty())
    {
        std::wstring msg = L"Fattura " + std::to_wstring(g_currentHtmlIndex + 1) +
            L" di " + std::to_wstring(g_multipleHtmlFiles.size());
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)msg.c_str());
    }
}

void OnZoom(HWND hWnd, bool zoomIn)
{
    if (!g_hBrowserWnd)
        return;

    if (zoomIn)
    {
        if (g_zoomLevel < 200)
            g_zoomLevel += 25;
    }
    else
    {
        if (g_zoomLevel > 50)
            g_zoomLevel -= 25;
    }

    FatturaViewer::SetZoom(g_hBrowserWnd, g_zoomLevel);

    if (g_hLabelZoom)
        SetWindowTextW(g_hLabelZoom, (std::to_wstring(g_zoomLevel) + L"%").c_str());

    if (g_hBtnZoomIn)  EnableWindow(g_hBtnZoomIn,  g_zoomLevel < 200 ? TRUE : FALSE);
    if (g_hBtnZoomOut) EnableWindow(g_hBtnZoomOut, g_zoomLevel > 50  ? TRUE : FALSE);
}

// Visualizza più fatture come file HTML separati, navigabili con Prec/Succ
void OnApplyStylesheetMultiple(HWND hWnd, const std::wstring& xsltPath, const std::vector<int>& indices)
{
    if (!g_pParser || indices.empty())
        return;

    HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));

    // Salva la posizione corrente per mantenerla dopo il re-render
    int savedIndex = g_currentHtmlIndex;

    g_multipleHtmlFiles.clear();
    g_currentHtmlIndex = -1;

    size_t total = indices.size();

    for (size_t i = 0; i < total; i++)
    {
        int fatIdx = indices[i];

        if (g_hStatusBar)
        {
            std::wstring msg = L"Elaborazione " + std::to_wstring(i + 1) + L" di " + std::to_wstring(total) + L"...";
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)msg.c_str());
            UpdateWindow(g_hStatusBar);
        }

        if (!g_pParser->LoadXmlFile(g_fattureInfo[fatIdx].filePath))
            continue;

        std::wstring htmlOutput;
        if (!g_pParser->ApplyXsltTransform(xsltPath, htmlOutput))
            continue;

        // Estrai allegati per questa fattura e iniettali nell'HTML
        std::wstring attachFolder = g_extractPath + L"allegati_seq_" + std::to_wstring(i) + L"\\";
        std::vector<AllegatoInfo> allegati = g_pParser->ExtractAttachments(g_fattureInfo[fatIdx].filePath, attachFolder);
        InjectAttachmentsIntoHtml(htmlOutput, allegati);

        // Salva ogni fattura in un file HTML separato
        std::wstring htmlPath = g_extractPath + L"fattura_seq_" + std::to_wstring(i) + L".html";
        FatturaViewer::SaveHTMLToFile(htmlOutput, htmlPath);
        g_multipleHtmlFiles.push_back(htmlPath);
    }

    SetCursor(hOldCursor);

    if (g_multipleHtmlFiles.empty())
    {
        MessageBoxW(hWnd, L"Nessuna fattura elaborata correttamente!", L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

    // Ripristina la posizione precedente se valida, altrimenti vai alla prima
    g_currentHtmlIndex = (savedIndex >= 0 && savedIndex < (int)g_multipleHtmlFiles.size()) ? savedIndex : 0;
    if (g_hBrowserWnd)
        FatturaViewer::NavigateToHTML(g_hBrowserWnd, g_multipleHtmlFiles[g_currentHtmlIndex]);

    UpdateNavButtons();
}

void OnApplySpecificStylesheet(HWND hWnd, const std::wstring& keyword)
{
    // Ottieni tutte le fatture selezionate
    std::vector<int> selectedIndices = GetSelectedFattureIndices();

    if (selectedIndices.empty())
    {
        MessageBoxW(hWnd, L"Seleziona prima una fattura dalla lista!", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    // Cerca il foglio di stile per keyword nella cartella Resources
    std::wstring resourcesPath = GetResourcesPath();
    std::wstring searchPattern = resourcesPath + L"*" + keyword + L"*.xsl";

    WIN32_FIND_DATAW findData;
    HANDLE hFind = FindFirstFileW(searchPattern.c_str(), &findData);

    std::wstring xsltPath;
    if (hFind != INVALID_HANDLE_VALUE)
    {
        xsltPath = resourcesPath + findData.cFileName;
        FindClose(hFind);
    }

    if (xsltPath.empty() || GetFileAttributesW(xsltPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        std::wstring msg = L"Foglio di stile '" + keyword + L"' non trovato nella cartella Resources!";
        MessageBoxW(hWnd, msg.c_str(), L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

    g_lastXsltPath = xsltPath;
    SaveStylesheetPreference(xsltPath);

    if (g_hStatusBar)
    {
        std::wstring statusMsg = L"Foglio di stile: " + std::wstring(findData.cFileName);
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)statusMsg.c_str());
    }

    if (selectedIndices.size() == 1)
    {
        // Selezione singola - comportamento normale
        g_currentXmlPath = g_fattureInfo[selectedIndices[0]].filePath;
        OnApplyStylesheet(hWnd);
    }
    else
    {
        // Selezione multipla - visualizza tutte in sequenza
        OnApplyStylesheetMultiple(hWnd, xsltPath, selectedIndices);
    }
}

std::wstring FindEdgePath()
{
    WCHAR buffer[MAX_PATH];
    const wchar_t* paths[] = {
        L"%ProgramFiles(x86)%\\Microsoft\\Edge\\Application\\msedge.exe",
        L"%ProgramFiles%\\Microsoft\\Edge\\Application\\msedge.exe",
        L"%LocalAppData%\\Microsoft\\Edge\\Application\\msedge.exe"
    };
    for (const auto& p : paths)
    {
        ExpandEnvironmentStringsW(p, buffer, MAX_PATH);
        if (GetFileAttributesW(buffer) != INVALID_FILE_ATTRIBUTES)
            return std::wstring(buffer);
    }
    return L"";
}

std::wstring ExtractBodyContent(const std::wstring& html)
{
    size_t bodyStart = html.find(L"<body");
    if (bodyStart == std::wstring::npos)
        bodyStart = html.find(L"<BODY");
    if (bodyStart == std::wstring::npos)
        return html;
    size_t gt = html.find(L">", bodyStart);
    if (gt == std::wstring::npos)
        return html;
    size_t bodyEnd = html.find(L"</body>");
    if (bodyEnd == std::wstring::npos)
        bodyEnd = html.find(L"</BODY>");
    if (bodyEnd == std::wstring::npos)
        return html;
    return html.substr(gt + 1, bodyEnd - gt - 1);
}

void OnPrintToPDF(HWND hWnd)
{
    std::vector<int> selectedIndices = GetSelectedFattureIndices();
    if (selectedIndices.empty())
    {
        MessageBoxW(hWnd, L"Seleziona almeno una fattura dalla lista!", L"Attenzione", MB_OK | MB_ICONWARNING);
        return;
    }

    if (!g_pParser)
        return;

    if (g_lastXsltPath.empty() || GetFileAttributesW(g_lastXsltPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        if (DialogBox(hInst, MAKEINTRESOURCE(IDD_STYLESHEET_SELECTOR), hWnd, StylesheetSelector) != IDOK || g_lastXsltPath.empty())
            return;
    }

    HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));
    if (g_hStatusBar)
    {
        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Generazione PDF in corso...");
        UpdateWindow(g_hStatusBar);
    }

    std::vector<std::wstring> htmlContents;
    size_t total = selectedIndices.size();

    for (size_t i = 0; i < total; i++)
    {
        int fatIdx = selectedIndices[i];
        if (g_hStatusBar)
        {
            std::wstring msg = L"Elaborazione fattura " + std::to_wstring(i + 1) + L" di " + std::to_wstring(total) + L"...";
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)msg.c_str());
            UpdateWindow(g_hStatusBar);
        }
        if (!g_pParser->LoadXmlFile(g_fattureInfo[fatIdx].filePath))
            continue;
        std::wstring htmlOutput;
        if (!g_pParser->ApplyXsltTransform(g_lastXsltPath, htmlOutput))
            continue;
        htmlContents.push_back(htmlOutput);
    }

    SetCursor(hOldCursor);

    if (htmlContents.empty())
    {
        MessageBoxW(hWnd, L"Errore nella generazione HTML per le fatture selezionate!", L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

    // Genera HTML singolo o combinato con page-break
    std::wstring htmlPath = g_extractPath + L"fatture_pdf_temp.html";

    if (htmlContents.size() == 1)
    {
        FatturaViewer::SaveHTMLToFile(htmlContents[0], htmlPath);
    }
    else
    {
        const std::wstring& first = htmlContents[0];
        size_t headStart = first.find(L"<head");
        size_t headEnd   = first.find(L"</head>");
        std::wstring combined = L"<!DOCTYPE html><html lang=\"it\">";
        if (headStart != std::wstring::npos && headEnd != std::wstring::npos)
            combined += first.substr(headStart, headEnd - headStart + 7);
        else
            combined += L"<head><meta charset=\"UTF-8\"></head>";
        combined += L"<body style=\"margin:0;padding:0;\">";
        combined += L"<style>@media print{.pfb{page-break-after:always;}}</style>";
        for (size_t i = 0; i < htmlContents.size(); i++)
        {
            combined += L"<div>" + ExtractBodyContent(htmlContents[i]) + L"</div>";
            if (i < htmlContents.size() - 1)
                combined += L"<div class=\"pfb\"></div>";
        }
        combined += L"</body></html>";
        FatturaViewer::SaveHTMLToFile(combined, htmlPath);
    }

    // Converti HTML in PDF tramite Microsoft Edge (headless)
    std::wstring edgePath = FindEdgePath();
    std::wstring pdfPath  = g_extractPath + L"FattureSelezionate.pdf";
    DeleteFileW(pdfPath.c_str());

    if (!edgePath.empty())
    {
        std::wstring fileUrl = L"file:///" + htmlPath;
        for (auto& c : fileUrl)
            if (c == L'\\') c = L'/';

        // Parametri Edge con metadati PDF per ridurre avvisi di sicurezza
        std::wstring params = L"--headless=new --disable-gpu --no-margins --no-pdf-header-footer "
            L"--enable-features=NetworkService,NetworkServiceInProcess "
            L"--disable-features=IsolateOrigins,site-per-process "
            L"--print-to-pdf=\"" + pdfPath + L"\" \"" + fileUrl + L"\"";

        SHELLEXECUTEINFOW sei = { sizeof(sei) };
        sei.fMask        = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NO_CONSOLE;
        sei.lpVerb       = NULL;
        sei.lpFile       = edgePath.c_str();
        sei.lpParameters = params.c_str();
        sei.nShow        = SW_HIDE;

        hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));
        if (ShellExecuteExW(&sei))
        {
            WaitForSingleObject(sei.hProcess, 30000);
            CloseHandle(sei.hProcess);
        }
        SetCursor(hOldCursor);

        if (GetFileAttributesW(pdfPath.c_str()) != INVALID_FILE_ATTRIBUTES)
        {
            // Firma il PDF se configurato
            PdfSignConfig signConfig;
            if (PdfSigner::LoadConfig(signConfig) && signConfig.enabled)
            {
                if (g_hStatusBar)
                {
                    SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Firma digitale del PDF in corso...");
                    UpdateWindow(g_hStatusBar);
                }

                std::wstring errorMsg;
                hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));
                bool signSuccess = PdfSigner::SignPdf(pdfPath, signConfig, errorMsg);
                SetCursor(hOldCursor);

                if (!signSuccess)
                {
                    std::wstring msg = L"Errore durante la firma del PDF:\n" + errorMsg + 
                                      L"\n\nIl PDF non firmato sarà comunque aperto.";
                    MessageBoxW(hWnd, msg.c_str(), L"Avviso", MB_OK | MB_ICONWARNING);
                }
                else
                {
                    if (g_hStatusBar)
                        SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"PDF firmato digitalmente");
                    if (g_hBrowserWnd)
                        FatturaViewer::ShowNotification(g_hBrowserWnd, L"PDF firmato digitalmente", L"success");
                }
            }

            ShellExecuteW(hWnd, L"open", pdfPath.c_str(), NULL, NULL, SW_SHOW);
            if (g_hStatusBar)
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"PDF generato e aperto");
            if (g_hBrowserWnd)
                FatturaViewer::ShowNotification(g_hBrowserWnd, L"PDF generato e aperto con successo", L"success");
        }
        else
        {
            ShellExecuteW(hWnd, L"open", htmlPath.c_str(), NULL, NULL, SW_SHOW);
            if (g_hStatusBar)
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Apertura HTML nel browser predefinito");
        }
    }
    else
    {
        ShellExecuteW(hWnd, L"open", htmlPath.c_str(), NULL, NULL, SW_SHOW);
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Edge non trovato: apertura HTML nel browser");
    }
}

// OnPrintAll rimossa: usare OnPrintToPDF (seleziona tutte le fatture per stampare tutto)
void OnPrintAll(HWND hWnd)
{
    // Seleziona tutte le fatture e stampa come PDF
    if (g_hListBox && !g_fattureInfo.empty())
    {
        SendMessage(g_hListBox, LB_SETSEL, TRUE, -1);
        OnPrintToPDF(hWnd);
    }
}

    // Crea un'icona badge moderna: rettangolo arrotondato con gradiente, bordo e ombra testo
static HICON CreateBadgeIcon(int size, COLORREF bgColor, const wchar_t* text)
{
    HDC hScreenDC = GetDC(NULL);

    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = size;
    bmi.bmiHeader.biHeight = size; // bottom-up per CreateIconIndirect
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;

    DWORD* pBits = NULL;
    HBITMAP hColor = CreateDIBSection(hScreenDC, &bmi, DIB_RGB_COLORS, (void**)&pBits, NULL, 0);
    if (!hColor) { ReleaseDC(NULL, hScreenDC); return NULL; }

    HDC hMemDC = CreateCompatibleDC(hScreenDC);
    HBITMAP hOldBmp = (HBITMAP)SelectObject(hMemDC, hColor);
    memset(pBits, 0, size * size * 4);

    int rBase = GetRValue(bgColor);
    int gBase = GetGValue(bgColor);
    int bBase = GetBValue(bgColor);

    // Forma: rettangolo arrotondato stile squircle moderno
    int margin = 2;
    int cornerRadius = size / 5;

    // Gradiente verticale clippato alla forma arrotondata (chiaro in alto, scuro in basso)
    HRGN hRgn = CreateRoundRectRgn(margin, margin, size - margin + 1, size - margin + 1,
                                     cornerRadius * 2, cornerRadius * 2);
    SelectClipRgn(hMemDC, hRgn);

    TRIVERTEX vert[2] = {};
    vert[0].x = 0;
    vert[0].y = 0;
    vert[0].Red   = (USHORT)(min(255, rBase + 40) << 8);
    vert[0].Green = (USHORT)(min(255, gBase + 40) << 8);
    vert[0].Blue  = (USHORT)(min(255, bBase + 40) << 8);
    vert[1].x = size;
    vert[1].y = size;
    vert[1].Red   = (USHORT)(max(0, rBase - 25) << 8);
    vert[1].Green = (USHORT)(max(0, gBase - 25) << 8);
    vert[1].Blue  = (USHORT)(max(0, bBase - 25) << 8);
    GRADIENT_RECT gRect = { 0, 1 };
    GradientFill(hMemDC, vert, 2, &gRect, 1, GRADIENT_FILL_RECT_V);

    SelectClipRgn(hMemDC, NULL);
    DeleteObject(hRgn);

    // Bordo sottile più scuro per definizione
    HBRUSH hOldBr = (HBRUSH)SelectObject(hMemDC, GetStockObject(NULL_BRUSH));
    HPEN hBorderPen = CreatePen(PS_SOLID, 1,
        RGB(max(0, rBase - 50), max(0, gBase - 50), max(0, bBase - 50)));
    HPEN hOldPn = (HPEN)SelectObject(hMemDC, hBorderPen);
    RoundRect(hMemDC, margin, margin, size - margin, size - margin,
              cornerRadius * 2, cornerRadius * 2);
    SelectObject(hMemDC, hOldPn);
    SelectObject(hMemDC, hOldBr);
    DeleteObject(hBorderPen);

    // Lettera con ombra sottile per profondità
    SetBkMode(hMemDC, TRANSPARENT);
    int textLen = lstrlenW(text);
    int fontSize = (textLen > 1) ? -(size * 3 / 8) : -(size * 5 / 8);
    HFONT hFont = CreateFontW(fontSize, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    HFONT hOldFont = (HFONT)SelectObject(hMemDC, hFont);

    // Ombra della lettera (1px offset in basso a destra)
    SetTextColor(hMemDC, RGB(max(0, rBase - 70), max(0, gBase - 70), max(0, bBase - 70)));
    RECT rcShadow = { 1, 1, size + 1, size + 1 };
    DrawTextW(hMemDC, text, -1, &rcShadow, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    // Lettera bianca
    SetTextColor(hMemDC, RGB(255, 255, 255));
    RECT rc = { 0, 0, size, size };
    DrawTextW(hMemDC, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    SelectObject(hMemDC, hOldFont);
    DeleteObject(hFont);

    SelectObject(hMemDC, hOldBmp);
    DeleteDC(hMemDC);

    // Imposta alpha a 255 per i pixel disegnati (pre-multiplied alpha)
    for (int i = 0; i < size * size; i++)
    {
        if (pBits[i] & 0x00FFFFFF)
            pBits[i] |= 0xFF000000;
    }

    // Maschera monocromatica (tutto nero = completamente opaco)
    HBITMAP hMask = CreateBitmap(size, size, 1, 1, NULL);
    HDC hMaskDC = CreateCompatibleDC(hScreenDC);
    HBITMAP hOldMask = (HBITMAP)SelectObject(hMaskDC, hMask);
    RECT rcMask = { 0, 0, size, size };
    FillRect(hMaskDC, &rcMask, (HBRUSH)GetStockObject(BLACK_BRUSH));
    SelectObject(hMaskDC, hOldMask);
    DeleteDC(hMaskDC);
    ReleaseDC(NULL, hScreenDC);

    ICONINFO ii = {};
    ii.fIcon = TRUE;
    ii.hbmColor = hColor;
    ii.hbmMask = hMask;
    HICON hIcon = CreateIconIndirect(&ii);

    DeleteObject(hColor);
    DeleteObject(hMask);
    return hIcon;
}

// Crea un'icona stile documento Office: pagina bianca con fascia colorata a sinistra e lettera
static HICON CreateDocumentBadgeIcon(int size, COLORREF accentColor, const wchar_t* letter)
{
    HDC hScreenDC = GetDC(NULL);

    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = size;
    bmi.bmiHeader.biHeight = size;
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;

    DWORD* pBits = NULL;
    HBITMAP hColor = CreateDIBSection(hScreenDC, &bmi, DIB_RGB_COLORS, (void**)&pBits, NULL, 0);
    if (!hColor) { ReleaseDC(NULL, hScreenDC); return NULL; }

    HDC hMemDC = CreateCompatibleDC(hScreenDC);
    HBITMAP hOldBmp = (HBITMAP)SelectObject(hMemDC, hColor);
    memset(pBits, 0, size * size * 4);

    int rBase = GetRValue(accentColor);
    int gBase = GetGValue(accentColor);
    int bBase = GetBValue(accentColor);

    int margin = 1;
    int stripeW = size * 2 / 5;
    int foldSize = size / 7;

    // 1. Corpo documento (sfondo bianco/grigio chiaro con bordo sottile)
    HBRUSH hDocBr = CreateSolidBrush(RGB(250, 250, 254));
    HPEN hDocPen = CreatePen(PS_SOLID, 1, RGB(195, 200, 212));
    HBRUSH hOldBr = (HBRUSH)SelectObject(hMemDC, hDocBr);
    HPEN hOldPn = (HPEN)SelectObject(hMemDC, hDocPen);
    RoundRect(hMemDC, stripeW - 2, margin, size - margin, size - margin, 2, 2);
    SelectObject(hMemDC, hOldPn);
    SelectObject(hMemDC, hOldBr);
    DeleteObject(hDocPen);
    DeleteObject(hDocBr);

    // 2. Righe di testo simulate (linee grigie orizzontali)
    HPEN hLinePen = CreatePen(PS_SOLID, 1, RGB(215, 218, 228));
    HPEN hOldLinePen = (HPEN)SelectObject(hMemDC, hLinePen);
    int lineX0 = stripeW + 3;
    int lineX1 = size - margin - 4;
    int lineY = margin + foldSize + 4;
    for (int i = 0; i < 4 && lineY < size - margin - 5; i++)
    {
        MoveToEx(hMemDC, lineX0, lineY, NULL);
        LineTo(hMemDC, lineX1 - (i == 2 ? 5 : 0), lineY);
        lineY += 5;
    }
    SelectObject(hMemDC, hOldLinePen);
    DeleteObject(hLinePen);

    // 3. Angolo piegato (dog-ear) in alto a destra
    POINT foldPts[3] = {
        { size - margin - foldSize, margin },
        { size - margin,            margin + foldSize },
        { size - margin - foldSize, margin + foldSize }
    };
    HBRUSH hFoldBr = CreateSolidBrush(RGB(230, 232, 242));
    HPEN hFoldPen = CreatePen(PS_SOLID, 1, RGB(195, 200, 212));
    SelectObject(hMemDC, hFoldBr);
    SelectObject(hMemDC, hFoldPen);
    Polygon(hMemDC, foldPts, 3);
    DeleteObject(hFoldBr);
    DeleteObject(hFoldPen);

    // 4. Fascia colorata a sinistra con gradiente
    HRGN hStripeRgn = CreateRoundRectRgn(margin, margin + 2, stripeW + 1, size - margin - 1, 5, 5);
    SelectClipRgn(hMemDC, hStripeRgn);

    TRIVERTEX vert[2] = {};
    vert[0].x = margin;
    vert[0].y = margin;
    vert[0].Red   = (USHORT)(min(255, rBase + 35) << 8);
    vert[0].Green = (USHORT)(min(255, gBase + 35) << 8);
    vert[0].Blue  = (USHORT)(min(255, bBase + 35) << 8);
    vert[1].x = stripeW + 1;
    vert[1].y = size - margin;
    vert[1].Red   = (USHORT)(max(0, rBase - 20) << 8);
    vert[1].Green = (USHORT)(max(0, gBase - 20) << 8);
    vert[1].Blue  = (USHORT)(max(0, bBase - 20) << 8);
    GRADIENT_RECT gRect = { 0, 1 };
    GradientFill(hMemDC, vert, 2, &gRect, 1, GRADIENT_FILL_RECT_V);

    SelectClipRgn(hMemDC, NULL);
    DeleteObject(hStripeRgn);

    // 5. Lettera bianca sulla fascia colorata
    SetBkMode(hMemDC, TRANSPARENT);
    SetTextColor(hMemDC, RGB(255, 255, 255));
    HFONT hFont = CreateFontW(-(stripeW * 3 / 4), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    HFONT hOldFont = (HFONT)SelectObject(hMemDC, hFont);
    RECT rcLetter = { margin, margin, stripeW + 1, size - margin };
    DrawTextW(hMemDC, letter, -1, &rcLetter, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hMemDC, hOldFont);
    DeleteObject(hFont);

    SelectObject(hMemDC, hOldBmp);
    DeleteDC(hMemDC);

    for (int i = 0; i < size * size; i++)
    {
        if (pBits[i] & 0x00FFFFFF)
            pBits[i] |= 0xFF000000;
    }

    HBITMAP hMask = CreateBitmap(size, size, 1, 1, NULL);
    HDC hMaskDC = CreateCompatibleDC(hScreenDC);
    HBITMAP hOldMask = (HBITMAP)SelectObject(hMaskDC, hMask);
    RECT rcMask = { 0, 0, size, size };
    FillRect(hMaskDC, &rcMask, (HBRUSH)GetStockObject(BLACK_BRUSH));
    SelectObject(hMaskDC, hOldMask);
    DeleteDC(hMaskDC);
    ReleaseDC(NULL, hScreenDC);

    ICONINFO ii = {};
    ii.fIcon = TRUE;
    ii.hbmColor = hColor;
    ii.hbmMask = hMask;
    HICON hIcon = CreateIconIndirect(&ii);

    DeleteObject(hColor);
    DeleteObject(hMask);
    return hIcon;
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

    SendMessage(hToolbar, TB_BUTTONSTRUCTSIZE, (WPARAM)sizeof(TBBUTTON), 0);

    // Crea image list 32x32 con supporto alpha per icone moderne
    const int iconSize = 32;
    HIMAGELIST hImageList = ImageList_Create(iconSize, iconSize, ILC_COLOR32 | ILC_MASK, 4, 0);

    // Indice 0: Apri - icona cartella aperta dal sistema (Windows 10/11 nativa)
    SHSTOCKICONINFO sii = {};
    sii.cbSize = sizeof(sii);
    if (SUCCEEDED(SHGetStockIconInfo(SIID_FOLDEROPEN, SHGSI_ICON | SHGSI_LARGEICON, &sii)))
    {
        ImageList_AddIcon(hImageList, sii.hIcon);
        DestroyIcon(sii.hIcon);
    }
    else
    {
        HICON hFallback = (HICON)LoadImage(NULL, IDI_APPLICATION, IMAGE_ICON, iconSize, iconSize, LR_SHARED);
        ImageList_AddIcon(hImageList, hFallback);
    }

    // Indice 1: Ministero - icona stile documento con fascia indigo e "M"
    HICON hIconM = CreateDocumentBadgeIcon(iconSize, RGB(79, 70, 229), L"M");
    if (hIconM) { ImageList_AddIcon(hImageList, hIconM); DestroyIcon(hIconM); }

    // Indice 2: Assosoftware - icona stile documento con fascia emerald e "A"
    HICON hIconA = CreateDocumentBadgeIcon(iconSize, RGB(16, 185, 129), L"A");
    if (hIconA) { ImageList_AddIcon(hImageList, hIconA); DestroyIcon(hIconA); }

    // Indice 3: Stampa PDF - badge amber/arancio con "PDF"
    HICON hIconPdf = CreateBadgeIcon(iconSize, RGB(245, 158, 11), L"PDF");
    if (hIconPdf) { ImageList_AddIcon(hImageList, hIconPdf); DestroyIcon(hIconPdf); }

    SendMessage(hToolbar, TB_SETIMAGELIST, 0, (LPARAM)hImageList);
    SendMessage(hToolbar, TB_SETBUTTONSIZE, 0, MAKELONG(48, 48));

    // Pulsanti con indici image list (0=Apri, 1=Ministero, 2=Assosoftware, 3=Stampa)
    TBBUTTON tbb[] = {
        {0, IDM_OPEN_ZIP, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Apri"},
        {0, 0, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0},
        {1, IDM_APPLY_MINISTERO, 0, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Ministero"},
        {2, IDM_APPLY_ASSOSOFTWARE, 0, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Assosoftware"},
        {0, 0, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0},
        {3, IDM_PRINT_SELECTED, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Stampa PDF"},
    };

    SendMessage(hToolbar, TB_ADDBUTTONS, sizeof(tbb) / sizeof(TBBUTTON), (LPARAM)&tbb);
    SendMessage(hToolbar, TB_AUTOSIZE, 0, 0);

    // Imposta la toolbar a tutta larghezza come una web nav bar
    RECT rcParent;
    GetClientRect(hParent, &rcParent);
    SetWindowPos(hToolbar, NULL, 0, 0, rcParent.right, TOOLBAR_HEIGHT,
        SWP_NOZORDER | SWP_NOACTIVATE);

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

// Applica tema moderno con palette Indigo brand alla finestra
void ApplyModernWindowTheme(HWND hWnd)
{
    // 1. BORDI ARROTONDATI (Windows 11)
    DWM_WINDOW_CORNER_PREFERENCE cornerPref = DWMWCP_ROUND;
    DwmSetWindowAttribute(hWnd, DWMWA_WINDOW_CORNER_PREFERENCE, &cornerPref, sizeof(cornerPref));

    // 2. TITLE BAR - Indigo brand
    COLORREF captionColor = RGB(99, 102, 241);  // Indigo 500
    DwmSetWindowAttribute(hWnd, DWMWA_CAPTION_COLOR, &captionColor, sizeof(captionColor));

    // 3. TESTO TITLE BAR BIANCO
    COLORREF textColor = RGB(255, 255, 255);
    DwmSetWindowAttribute(hWnd, DWMWA_TEXT_COLOR, &textColor, sizeof(textColor));

    // 4. BORDO - Indigo scuro
    COLORREF borderColor = RGB(79, 70, 229);
    DwmSetWindowAttribute(hWnd, DWMWA_BORDER_COLOR, &borderColor, sizeof(borderColor));

    // 5. LIGHT MODE ATTIVO
    BOOL useDarkMode = FALSE;
    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDarkMode, sizeof(useDarkMode));

    // 6. Bordi arrotondati moderni
    MARGINS margins = { 0, 0, 0, 0 };
    DwmExtendFrameIntoClientArea(hWnd, &margins);
}

// Create an icon from resources replacing a solid background color with transparency
HICON CreateIconWithTransparentBackground(HINSTANCE hInst, int resId, COLORREF bgColor)
{
    HICON hIcon = (HICON)LoadImageW(hInst, MAKEINTRESOURCEW(resId), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR);
    if (!hIcon)
        return NULL;

    ICONINFO ii;
    if (!GetIconInfo(hIcon, &ii))
        return hIcon;

    BITMAP bmpColor = {0};
    if (ii.hbmColor)
        GetObject(ii.hbmColor, sizeof(bmpColor), &bmpColor);

    HDC hdc = GetDC(NULL);

    if (ii.hbmColor)
    {
        int w = bmpColor.bmWidth;
        int h = bmpColor.bmHeight;
        BITMAPINFO bi = {0};
        bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        bi.bmiHeader.biWidth = w;
        bi.bmiHeader.biHeight = -h; // top-down
        bi.bmiHeader.biPlanes = 1;
        bi.bmiHeader.biBitCount = 32;
        bi.bmiHeader.biCompression = BI_RGB;

        void* pvBits = NULL;
        HBITMAP hbm32 = CreateDIBSection(hdc, &bi, DIB_RGB_COLORS, &pvBits, NULL, 0);
        if (hbm32 && pvBits)
        {
            HDC hMem = CreateCompatibleDC(hdc);
            HBITMAP hOld = (HBITMAP)SelectObject(hMem, hbm32);
            HDC hIconDC = CreateCompatibleDC(hdc);
            HBITMAP hOld2 = (HBITMAP)SelectObject(hIconDC, ii.hbmColor);

            BitBlt(hMem, 0, 0, w, h, hIconDC, 0, 0, SRCCOPY);

            DWORD* pixels = (DWORD*)pvBits;
            COLORREF target = bgColor & 0x00FFFFFF;
            for (int y = 0; y < h; ++y)
            {
                for (int x = 0; x < w; ++x)
                {
                    DWORD px = pixels[y * w + x];
                    COLORREF col = RGB(GetRValue(px), GetGValue(px), GetBValue(px));
                    if ((col & 0x00FFFFFF) == target)
                        pixels[y * w + x] = 0x00000000; // fully transparent
                    else
                        pixels[y * w + x] = 0xFF000000 | (px & 0x00FFFFFF);
                }
            }

            ICONINFO newii = {0};
            newii.fIcon = ii.fIcon;
            newii.xHotspot = ii.xHotspot;
            newii.yHotspot = ii.yHotspot;
            newii.hbmMask = ii.hbmMask;
            newii.hbmColor = hbm32;

            HICON hNewIcon = CreateIconIndirect(&newii);

            SelectObject(hIconDC, hOld2);
            DeleteDC(hIconDC);
            SelectObject(hMem, hOld);
            DeleteDC(hMem);
            DeleteObject(hbm32);

            ReleaseDC(NULL, hdc);
            if (ii.hbmColor) DeleteObject(ii.hbmColor);
            if (ii.hbmMask) DeleteObject(ii.hbmMask);
            DestroyIcon(hIcon);

            return hNewIcon;
        }
    }

    if (ii.hbmColor) DeleteObject(ii.hbmColor);
    if (ii.hbmMask) DeleteObject(ii.hbmMask);
    ReleaseDC(NULL, hdc);
    return hIcon;
}

INT_PTR CALLBACK LicenseDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    static HFONT hLicenseFont = NULL;
    switch (message)
    {
    case WM_INITDIALOG:
    {
        // Cerca LICENSE prima in g_appPath, poi in working dir
        std::wstring text;
        std::wstring licensePath = g_appPath + L"LICENSE";
        FILE* f = nullptr;
        _wfopen_s(&f, licensePath.c_str(), L"rt, ccs=UTF-8");
        if (!f) {
            // Prova anche working directory
            _wfopen_s(&f, L"LICENSE", L"rt, ccs=UTF-8");
        }
        if (f) {
            wchar_t buf[1024];
            while (fgetws(buf, 1024, f))
                text += buf;
            fclose(f);
        } else {
            // Versione inglese hardcoded
            text =
                L"MIT License\r\n\r\n"
                L"Copyright (c) 2026 Roberto Ferri\r\n\r\n"
                L"Permission is hereby granted, free of charge, to any person obtaining a copy\r\n"
                L"of this software and associated documentation files (the \"Software\"), to deal\r\n"
                L"in the Software without restriction, including without limitation the rights\r\n"
                L"to use, copy, modify, merge, publish, distribute, sublicense, and/or sell\r\n"
                L"copies of the Software, and to permit persons to whom the Software is\r\n"
                L"furnished to do so, subject to the following conditions:\r\n\r\n"
                L"The above copyright notice and this permission notice shall be included in all\r\n"
                L"copies or substantial portions of the Software.\r\n\r\n"
                L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR\r\n"
                L"IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,\r\n"
                L"FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE\r\n"
                L"AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER\r\n"
                L"LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,\r\n"
                L"OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE\r\n"
                L"SOFTWARE.\r\n";
        }
        // Aggiungi versione italiana (solo ASCII)
        text += L"\r\n\r\n--- Versione italiana non ufficiale ---\r\n";
        text += L"Licenza MIT\r\n\r\n";
        text += L"Copyright (c) 2026 Roberto Ferri\r\n\r\n";
        text += L"Con la presente si concede gratuitamente a chiunque ottenga una copia di questo software e dei relativi file di documentazione (il 'Software'), il permesso di trattare il Software senza restrizioni, inclusi senza limitazione i diritti di usare, copiare, modificare, unire, pubblicare, distribuire, concedere in sublicenza e/o vendere copie del Software, e di permettere alle persone a cui il Software e' fornito di farlo, alle seguenti condizioni:\r\n\r\n";
        text += L"L'avviso di copyright di cui sopra e questo avviso di autorizzazione devono essere inclusi in tutte le copie o parti sostanziali del Software.\r\n\r\n";
        text += L"IL SOFTWARE VIENE FORNITO 'COSI' COM'E'', SENZA GARANZIE DI ALCUN TIPO, ESPLICITE O IMPLICITE, INCLUSE MA NON LIMITATE ALLE GARANZIE DI COMMERCIABILITA', IDONEITA' PER UNO SCOPO PARTICOLARE E NON VIOLAZIONE. IN NESSUN CASO GLI AUTORI O I DETENTORI DEL COPYRIGHT SARANNO RESPONSABILI PER QUALSIASI RECLAMO, DANNO O ALTRA RESPONSABILITA', SIA IN UN'AZIONE DI CONTRATTO, ILLECITO O ALTRO, DERIVANTE DA, O IN CONNESSIONE CON IL SOFTWARE O L'USO O ALTRI RAPPORTI NEL SOFTWARE.\r\n";
        // Format nicer: add clear title, separators and paragraph spacing
        std::wstring formatted;
        formatted += L"MIT License\r\n";
        formatted += L"\r\n";
        formatted += L"Copyright (c) 2026 Roberto Ferri\r\n";
        // Use plain ASCII separator to avoid rendering issues with some fonts
        formatted += L"------------------------------------------------------------\r\n\r\n";

        // Ensure paragraphs are separated by an empty line for readability
        auto appendWithParagraphs = [&formatted](const std::wstring& src)
        {
            // split on \r\n and rejoin ensuring single blank line between paragraphs
            size_t pos = 0;
            size_t len = src.length();
            std::wstring paragraph;
            for (size_t i = 0; i < len; ++i)
            {
                paragraph += src[i];
                if (i + 1 < len && src[i] == L'\r' && src[i+1] == L'\n')
                {
                    // consume the \n as well
                    ++i;
                    // trim trailing spaces
                    while (!paragraph.empty() && (paragraph.back() == L'\r' || paragraph.back() == L'\n'))
                        paragraph.pop_back();
                    // add paragraph and a blank line
                    formatted += paragraph;
                    formatted += L"\r\n\r\n";
                    paragraph.clear();
                }
            }
            if (!paragraph.empty())
            {
                formatted += paragraph;
                formatted += L"\r\n\r\n";
            }
        };

        appendWithParagraphs(text);

        // Italian section header
        formatted += L"--- Versione italiana (non ufficiale) ---\r\n\r\n";
        // Build italian text from existing appended portion (we appended italian already in 'text')
        // To avoid duplication if italian already included, append nothing extra here.

        SetDlgItemTextW(hDlg, IDC_LICENSE_TEXT, formatted.c_str());
        // Migliora la formattazione della casella di testo: usa un font leggibile e rendi read-only
        if (!hLicenseFont)
        {
            HDC hdc = GetDC(hDlg);
            int logPixY = GetDeviceCaps(hdc, LOGPIXELSY);
            ReleaseDC(hDlg, hdc);

            // 10pt Segoe UI
            hLicenseFont = CreateFontW(-MulDiv(10, logPixY, 72), 0, 0, 0, FW_NORMAL,
                FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
                L"Segoe UI");
        }

        HWND hText = GetDlgItem(hDlg, IDC_LICENSE_TEXT);
        if (hText)
        {
            SendMessageW(hText, WM_SETFONT, (WPARAM)hLicenseFont, TRUE);
            // Make sure the control is read-only (if it's an edit control)
            SendMessageW(hText, EM_SETREADONLY, TRUE, 0);
            // Move caret to start and scroll to top
            SendMessageW(hText, EM_SETSEL, 0, 0);
            SendMessageW(hText, EM_SCROLLCARET, 0, 0);
        }

        return (INT_PTR)TRUE;
    }
    case WM_CTLCOLORDLG:
    case WM_CTLCOLORSTATIC:
        {
            HDC hdcCtl = (HDC)wParam;
            SetTextColor(hdcCtl, RGB(30, 41, 59));
            SetBkColor(hdcCtl, RGB(248, 250, 252));
            if (!g_hBrushDialogBg)
                g_hBrushDialogBg = CreateSolidBrush(RGB(248, 250, 252));
            return (INT_PTR)g_hBrushDialogBg;
        }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            if (hLicenseFont)
            {
                DeleteObject(hLicenseFont);
                hLicenseFont = NULL;
            }
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
