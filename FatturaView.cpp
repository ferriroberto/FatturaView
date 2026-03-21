// xmlRead.cpp : Definisce il punto di ingresso dell'applicazione.
//

#include "framework.h"
#include "FatturaView.h"
#include "FatturaElettronicaParser.h"
#include "FatturaViewer.h"
#include "PdfSigner.h"
#include <commdlg.h>
#include <shellapi.h>
#include <vector>
#include <string>
#include <algorithm>
#include <commctrl.h>
#include <windowsx.h>  // Per GET_X_LPARAM e GET_Y_LPARAM

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
HFONT g_hFontNormal = NULL;                     // Font normale per ListBox
HFONT g_hFontBold = NULL;                       // Font grassetto per intestazioni

#define APP_VERSION L"1.0.0"
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
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    StylesheetSelector(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    PdfSignConfigDialog(HWND, UINT, WPARAM, LPARAM);
void                OnOpenFile(HWND hWnd);
void                OnApplyStylesheet(HWND hWnd);
void                OnApplySpecificStylesheet(HWND hWnd, const std::wstring& keyword);
void                OnApplyStylesheetMultiple(HWND hWnd, const std::wstring& xsltPath, const std::vector<int>& indices);
std::vector<int>    GetSelectedFattureIndices();
void                UpdateNavButtons();
void                OnZoom(HWND hWnd, bool zoomIn);
void                OnPrintToPDF(HWND hWnd);
void                LoadFattureList(HWND hWnd);
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
    wcex.hbrBackground  = CreateSolidBrush(RGB(250, 250, 250)); // Sfondo chiaro
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
   int margin = 20;  // Margine uniforme per design allineato
   int topOffset = TOOLBAR_HEIGHT + margin;

   // Crea una ListBox per mostrare le fatture (pannello sinistro) con selezione multipla
   // Usa LBS_OWNERDRAWFIXED per disegnare manualmente gli elementi con colori
   g_hListBox = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", NULL,
      WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY | LBS_EXTENDEDSEL | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
      margin, topOffset,  // Allineato con la toolbar (20px dal bordo)
      leftPanelWidth, 
      rcClient.bottom - topOffset - margin - statusHeight,
      hWnd, (HMENU)IDC_LISTBOX_FATTURE, hInstance, NULL);

   // Imposta l'altezza degli elementi (più alta per stile moderno card-like)
   SendMessage(g_hListBox, LB_SETITEMHEIGHT, 0, (LPARAM)32);

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
      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_DISABLED,
      navBaseX, navY, 125, 30,
      hWnd, (HMENU)IDM_PREV_FATTURA, hInstance, NULL);
   g_hBtnNext = CreateWindowExW(0, L"BUTTON", L"Fattura Succ \u25BA",
      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_DISABLED,
      navBaseX + 135, navY, 125, 30,
      hWnd, (HMENU)IDM_NEXT_FATTURA, hInstance, NULL);
   g_hBtnZoomOut = CreateWindowExW(0, L"BUTTON", L"\u2212",
      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
      navBaseX + 275, navY, 28, 30,
      hWnd, (HMENU)IDM_ZOOM_OUT, hInstance, NULL);
   g_hLabelZoom = CreateWindowExW(WS_EX_CLIENTEDGE, L"STATIC", L"100%",
      WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
      navBaseX + 308, navY, 55, 30,
      hWnd, (HMENU)IDC_ZOOM_LABEL, hInstance, NULL);
   g_hBtnZoomIn = CreateWindowExW(0, L"BUTTON", L"+",
      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
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
       SetTimer(hWnd, 1, 500, NULL); // Timer ID 1, 500ms delay (aumentato)
   }

   // Applica tema moderno Windows 11
   ApplyModernWindowTheme(hWnd);

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
                    clrBackground = RGB(245, 245, 245);
                    clrText = RGB(0, 0, 0);
                    hFont = g_hFontNormal;
                }
                else if (isHeader)
                {
                    // Intestazione emittente - Blu Windows 11
                    if (pDIS->itemState & ODS_SELECTED)
                    {
                        clrBackground = RGB(0, 90, 158);  // Blu scuro se selezionato
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
                    // Fattura normale - Effetto card
                    if (pDIS->itemState & ODS_SELECTED)
                    {
                        // Selezionato - Azzurro chiaro
                        clrBackground = RGB(204, 230, 255);
                        clrText = RGB(0, 0, 0);
                    }
                    else
                    {
                        // Alternanza bianco/grigio chiaro
                        clrBackground = (pDIS->itemID % 2 == 0) ? RGB(255, 255, 255) : RGB(248, 251, 255);
                        clrText = RGB(32, 32, 32);
                    }
                    hFont = g_hFontNormal;
                }

                // Disegna lo sfondo con gradiente per le intestazioni
                if (isHeader && !(pDIS->itemState & ODS_SELECTED))
                {
                    // Crea un gradiente orizzontale blu per le intestazioni
                    TRIVERTEX vertex[2];
                    vertex[0].x = pDIS->rcItem.left;
                    vertex[0].y = pDIS->rcItem.top;
                    vertex[0].Red = 0x0000;      // RGB(0, 120, 215)
                    vertex[0].Green = 0x7800;
                    vertex[0].Blue = 0xD700;
                    vertex[0].Alpha = 0x0000;

                    vertex[1].x = pDIS->rcItem.right;
                    vertex[1].y = pDIS->rcItem.bottom;
                    vertex[1].Red = 0x0000;      // RGB(0, 90, 158)
                    vertex[1].Green = 0x5A00;
                    vertex[1].Blue = 0x9E00;
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
                        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(230, 230, 230));
                        HPEN hOldPen = (HPEN)SelectObject(pDIS->hDC, hPen);
                        MoveToEx(pDIS->hDC, pDIS->rcItem.left + 8, pDIS->rcItem.bottom - 1, NULL);
                        LineTo(pDIS->hDC, pDIS->rcItem.right - 8, pDIS->rcItem.bottom - 1);
                        SelectObject(pDIS->hDC, hOldPen);
                        DeleteObject(hPen);
                    }
                }

                // Disegna bordo focus con colore blu
                if (pDIS->itemState & ODS_FOCUS)
                {
                    RECT rcFocus = pDIS->rcItem;
                    InflateRect(&rcFocus, -2, -2);
                    HPEN hPenFocus = CreatePen(PS_SOLID, 2, RGB(0, 120, 215));
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

                // Mostra notifica popup
                if (g_hBrowserWnd)
                    FatturaViewer::ShowNotification(g_hBrowserWnd, L"Caricamento interfaccia...", L"info");

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

                // Mostra notifica di benvenuto
                if (g_hBrowserWnd)
                    FatturaViewer::ShowNotification(g_hBrowserWnd, L"Interfaccia caricata - Pronto all'uso", L"success");
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
            int margin = 20;  // Margine uniforme allineato con toolbar
            int topOffset = TOOLBAR_HEIGHT + margin;

            // Ridimensiona la toolbar con padding sinistro
            if (g_hToolbar)
            {
                SendMessage(g_hToolbar, TB_AUTOSIZE, 0, 0);

                // Applica padding sinistro allineato
                RECT rcToolbar;
                GetWindowRect(g_hToolbar, &rcToolbar);
                int toolbarHeight = rcToolbar.bottom - rcToolbar.top;
                SetWindowPos(g_hToolbar, NULL, margin, 0, 0, toolbarHeight, 
                    SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
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

void OnOpenFile(HWND hWnd)
{
    std::vector<std::wstring> paths = OpenMultipleFilesDialog(hWnd,
        L"File Fatture (ZIP, XML)\0*.zip;*.xml;*.XML\0File ZIP\0*.zip\0File XML Fattura\0*.xml;*.XML\0Tutti i file\0*.*\0",
        L"Apri file fattura (Ctrl+Click per selezione multipla XML)");

    if (paths.empty())
        return;

    if (!g_pParser)
        return;

    // Valida anticipatamente tutti i file XML (non ZIP): se anche uno solo non e' FatturaPA blocca tutto
    {
        std::vector<std::wstring> invalidFiles;
        for (const auto& path : paths)
        {
            size_t dotPos = path.find_last_of(L'.');
            bool isZip = (dotPos != std::wstring::npos &&
                _wcsicmp(path.c_str() + dotPos, L".zip") == 0);
            if (!isZip && !g_pParser->IsFatturaPA(path))
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

    // Reset completo: svuota il browser, la lista e lo stato di navigazione
    g_currentXmlPath.clear();
    g_multipleHtmlFiles.clear();
    g_currentHtmlIndex = -1;
    UpdateNavButtons();
    if (g_hBrowserWnd)
        FatturaViewer::NavigateToString(g_hBrowserWnd, FatturaViewer::GetWelcomePageHTML());
    if (g_hListBox)
        SendMessage(g_hListBox, LB_RESETCONTENT, 0, 0);

    WCHAR titleBuffer[256];
    GetWindowTextW(hWnd, titleBuffer, 256);
    std::wstring oldTitle = titleBuffer;
    SetWindowTextW(hWnd, L"Caricamento in corso - Attendere...");
    UpdateWindow(hWnd);

    // Svuota completamente la cartella temporanea (file e sottocartelle di sessioni precedenti)
    ClearFolderContents(g_extractPath);
    CreateDirectoryW(g_extractPath.c_str(), NULL); // ricrea la cartella radice

    int successCount = 0;
    int errorCount = 0;

    for (const auto& path : paths)
    {
        size_t dotPos = path.find_last_of(L'.');
        bool isZip = (dotPos != std::wstring::npos &&
            _wcsicmp(path.c_str() + dotPos, L".zip") == 0);

        if (isZip)
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
        else
        {
            if (g_hStatusBar)
            {
                SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Caricamento XML...");
                UpdateWindow(g_hStatusBar);
            }
            if (g_hBrowserWnd)
                FatturaViewer::ShowNotification(g_hBrowserWnd, L"Caricamento file XML...", L"info");
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

    if (successCount == 0)
    {
        if (g_hStatusBar)
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Errore caricamento file");
        if (g_hBrowserWnd)
            FatturaViewer::ShowNotification(g_hBrowserWnd, L"Errore nel caricamento dei file", L"error");
        MessageBoxW(hWnd, L"Errore nel caricamento dei file selezionati!", L"Errore", MB_OK | MB_ICONERROR);
        return;
    }

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

    // Abilita/Disabilita i pulsanti Ministero e Assosoftware in base alle fatture caricate
    if (g_hToolbar)
    {
        BOOL hasFatture = !g_fattureInfo.empty() ? TRUE : FALSE;
        SendMessage(g_hToolbar, TB_ENABLEBUTTON, IDM_APPLY_MINISTERO, MAKELONG(hasFatture, 0));
        SendMessage(g_hToolbar, TB_ENABLEBUTTON, IDM_APPLY_ASSOSOFTWARE, MAKELONG(hasFatture, 0));
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
            SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Fattura visualizzata");
        if (g_hBrowserWnd)
            FatturaViewer::ShowNotification(g_hBrowserWnd, L"Fattura visualizzata correttamente", L"success");
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
    SendMessage(hToolbar, TB_SETBUTTONSIZE, 0, MAKELONG(48, 48));  // Pulsanti più grandi

    // Usa icone standard di Windows (versione large per aspetto moderno)
    TBADDBITMAP tbab;
    tbab.hInst = HINST_COMMCTRL;
    tbab.nID = IDB_STD_LARGE_COLOR;
    SendMessage(hToolbar, TB_ADDBITMAP, 0, (LPARAM)&tbab);

    // Definisci i pulsanti della toolbar con testo moderno
    TBBUTTON tbb[] = {
        {STD_FILEOPEN, IDM_OPEN_ZIP, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Apri"},
        {0, 0, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0},
        {STD_PROPERTIES, IDM_APPLY_MINISTERO, 0, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Ministero"},
        {STD_PROPERTIES, IDM_APPLY_ASSOSOFTWARE, 0, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Assosoftware"},
        {0, 0, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0},
        {STD_PRINT, IDM_PRINT_SELECTED, TBSTATE_ENABLED, BTNS_AUTOSIZE | BTNS_SHOWTEXT, {0}, 0, (INT_PTR)L"Stampa PDF"},
    };

    SendMessage(hToolbar, TB_ADDBUTTONS, sizeof(tbb) / sizeof(TBBUTTON), (LPARAM)&tbb);
    SendMessage(hToolbar, TB_AUTOSIZE, 0, 0);

    // Aggiungi padding sinistro per centrare meglio la toolbar
    RECT rcToolbar;
    GetWindowRect(hToolbar, &rcToolbar);
    int toolbarWidth = rcToolbar.right - rcToolbar.left;
    int leftPadding = 20; // Sposta 20 pixel a destra

    // Riposiziona la toolbar con offset sinistro
    SetWindowPos(hToolbar, NULL, leftPadding, 0, toolbarWidth, 0, 
        SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);

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

// Applica tema moderno Light Mode alla finestra
void ApplyModernWindowTheme(HWND hWnd)
{
    // 1. BORDI ARROTONDATI (Windows 11)
    DWM_WINDOW_CORNER_PREFERENCE cornerPref = DWMWCP_ROUND;
    DwmSetWindowAttribute(hWnd, DWMWA_WINDOW_CORNER_PREFERENCE, &cornerPref, sizeof(cornerPref));

    // 2. TITLE BAR - Blu Windows 11
    COLORREF captionColor = RGB(0, 120, 215);  // Blu Accent
    DwmSetWindowAttribute(hWnd, DWMWA_CAPTION_COLOR, &captionColor, sizeof(captionColor));

    // 3. TESTO TITLE BAR BIANCO
    COLORREF textColor = RGB(255, 255, 255);
    DwmSetWindowAttribute(hWnd, DWMWA_TEXT_COLOR, &textColor, sizeof(textColor));

    // 4. BORDO - Blu scuro
    COLORREF borderColor = RGB(0, 90, 158);
    DwmSetWindowAttribute(hWnd, DWMWA_BORDER_COLOR, &borderColor, sizeof(borderColor));

    // 5. LIGHT MODE ATTIVO
    BOOL useDarkMode = FALSE;
    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDarkMode, sizeof(useDarkMode));

    // 6. Bordi arrotondati moderni
    MARGINS margins = { 0, 0, 0, 0 };
    DwmExtendFrameIntoClientArea(hWnd, &margins);
}
