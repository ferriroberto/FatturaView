# FatturaView — Linux (Qt6)

Versione Linux di **FatturaView** scritta con **Qt 6 + CMake**.  
Porta completa dell'applicazione Windows (Win32/COM/MSXML) usando solo librerie open source.

---

## Dipendenze

| Libreria | Sostituisce | Pacchetto Ubuntu/Debian |
|---|---|---|
| **Qt6 Widgets** | Win32 API (finestra, toolbar, listbox) | `qt6-base-dev` |
| **Qt6 WebEngineWidgets** | COM `IWebBrowser2` (IE) | `qt6-webengine-dev` |
| **Qt6 PrintSupport** | Win32 Print API | incluso in `qt6-base-dev` |
| **libxml2** | MSXML `DOMDocument60` | `libxml2-dev` |
| **libxslt** | MSXML `transformNode` | `libxslt1-dev` |
| **libzip** | Shell32 ZIP extraction | `libzip-dev` |

### Installazione dipendenze (Ubuntu 22.04+)

```bash
sudo apt install cmake ninja-build \
     qt6-base-dev qt6-webengine-dev \
     libxml2-dev libxslt1-dev libzip-dev
```

---

## Build

```bash
# Clona e vai nella cartella Linux
cd FatturaView/linux

# Build automatica (script)
chmod +x build-linux.sh
./build-linux.sh

# Oppure manuale
mkdir build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
cmake --build . --parallel
```

### Installazione in ~/.local

```bash
./build-linux.sh --install
# Eseguibile disponibile in ~/.local/bin/fatturaview
```

---

## Struttura

```
linux/
├── CMakeLists.txt          ← build system
├── build-linux.sh          ← script build automatica
├── README.md
└── src/
    ├── main.cpp            ← QApplication entry point
    ├── MainWindow.h/.cpp   ← QMainWindow (porta Win32 WndProc)
    ├── FatturaParser.h/.cpp← libxml2+libxslt+libzip (porta MSXML/COM)
    └── PdfSigner.h/.cpp    ← config cross-platform (porta Shell+Registry)
```

---

## Funzionalità portate

| Feature Windows | Equivalente Linux |
|---|---|
| Win32 ListBox owner-draw | `QListWidget` con stile CSS Qt |
| `IWebBrowser2` (IE engine) | `QWebEngineView` (Chromium) |
| MSXML `DOMDocument60` | `libxml2` + `xmlXPathEvalExpression` |
| MSXML `transformNode` | `libxslt` + `xsltApplyStylesheet` |
| Shell32 ZIP extraction | `libzip` |
| Registry `HKCU\Software\...` | `QSettings` (`~/.config/FatturaView/`) |
| `SHGetFolderPathW(APPDATA)` | `QStandardPaths::AppConfigLocation` |
| `SHGetFolderPathW(TEMP)` | `QStandardPaths::TempLocation` |
| Firma PDF (PowerShell+iTextSharp) | `pyhanko` (`pip3 install pyhanko`) |

---

## Note runtime

- Il foglio di stile preferito viene salvato in `~/.config/FatturaView/FatturaView.conf`
- La configurazione firma PDF è in `~/.config/FatturaView/pdfsign.cfg`
- I file temporanei estratti finiscono in `/tmp/FatturaView/`
