// header.h: file di inclusione per file di inclusione di sistema standard
// o file di inclusione specifici del progetto
//

#pragma once

#include "targetver.h"
#define WIN32_LEAN_AND_MEAN             // Escludere gli elementi usati raramente dalle intestazioni di Windows
// File di intestazione di Windows
#include <windows.h>
#include <dwmapi.h>                     // Per effetti DWM Windows 11
// File di intestazione Runtime C
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>

#pragma comment(lib, "dwmapi.lib")      // Linker per DWM
