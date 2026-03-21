#!/bin/bash
# ===========================================================================
#  build-linux.sh — Compila e installa FatturaView su Linux
#  Uso: ./linux/build-linux.sh [--install] [--clean]
# ===========================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
BUILD_DIR="$SCRIPT_DIR/build"
INSTALL_PREFIX="$HOME/.local"

CLEAN=0
INSTALL=0
for arg in "$@"; do
  case $arg in
    --clean)   CLEAN=1 ;;
    --install) INSTALL=1 ;;
  esac
done

echo "==========================================="
echo "  FatturaView - Build Linux (Qt6 + CMake)"
echo "==========================================="

# ---------------------------------------------------------------------------
# 1. Verifica dipendenze
# ---------------------------------------------------------------------------
echo ""
echo "[1/4] Verifica dipendenze..."

check_dep() {
    if ! command -v "$1" &>/dev/null && ! pkg-config --exists "$2" 2>/dev/null; then
        echo "  MANCANTE: $3"
        echo "  Installa con: $4"
        MISSING=1
    else
        echo "  OK: $3"
    fi
}

MISSING=0
command -v cmake  &>/dev/null && echo "  OK: cmake"    || { echo "  MANCANTE: cmake — apt install cmake";          MISSING=1; }
command -v ninja  &>/dev/null && echo "  OK: ninja"    || echo "  (opzionale) ninja — apt install ninja-build"
pkg-config --exists Qt6Widgets   2>/dev/null && echo "  OK: Qt6 Widgets"          || { echo "  MANCANTE: Qt6 — apt install qt6-base-dev qt6-webengine-dev"; MISSING=1; }
pkg-config --exists libxml-2.0   2>/dev/null && echo "  OK: libxml2"              || { echo "  MANCANTE: libxml2  — apt install libxml2-dev";    MISSING=1; }
pkg-config --exists libxslt      2>/dev/null && echo "  OK: libxslt"              || { echo "  MANCANTE: libxslt  — apt install libxslt1-dev";   MISSING=1; }
pkg-config --exists libzip       2>/dev/null && echo "  OK: libzip"               || { echo "  MANCANTE: libzip   — apt install libzip-dev";     MISSING=1; }

if [ "$MISSING" -eq 1 ]; then
    echo ""
    echo "Installa le dipendenze mancanti e riprova."
    echo "Comando completo (Ubuntu/Debian):"
    echo "  sudo apt install cmake ninja-build qt6-base-dev qt6-webengine-dev \\"
    echo "       libxml2-dev libxslt1-dev libzip-dev"
    exit 1
fi

# ---------------------------------------------------------------------------
# 2. Pulizia (opzionale)
# ---------------------------------------------------------------------------
if [ "$CLEAN" -eq 1 ] && [ -d "$BUILD_DIR" ]; then
    echo ""
    echo "[2/4] Pulizia build precedente..."
    rm -rf "$BUILD_DIR"
fi

# ---------------------------------------------------------------------------
# 3. Configurazione CMake
# ---------------------------------------------------------------------------
echo ""
echo "[2/4] Configurazione CMake..."
mkdir -p "$BUILD_DIR"

CMAKE_GENERATOR="Unix Makefiles"
command -v ninja &>/dev/null && CMAKE_GENERATOR="Ninja"

cmake -S "$SCRIPT_DIR" -B "$BUILD_DIR" \
      -G "$CMAKE_GENERATOR" \
      -DCMAKE_BUILD_TYPE=Release \
      -DCMAKE_INSTALL_PREFIX="$INSTALL_PREFIX"

# ---------------------------------------------------------------------------
# 4. Compilazione
# ---------------------------------------------------------------------------
echo ""
echo "[3/4] Compilazione ($(nproc) core)..."
cmake --build "$BUILD_DIR" --config Release --parallel "$(nproc)"

echo ""
echo "[4/4] Build completata!"
echo "  Eseguibile: $BUILD_DIR/fatturaview"

# ---------------------------------------------------------------------------
# 5. Installazione (opzionale)
# ---------------------------------------------------------------------------
if [ "$INSTALL" -eq 1 ]; then
    echo ""
    echo "Installazione in $INSTALL_PREFIX ..."
    cmake --install "$BUILD_DIR"
    echo "  Installato: $INSTALL_PREFIX/bin/fatturaview"
    echo "  Risorse:    $INSTALL_PREFIX/share/fatturaview/"
fi

echo ""
echo "==========================================="
echo "  Completato con successo!"
echo "==========================================="
echo ""
echo "Per eseguire:  $BUILD_DIR/fatturaview"
if [ "$INSTALL" -eq 1 ]; then
echo "Oppure:        fatturaview   (se ~/.local/bin è nel PATH)"
fi
