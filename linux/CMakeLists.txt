cmake_minimum_required(VERSION 3.20)
project(FatturaView VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

# ---------------------------------------------------------------------------
# Dipendenze Qt6
# ---------------------------------------------------------------------------
find_package(Qt6 REQUIRED COMPONENTS Core Widgets WebEngineWidgets PrintSupport)

# ---------------------------------------------------------------------------
# Dipendenze sistema — usiamo pkg-config per avere variabili uniformi *_INCLUDE_DIRS
# ---------------------------------------------------------------------------
find_package(PkgConfig REQUIRED)
pkg_check_modules(LIBXML2  REQUIRED libxml-2.0)
pkg_check_modules(LIBXSLT  REQUIRED libxslt)
pkg_check_modules(LIBZIP   REQUIRED libzip)

# ---------------------------------------------------------------------------
# Sorgenti
# ---------------------------------------------------------------------------
set(SOURCES
    src/main.cpp
    src/MainWindow.cpp
    src/FatturaParser.cpp
    src/PdfSigner.cpp
)

set(HEADERS
    src/MainWindow.h
    src/FatturaParser.h
    src/PdfSigner.h
)

add_executable(fatturaview ${SOURCES} ${HEADERS})

target_include_directories(fatturaview PRIVATE
    src/
    ${LIBXML2_INCLUDE_DIRS}
    ${LIBXSLT_INCLUDE_DIRS}
    ${LIBZIP_INCLUDE_DIRS}
)

target_link_libraries(fatturaview PRIVATE
    Qt6::Core
    Qt6::Widgets
    Qt6::WebEngineWidgets
    Qt6::PrintSupport
    ${LIBXML2_LIBRARIES}
    ${LIBXSLT_LIBRARIES}
    ${LIBZIP_LIBRARIES}
)

target_compile_options(fatturaview PRIVATE
    ${LIBXML2_CFLAGS_OTHER}
    ${LIBXSLT_CFLAGS_OTHER}
)

# ---------------------------------------------------------------------------
# Installazione
# ---------------------------------------------------------------------------
include(GNUInstallDirs)

install(TARGETS fatturaview DESTINATION ${CMAKE_INSTALL_BINDIR})

install(DIRECTORY ${CMAKE_SOURCE_DIR}/../Resources/
        DESTINATION ${CMAKE_INSTALL_DATADIR}/fatturaview/Resources)

install(FILES ${CMAKE_SOURCE_DIR}/fatturaview.desktop
        DESTINATION ${CMAKE_INSTALL_DATADIR}/applications)

# ---------------------------------------------------------------------------
# .desktop file per il menu di sistema (generato a configure-time)
# ---------------------------------------------------------------------------
file(WRITE ${CMAKE_CURRENT_BINARY_DIR}/fatturaview.desktop
"[Desktop Entry]
Type=Application
Name=FatturaView
GenericName=Visualizzatore Fatture Elettroniche
Comment=Lettore di FatturaPA XML con fogli di stile
Exec=fatturaview %f
Icon=fatturaview
Categories=Office;Finance;
MimeType=application/xml;application/zip;
Terminal=false
")
