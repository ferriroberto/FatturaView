#pragma once
#include <QMainWindow>
#include <QListWidget>
#include <QWebEngineView>
#include <QPushButton>
#include <QLabel>
#include <QSplitter>
#include <QSettings>

#include "FatturaParser.h"
#include "PdfSigner.h"

// ---------------------------------------------------------------------------
// MainWindow — porta Qt6 della finestra Win32 principale
// Layout:  QToolBar | QSplitter(QListWidget + QWebEngineView) | nav bar | QStatusBar
// ---------------------------------------------------------------------------
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

private slots:
    void onOpenFile();
    void onFatturaSelectionChanged();
    void onApplyStylesheetMinistero();
    void onApplyStylesheetAsso();
    void onSelectStylesheet();
    void onPrevFattura();
    void onNextFattura();
    void onZoomIn();
    void onZoomOut();
    void onPrintToPdf();
    void onAbout();
    void onPdfSignConfig();
    void onContextMenu(const QPoint& pos);

private:
    // Setup UI
    void setupMenuBar();
    void setupToolBar();
    void setupCentralWidget();
    void setupNavBar(QWidget* parent, QVBoxLayout* layout);
    void setupStatusBar();

    // Logica
    void loadWelcomePage();
    void loadFattureList();
    void applyStylesheetToIndex(int fatturaIndex, const QString& xsltPath);
    void applyStylesheetMultiple(const QList<int>& indices, const QString& xsltPath);
    void updateNavButtons();
    QString getResourcesPath() const;
    QString findStylesheet(const QString& keyword) const;
    QList<int> getSelectedFatturaIndices() const;
    void saveStylesheetPreference(const QString& path);
    QString loadStylesheetPreference() const;

    // Rendering lista con stile Windows 11
    void populateListWidget();
    void styleListItem(QListWidgetItem* item, bool isHeader) const;

    // Widget
    QSplitter*      m_splitter    = nullptr;
    QListWidget*    m_listWidget  = nullptr;
    QWebEngineView* m_webView     = nullptr;
    QPushButton*    m_btnPrev     = nullptr;
    QPushButton*    m_btnNext     = nullptr;
    QPushButton*    m_btnZoomIn   = nullptr;
    QPushButton*    m_btnZoomOut  = nullptr;
    QLabel*         m_labelZoom   = nullptr;

    // Dati
    FatturaParser*           m_parser           = nullptr;
    std::vector<FatturaInfo> m_fattureInfo;
    std::vector<int>         m_listToFatturaIdx; // -1 = intestazione, -2 = separatore
    QString                  m_extractPath;
    QString                  m_lastXsltPath;
    QStringList              m_multipleHtmlFiles;
    int                      m_currentHtmlIndex = -1;
    int                      m_zoomLevel        = 100;

    static constexpr int LEFT_PANEL_WIDTH = 550;
};
