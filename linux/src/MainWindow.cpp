#include "MainWindow.h"

#include <QApplication>
#include <QMenuBar>
#include <QToolBar>
#include <QStatusBar>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QFileDialog>
#include <QMessageBox>
#include <QListWidgetItem>
#include <QStandardPaths>
#include <QDir>
#include <QTemporaryDir>
#include <QSplitter>
#include <QWebEngineView>
#include <QWebEnginePage>
#include <QPrintDialog>
#include <QPrinter>
#include <QDialog>
#include <QFormLayout>
#include <QLineEdit>
#include <QCheckBox>
#include <QDialogButtonBox>
#include <QLabel>
#include <QGroupBox>
#include <QListView>
#include <QStringListModel>
#include <QFont>
#include <QFontMetrics>
#include <QPalette>

#include <filesystem>

namespace fs = std::filesystem;

// ---------------------------------------------------------------------------
MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
    , m_parser(new FatturaParser())
{
    setWindowTitle("FatturaView - Visualizzatore Fatture Elettroniche");
    resize(1200, 700);
    showMaximized();

    // Cartella temporanea per l'estrazione
    m_extractPath = QStandardPaths::writableLocation(QStandardPaths::TempLocation)
                    + "/FatturaView/";
    QDir().mkpath(m_extractPath);

    m_lastXsltPath = loadStylesheetPreference();

    setupMenuBar();
    setupToolBar();
    setupCentralWidget();
    setupStatusBar();

    loadWelcomePage();
}

MainWindow::~MainWindow()
{
    delete m_parser;
}

// ---------------------------------------------------------------------------
// MENU BAR
// ---------------------------------------------------------------------------
void MainWindow::setupMenuBar()
{
    QMenu* fileMenu = menuBar()->addMenu("&File");
    auto* actOpen = fileMenu->addAction(QIcon::fromTheme("document-open"), "Apri ZIP / XML...",
                                        this, &MainWindow::onOpenFile);
    actOpen->setShortcut(QKeySequence::Open);
    fileMenu->addSeparator();
    auto* actExit = fileMenu->addAction("E&sci", this, &QMainWindow::close);
    actExit->setShortcut(QKeySequence::Quit);

    QMenu* viewMenu = menuBar()->addMenu("&Visualizza");
    viewMenu->addAction("Foglio &Ministero",    this, &MainWindow::onApplyStylesheetMinistero);
    viewMenu->addAction("Foglio &Assosoftware", this, &MainWindow::onApplyStylesheetAsso);
    viewMenu->addSeparator();
    viewMenu->addAction("&Seleziona foglio di stile...", this, &MainWindow::onSelectStylesheet);
    viewMenu->addSeparator();
    auto* actZoomIn  = viewMenu->addAction("Zoom +", this, &MainWindow::onZoomIn);
    auto* actZoomOut = viewMenu->addAction("Zoom -", this, &MainWindow::onZoomOut);
    actZoomIn->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_Plus));
    actZoomOut->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_Minus));

    QMenu* toolsMenu = menuBar()->addMenu("&Strumenti");
    toolsMenu->addAction("Stampa / Salva PDF...", this, &MainWindow::onPrintToPdf);
    toolsMenu->addSeparator();
    toolsMenu->addAction("Configurazione firma PDF...", this, &MainWindow::onPdfSignConfig);

    QMenu* helpMenu = menuBar()->addMenu("&?");
    helpMenu->addAction("Informazioni su FatturaView...", this, &MainWindow::onAbout);
}

// ---------------------------------------------------------------------------
// TOOLBAR
// ---------------------------------------------------------------------------
void MainWindow::setupToolBar()
{
    QToolBar* tb = addToolBar("Principale");
    tb->setIconSize(QSize(32, 32));
    tb->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);

    tb->addAction(QIcon::fromTheme("document-open"), "Apri",
                  this, &MainWindow::onOpenFile);
    tb->addSeparator();
    tb->addAction(QIcon::fromTheme("format-justify-fill"), "Ministero",
                  this, &MainWindow::onApplyStylesheetMinistero);
    tb->addAction(QIcon::fromTheme("text-x-generic"), "Assosoftware",
                  this, &MainWindow::onApplyStylesheetAsso);
    tb->addAction(QIcon::fromTheme("preferences-other"), "Foglio...",
                  this, &MainWindow::onSelectStylesheet);
    tb->addSeparator();
    tb->addAction(QIcon::fromTheme("document-print"), "Stampa PDF",
                  this, &MainWindow::onPrintToPdf);
}

// ---------------------------------------------------------------------------
// WIDGET CENTRALE  (splitter: lista sinistra | browser destra)
// ---------------------------------------------------------------------------
void MainWindow::setupCentralWidget()
{
    QWidget*     central = new QWidget(this);
    QVBoxLayout* vlay    = new QVBoxLayout(central);
    vlay->setContentsMargins(12, 12, 12, 8);
    vlay->setSpacing(6);

    // Splitter orizzontale
    m_splitter = new QSplitter(Qt::Horizontal, central);

    // — Pannello sinistro: lista fatture —
    m_listWidget = new QListWidget(m_splitter);
    m_listWidget->setSelectionMode(QAbstractItemView::ExtendedSelection);
    m_listWidget->setContextMenuPolicy(Qt::CustomContextMenu);
    m_listWidget->setAlternatingRowColors(false);
    m_listWidget->setUniformItemSizes(false);
    m_listWidget->setFont(QFont("Segoe UI", 10));
    m_listWidget->setMinimumWidth(LEFT_PANEL_WIDTH);
    m_listWidget->setMaximumWidth(LEFT_PANEL_WIDTH);

    // Stile lista (simile a Windows 11)
    m_listWidget->setStyleSheet(
        "QListWidget { border: 1px solid #d0d0d0; border-radius: 6px; background: #ffffff; }"
        "QListWidget::item { padding: 6px 10px; border-bottom: 1px solid #f0f0f0; }"
        "QListWidget::item:selected { background: #cce5ff; color: #000000; }"
    );

    connect(m_listWidget, &QListWidget::itemSelectionChanged,
            this, &MainWindow::onFatturaSelectionChanged);
    connect(m_listWidget, &QListWidget::customContextMenuRequested,
            this, &MainWindow::onContextMenu);

    // — Pannello destro: browser HTML —
    QWidget*     rightPane = new QWidget(m_splitter);
    QVBoxLayout* rightLay  = new QVBoxLayout(rightPane);
    rightLay->setContentsMargins(0, 0, 0, 0);
    rightLay->setSpacing(4);

    m_webView = new QWebEngineView(rightPane);
    m_webView->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    rightLay->addWidget(m_webView, 1);

    // Barra navigazione + zoom sotto il browser
    setupNavBar(rightPane, rightLay);

    m_splitter->addWidget(m_listWidget);
    m_splitter->addWidget(rightPane);
    m_splitter->setStretchFactor(0, 0);
    m_splitter->setStretchFactor(1, 1);

    vlay->addWidget(m_splitter, 1);
    setCentralWidget(central);
}

// ---------------------------------------------------------------------------
// BARRA NAVIGAZIONE + ZOOM
// ---------------------------------------------------------------------------
void MainWindow::setupNavBar(QWidget* parent, QVBoxLayout* layout)
{
    QHBoxLayout* navLay = new QHBoxLayout();
    navLay->setContentsMargins(4, 2, 4, 2);

    m_btnPrev = new QPushButton("◄ Fattura Prec", parent);
    m_btnNext = new QPushButton("Fattura Succ ►", parent);
    m_btnPrev->setEnabled(false);
    m_btnNext->setEnabled(false);
    m_btnPrev->setFixedSize(130, 30);
    m_btnNext->setFixedSize(130, 30);

    m_btnZoomOut  = new QPushButton("−", parent);
    m_labelZoom   = new QLabel("100%", parent);
    m_btnZoomIn   = new QPushButton("+", parent);
    m_btnZoomOut->setFixedSize(28, 30);
    m_btnZoomIn->setFixedSize(28, 30);
    m_labelZoom->setFixedSize(55, 30);
    m_labelZoom->setAlignment(Qt::AlignCenter);
    m_labelZoom->setStyleSheet("border: 1px solid #c0c0c0; border-radius: 3px;");

    navLay->addStretch();
    navLay->addWidget(m_btnPrev);
    navLay->addSpacing(10);
    navLay->addWidget(m_btnNext);
    navLay->addSpacing(20);
    navLay->addWidget(m_btnZoomOut);
    navLay->addWidget(m_labelZoom);
    navLay->addWidget(m_btnZoomIn);
    navLay->addStretch();

    layout->addLayout(navLay);

    connect(m_btnPrev,    &QPushButton::clicked, this, &MainWindow::onPrevFattura);
    connect(m_btnNext,    &QPushButton::clicked, this, &MainWindow::onNextFattura);
    connect(m_btnZoomIn,  &QPushButton::clicked, this, &MainWindow::onZoomIn);
    connect(m_btnZoomOut, &QPushButton::clicked, this, &MainWindow::onZoomOut);
}

// ---------------------------------------------------------------------------
// STATUS BAR
// ---------------------------------------------------------------------------
void MainWindow::setupStatusBar()
{
    QStatusBar* sb = statusBar();
    sb->showMessage("Pronto");

    QLabel* lAuthor  = new QLabel("Autore: Roberto Ferri");
    QLabel* lVersion = new QLabel("Versione 1.0.0");
    QLabel* lSite    = new QLabel("<a href='https://www.fatturaview.it'>www.fatturaview.it</a>");
    lSite->setOpenExternalLinks(true);
    sb->addPermanentWidget(lAuthor);
    sb->addPermanentWidget(lVersion);
    sb->addPermanentWidget(lSite);
}

// ---------------------------------------------------------------------------
// WELCOME PAGE
// ---------------------------------------------------------------------------
void MainWindow::loadWelcomePage()
{
    const QString html = R"HTML(
<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body { font-family: 'Segoe UI', sans-serif; background: #f3f4f6;
         display: flex; align-items: center; justify-content: center;
         min-height: 100vh; margin: 0; }
  .card { background: white; border-radius: 16px; padding: 48px 56px;
          box-shadow: 0 8px 32px rgba(0,0,0,0.12); max-width: 560px;
          text-align: center; }
  h1 { color: #0078d4; font-size: 2rem; margin-bottom: 0.25em; }
  .subtitle { color: #555; font-size: 1.05rem; margin-bottom: 1.5em; }
  .badge { display: inline-block; background: #e7f3ff; color: #0078d4;
           border-radius: 20px; padding: 4px 14px; font-size: 0.85rem;
           margin: 4px; }
  .hint { color: #888; font-size: 0.9rem; margin-top: 2em; }
  .step { background: #f8f9fa; border-left: 3px solid #0078d4;
          padding: 10px 16px; margin: 8px 0; border-radius: 4px;
          text-align: left; font-size: 0.92rem; }
</style></head><body>
<div class="card">
  <h1>📄 FatturaView</h1>
  <p class="subtitle">Visualizzatore Fatture Elettroniche (FatturaPA)</p>
  <span class="badge">XML FatturaPA</span>
  <span class="badge">Archivi ZIP</span>
  <span class="badge">XSLT Fogli di stile</span>
  <div style="margin-top:1.5em">
    <div class="step">1. Clicca <strong>Apri</strong> per caricare un file ZIP o XML</div>
    <div class="step">2. Seleziona una fattura dalla lista a sinistra</div>
    <div class="step">3. Usa <strong>Visualizza → Foglio Ministero/Assosoftware</strong></div>
  </div>
  <p class="hint">Versione 1.0.0 — Roberto Ferri — <a href="https://www.fatturaview.it" style="color:#0078d4;text-decoration:none">www.fatturaview.it</a></p>
</div>
</body></html>
)HTML";

    m_webView->setHtml(html);
}

// ---------------------------------------------------------------------------
// APERTURA FILE (ZIP / XML)
// ---------------------------------------------------------------------------
void MainWindow::onOpenFile()
{
    const QStringList files = QFileDialog::getOpenFileNames(
        this, "Apri Fatture",
        QStandardPaths::writableLocation(QStandardPaths::HomeLocation),
        "File Fatture (*.zip *.xml);;File ZIP (*.zip);;File XML Fattura (*.xml);;Tutti i file (*)");

    if (files.isEmpty()) return;

    statusBar()->showMessage("Caricamento fatture...");

    // Svuota la cartella di estrazione
    QDir extractDir(m_extractPath);
    for (const QString& f : extractDir.entryList(QDir::Files))
        extractDir.remove(f);

    m_fattureInfo.clear();
    m_listToFatturaIdx.clear();
    m_multipleHtmlFiles.clear();
    m_currentHtmlIndex = -1;

    for (const QString& filePath : files)
    {
        if (filePath.endsWith(".zip", Qt::CaseInsensitive))
        {
            m_parser->ExtractZipFile(filePath.toStdString(),
                                     m_extractPath.toStdString());
        }
        else if (filePath.endsWith(".xml", Qt::CaseInsensitive))
        {
            // Copia il singolo XML nella cartella di estrazione
            QFile::copy(filePath, m_extractPath + QFileInfo(filePath).fileName());
        }
    }

    // Carica le informazioni delle fatture estratte
    auto infos = m_parser->GetFattureInfoFromFolder(m_extractPath.toStdString());
    m_fattureInfo.assign(infos.begin(), infos.end());

    loadFattureList();

    const QString msg = QString("%1 fattura/e caricata/e").arg(m_fattureInfo.size());
    statusBar()->showMessage(msg);
}

// ---------------------------------------------------------------------------
// POPOLAMENTO LISTA (replica la logica Windows con intestazioni per emittente)
// ---------------------------------------------------------------------------
void MainWindow::loadFattureList()
{
    populateListWidget();
    updateNavButtons();
}

void MainWindow::populateListWidget()
{
    m_listWidget->clear();
    m_listToFatturaIdx.clear();

    if (m_fattureInfo.empty()) return;

    // Raggruppa per cedente (emittente)
    std::string currentCedente;

    for (int i = 0; i < static_cast<int>(m_fattureInfo.size()); ++i)
    {
        const auto& f = m_fattureInfo[i];

        if (f.cedenteDenominazione != currentCedente)
        {
            currentCedente = f.cedenteDenominazione;

            // Riga separatore vuota (tranne la prima)
            if (!m_listToFatturaIdx.empty())
            {
                auto* sep = new QListWidgetItem("", m_listWidget);
                sep->setFlags(Qt::NoItemFlags);
                sep->setSizeHint(QSize(0, 8));
                sep->setBackground(QColor(245, 245, 245));
                m_listToFatturaIdx.push_back(-2);
            }

            // Intestazione emittente
            auto* header = new QListWidgetItem(
                "► " + QString::fromStdString(currentCedente), m_listWidget);
            styleListItem(header, true);
            header->setFlags(Qt::ItemIsEnabled); // non selezionabile
            m_listToFatturaIdx.push_back(-1);
        }

        // Voce fattura
        auto* item = new QListWidgetItem(
            QString::fromStdString(f.GetDisplayText()), m_listWidget);
        styleListItem(item, false);
        m_listToFatturaIdx.push_back(i);
    }
}

void MainWindow::styleListItem(QListWidgetItem* item, bool isHeader) const
{
    if (isHeader)
    {
        item->setBackground(QColor(0, 120, 215));
        item->setForeground(Qt::white);
        QFont f = item->font();
        f.setBold(true);
        f.setPointSize(10);
        item->setFont(f);
        item->setSizeHint(QSize(0, 32));
    }
    else
    {
        item->setBackground(Qt::white);
        item->setForeground(QColor(32, 32, 32));
        QFont f = item->font();
        f.setPointSize(10);
        item->setFont(f);
        item->setSizeHint(QSize(0, 32));
    }
}

// ---------------------------------------------------------------------------
// SELEZIONE ELEMENTO LISTA
// ---------------------------------------------------------------------------
void MainWindow::onFatturaSelectionChanged()
{
    const QList<QListWidgetItem*> selected = m_listWidget->selectedItems();

    if (selected.size() == 1)
    {
        const int row = m_listWidget->row(selected.first());
        if (row < 0 || row >= static_cast<int>(m_listToFatturaIdx.size())) return;
        const int fatturaIdx = m_listToFatturaIdx[row];
        if (fatturaIdx < 0) return;

        applyStylesheetToIndex(fatturaIdx, m_lastXsltPath.isEmpty() ? findStylesheet("Asso") : m_lastXsltPath);
    }
    else if (selected.size() > 1)
    {
        applyStylesheetMultiple(getSelectedFatturaIndices(),
                                findStylesheet(m_lastXsltPath.isEmpty() ? "Asso" : ""));
    }
}

// ---------------------------------------------------------------------------
// APPLICAZIONE FOGLIO DI STILE
// ---------------------------------------------------------------------------
QString MainWindow::findStylesheet(const QString& keyword) const
{
    if (!keyword.isEmpty())
    {
        // Se è un percorso file valido, usalo direttamente
        if (QFile::exists(keyword)) return keyword;
    }

    const QString resPath = getResourcesPath();
    if (resPath.isEmpty()) return {};

    QDir dir(resPath);
    dir.setNameFilters({"*.xsl"});

    if (!keyword.isEmpty())
    {
        for (const QString& f : dir.entryList())
            if (f.contains(keyword, Qt::CaseInsensitive))
                return dir.absoluteFilePath(f);
    }

    // Fallback: primo XSL trovato
    const auto files = dir.entryList();
    if (!files.isEmpty()) return dir.absoluteFilePath(files.first());
    return {};
}

void MainWindow::applyStylesheetToIndex(int fatturaIdx, const QString& xsltPath)
{
    if (fatturaIdx < 0 || fatturaIdx >= static_cast<int>(m_fattureInfo.size())) return;

    const std::string xmlPath = m_fattureInfo[fatturaIdx].filePath;
    if (!m_parser->LoadXmlFile(xmlPath)) return;

    const QString xslt = !xsltPath.isEmpty()    ? xsltPath
                        : !m_lastXsltPath.isEmpty() ? m_lastXsltPath
                        : findStylesheet("Asso");
    if (xslt.isEmpty()) return;

    std::string htmlOutput;
    if (!m_parser->ApplyXsltTransform(xslt.toStdString(), htmlOutput)) return;

    m_lastXsltPath = xslt;
    saveStylesheetPreference(xslt);

    // Salva l'HTML nella cartella temporanea e naviga
    const QString htmlPath = m_extractPath + "fattura_visualizzata.html";
    QFile f(htmlPath);
    if (f.open(QIODevice::WriteOnly | QIODevice::Text))
    {
        f.write(htmlOutput.c_str());
        f.close();
    }

    m_webView->setUrl(QUrl::fromLocalFile(htmlPath));
    statusBar()->showMessage("Fattura visualizzata");
}

void MainWindow::applyStylesheetMultiple(const QList<int>& indices, const QString& xsltPath)
{
    if (indices.isEmpty()) return;

    const QString xslt = xsltPath.isEmpty() ? findStylesheet("Asso") : xsltPath;
    if (xslt.isEmpty()) return;

    m_multipleHtmlFiles.clear();
    m_currentHtmlIndex = -1;

    for (int idx : indices)
    {
        if (idx < 0 || idx >= static_cast<int>(m_fattureInfo.size())) continue;
        if (!m_parser->LoadXmlFile(m_fattureInfo[idx].filePath)) continue;

        std::string htmlOutput;
        if (!m_parser->ApplyXsltTransform(xslt.toStdString(), htmlOutput)) continue;

        const QString htmlPath = m_extractPath
            + QString("fattura_%1.html").arg(idx);
        QFile f(htmlPath);
        if (f.open(QIODevice::WriteOnly | QIODevice::Text))
        {
            f.write(htmlOutput.c_str());
            f.close();
            m_multipleHtmlFiles.append(htmlPath);
        }
    }

    if (!m_multipleHtmlFiles.isEmpty())
    {
        m_currentHtmlIndex = 0;
        m_webView->setUrl(QUrl::fromLocalFile(m_multipleHtmlFiles.first()));
        updateNavButtons();
        statusBar()->showMessage(
            QString("%1 fatture visualizzate").arg(m_multipleHtmlFiles.size()));
    }
}

// ---------------------------------------------------------------------------
// AZIONI FOGLIO DI STILE
// ---------------------------------------------------------------------------
void MainWindow::onApplyStylesheetMinistero()
{
    const QString xslt = findStylesheet("Ministero");
    if (xslt.isEmpty()) { statusBar()->showMessage("Foglio Ministero non trovato"); return; }
    for (int idx : getSelectedFatturaIndices())
        applyStylesheetToIndex(idx, xslt);
}

void MainWindow::onApplyStylesheetAsso()
{
    const QString xslt = findStylesheet("Asso");
    if (xslt.isEmpty()) { statusBar()->showMessage("Foglio Assosoftware non trovato"); return; }
    for (int idx : getSelectedFatturaIndices())
        applyStylesheetToIndex(idx, xslt);
}

void MainWindow::onSelectStylesheet()
{
    const QString resPath = getResourcesPath();
    const QString xslt = QFileDialog::getOpenFileName(
        this, "Seleziona foglio di stile",
        resPath.isEmpty() ? QDir::homePath() : resPath,
        "Fogli di stile XSLT (*.xsl *.xslt);;Tutti i file (*)");

    if (xslt.isEmpty()) return;

    m_lastXsltPath = xslt;
    saveStylesheetPreference(xslt);

    const auto indices = getSelectedFatturaIndices();
    if (indices.size() > 1)
        applyStylesheetMultiple(indices, xslt);
    else if (!indices.isEmpty())
        applyStylesheetToIndex(indices.first(), xslt);
}

// ---------------------------------------------------------------------------
// NAVIGAZIONE MULTI-FATTURA
// ---------------------------------------------------------------------------
void MainWindow::onPrevFattura()
{
    if (m_currentHtmlIndex > 0)
    {
        --m_currentHtmlIndex;
        m_webView->setUrl(QUrl::fromLocalFile(m_multipleHtmlFiles[m_currentHtmlIndex]));
        updateNavButtons();
    }
}

void MainWindow::onNextFattura()
{
    if (m_currentHtmlIndex >= 0
        && m_currentHtmlIndex < m_multipleHtmlFiles.size() - 1)
    {
        ++m_currentHtmlIndex;
        m_webView->setUrl(QUrl::fromLocalFile(m_multipleHtmlFiles[m_currentHtmlIndex]));
        updateNavButtons();
    }
}

void MainWindow::updateNavButtons()
{
    const bool multi = m_multipleHtmlFiles.size() > 1;
    m_btnPrev->setEnabled(multi && m_currentHtmlIndex > 0);
    m_btnNext->setEnabled(multi && m_currentHtmlIndex < m_multipleHtmlFiles.size() - 1);
}

// ---------------------------------------------------------------------------
// ZOOM
// ---------------------------------------------------------------------------
void MainWindow::onZoomIn()
{
    m_zoomLevel = qMin(m_zoomLevel + 10, 200);
    m_webView->setZoomFactor(m_zoomLevel / 100.0);
    m_labelZoom->setText(QString::number(m_zoomLevel) + "%");
}

void MainWindow::onZoomOut()
{
    m_zoomLevel = qMax(m_zoomLevel - 10, 50);
    m_webView->setZoomFactor(m_zoomLevel / 100.0);
    m_labelZoom->setText(QString::number(m_zoomLevel) + "%");
}

// ---------------------------------------------------------------------------
// STAMPA / SALVA PDF
// ---------------------------------------------------------------------------
void MainWindow::onPrintToPdf()
{
    const QString destFile = QFileDialog::getSaveFileName(
        this, "Salva come PDF",
        QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation),
        "File PDF (*.pdf)");

    if (destFile.isEmpty()) return;

    m_webView->page()->printToPdf(destFile);
    statusBar()->showMessage("PDF salvato: " + destFile);
}

// ---------------------------------------------------------------------------
// MENU CONTESTUALE
// ---------------------------------------------------------------------------
void MainWindow::onContextMenu(const QPoint& pos)
{
    const auto indices = getSelectedFatturaIndices();
    if (indices.isEmpty()) return;

    QMenu menu(this);

    if (indices.size() == 1)
    {
        menu.addAction("Visualizza con foglio Ministero",    this, &MainWindow::onApplyStylesheetMinistero);
        menu.addAction("Visualizza con foglio Assosoftware", this, &MainWindow::onApplyStylesheetAsso);
        menu.addSeparator();
        menu.addAction("Cambia foglio di stile...", this, &MainWindow::onSelectStylesheet);
        menu.addSeparator();
        menu.addAction("Stampa fattura", this, &MainWindow::onPrintToPdf);
    }
    else
    {
        menu.addAction(QString("Stampa %1 fatture selezionate").arg(indices.size()),
                       this, &MainWindow::onPrintToPdf);
        menu.addSeparator();
        menu.addAction("Cambia foglio di stile...", this, &MainWindow::onSelectStylesheet);
    }

    menu.exec(m_listWidget->mapToGlobal(pos));
}

// ---------------------------------------------------------------------------
// CONFIGURAZIONE FIRMA PDF
// ---------------------------------------------------------------------------
void MainWindow::onPdfSignConfig()
{
    PdfSignConfig cfg;
    PdfSigner::LoadConfig(cfg);

    QDialog dlg(this);
    dlg.setWindowTitle("Configurazione Firma Digitale PDF");
    dlg.setMinimumWidth(440);

    QFormLayout* form = new QFormLayout(&dlg);

    QCheckBox* chkEnable = new QCheckBox("Abilita firma digitale automatica dei PDF", &dlg);
    chkEnable->setChecked(cfg.enabled);
    form->addRow(chkEnable);

    QLineEdit* edCert = new QLineEdit(QString::fromStdString(cfg.certPath), &dlg);
    QPushButton* btnBrowse = new QPushButton("Sfoglia...", &dlg);
    QHBoxLayout* certLay = new QHBoxLayout();
    certLay->addWidget(edCert);
    certLay->addWidget(btnBrowse);
    form->addRow("Certificato (.pfx/.p12):", certLay);

    QLineEdit* edPass = new QLineEdit(QString::fromStdString(cfg.certPassword), &dlg);
    edPass->setEchoMode(QLineEdit::Password);
    form->addRow("Password:", edPass);

    QLineEdit* edReason   = new QLineEdit(QString::fromStdString(cfg.signReason),   &dlg);
    QLineEdit* edLocation = new QLineEdit(QString::fromStdString(cfg.signLocation), &dlg);
    QLineEdit* edContact  = new QLineEdit(QString::fromStdString(cfg.signContact),  &dlg);
    form->addRow("Motivo:", edReason);
    form->addRow("Localita:", edLocation);
    form->addRow("Contatto:", edContact);

    QDialogButtonBox* btns = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
    form->addRow(btns);

    connect(btnBrowse, &QPushButton::clicked, [&]() {
        const QString f = QFileDialog::getOpenFileName(
            &dlg, "Seleziona certificato", QDir::homePath(),
            "Certificati (*.pfx *.p12);;Tutti i file (*)");
        if (!f.isEmpty()) edCert->setText(f);
    });
    connect(btns, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
    connect(btns, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);

    if (dlg.exec() == QDialog::Accepted)
    {
        cfg.enabled      = chkEnable->isChecked();
        cfg.certPath     = edCert->text().toStdString();
        cfg.certPassword = edPass->text().toStdString();
        cfg.signReason   = edReason->text().toStdString();
        cfg.signLocation = edLocation->text().toStdString();
        cfg.signContact  = edContact->text().toStdString();
        PdfSigner::SaveConfig(cfg);
        statusBar()->showMessage("Configurazione firma salvata");
    }
}

// ---------------------------------------------------------------------------
// INFO
// ---------------------------------------------------------------------------
void MainWindow::onAbout()
{
    QMessageBox dlg(this);
    dlg.setWindowTitle("Informazioni su FatturaView");
    dlg.setTextFormat(Qt::RichText);
    dlg.setText(
        "<b>FatturaView</b> v1.0.0<br>"
        "Visualizzatore Fatture Elettroniche (FatturaPA)<br><br>"
        "Autore: Roberto Ferri<br>"
        "Copyright &copy; 2026<br><br>"
        "Sito web: <a href='https://www.fatturaview.it'>www.fatturaview.it</a><br><br>"
        "Lettore di FatturaPA XML con fogli di stile XSLT"
    );
    dlg.setTextInteractionFlags(Qt::TextBrowserInteraction);
    dlg.exec();
}

// ---------------------------------------------------------------------------
// HELPER: indici fattureInfo degli elementi selezionati
// ---------------------------------------------------------------------------
QList<int> MainWindow::getSelectedFatturaIndices() const
{
    QList<int> result;
    for (QListWidgetItem* item : m_listWidget->selectedItems())
    {
        const int row = m_listWidget->row(item);
        if (row < 0 || row >= static_cast<int>(m_listToFatturaIdx.size())) continue;
        const int idx = m_listToFatturaIdx[row];
        if (idx >= 0) result.append(idx);
    }
    return result;
}

// ---------------------------------------------------------------------------
// HELPER: percorso cartella Resources
// ---------------------------------------------------------------------------
QString MainWindow::getResourcesPath() const
{
    // Release: <exe_dir>/Resources/
    const QString exeDir = QCoreApplication::applicationDirPath();
    for (const QString& rel : {"Resources", "../Resources",
                                "../share/fatturaview/Resources"})
    {
        const QString p = QDir::cleanPath(exeDir + "/" + rel);
        if (QDir(p).exists()) return p + "/";
    }
    return {};
}

// ---------------------------------------------------------------------------
// PREFERENZA FOGLIO DI STILE  (QSettings → ~/.config/FatturaView/FatturaView.conf)
// ---------------------------------------------------------------------------
void MainWindow::saveStylesheetPreference(const QString& path)
{
    QSettings s("FatturaView", "FatturaView");
    s.setValue("StylesheetPath", path);
}

QString MainWindow::loadStylesheetPreference() const
{
    QSettings s("FatturaView", "FatturaView");
    return s.value("StylesheetPath").toString();
}
