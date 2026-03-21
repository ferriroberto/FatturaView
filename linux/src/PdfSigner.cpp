#include "PdfSigner.h"

#include <QStandardPaths>
#include <QDir>
#include <QProcess>
#include <QSettings>

#include <fstream>
#include <sstream>

// ---------------------------------------------------------------------------
// Percorso file di configurazione: ~/.config/FatturaView/pdfsign.cfg
// ---------------------------------------------------------------------------
std::string PdfSigner::GetConfigFilePath()
{
    QString configDir = QStandardPaths::writableLocation(QStandardPaths::AppConfigLocation);
    QDir().mkpath(configDir);
    return (configDir + "/pdfsign.cfg").toStdString();
}

// ---------------------------------------------------------------------------
// Carica la configurazione (formato key=value — compatibile con Windows)
// ---------------------------------------------------------------------------
bool PdfSigner::LoadConfig(PdfSignConfig& config)
{
    std::ifstream file(GetConfigFilePath());
    if (!file.is_open()) return false;

    std::string line;
    while (std::getline(file, line))
    {
        const auto pos = line.find('=');
        if (pos == std::string::npos) continue;

        const std::string key   = line.substr(0, pos);
        const std::string value = line.substr(pos + 1);

        if      (key == "enabled")       config.enabled      = (value == "1" || value == "true");
        else if (key == "certPath")      config.certPath     = value;
        else if (key == "certPassword")  config.certPassword = value;
        else if (key == "signReason")    config.signReason   = value;
        else if (key == "signLocation")  config.signLocation = value;
        else if (key == "signContact")   config.signContact  = value;
    }
    return true;
}

// ---------------------------------------------------------------------------
// Salva la configurazione
// ---------------------------------------------------------------------------
bool PdfSigner::SaveConfig(const PdfSignConfig& config)
{
    std::ofstream file(GetConfigFilePath());
    if (!file.is_open()) return false;

    file << "enabled="      << (config.enabled ? "1" : "0") << "\n"
         << "certPath="     << config.certPath     << "\n"
         << "certPassword=" << config.certPassword << "\n"
         << "signReason="   << config.signReason   << "\n"
         << "signLocation=" << config.signLocation << "\n"
         << "signContact="  << config.signContact  << "\n";

    return true;
}

// ---------------------------------------------------------------------------
// Verifica prerequisiti: cerca pyhanko (python3) o java
// ---------------------------------------------------------------------------
bool PdfSigner::CheckPrerequisites(std::string& errorMsg)
{
    // Prova pyhanko (pyhanko sign)
    QProcess proc;
    proc.start("python3", {"-m", "pyhanko", "--version"});
    if (proc.waitForFinished(3000) && proc.exitCode() == 0)
        return true;

    // Prova java
    proc.start("java", {"-version"});
    if (proc.waitForFinished(3000) && proc.exitCode() == 0)
        return true;

    errorMsg = "Nessun motore di firma trovato.\n"
               "Installa pyhanko (pip3 install pyhanko) oppure Java Runtime.";
    return false;
}

// ---------------------------------------------------------------------------
// Firma PDF tramite pyhanko o java+iTextSharp
// ---------------------------------------------------------------------------
bool PdfSigner::SignPdf(const std::string& pdfPath,
                         const PdfSignConfig& config,
                         std::string& errorMsg)
{
    if (!config.enabled)
    {
        errorMsg = "Firma digitale non abilitata nella configurazione.";
        return false;
    }

    // Controlla prerequisiti
    if (!CheckPrerequisites(errorMsg)) return false;

    // Output firmato: same name with _signed suffix
    std::string signedPath = pdfPath;
    const auto  dotPos     = signedPath.rfind('.');
    if (dotPos != std::string::npos)
        signedPath.insert(dotPos, "_signed");
    else
        signedPath += "_signed.pdf";

    QProcess proc;

    // Preferisce pyhanko se disponibile
    QStringList pyArgs = {
        "-m", "pyhanko", "sign", "addsig", "pkcs12",
        "--field",    "Sig1",
        "--passfile", "-",
        "--p12-file", QString::fromStdString(config.certPath),
        "--reason",   QString::fromStdString(config.signReason),
        "--location", QString::fromStdString(config.signLocation),
        "--contact",  QString::fromStdString(config.signContact),
        QString::fromStdString(pdfPath),
        QString::fromStdString(signedPath)
    };

    proc.start("python3", pyArgs);
    proc.write((config.certPassword + "\n").c_str());
    proc.closeWriteChannel();

    if (!proc.waitForFinished(30000))
    {
        errorMsg = "Timeout durante la firma.";
        return false;
    }

    if (proc.exitCode() != 0)
    {
        errorMsg = proc.readAllStandardError().toStdString();
        return false;
    }

    return true;
}
