# Guida alla Firma Digitale dei PDF

## Panoramica

FatturaView supporta la firma digitale automatica dei PDF generati dalle fatture elettroniche. Questa funzionalità elimina gli avvisi di sicurezza mostrati dai lettori PDF e garantisce l'autenticità e l'integrità del documento.

## Requisiti

Per utilizzare la firma digitale dei PDF è necessario:

1. **PowerShell 3.0 o superiore** (già presente in Windows 8/10/11)
2. **.NET Framework 4.5 o superiore** (già presente in Windows 8/10/11)
3. **Un certificato digitale** in formato `.pfx` o `.p12`

### Ottenere un Certificato Digitale

Puoi ottenere un certificato digitale in diversi modi:

#### Opzione 1: Certificato Self-Signed (Solo per test)
Per creare un certificato di test usando PowerShell:

```powershell
$cert = New-SelfSignedCertificate -Subject "CN=Il Tuo Nome" `
    -Type CodeSigningCert `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -NotAfter (Get-Date).AddYears(5)

$certPath = "C:\Certificati\MioCertificato.pfx"
$password = ConvertTo-SecureString -String "TuaPassword" -Force -AsPlainText

Export-PfxCertificate -Cert $cert -FilePath $certPath -Password $password
```

**⚠️ IMPORTANTE**: I certificati self-signed vanno bene solo per test interni. Per uso professionale, acquista un certificato da una CA riconosciuta.

#### Opzione 2: Certificato da Autorità di Certificazione (Raccomandato)

Per uso professionale, acquista un certificato digitale da:
- **InfoCert** (Italia)
- **Aruba PEC**
- **Poste Italiane**
- **Namirial**
- Altre CA italiane riconosciute

Questi certificati:
- Sono riconosciuti come attendibili dai lettori PDF
- Non mostrano avvisi di sicurezza
- Hanno validità legale in Italia
- Costano generalmente €50-200/anno

## Configurazione

### 1. Accedere alle Impostazioni

1. Apri FatturaView
2. Vai su **Menu → Impostazioni → Configurazione firma PDF...**

### 2. Configurare la Firma

Nella finestra di configurazione:

#### Abilitazione
- ✅ Spunta **"Abilita firma digitale automatica dei PDF"**

#### Certificato
- **Percorso certificato**: Clicca su "Sfoglia" e seleziona il tuo file `.pfx` o `.p12`
- **Password certificato**: Inserisci la password del certificato

#### Informazioni Firma (Opzionali)
- **Motivo della firma**: Es. "Fattura Elettronica Italiana" (predefinito)
- **Località**: Es. "Italia", "Milano", ecc.
- **Contatto**: Es. email o numero di telefono

### 3. Salvare la Configurazione

Clicca su **"Salva"** per salvare le impostazioni.

## Utilizzo

Una volta configurata la firma digitale:

1. Seleziona una o più fatture dalla lista
2. Clicca su **"Stampa fatture selezionate (PDF)"** dal menu File
3. Il PDF verrà generato e **automaticamente firmato** digitalmente
4. Il PDF firmato verrà aperto nel tuo lettore PDF predefinito

### Processo di Firma

Quando generi un PDF:

1. FatturaView genera il PDF dall'HTML
2. Se la firma è abilitata, compare il messaggio "Firma digitale del PDF in corso..."
3. PowerShell firma il PDF usando iTextSharp
4. Il PDF firmato viene salvato e aperto

**Tempo richiesto**: 3-10 secondi per la prima firma (download iTextSharp), poi 1-2 secondi per le successive.

## Verifica della Firma

Per verificare che il PDF sia correttamente firmato:

### Adobe Acrobat Reader
1. Apri il PDF
2. Dovresti vedere un banner blu in alto: "Firmato e tutte le firme sono valide"
3. Clicca sulla firma per vedere i dettagli

### Altre Lettori PDF
- **Foxit Reader**: Mostra l'icona firma nel pannello laterale
- **PDF-XChange**: Verifica automatica della firma
- **Browser (Chrome, Edge)**: Può mostrare "Documento firmato"

## Risoluzione Problemi

### Il PDF non viene firmato

**Verifica:**
1. Che PowerShell sia disponibile (normalmente già presente in Windows)
2. Che il certificato esista nel percorso specificato
3. Che la password del certificato sia corretta
4. Che hai diritti di scrittura nella cartella temporanea

### Errore "PowerShell 3.0 non trovato"

**Soluzione**: Aggiorna Windows o installa Windows Management Framework 4.0+

### Il lettore PDF mostra ancora avvisi

**Possibili cause:**
1. **Certificato self-signed**: È normale. I certificati self-signed non sono riconosciuti come attendibili. Usa un certificato da una CA per eliminare completamente gli avvisi.
2. **Certificato scaduto**: Verifica la validità del certificato
3. **Catena di certificazione incompleta**: Assicurati che il certificato includa la catena completa

### Messaggio "Download iTextSharp..."

**Cosa succede**: Alla prima firma, FatturaView scarica automaticamente la libreria iTextSharp (circa 2MB) da NuGet. Questo accade una sola volta.

**Se il download fallisce**:
- Verifica la connessione Internet
- Controlla che Windows Firewall non blocchi PowerShell
- Prova a disabilitare temporaneamente l'antivirus

## Sicurezza

### Password del Certificato

La password del certificato viene salvata in:
```
%APPDATA%\FatturaView\pdfsign.cfg
```

**⚠️ IMPORTANTE**: Questo file contiene la password in chiaro. Proteggi l'accesso al tuo computer.

**Raccomandazioni**:
- Usa password complesse per il certificato
- Non condividere il file di configurazione
- Mantieni il certificato in un luogo sicuro
- Considera l'uso di password manager aziendali

### Libreria iTextSharp

FatturaView usa **iTextSharp 5.5.13.3**, una libreria open source affidabile per la manipolazione di PDF. La libreria viene scaricata da:
- **NuGet.org**: Repository ufficiale Microsoft per pacchetti .NET
- **Verifica**: La libreria viene scaricata solo se necessaria

## Disabilitazione

Per disabilitare la firma digitale:

1. Menu → **Impostazioni → Configurazione firma PDF...**
2. ❌ Deseleziona **"Abilita firma digitale automatica dei PDF"**
3. Clicca su **"Salva"**

I PDF verranno generati senza firma (con possibili avvisi di sicurezza).

## FAQ

### Q: Posso usare lo stesso certificato per più computer?
**R**: Sì, copia il file `.pfx` su ogni computer e configura FatturaView con la stessa password.

### Q: La firma ha validità legale?
**R**: Sì, se usi un certificato da una CA riconosciuta. I certificati self-signed non hanno validità legale.

### Q: Posso firmare fatture già generate?
**R**: No, attualmente FatturaView firma solo durante la generazione. Per firmare PDF esistenti, usa Adobe Acrobat o altri software di firma PDF.

### Q: La firma rallenta la generazione PDF?
**R**: La prima firma richiede 3-10 secondi (download iTextSharp). Le successive richiedono 1-2 secondi aggiuntivi.

### Q: Posso vedere l'aspetto visibile della firma nel PDF?
**R**: Attualmente la firma è invisibile (non occupa spazio nel documento). In versioni future sarà possibile configurare una firma visibile.

## Supporto

Per problemi o domande:
- Email: roberto.ferri@example.com
- GitHub: https://github.com/ferriroberto/FatturaView/issues

## Aggiornamenti Futuri

Funzionalità pianificate:
- [ ] Firma visibile con immagine personalizzata
- [ ] Supporto per timestamp server (RFC 3161)
- [ ] Firma con smart card
- [ ] Verifica firma esistente
- [ ] Batch signing (firma massiva)

---

**FatturaView** - Versione 1.0.0  
© 2026 Roberto Ferri
