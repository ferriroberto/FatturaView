# Risoluzione Avviso Foxit PDF Reader

## 🔍 Problema

Foxit PDF Reader mostra il messaggio:

> **"È stato aperto dall'app riportata di seguito senza una firma valida. Questo potrebbe indicare un problema di protezione..."**

## 📋 Causa

Questo avviso appare perché:

1. **L'eseguibile FatturaView.exe NON è firmato digitalmente**
2. Foxit controlla la firma dell'applicazione che genera/apre il PDF
3. Applicazioni non firmate sono considerate "non trusted"

**IMPORTANTE**: Questo avviso è **DIVERSO** dalla firma del contenuto del PDF. Si riferisce alla firma del file `.exe`.

## ✅ Soluzioni

### Soluzione 1: Firmare l'Eseguibile FatturaView.exe (RACCOMANDATO)

Questa è la soluzione definitiva che elimina completamente l'avviso.

#### Requisiti

- **Certificato code-signing** (.pfx o .p12)
- **Windows SDK** (contiene `signtool.exe`)

#### Download Windows SDK

Se non hai signtool.exe installato:

1. Vai su: https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/
2. Scarica **Windows SDK**
3. Durante l'installazione seleziona: **"Windows SDK Signing Tools for Desktop Apps"**

#### Passaggi

**Metodo A - Script Automatico (Più Facile)**:

```powershell
# Esegui nella directory del progetto
.\SignExecutable.ps1

# Oppure specifica il certificato manualmente:
.\SignExecutable.ps1 -CertPath "C:\Certificati\MioCert.pfx" -CertPassword "password123"
```

Lo script:
- ✅ Cerca automaticamente signtool.exe
- ✅ Usa il certificato configurato in FatturaView
- ✅ Aggiunge timestamp per validità a lungo termine
- ✅ Verifica la firma dopo averla applicata

**Metodo B - Manuale con signtool**:

```powershell
# Cerca signtool.exe
$signtool = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin" -Filter "signtool.exe" -Recurse | Select-Object -First 1

# Firma l'eseguibile
& $signtool.FullName sign /f "C:\Cert\MioCert.pfx" /p "password" /t "http://timestamp.digicert.com" /fd SHA256 ".\x64\Debug\FatturaView.exe"
```

**Metodo C - PowerShell Set-AuthenticodeSignature**:

```powershell
$cert = Get-PfxCertificate -FilePath "C:\Cert\MioCert.pfx"
Set-AuthenticodeSignature -FilePath ".\x64\Debug\FatturaView.exe" -Certificate $cert -TimestampServer "http://timestamp.digicert.com"
```

#### Verifica

```powershell
# Verifica che la firma sia valida
Get-AuthenticodeSignature -FilePath ".\x64\Debug\FatturaView.exe" | Format-List
```

Output atteso:
```
Status        : Valid
SignerCertificate : ...
TimeStamperCertificate : ...
```

### Soluzione 2: Configurare Foxit come "Trusted"

Se non puoi firmare l'eseguibile (es. certificato self-signed), puoi ridurre gli avvisi:

#### In Foxit PDF Reader:

1. Apri **Foxit PDF Reader**
2. Vai su **File** → **Preferences** (o Ctrl+K)
3. Seleziona **Trust Manager** nella lista a sinistra
4. Nella sezione **"Certified Documents"**:
   - ☑ Spunta **"Always trust documents certified by..."**
5. Aggiungi FatturaView alla lista delle applicazioni trusted:
   - Clicca **"Manage Trusted Identities"**
   - Importa il certificato usato
   - Marca come **"Trust this certificate as trusted root"**
6. Clicca **OK**

### Soluzione 3: Whitelist Applicazione in Foxit

1. In Foxit, vai su **File** → **Preferences** → **File Associations**
2. Aggiungi FatturaView.exe alla lista delle applicazioni attendibili
3. Clicca **"Trust"** quando richiesto

### Soluzione 4: Disabilitare il Controllo in Foxit (Non Raccomandato)

⚠️ **Non consigliato per sicurezza**, ma funziona:

1. Foxit PDF Reader → **File** → **Preferences**
2. **Trust Manager** → **Internet Security**
3. Disabilita **"Enable Enhanced Security"**

## 🎯 Raccomandazioni per Tipo di Certificato

### Certificato Self-Signed (Test/Personale)

✅ **Firma l'eseguibile**: Riduce avvisi ma Foxit mostrerà "firma non trusted"
✅ **Aggiungi certificato a Foxit**: Marca come trusted manualmente
❌ **Avviso rimane** ma meno prominente

**Procedura**:
1. Firma FatturaView.exe con certificato self-signed
2. In Foxit, importa il certificato e marcalo come trusted
3. Riavvia Foxit

### Certificato da CA Riconosciuta (Produzione)

✅ **Firma l'eseguibile**: **ELIMINA COMPLETAMENTE** l'avviso
✅ **Nessuna configurazione** Foxit necessaria
✅ **Funziona immediatamente** per tutti gli utenti

**Certificati raccomandati**:
- **DigiCert Code Signing** (~€350/anno)
- **GlobalSign Code Signing** (~€300/anno)
- **Comodo Code Signing** (~€150/anno)
- **Sectigo Code Signing** (~€200/anno)

## 📊 Confronto Soluzioni

| Soluzione | Costo | Efficacia | Difficoltà |
|-----------|-------|-----------|-----------|
| Certificato CA + Firma EXE | €150-350/anno | ⭐⭐⭐⭐⭐ | ⭐⭐ |
| Self-signed + Firma EXE | Gratis | ⭐⭐⭐ | ⭐⭐ |
| Whitelist in Foxit | Gratis | ⭐⭐ | ⭐ |
| Disabilita controllo | Gratis | ⭐⭐⭐⭐ | ⭐ (⚠️ Non sicuro) |

## 🔧 Script di Build Automatico

Per firmare automaticamente ad ogni build, aggiungi al progetto un **Post-Build Event**:

1. Visual Studio → **Proprietà Progetto**
2. **Build Events** → **Post-Build Event**
3. **Command Line**:
   ```
   powershell.exe -ExecutionPolicy Bypass -File "$(ProjectDir)SignExecutable.ps1" -ExePath "$(TargetPath)"
   ```

Questo firmerà automaticamente l'eseguibile dopo ogni compilazione.

## 🐛 Troubleshooting

### "signtool.exe non trovato"

**Soluzione**: Installa Windows SDK o specifica il percorso:

```powershell
$env:PATH += ";C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64"
```

### "La firma è presente ma non trusted"

**Causa**: Certificato self-signed o scaduto

**Soluzione**:
1. Verifica validità certificato:
   ```powershell
   $cert = Get-PfxCertificate "C:\Cert\Cert.pfx"
   $cert.NotBefore
   $cert.NotAfter
   ```
2. Per produzione, acquista certificato da CA riconosciuta

### "Timestamp server non risponde"

**Soluzione**: Prova server alternativi:

```powershell
# DigiCert
-TimestampServer "http://timestamp.digicert.com"

# GlobalSign
-TimestampServer "http://timestamp.globalsign.com/scripts/timstamp.dll"

# Comodo
-TimestampServer "http://timestamp.comodoca.com"

# Sectigo
-TimestampServer "http://timestamp.sectigo.com"
```

### "Certificato non contiene chiave privata"

**Causa**: Il file .pfx non include la chiave privata

**Soluzione**:
1. Quando esporti il certificato, seleziona **"Export the private key"**
2. Formato: **Personal Information Exchange (.pfx)**
3. ☑ **"Include all certificates in the certification path if possible"**
4. Imposta una password

## 📚 Risorse Aggiuntive

### Documentazione Microsoft

- **Code Signing Best Practices**: https://docs.microsoft.com/en-us/windows-hardware/drivers/dashboard/code-signing-best-practices
- **Signtool.exe**: https://docs.microsoft.com/en-us/windows/win32/seccrypto/signtool

### Acquistare Certificati

- **DigiCert**: https://www.digicert.com/code-signing/
- **GlobalSign**: https://www.globalsign.com/en/code-signing-certificate
- **Sectigo**: https://sectigo.com/ssl-certificates-tls/code-signing

### Alternative Gratuite (Solo Test)

- **Self-Signed Certificate** (PowerShell):
  ```powershell
  $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=FatturaView" -CertStoreLocation "Cert:\CurrentUser\My"
  ```

## 📞 Supporto

Per ulteriori domande:
- **Email**: roberto.ferri@example.com
- **GitHub Issues**: https://github.com/ferriroberto/FatturaView/issues

---

## ✅ Checklist Rapida

**Per eliminare completamente l'avviso di Foxit**:

- [ ] Acquista certificato code-signing da CA riconosciuta (~€150-350/anno)
- [ ] Installa Windows SDK (per signtool.exe)
- [ ] Esegui `.\SignExecutable.ps1`
- [ ] Verifica firma: `Get-AuthenticodeSignature .\x64\Debug\FatturaView.exe`
- [ ] Distribuisci l'eseguibile firmato
- [ ] ✅ Avviso Foxit **ELIMINATO**!

**Alternativa per test/uso personale**:

- [ ] Crea certificato self-signed
- [ ] Firma eseguibile con script
- [ ] Importa certificato in Foxit come trusted
- [ ] ⚠️ Avviso ridotto ma presente

---

**FatturaView** v1.0.0  
© 2026 Roberto Ferri
