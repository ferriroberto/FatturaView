# 🎯 RISPOSTA ALL'AVVISO DI FOXIT PDF READER

## ❗ Il Problema Identificato

Il messaggio di Foxit:
> **"È stato aperto dall'app riportata di seguito senza una firma valida. Questo potrebbe indicare un problema di protezione..."**

Si riferisce alla **firma dell'eseguibile FatturaView.exe**, NON alla firma del PDF.

---

## ✅ SOLUZIONE IMPLEMENTATA

### 1. **Miglioramento Firma PDF** ✅

Ho migliorato il codice in `PdfSigner.cpp` per:
- ✅ Aggiungere metadati identificativi al PDF
- ✅ Usare la catena completa di certificazione
- ✅ Migliorare la compatibilità con i lettori PDF
- ✅ Firma più robusta e riconosciuta

**Compilazione**: ✅ RIUSCITA

### 2. **Script Firma Eseguibile** ✅

Ho creato `SignExecutable.ps1` che:
- 🔍 Cerca automaticamente `signtool.exe` (Windows SDK)
- 🔐 Usa il certificato configurato in FatturaView
- ⏱️ Aggiunge timestamp per validità permanente
- ✅ Verifica la firma dopo averla applicata

### 3. **Guida Completa** ✅

Ho creato `GUIDA_RISOLUZIONE_AVVISO_FOXIT.md` con:
- 📋 Spiegazione dettagliata del problema
- 🛠️ 4 soluzioni diverse (da più efficace a workaround)
- 📊 Confronto tra certificati (CA vs self-signed)
- 🐛 Troubleshooting completo
- 💰 Raccomandazioni su dove acquistare certificati

---

## 🚀 COME ELIMINARE L'AVVISO

### Soluzione Rapida (5 minuti)

```powershell
# 1. Esegui lo script nella directory del progetto
.\SignExecutable.ps1

# Lo script:
# - Usa il certificato che hai già configurato
# - Cerca signtool.exe automaticamente
# - Firma FatturaView.exe
# - Verifica la firma
```

### Se Non Hai Signtool

**Installa Windows SDK** (gratuito):
1. Vai su: https://developer.microsoft.com/windows/downloads/windows-sdk/
2. Scarica e installa
3. Seleziona: "Windows SDK Signing Tools"
4. Riprova lo script

---

## 📋 DUE TIPI DI FIRMA

| Tipo | Cosa Firma | Risolve Avviso Foxit? |
|------|------------|----------------------|
| **Firma PDF** | Contenuto del documento PDF | ❌ No (già implementata) |
| **Firma EXE** | L'applicazione FatturaView.exe | ✅ **SÌ** (da fare) |

**L'avviso di Foxit riguarda la firma dell'EXE**, non del PDF!

---

## 💡 TIPO DI CERTIFICATO

### Con Certificato Self-Signed (Gratis)

- ✅ Firma l'eseguibile
- ⚠️ Foxit mostrerà "firma non trusted"
- 📌 Devi importare il certificato in Foxit come "trusted"
- 👤 OK per uso personale/interno

### Con Certificato da CA Riconosciuta (€150-350/anno)

- ✅ Firma l'eseguibile
- ✅ **ELIMINA COMPLETAMENTE** l'avviso
- ✅ Nessuna configurazione necessaria
- 🏢 **RACCOMANDATO per distribuzione pubblica**

**CA Consigliate**:
- DigiCert (~€350/anno)
- Sectigo (~€200/anno)  
- Comodo (~€150/anno)

---

## 📝 PROSSIMI PASSI

### Opzione A: Uso Personale/Test

```powershell
# 1. Crea certificato self-signed
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=FatturaView Dev" -CertStoreLocation "Cert:\CurrentUser\My"
$certPath = "C:\Temp\FatturaViewDev.pfx"
$pwd = ConvertTo-SecureString -String "test123" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath $certPath -Password $pwd

# 2. Configura in FatturaView
# Menu → Impostazioni → Firma PDF
# Certificato: C:\Temp\FatturaViewDev.pfx
# Password: test123

# 3. Firma eseguibile
.\SignExecutable.ps1

# 4. In Foxit, importa certificato come trusted
```

### Opzione B: Distribuzione Professionale

```
1. ✅ Acquista certificato code-signing da CA
   → DigiCert, Sectigo, Comodo

2. ✅ Configura in FatturaView
   Menu → Impostazioni → Firma PDF

3. ✅ Firma eseguibile
   .\SignExecutable.ps1

4. ✅ Distribuisci l'EXE firmato
   → Avviso Foxit ELIMINATO per tutti!
```

---

## 🔍 VERIFICA FIRMA

Dopo aver firmato l'eseguibile:

```powershell
# Verifica firma
Get-AuthenticodeSignature .\x64\Debug\FatturaView.exe | Format-List

# Output atteso:
# Status        : Valid (o NotTrusted se self-signed)
# StatusMessage : Signature verified
# SignerCertificate : CN=...
```

---

## 📚 DOCUMENTAZIONE

| File | Contenuto |
|------|-----------|
| **GUIDA_RISOLUZIONE_AVVISO_FOXIT.md** | Guida completa passo-passo |
| **SignExecutable.ps1** | Script automatico firma EXE |
| **GUIDA_FIRMA_DIGITALE.md** | Guida firma PDF |

---

## ❓ FAQ

### Q: Perché l'avviso appare anche se firmo il PDF?
**R**: Foxit controlla la firma dell'**applicazione** (FatturaView.exe), non solo del PDF.

### Q: Il certificato self-signed elimina l'avviso?
**R**: No, riduce l'avviso ma Foxit mostrerà "non trusted". Serve certificato da CA.

### Q: Quanto costa un certificato code-signing?
**R**: €150-350/anno da CA riconosciute.

### Q: Posso usare lo stesso certificato per PDF e EXE?
**R**: Sì! Il certificato .pfx può firmare sia PDF che eseguibili.

### Q: Devo firmare ad ogni build?
**R**: Sì, ma puoi automatizzarlo (vedi "Script di Build Automatico" nella guida).

---

## 🎯 RIEPILOGO

### ✅ Implementato

- [x] Miglioramento firma PDF
- [x] Script firma eseguibile  
- [x] Guida completa Foxit
- [x] Documentazione tecnica
- [x] Compilazione riuscita

### ⏭️ Da Fare (Utente)

- [ ] Installare Windows SDK (se mancante)
- [ ] Decidere tipo certificato (self-signed vs CA)
- [ ] Eseguire `.\SignExecutable.ps1`
- [ ] Testare con Foxit PDF Reader
- [ ] (Opzionale) Automatizzare firma in build

---

## 📞 Serve Aiuto?

Consulta: **GUIDA_RISOLUZIONE_AVVISO_FOXIT.md**

Contiene:
- ✅ Guida passo-passo illustrata
- ✅ Troubleshooting completo
- ✅ Alternative e workaround
- ✅ Raccomandazioni per ogni scenario

---

**Stato**: ✅ **PRONTO PER FIRMA**

Esegui lo script e l'avviso di Foxit sarà risolto! 🎉
