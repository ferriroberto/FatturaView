# Troubleshooting - Visualizzatore Fatture Elettroniche

## Problema: Non si vede la fattura nel pannello destro

### Possibili cause e soluzioni:

#### 1. Internet Explorer non configurato correttamente
**Sintomo**: Il pannello destro rimane bianco o mostra un errore.

**Soluzione**:
- Apri Internet Explorer manualmente
- Vai su Strumenti → Opzioni Internet → Avanzate
- Abilita:
  - ✓ Usa HTTP 1.1
  - ✓ Usa HTTP 1.1 tramite connessioni proxy
  - ✓ Non salvare le pagine crittografate su disco
- Disabilita:
  - ☐ Visualizza notifica per ogni errore di script

#### 2. Controllo ActiveX bloccato
**Sintomo**: Il browser non si carica nell'applicazione.

**Soluzione**:
```
1. Esegui come Amministratore
2. Menu Start → cmd (come amministratore)
3. Digita: regsvr32 /i "C:\Windows\System32\ieframe.dll"
4. Riavvia l'applicazione
```

#### 3. File HTML non creato correttamente
**Sintomo**: Messaggio "Errore durante l'applicazione del foglio di stile".

**Verifica**:
1. Controlla che il foglio XSLT sia valido
2. Verifica i permessi sulla cartella temporanea
3. Percorso temporaneo: `%TEMP%\FattureElettroniche\`

**Test manuale**:
```
1. Naviga in: %TEMP%\FattureElettroniche\
2. Cerca il file: fattura_visualizzata.html
3. Aprilo con un browser esterno
4. Se funziona → problema del controllo WebBrowser
5. Se non funziona → problema XSLT o XML
```

#### 4. Encoding del file HTML
**Sintomo**: Caratteri strani o non visualizzati correttamente.

**Soluzione**:
- Il file HTML viene salvato con encoding UTF-8
- Verifica che il foglio XSLT specifichi: `encoding="UTF-8"`
- Aggiungi al file XSLT: `<meta charset="UTF-8"/>`

#### 5. Percorso file con caratteri speciali
**Sintomo**: La navigazione fallisce silenziosamente.

**Soluzione**:
- Evita cartelle con caratteri accentati nel percorso
- Usa percorsi brevi (< 260 caratteri)
- Evita spazi multipli nei nomi delle cartelle

## Problema: Il file ZIP non viene estratto

### Possibili cause e soluzioni:

#### 1. Archivio ZIP corrotto
**Verifica**:
```
1. Apri il file ZIP con 7-Zip o WinRAR
2. Controlla se ci sono errori
3. Prova a estrarre manualmente
```

#### 2. Permessi insufficienti
**Soluzione**:
```
1. Esegui l'applicazione come Amministratore
2. Verifica i permessi su %TEMP%
3. Cambia cartella di estrazione temporanea
```

#### 3. File ZIP protetto da password
**Nota**: Attualmente non supportato.

**Workaround**:
```
1. Estrai manualmente il file ZIP
2. Usa: File → Apri cartella (se implementato)
3. O comprimi i file XML senza password
```

## Problema: Foglio di stile non funziona

### Diagnostica:

#### Test XSLT base
Crea un file `test.xsl` semplice:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:output method="html" encoding="UTF-8"/>
    <xsl:template match="/">
        <html>
            <head><title>Test</title></head>
            <body>
                <h1>Test Trasformazione</h1>
                <p>Se vedi questo messaggio, XSLT funziona!</p>
            </body>
        </html>
    </xsl:template>
</xsl:stylesheet>
```

#### Verifica namespace XML
Le fatture elettroniche usano namespace. Il foglio XSLT deve dichiarare:

```xml
<xsl:stylesheet version="1.0" 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:fe="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2">
```

E usare il prefisso nei selettori:
```xml
<xsl:value-of select="//fe:CedentePrestatore/fe:DatiAnagrafici/fe:Anagrafica/fe:Denominazione"/>
```

## Problema: L'applicazione si blocca

### Cause comuni:

#### 1. File XML molto grande
**Soluzione**:
- Limita le fatture a < 5MB ciascuna
- Processa un file alla volta
- Chiudi e riapri l'applicazione periodicamente

#### 2. Troppi file nell'archivio
**Soluzione**:
- Dividi gli archivi ZIP in gruppi più piccoli
- Max consigliato: 50 fatture per ZIP

#### 3. Memoria insufficiente
**Verifica**:
```
Task Manager → scheda Prestazioni
Controlla uso RAM e CPU
```

## Log e Debug

### Abilitare modalità verbosa

**Modifica temporanea per debug**:

Nel codice, aggiungi logging:
```cpp
// In OnApplyStylesheet, aggiungi:
MessageBoxW(hWnd, htmlPath.c_str(), L"Percorso HTML generato", MB_OK);
```

### File di log
Attualmente non implementato. Per aggiungere:

1. Crea un file `log.txt` in `%TEMP%\FattureElettroniche\`
2. Scrivi timestamp e azioni
3. Consulta in caso di errore

## Requisiti di sistema verificati

✅ **Sistemi testati**:
- Windows 10 (versione 1809+)
- Windows 11

⚠️ **Sistemi con limitazioni**:
- Windows 7/8.1: Richiede IE11 installato
- Windows Server: Potrebbe richiedere configurazione aggiuntiva

❌ **Non supportato**:
- Windows XP/Vista
- Windows senza Internet Explorer

## Contatti e supporto

Per problemi persistenti:

1. Verifica la sezione GUIDA_USO.md
2. Controlla il README.md per requisiti
3. Testa con il foglio XSLT di esempio incluso
4. Verifica i log di Windows Event Viewer

## Checklist rapida

Prima di segnalare un problema, verifica:

- [ ] Internet Explorer 11 installato
- [ ] Esecuzione con privilegi amministratore (se necessario)
- [ ] Cartella %TEMP% accessibile in scrittura
- [ ] File ZIP valido e non corrotto
- [ ] File XML conforme allo standard FatturaPA
- [ ] Foglio XSLT valido e compatibile
- [ ] Spazio su disco sufficiente (>100MB liberi)
- [ ] Antivirus non blocca l'applicazione
- [ ] Firewall non blocca l'applicazione (se scarica risorse online)

## Risoluzione problemi avanzata

### Reset completo

Se nulla funziona:

```batch
1. Chiudi l'applicazione
2. Elimina: %TEMP%\FattureElettroniche\
3. Riavvia Windows
4. Esegui l'applicazione come Amministratore
5. Prova con un singolo file XML di test
```

### Verifica installazione componenti

Controlla che questi componenti siano presenti:

```
C:\Windows\System32\msxml6.dll
C:\Windows\System32\ieframe.dll
C:\Windows\System32\jscript.dll
```

Se mancanti, reinstalla Internet Explorer 11.
