# Lettore Fatture Elettroniche

## Descrizione
Applicazione Windows per la visualizzazione di fatture elettroniche XML italiane con interfaccia a due pannelli.

## Funzionalità
- **Importazione archivi ZIP**: Estrae automaticamente file XML da archivi compressi
- **Interfaccia a due pannelli**: 
  - **Pannello sinistro**: Lista delle fatture estratte
  - **Pannello destro**: Visualizzazione HTML integrata della fattura selezionata
- **Applicazione fogli di stile**: 
  - Foglio di stile Ministero dell'Economia
  - Foglio di stile Assosoftware
- **Visualizzazione integrata**: WebBrowser incorporato per vedere le fatture direttamente nell'applicazione
- **Generazione PDF**: Converti fatture in PDF usando Microsoft Edge headless
- **Firma digitale PDF**: Firma automaticamente i PDF generati con certificato digitale (elimina avvisi di sicurezza)
- **Selezione multipla**: Genera PDF da più fatture contemporaneamente
- **Memoria foglio di stile**: Ricorda l'ultimo foglio di stile usato per applicazioni successive

## Utilizzo

### 1. Aprire un archivio ZIP
- Menu: **File → Apri archivio ZIP...**
- Selezionare il file ZIP contenente le fatture elettroniche
- Le fatture verranno estratte e mostrate nella lista a sinistra

### 2. Visualizzare una fattura
**Metodo 1 - Doppio click:**
- Fare doppio click su una fattura nella lista
- Se è la prima volta, selezionare il foglio di stile XSLT
- La fattura viene visualizzata nel pannello destro

**Metodo 2 - Menu:**
- Selezionare una fattura dalla lista (singolo click)
- Menu: **Visualizza → Applica foglio stile Ministero** o **Assosoftware**
- La fattura viene visualizzata nel pannello destro

### 3. Cambiare foglio di stile
- Menu: **Visualizza → Cambia foglio di stile...**
- Selezionare un nuovo file XSLT
- La fattura corrente viene ricaricata con il nuovo stile

### 4. Generare PDF
- Selezionare una o più fatture dalla lista
- Menu: **File → Stampa fatture selezionate (PDF)**
- Il PDF viene generato e aperto automaticamente

### 5. Configurare la firma digitale (opzionale)
- Menu: **Impostazioni → Configurazione firma PDF...**
- Abilita la firma digitale e configura il certificato
- I PDF generati saranno automaticamente firmati
- **Vedi GUIDA_FIRMA_DIGITALE.md per istruzioni dettagliate**

## Layout dell'interfaccia

```
┌─────────────────────────────────────────────────┐
│ File  Visualizza  ?                             │
├──────────────┬──────────────────────────────────┤
│              │                                   │
│  LISTA       │     VISUALIZZAZIONE              │
│  FATTURE     │     FATTURA HTML                 │
│              │                                   │
│ • Fattura1   │  ┌─────────────────────────┐    │
│ • Fattura2   │  │                         │    │
│ • Fattura3   │  │  Contenuto fattura      │    │
│ • ...        │  │  trasformata con XSLT   │    │
│              │  │                         │    │
│              │  └─────────────────────────┘    │
│              │                                   │
└──────────────┴──────────────────────────────────┘
```

## Fogli di stile XSLT

### Dove scaricare i fogli di stile:

**Foglio di stile Ministero:**
- Disponibile sul sito dell'Agenzia delle Entrate
- URL: https://www.fatturapa.gov.it/

**Foglio di stile Assosoftware:**
- Disponibile sul repository GitHub di Assosoftware
- URL: https://github.com/assosoftware/

## Requisiti
- Windows 7 o superiore (Windows 10/11 raccomandato)
- MSXML 6.0 (incluso in Windows)
- Internet Explorer 11+ (per il controllo WebBrowser)
- Microsoft Edge (per generazione PDF)
- **Per firma digitale PDF**:
  - PowerShell 3.0+ (incluso in Windows 8+)
  - .NET Framework 4.5+ (incluso in Windows 8+)
  - Certificato digitale in formato .pfx o .p12
- **Importante**: Se il pannello destro non visualizza le fatture, consultare il file TROUBLESHOOTING.md

## Risoluzione problemi comuni

### Il pannello destro è vuoto
1. Verifica che Internet Explorer 11 sia installato
2. Esegui l'applicazione come Amministratore
3. Controlla che il file XSLT sia valido
4. Consulta TROUBLESHOOTING.md per dettagli

### L'applicazione si blocca
1. Chiudi e riapri l'applicazione
2. Usa file ZIP più piccoli (max 50 fatture)
3. Verifica memoria disponibile

Per tutti i problemi, consulta il file **TROUBLESHOOTING.md** incluso nella distribuzione.

## Note tecniche
- I file vengono estratti in una cartella temporanea
- Formato supportato: FatturaPA XML (standard italiano)
- Trasformazione XSLT per rendering HTML
- Il pannello destro usa il controllo WebBrowser nativo di Windows

## Compilazione
Aprire la soluzione in Visual Studio 2019 o superiore e compilare in modalità Release.
