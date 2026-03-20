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
- Windows 7 o superiore
- MSXML 6.0 (incluso in Windows)
- Internet Explorer 11+ (per il controllo WebBrowser)
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
