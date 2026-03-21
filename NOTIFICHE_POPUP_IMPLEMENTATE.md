# Notifiche Popup Implementate

## Descrizione
È stato implementato un sistema di notifiche popup che appaiono nell'angolo superiore destro del frame HTML del browser, mostrando messaggi informativi sullo stato delle operazioni in corso.

## Funzionalità

### Nuova Funzione: `FatturaViewer::ShowNotification()`
```cpp
void FatturaViewer::ShowNotification(HWND hBrowser, const std::wstring& message, const std::wstring& type = L"info");
```

**Parametri:**
- `hBrowser`: Handle del controllo browser
- `message`: Testo del messaggio da visualizzare
- `type`: Tipo di notifica (`"info"`, `"success"`, `"error"`, `"warning"`)

**Caratteristiche:**
- Notifiche animate con slide-in da destra
- Scompaiono automaticamente dopo 3 secondi con animazione slide-out
- Colori e icone differenziati per tipo:
  - **Info** (blu #0078d4): ℹ️ Operazioni in corso
  - **Success** (verde #10b981): ✓ Operazioni completate
  - **Error** (rosso #ef4444): ✗ Errori
  - **Warning** (arancione #f59e0b): ⚠️ Avvisi
- Design moderno con ombre e bordi arrotondati
- Font Segoe UI per coerenza con Windows 11
- Z-index elevato (999999) per visibilità sopra altri elementi

## Notifiche Implementate

### Caricamento Iniziale
- ℹ️ "Caricamento interfaccia..."
- ✓ "Interfaccia caricata - Pronto all'uso"

### Apertura File
- ℹ️ "Estrazione archivio ZIP..."
- ℹ️ "Caricamento file XML..."
- ℹ️ "Lettura fatture in corso..."
- ✓ "X fattura/e caricata/e con successo"
- ⚠️ "Nessuna fattura trovata"
- ✗ "Errore nel caricamento dei file"

### Visualizzazione Fatture
- ℹ️ "Applicazione foglio di stile..."
- ℹ️ "Generazione visualizzazione..."
- ✓ "Fattura visualizzata correttamente"
- ✗ "Errore trasformazione XSLT"

### Generazione PDF
- ℹ️ "Generazione PDF per X fattura/e..."
- ✓ "PDF firmato digitalmente"
- ✓ "PDF generato e aperto con successo"

## Implementazione Tecnica

### Injection JavaScript
La funzione inietta JavaScript dinamicamente nel documento HTML caricato nel browser:
- Crea un container fisso nell'angolo superiore destro
- Genera elementi div per ogni notifica
- Applica CSS inline con animazioni
- Imposta timer per rimozione automatica

### Compatibilità
- Funziona con qualsiasi contenuto HTML caricato nel browser
- Non interferisce con il contenuto della pagina
- Stile CSS isolato per evitare conflitti
- Compatibile con Internet Explorer (WebBrowser control)

### Posizionamento
```
┌─────────────────────────────┐
│                    ┌──────┐ │
│                    │ ℹ️ msg│ │
│                    └──────┘ │
│                             │
│      Contenuto HTML         │
│                             │
└─────────────────────────────┘
```

## Vantaggi
1. **Feedback visivo immediato**: L'utente vede subito lo stato dell'operazione
2. **Non invasivo**: Le notifiche non bloccano l'interfaccia
3. **Contestuale**: Appare nel frame dove avviene l'azione
4. **Professionale**: Design moderno e coerente con Windows 11
5. **Automatico**: Scompaiono da sole senza richiedere azione dell'utente
6. **Complementare**: Si affianca alla status bar tradizionale

## File Modificati
- **FatturaViewer.h**: Aggiunta dichiarazione `ShowNotification()`
- **FatturaViewer.cpp**: Implementazione completa della funzione (90 righe)
- **FatturaView.cpp**: Aggiunte chiamate a `ShowNotification()` in 15+ punti strategici

## Test Consigliati
1. Aprire un archivio ZIP → Verificare notifiche di estrazione e caricamento
2. Visualizzare una fattura → Verificare notifiche di applicazione stile
3. Generare PDF → Verificare notifiche di generazione e firma
4. Causare errori (file non valido) → Verificare notifiche di errore
5. Testare con più operazioni consecutive → Verificare stack di notifiche

## Note
- La barra di stato rimane funzionante e complementare alle notifiche popup
- Le notifiche non vengono mostrate se il browser non è ancora inizializzato
- Il sistema è thread-safe e gestisce correttamente operazioni asincrone
