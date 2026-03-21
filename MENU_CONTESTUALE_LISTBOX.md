# Menu Contestuale ListBox Fatture

## Descrizione
È stato implementato un menu contestuale (tasto destro del mouse) sulla ListBox delle fatture nel pannello sinistro dell'applicazione. Il menu permette di accedere rapidamente alle funzioni di visualizzazione e stampa senza dover passare dal menu principale.

## Funzionalità

### Menu per Singola Fattura Selezionata
Quando è selezionata una sola fattura, il menu contestuale offre:

1. **Visualizza con foglio Ministero** - Applica immediatamente il foglio di stile del Ministero dell'Economia
2. **Visualizza con foglio Assosoftware** - Applica immediatamente il foglio di stile Assosoftware
3. **--- separatore ---**
4. **Cambia foglio di stile...** - Apre il dialogo di selezione per scegliere un foglio personalizzato
5. **--- separatore ---**
6. **Stampa fattura** - Genera il PDF della fattura selezionata

### Menu per Multiple Fatture Selezionate
Quando sono selezionate più fatture, il menu contestuale si adatta mostrando:

1. **Stampa X fatture selezionate** - Genera i PDF di tutte le fatture selezionate (X = numero)
2. **--- separatore ---**
3. **Cambia foglio di stile...** - Apre il dialogo per applicare uno stile a tutte le fatture selezionate

## Utilizzo

### Attivazione del Menu
1. Seleziona una o più fatture nella lista a sinistra
2. Click con il **tasto destro del mouse** sulla selezione
3. Il menu contestuale apparirà nella posizione del cursore

### Selezione Multipla
Per selezionare più fatture:
- **Ctrl + Click** per aggiungere/rimuovere fatture alla selezione
- **Shift + Click** per selezionare un intervallo
- Poi click destro per aprire il menu contestuale

## Implementazione Tecnica

### Nuovi ID Comando (Resource.h)
```cpp
#define IDM_CONTEXT_PRINT_SELECTED  126    // Stampa dal menu contestuale
#define IDM_CONTEXT_CHANGE_STYLE    127    // Cambia stile dal contestuale
#define IDM_CONTEXT_VIEW_MINISTERO  128    // Vista rapida Ministero
#define IDM_CONTEXT_VIEW_ASSOSOFTWARE 129  // Vista rapida Assosoftware
```

### Gestione Eventi (FatturaView.cpp)

#### WM_CONTEXTMENU Handler
- Intercetta il click destro sulla ListBox
- Verifica che ci siano fatture selezionate
- Crea dinamicamente il menu in base al numero di selezioni
- Mostra il menu nella posizione del cursore
- Invia il comando WM_COMMAND per l'azione selezionata

#### WM_COMMAND Handler
Gestisce i nuovi comandi:
- `IDM_CONTEXT_PRINT_SELECTED` → Chiama `OnPrintToPDF()`
- `IDM_CONTEXT_CHANGE_STYLE` → Apre dialogo selezione stile
- `IDM_CONTEXT_VIEW_MINISTERO` → Chiama `OnApplySpecificStylesheet(, "Ministero")`
- `IDM_CONTEXT_VIEW_ASSOSOFTWARE` → Chiama `OnApplySpecificStylesheet(, "Asso")`

### Funzioni Utilizzate
- `CreatePopupMenu()` - Crea il menu dinamico
- `AppendMenuW()` - Aggiunge voci al menu
- `TrackPopupMenu()` - Mostra il menu e ottiene la selezione
- `GetSelectedFattureIndices()` - Ottiene gli indici delle fatture selezionate
- `DestroyMenu()` - Libera le risorse del menu

## Vantaggi

1. **Accesso Rapido**: Non serve navigare nei menu principali
2. **Contestuale**: Il menu si adatta alla selezione (singola o multipla)
3. **Efficienza**: Riduce i click necessari per operazioni comuni
4. **UX Migliorata**: Comportamento standard Windows che gli utenti conoscono
5. **Flessibilità**: Visualizzazione rapida con fogli diversi senza dialoghi

## Integrazione con Notifiche

Il menu contestuale è integrato con il sistema di notifiche popup:
- Quando si applica un foglio di stile, appare la notifica "Applicazione foglio di stile..."
- Al completamento: "Fattura visualizzata correttamente" (success)
- In caso di errore: notifica rossa con messaggio di errore

## Test Consigliati

1. **Singola fattura**:
   - Seleziona una fattura → Click destro
   - Prova "Visualizza con foglio Ministero"
   - Prova "Visualizza con foglio Assosoftware"
   - Prova "Stampa fattura"
   
2. **Multiple fatture**:
   - Seleziona 3-5 fatture (Ctrl+Click)
   - Click destro → Verifica testo "Stampa X fatture"
   - Prova stampa multipla
   - Prova cambio stile per tutte

3. **Nessuna selezione**:
   - Deseleziona tutto → Click destro
   - Verifica che il menu NON appaia

4. **Aree diverse**:
   - Click destro su altre aree (browser, toolbar)
   - Verifica che il menu appaia solo sulla ListBox

## Note Implementative

- Il menu viene creato dinamicamente ad ogni click destro
- Le risorse vengono liberate automaticamente con `DestroyMenu()`
- Il menu utilizza `TPM_RETURNCMD` per ottenere direttamente il comando selezionato
- Se l'utente clicca fuori dal menu, questo si chiude senza azione
- Il testo "Stampa X fatture" viene generato dinamicamente con `std::to_wstring()`

## File Modificati

- **Resource.h**: Aggiunti 4 nuovi ID comando (126-129)
- **FatturaView.cpp**: 
  - Aggiunto handler `WM_CONTEXTMENU` (~60 righe)
  - Aggiunti 4 case in `WM_COMMAND` (~40 righe)

## Compatibilità

- ✅ Windows 7 e successivi
- ✅ Supporto multi-selezione con Ctrl/Shift
- ✅ Keyboard friendly (Shift+F10 o tasto Menu)
- ✅ Integrato con sistema notifiche esistente
