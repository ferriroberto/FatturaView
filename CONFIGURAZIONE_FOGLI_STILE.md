# Configurazione Fogli di Stile

## Posizionamento dei File XSLT

L'applicazione cerca automaticamente i fogli di stile nelle seguenti posizioni (in ordine):

1. **Cartella dell'eseguibile**: `xmlRead.exe` nella stessa directory
2. **Sottocartella Resources**: `Resources\FoglioStileMinistero.xsl`
3. **Sottocartella Stylesheets**: `Stylesheets\FoglioStileMinistero.xsl`
4. **Cartella temporanea**: (se copiati lì durante l'estrazione)

## Nomi dei File

I file devono essere nominati esattamente come segue:

- **Ministero**: `FoglioStileMinistero.xsl`
- **Assosoftware**: `FoglioStileAssoSoftware.xsl` (nota: AssoSoftware con la S maiuscola)

⚠️ **Attenzione**: Il nome è case-sensitive! Usa esattamente `AssoSoftware` (non `Assosoftware`).

## Esempio Struttura Cartelle

```
xmlRead.exe
├── FoglioStileMinistero.xsl        ← Opzione 1 (consigliata)
├── FoglioStileAssoSoftware.xsl     ← (nota: AssoSoftware con S maiuscola!)
└── Resources\                       ← Opzione 2 (già presente nel progetto)
    ├── FoglioStileMinistero.xsl
    └── FoglioStileAssoSoftware.xsl
```

## Dove Trovare i Fogli di Stile

### Foglio di Stile Ministero (Agenzia delle Entrate)
- **Sito ufficiale**: https://www.agenziaentrate.gov.it/portale/web/guest/schede/comunicazioni/fattura-elettronica
- **Nome file**: Cerca "FoglioStileAssoSoftware.xsl" o "fatturaordinaria_v1.2.1.xsl"
- **Versione**: Usa la versione più recente compatibile con FatturaPA 1.2+

### Foglio di Stile Assosoftware
- **Sito ufficiale**: http://www.assosoftware.it/fatturazione-elettronica
- **Nome file**: Cerca "FoglioStileAssoSoftware.xsl"
- **Caratteristiche**: Layout più compatto e leggibile

## Menu Impostazioni

Usa il menu **Impostazioni** per scegliere quale foglio di stile usare di default:

1. **Foglio di stile Ministero** ✓ (predefinito)
2. **Foglio di stile Assosoftware**
3. **Foglio di stile personalizzato...** (seleziona un file .xsl/.xslt a piacere)

La preferenza viene salvata automaticamente nel **Registry di Windows**:
```
HKEY_CURRENT_USER\Software\FatturaElettronicaViewer
├── StylesheetPreference (DWORD)  → 0=Ministero, 1=Assosoftware, 2=Custom
└── CustomStylesheetPath (STRING) → Percorso personalizzato (se tipo=2)
```

## Primo Avvio

Al primo avvio, se nessun foglio di stile è trovato:
1. Viene mostrato un messaggio informativo
2. Si apre la finestra di selezione file
3. Il file selezionato viene salvato come "Personalizzato"
4. Puoi cambiarlo in seguito dal menu **Impostazioni**

## Suggerimenti

✅ **Consigliato**: Lascia i fogli di stile nella cartella `Resources\` fornita con l'applicazione
✅ **Rapido**: L'applicazione carica automaticamente Assosoftware al primo avvio
✅ **Flessibile**: Aggiungi nuovi file .xsl nella cartella Resources\ e appariranno automaticamente nella lista
✅ **Zero configurazione**: Nessun file da selezionare manualmente, tutto automatico!

⚠️ **Importante**: Tutti i file devono essere nella cartella `Resources\`
⚠️ **Compatibilità**: Assicurati che i file .xsl siano compatibili con FatturaPA v1.2+

## Aggiungere Nuovi Fogli di Stile

1. Scarica il file `.xsl` desiderato
2. Copialo nella cartella `Resources\`
3. Riavvia l'applicazione (o vai su Impostazioni → Seleziona foglio di stile)
4. Il nuovo file apparirà automaticamente nella lista!

Nessuna configurazione aggiuntiva richiesta.
