# Guida all'uso del Lettore di Fatture Elettroniche

## Panoramica
Questo programma Windows consente di importare, leggere e visualizzare fatture elettroniche italiane in formato XML (FatturaPA) applicando i fogli di stile XSLT ufficiali.

## Caratteristiche principali

### 1. Estrazione da archivi ZIP
- Supporto nativo per file ZIP
- Estrazione automatica in cartella temporanea
- Gestione di archivi multipli

### 2. Parsing XML
- Utilizza MSXML 6.0 per il parsing
- Supporto completo dello standard FatturaPA
- Validazione automatica del formato

### 3. Trasformazione XSLT
- Applicazione di fogli di stile personalizzati
- Due modalità di visualizzazione:
  - **Ministero**: Foglio di stile ufficiale dell'Agenzia delle Entrate
  - **Assosoftware**: Foglio di stile della comunità Assosoftware

### 4. Visualizzazione
- Rendering HTML nel browser predefinito
- Layout responsive
- Stampa supportata

## Download dei Fogli di Stile XSLT

### Foglio di stile Ministero dell'Economia
1. Visita il sito ufficiale: https://www.fatturapa.gov.it/
2. Sezione "Strumenti" → "Fogli di Stile"
3. Scarica il file `fatturapa_v1.2.1.xsl` o versione più recente
4. Salva in una cartella facilmente accessibile

### Foglio di stile Assosoftware
1. Visita il repository GitHub: https://github.com/assosoftware/
2. Cerca il progetto relativo ai fogli di stile FatturaPA
3. Scarica il file XSLT più recente
4. Salva nella stessa cartella del foglio ministeriale

## Workflow d'uso

### Passo 1: Avviare l'applicazione
- Eseguire `xmlRead.exe`
- Appare la finestra principale con una lista vuota

### Passo 2: Importare le fatture
1. Menu **File → Apri archivio ZIP...**
2. Selezionare il file ZIP contenente le fatture
3. Attendere l'estrazione (messaggio di conferma)
4. Le fatture appaiono nella lista

### Passo 3: Selezionare una fattura
- Click su una fattura nella lista per selezionarla
- Il nome del file viene evidenziato

### Passo 4: Applicare un foglio di stile
**Opzione A - Foglio Ministero:**
1. Menu **Visualizza → Applica foglio stile Ministero**
2. Selezionare il file `.xsl` del Ministero
3. Il browser si apre con la fattura formattata

**Opzione B - Foglio Assosoftware:**
1. Menu **Visualizza → Applica foglio stile Assosoftware**
2. Selezionare il file `.xsl` di Assosoftware
3. Il browser si apre con la fattura formattata

### Passo 5: Visualizzare e stampare
- La fattura appare nel browser predefinito
- Utilizzare **Ctrl+P** per stampare
- Salvare come PDF se necessario

## Formato delle Fatture Elettroniche

### Standard FatturaPA
Il formato XML deve rispettare lo schema ufficiale:
- Namespace: `http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2`
- Versioni supportate: 1.0, 1.1, 1.2, 1.2.1
- Encoding: UTF-8

### Struttura tipica
```xml
<?xml version="1.0" encoding="UTF-8"?>
<p:FatturaElettronica xmlns:p="...">
  <FatturaElettronicaHeader>
    <DatiTrasmissione>...</DatiTrasmissione>
    <CedentePrestatore>...</CedentePrestatore>
    <CessionarioCommittente>...</CessionarioCommittente>
  </FatturaElettronicaHeader>
  <FatturaElettronicaBody>
    <DatiGenerali>...</DatiGenerali>
    <DatiBeniServizi>...</DatiBeniServizi>
  </FatturaElettronicaBody>
</p:FatturaElettronica>
```

## Risoluzione problemi

### Problema: "Errore durante l'estrazione del file ZIP"
**Soluzione:**
- Verificare che il file ZIP non sia corrotto
- Controllare i permessi di scrittura sulla cartella temporanea
- Eseguire come amministratore se necessario

### Problema: "Errore nel caricamento del file XML"
**Soluzione:**
- Verificare che il file sia un XML valido
- Controllare che rispetti lo schema FatturaPA
- Aprire il file XML in un editor per verificare errori

### Problema: "Errore durante l'applicazione del foglio di stile"
**Soluzione:**
- Verificare che il file XSLT sia valido
- Controllare la compatibilità della versione XSLT
- Provare con il foglio di stile di esempio incluso

### Problema: Il browser non si apre
**Soluzione:**
- Impostare un browser predefinito in Windows
- Verificare le associazioni dei file HTML
- Aprire manualmente il file nella cartella temporanea

## File temporanei

### Posizione
I file vengono estratti in:
```
%TEMP%\FattureElettroniche\
```

### Pulizia
- I file rimangono fino alla chiusura dell'applicazione
- Eliminare manualmente se necessario
- Windows pulisce automaticamente la cartella TEMP periodicamente

## Limitazioni note

1. **Dimensione archivi**: Non testato con ZIP > 100MB
2. **Numero fatture**: Ottimale fino a 50 fatture per ZIP
3. **Browser**: Richiede Internet Explorer 11+ o Edge per alcune funzionalità
4. **XSLT**: Solo versione 1.0 supportata nativamente

## Estensioni future

- [ ] Visualizzazione integrata (controllo browser embedded)
- [ ] Esportazione massiva in PDF
- [ ] Ricerca full-text nelle fatture
- [ ] Validazione schema XML automatica
- [ ] Supporto firma digitale (P7M)
- [ ] Database locale per archiviazione

## Supporto tecnico

Per problemi o suggerimenti:
- Verificare la documentazione ufficiale FatturaPA
- Consultare il forum di Assosoftware
- Controllare gli aggiornamenti del software

## Licenza
Questo software è fornito "così com'è" senza garanzie di alcun tipo.

## Credits
- MSXML 6.0: Microsoft Corporation
- Standard FatturaPA: Agenzia delle Entrate
- Fogli di stile: Ministero dell'Economia e Assosoftware
