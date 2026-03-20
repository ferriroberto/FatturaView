<?xml version="1.0" encoding="UTF-8"?>
<!--
  FOGLIO DI STILE ESEMPIO PER FATTURA ELETTRONICA
  Questo è un esempio semplificato. 
  Scaricare i fogli di stile ufficiali da:
  - Ministero: https://www.fatturapa.gov.it/
  - Assosoftware: https://github.com/assosoftware/
-->
<xsl:stylesheet version="1.0" 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:fe="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2">
    
    <xsl:output method="html" encoding="UTF-8" indent="yes"/>
    
    <xsl:template match="/">
        <html>
            <head>
                <title>Fattura Elettronica</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    h1 { color: #003366; }
                    table { border-collapse: collapse; width: 100%; margin: 20px 0; }
                    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                    th { background-color: #003366; color: white; }
                    .section { margin: 20px 0; padding: 10px; background-color: #f5f5f5; }
                </style>
            </head>
            <body>
                <h1>FATTURA ELETTRONICA</h1>
                
                <div class="section">
                    <h2>Dati Cedente/Prestatore</h2>
                    <p><strong>Denominazione:</strong> 
                        <xsl:value-of select="//CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione"/>
                    </p>
                    <p><strong>Partita IVA:</strong> 
                        <xsl:value-of select="//CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdCodice"/>
                    </p>
                    <p><strong>Indirizzo:</strong> 
                        <xsl:value-of select="//CedentePrestatore/Sede/Indirizzo"/>
                        <xsl:text> - </xsl:text>
                        <xsl:value-of select="//CedentePrestatore/Sede/CAP"/>
                        <xsl:text> </xsl:text>
                        <xsl:value-of select="//CedentePrestatore/Sede/Comune"/>
                        <xsl:text> (</xsl:text>
                        <xsl:value-of select="//CedentePrestatore/Sede/Provincia"/>
                        <xsl:text>)</xsl:text>
                    </p>
                </div>
                
                <div class="section">
                    <h2>Dati Cessionario/Committente</h2>
                    <p><strong>Denominazione:</strong> 
                        <xsl:value-of select="//CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione"/>
                    </p>
                    <p><strong>Partita IVA:</strong> 
                        <xsl:value-of select="//CessionarioCommittente/DatiAnagrafici/IdFiscaleIVA/IdCodice"/>
                    </p>
                </div>
                
                <div class="section">
                    <h2>Dati Generali</h2>
                    <p><strong>Numero:</strong> 
                        <xsl:value-of select="//DatiGeneraliDocumento/Numero"/>
                    </p>
                    <p><strong>Data:</strong> 
                        <xsl:value-of select="//DatiGeneraliDocumento/Data"/>
                    </p>
                    <p><strong>Tipo Documento:</strong> 
                        <xsl:value-of select="//DatiGeneraliDocumento/TipoDocumento"/>
                    </p>
                </div>
                
                <h2>Dettaglio Linee</h2>
                <table>
                    <tr>
                        <th>N.</th>
                        <th>Descrizione</th>
                        <th>Quantità</th>
                        <th>Prezzo Unit.</th>
                        <th>Aliquota IVA</th>
                        <th>Totale</th>
                    </tr>
                    <xsl:for-each select="//DettaglioLinee">
                        <tr>
                            <td><xsl:value-of select="NumeroLinea"/></td>
                            <td><xsl:value-of select="Descrizione"/></td>
                            <td><xsl:value-of select="Quantita"/></td>
                            <td><xsl:value-of select="PrezzoUnitario"/></td>
                            <td><xsl:value-of select="AliquotaIVA"/>%</td>
                            <td><xsl:value-of select="PrezzoTotale"/></td>
                        </tr>
                    </xsl:for-each>
                </table>
                
                <div class="section">
                    <h2>Riepilogo IVA</h2>
                    <table>
                        <tr>
                            <th>Aliquota</th>
                            <th>Imponibile</th>
                            <th>Imposta</th>
                        </tr>
                        <xsl:for-each select="//DatiRiepilogo">
                            <tr>
                                <td><xsl:value-of select="AliquotaIVA"/>%</td>
                                <td><xsl:value-of select="ImponibileImporto"/></td>
                                <td><xsl:value-of select="Imposta"/></td>
                            </tr>
                        </xsl:for-each>
                    </table>
                </div>
                
                <div class="section">
                    <h2>Totale Documento</h2>
                    <p><strong>Importo Totale Documento:</strong> 
                        <xsl:value-of select="//ImportoTotaleDo"/>
                    </p>
                </div>
            </body>
        </html>
    </xsl:template>
    
</xsl:stylesheet>
