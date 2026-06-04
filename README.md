# mappingXLSX – Tool CLI per mapping su fogli Excel

## Panoramica
un’utility da riga di comando (CLI) scritta in Python che automatizza il mapping di codici tra due fogli Excel (formati .xlsx, .xls).  
Il tool è stato sviluppato come supporto per un progetto interno, dove è necessario arricchire un foglio dati principale con informazioni contenute in un foglio di lookup (tabella di corrispondenza).

## Come funziona (step logici)

### Input – L’utente specifica:
- Il percorso del file Excel sorgente (foglio A) e il nome del foglio.
- Il percorso del file Excel di lookup (foglio B) e il nome del foglio.
- Il nome della colonna chiave da utilizzare per la corrispondenza (presente in entrambi i fogli).

### Estrazione – 
Il tool legge dal foglio A la colonna chiave e memorizza i relativi valori.

### Lookup – 
Legge dal foglio B la stessa colonna chiave e una o più colonne contenenti i codici da mappare (es. codice prodotto, ID anagrafica, etc.).

### Matching – 
Per ogni riga del foglio A, cerca una corrispondenza esatta (case‑sensitive o insensitive, a configurazione) con i valori della colonna chiave nel foglio B.

### Output – 
Quando trova una corrispondenza, aggiunge due nuove colonne nel foglio A:
- Colonna 1: il valore originale della colonna chiave (o un suo derivato).
- Colonna 2: il codice mappato ottenuto dal foglio B.
- (Opzionale) può aggiungere anche altre colonne di dettaglio presenti nel foglio B.

Il risultato viene salvato sovrascrivendo solamente il foglio specificato, lasciando inalterato il resto dei fogli presenti.
