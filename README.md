# mapSocietà 

## Panoramica
un’utility da riga di comando (CLI) scritta in Python che automatizza il mapping di codici tra due fogli Excel (formati .xlsx, .xls).  
Il tool è stato sviluppato come supporto per un progetto interno, dove è necessario mappare un foglio dati principale con informazioni contenute in un foglio di lookup (tabella di corrispondenza).

## Funzionamento

### Input – L’utente specifica:
- Il tipo di mapping da effettuare : società o motivazioni (a seconda di questa scelta partirà uno script diverso)
- Il percorso del file Excel sorgente (foglio A) e il nome del foglio.
- Il percorso del file Excel di lookup (foglio B) e il nome del foglio.
- Il nome della colonna chiave da utilizzare per la corrispondenza (presente in entrambi i fogli).

### Caricamento
- i fogli vengono caricati tramite `_load_sheet()` che utilizz calamine per la lettura dei dati con performance ottimali 

### Estrazione  
- Il tool legge dal foglio A la colonna chiave e memorizza i relativi valori. Tramite la funzione 

### Lookup 
- Legge dal foglio B la stessa colonna chiave (dove sono contenuti i nomi) e i loro codici associati

### Matching 
- Per ogni riga del foglio A, cerca una corrispondenza esatta con i valori della colonna chiave nel foglio B. Le parole vengono `trasformate` in minuscolo e `strippate` dei loro spazi cosi da permettere il corretto matching senza errori 

### Output 
- Quando trova una corrispondenza, aggiunge due nuove colonne nel foglio A con il nome specificato nell'input, altrimenti le aggiungerà alla costante specificata di default.
Queste colonne verranno popolate con i record dei valori associati ad ogni corrispondenza trovata nel foglio B

Il risultato viene salvato sovrascrivendo solamente il foglio specificato, lasciando inalterato il resto dei fogli presenti.

## Adattabilità 

### il tool è semplicemente modificabile ed adattabile alle proprie esigenze:

1 - Modifica del nome delle colonne del foglio B di lookup :
    modificando i valori "Key e Value" si può scegliere quale colonna selezionare dal foglio B
    
<img width="477" height="132" alt="image" src="https://github.com/user-attachments/assets/016ffb14-6d43-49f7-9172-eca67d0f0f2f" />

2 - Modifica del nome della colonna del foglio A: 
    modificando i valori "societa e societa_conferente" si può decidere quale colonna prendere dal foglio A 
<img width="1248" height="97" alt="image" src="https://github.com/user-attachments/assets/26d55e6e-3b26-4b66-9597-008823e311b9" />


la stessa logica si applica alle motivazioni :

1 - Modifica del nome delle colonne del foglio B di lookup:
n
<img width="597" height="171" alt="image" src="https://github.com/user-attachments/assets/236de0d8-5413-4d3b-916b-ae240f361f37" />

2 - Modifica del nome della colonna del foglio A:
<img width="1022" height="90" alt="image" src="https://github.com/user-attachments/assets/01fd3e7d-b406-43ac-a79b-1d5b8687d46e" />

