# Importazione carico da Excel

Questa versione aggiunge la funzione **Carico magazzino > Importa carico da Excel**.

## Formato file Excel

Il file deve essere `.xlsx` e deve avere una riga di intestazione con queste colonne consigliate:

- `codice_prodotto`
- `nome_prodotto`
- `descrizione`
- `taglia`
- `quantita`
- `prezzo_acquisto`
- `prezzo_vendita`
- `soglia_minima`
- `note`

Colonne obbligatorie:

- `codice_prodotto`
- `nome_prodotto`
- `taglia`
- `quantita`

## Logica di importazione

- Ogni riga Excel corrisponde a una variante/taglia.
- Se `codice_prodotto + taglia` esiste già, il gestionale registra un carico su quell'articolo.
- Se `codice_prodotto + taglia` non esiste, il gestionale crea l'articolo e poi registra il carico.
- L'operatore viene preso automaticamente dall'utente loggato.
- Le giacenze si aggiornano automaticamente tramite i movimenti di carico.
- Se ci sono errori nel file, l'importazione non viene eseguita e viene mostrato l'elenco delle righe da correggere.

## Esempio

| codice_prodotto | nome_prodotto | descrizione | taglia | quantita | prezzo_acquisto | prezzo_vendita | soglia_minima | note |
|---|---|---|---|---:|---:|---:|---:|---|
| TSH001 | T-shirt bianca | T-shirt cotone | S | 10 | 8 | 19.90 | 5 | primo carico |
| TSH001 | T-shirt bianca | T-shirt cotone | M | 15 | 8 | 19.90 | 5 | primo carico |
