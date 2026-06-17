# Gestionale Magazzino Abbigliamento

Applicazione web Flask per la gestione online di un magazzino di prodotti di abbigliamento.

## Funzioni incluse

- Login utenti con ruoli: amministratore e operatore.
- Schede articolo con codice prodotto, nome, descrizione, taglia, prezzo di acquisto, prezzo di vendita.
- Vincolo univoco su codice prodotto + taglia.
- Carico magazzino con aggiornamento automatico giacenze.
- Scarico magazzino con controllo disponibilità e blocco se la giacenza è insufficiente.
- Destinatario e operatore salvati nelle schede di scarico.
- Storico completo di carichi e scarichi con filtri.
- Stampa e salvataggio PDF tramite funzione di stampa del browser.
- Esportazione giacenze in CSV ed Excel.
- Report stampabile delle giacenze con valorizzazione economica.
- Dashboard con totali, valori economici e articoli sotto soglia minima.

## Avvio in locale

1. Entra nella cartella del progetto.

```bash
cd gestionale_magazzino
```

2. Crea un ambiente virtuale.

```bash
python -m venv .venv
```

3. Attiva l'ambiente virtuale.

Su Windows:

```bash
.venv\Scripts\activate
```

Su Mac/Linux:

```bash
source .venv/bin/activate
```

4. Installa le dipendenze.

```bash
pip install -r requirements.txt
```

5. Avvia l'app.

```bash
python app.py
```

6. Apri il browser su:

```text
http://127.0.0.1:5000
```

## Account demo

Amministratore:

```text
Email: admin@magazzino.it
Password: admin123
```

Operatore:

```text
Email: operatore@magazzino.it
Password: operatore123
```

## Pubblicazione online

Per uso reale online, modifica obbligatoriamente:

- `SECRET_KEY`, impostandola come variabile d'ambiente.
- Password degli utenti demo.
- Database: SQLite va bene per test e prototipo; per produzione è consigliabile migrare a PostgreSQL.
- Configurazione HTTPS sul server.

## Nota PDF

Le liste di scarico e le giacenze sono pagine stampabili. Dal browser puoi scegliere `Stampa` e poi `Salva come PDF`.


## Supabase

Questa versione supporta Supabase/PostgreSQL. Imposta la variabile ambiente `DATABASE_URL` con la stringa di connessione Supabase. Se `DATABASE_URL` non è impostata, l'app usa SQLite locale.
