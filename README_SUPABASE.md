# Gestionale magazzino con Supabase PostgreSQL

Questa versione può usare due database:

- in locale, se non imposti `DATABASE_URL`, usa SQLite (`warehouse.db`);
- online, se imposti `DATABASE_URL`, usa Supabase PostgreSQL.

## Variabili ambiente online

Imposta su Render/Railway:

- `DATABASE_URL`: connection string PostgreSQL di Supabase
- `SECRET_KEY`: chiave segreta Flask, può essere generata automaticamente su Render
- `SESSION_COOKIE_SECURE`: `1`

## Avvio locale

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Poi apri `http://127.0.0.1:5000`.

## Account iniziali

Admin:
- `admin@magazzino.it`
- `admin123`

Operatore:
- `operatore@magazzino.it`
- `operatore123`

Dopo il primo accesso, crea utenti reali e cambia le password demo.
