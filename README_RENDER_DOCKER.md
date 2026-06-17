# Deploy corretto su Render con Docker

Questa versione evita i problemi di Python/Gunicorn su Render usando Docker.

## File importanti

- `Dockerfile`: installa Python, requirements e avvia Gunicorn.
- `render.yaml`: indica a Render di usare Docker.
- `requirements.txt`: dipendenze Python.
- `app.py`: applicazione Flask.

## Passaggi rapidi

1. Carica su GitHub i file di questa cartella, direttamente nella root della repository.
   La prima pagina della repo deve mostrare subito `Dockerfile`, `app.py`, `requirements.txt`, `templates`, `static`.

2. Su Render crea un NUOVO servizio:
   `New` > `Web Service` > scegli la repository GitHub.

3. Quando Render chiede il runtime/linguaggio, scegli `Docker`.
   Non inserire Build Command e non inserire Start Command.

4. In Environment aggiungi:
   - `DATABASE_URL` = stringa PostgreSQL Supabase
   - `SECRET_KEY` = una frase lunga casuale
   - `SESSION_COOKIE_SECURE` = `1`

5. Avvia il deploy.

## Nota

Non usare più i comandi:

- `.venv/bin/gunicorn`
- `python3 -m gunicorn`
- `python3 app.py`
- `python3 -m pip install ...`

Con Docker non servono.
