# Pubblicazione online con Supabase + Render

Questa versione è predisposta per usare Supabase PostgreSQL tramite variabile `DATABASE_URL`.

## Architettura

Browser utenti → App Flask su Render → Database Supabase PostgreSQL

## Passaggi rapidi

1. Crea un progetto su Supabase.
2. Copia la connection string PostgreSQL da `Connect`.
3. Carica questa cartella su una repository GitHub.
4. Crea un Web Service su Render collegando la repository.
5. Imposta le variabili ambiente:
   - `DATABASE_URL`
   - `SECRET_KEY`
   - `SESSION_COOKIE_SECURE=1`
6. Avvia il deploy.
7. Apri il link pubblico creato da Render.

Al primo avvio l'app crea automaticamente le tabelle e gli utenti demo.
