# Spese app

## Avvio in locale
    pip install -r requirements.txt
    flask --app app crea-utente IL_TUO_NOME     # chiede la password
    flask --app app run --debug
Apri http://127.0.0.1:5000, accedi, poi Impostazioni > Importa da Excel.

## Online
Variabili d'ambiente: `SECRET_KEY` (stringa lunga casuale) e `DATABASE_URL` (PostgreSQL, es. Neon o Supabase).
Comando di avvio: `gunicorn app:app`
