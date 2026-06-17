@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo   Gestionale Magazzino Abbigliamento
echo ==========================================
echo.

echo Controllo Python...
py --version >nul 2>&1
if %errorlevel%==0 (
    set PYTHON_CMD=py
) else (
    python --version >nul 2>&1
    if %errorlevel%==0 (
        set PYTHON_CMD=python
    ) else (
        echo ERRORE: Python non risulta installato o non e' nel PATH.
        echo Installa Python da https://www.python.org/downloads/ e spunta "Add Python to PATH".
        pause
        exit /b 1
    )
)

echo Python trovato.
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creo ambiente virtuale...
    %PYTHON_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo ERRORE: impossibile creare l'ambiente virtuale.
        pause
        exit /b 1
    )
)

echo Attivo ambiente virtuale...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERRORE: impossibile attivare l'ambiente virtuale.
    pause
    exit /b 1
)

echo Installo/Aggiorno librerie necessarie...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERRORE: installazione librerie non riuscita.
    pause
    exit /b 1
)

echo.
echo Avvio applicazione...
echo Quando compare l'indirizzo, apri il browser su:
echo http://127.0.0.1:5000
echo.
echo Per chiudere il gestionale premi CTRL+C in questa finestra.
echo.
python app.py
pause
