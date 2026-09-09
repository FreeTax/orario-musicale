@echo off
setlocal
REM ============================================================
REM  Orario Musicale - crea l'installer per Windows (setup.exe)
REM  Da lanciare su un PC Windows. Fa tutto: ambiente, programma
REM  e installer. Risultato in:
REM     installer_output\OrarioMusicale-setup-1.0.exe
REM ============================================================
cd /d "%~dp0"
title Orario Musicale - creazione installer

set PY=
where py >nul 2>nul
if not errorlevel 1 set PY=py -3
if defined PY goto pythontrovato
where python >nul 2>nul
if not errorlevel 1 set PY=python
:pythontrovato
if not defined PY goto senzapython

REM ---- 1. ambiente Python ----
if not exist ".venv\Scripts\python.exe" goto prepara
goto cercainno

:prepara
echo.
echo  [1/3] Preparo l'ambiente Python: alcuni minuti...
echo.
%PY% -m venv .venv
if errorlevel 1 goto errore
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto errore

REM ---- 2. Inno Setup (il programma che crea l'installer) ----
:cercainno
set ISCC=
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe
if defined ISCC goto costruisci

echo.
echo  Inno Setup non risulta installato: provo a installarlo automaticamente...
echo.
where winget >nul 2>nul
if errorlevel 1 goto senzainno
winget install --id JRSoftware.InnoSetup --accept-source-agreements --accept-package-agreements
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe
if not defined ISCC goto senzainno

REM ---- 3. programma + installer ----
:costruisci
echo.
echo  [2/3] Creo il programma: alcuni minuti...
echo.
.venv\Scripts\python.exe -m pip install -q pyinstaller
if errorlevel 1 goto errore
.venv\Scripts\python.exe -m PyInstaller --noconfirm --windowed --name "Orario Musicale" ^
  --collect-all customtkinter --collect-all tksheet --collect-all ortools --collect-all tzdata ^
  --hidden-import orario.motore --hidden-import orario.export --hidden-import orario.export_excel ^
  --hidden-import orario.export_pdf --hidden-import orario.export_comune ^
  app.py
if errorlevel 1 goto errore

echo.
echo  [3/3] Creo l'installer...
echo.
"%ISCC%" "installer\installer.iss"
if errorlevel 1 goto errore

echo.
echo  ============================================================
echo   Fatto. L'installer da consegnare e' qui:
echo.
echo      installer_output\OrarioMusicale-setup-1.0.exe
echo.
echo   E' un unico file: si manda per email (se passa, pesa circa
echo   100 MB), su chiavetta o con un link di condivisione.
echo   Chi lo riceve fa doppio clic e segue le schermate: non
echo   serve Python ne' i diritti di amministratore.
echo  ============================================================
echo.
pause
exit /b 0

:senzapython
echo.
echo  Python non risulta installato su questo computer.
echo  Scaricalo da:  https://www.python.org/downloads/windows/
echo  Durante l'installazione spunta "Add python.exe to PATH".
echo.
pause
exit /b 1

:senzainno
echo.
echo  Non riesco a installare Inno Setup da solo.
echo  Scaricalo a mano (e' gratuito) da:  https://jrsoftware.org/isdl.php
echo  Installalo, poi rilancia questo file.
echo.
echo  In alternativa usa costruisci_app.bat, che crea il programma
echo  senza installer (va copiata tutta la cartella dist).
echo.
pause
exit /b 1

:errore
echo.
echo  Qualcosa non ha funzionato. Copia il messaggio qui sopra
echo  e mandalo a chi ti ha dato il programma.
echo.
pause
exit /b 1
