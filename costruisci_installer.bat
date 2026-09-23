@echo off
setlocal
REM ============================================================
REM  Orario Musicale - crea l'installer per Windows (setup.exe)
REM  Da lanciare su un PC Windows. Fa tutto: ambiente, programma
REM  e installer. Risultato in:
REM     installer_output\<nome_installer>.exe
REM  Nome, versione e icona si cambiano in impostazioni_installer.txt.
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

REM ---- impostazioni: nome, versione, icona (da impostazioni_installer.txt) ----
set IMPOSTAZIONI_OK=
for /f "usebackq delims=" %%l in (`%PY% leggi_impostazioni.py bat`) do %%l
if not defined IMPOSTAZIONI_OK goto erroreimpostazioni
set OPZ_ICONA=
if defined ICONA set OPZ_ICONA=--icon "%ICONA%"
echo  Programma: %NOME%  (versione %VERSIONE%)
if defined ICONA echo  Icona: %ICONA%
if not defined ICONA echo  Icona: predefinita

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
call :trovainno
if defined ISCC goto costruisci

echo.
echo  Inno Setup non risulta installato: provo a installarlo automaticamente...
echo.
where winget >nul 2>nul
if errorlevel 1 goto senzainno
winget install --id JRSoftware.InnoSetup --accept-source-agreements --accept-package-agreements
call :trovainno
if not defined ISCC goto senzainno

REM ---- 3. programma + installer ----
:costruisci
echo.
echo  [2/3] Creo il programma: alcuni minuti...
echo.
.venv\Scripts\python.exe -m pip install -q pyinstaller
if errorlevel 1 goto errore
.venv\Scripts\python.exe -m PyInstaller --noconfirm --windowed --name "%NOME%" %OPZ_ICONA% ^
  --collect-all customtkinter --collect-all tksheet --collect-all ortools --collect-all tzdata ^
  --hidden-import orario.motore --hidden-import orario.export --hidden-import orario.export_excel ^
  --hidden-import orario.export_pdf --hidden-import orario.export_comune ^
  app.py
if errorlevel 1 goto errore

echo.
echo  [3/3] Creo l'installer con %ISCC%
echo.
"%ISCC%" "installer\installer.iss"
if errorlevel 1 goto errore

echo.
echo  ============================================================
echo   Fatto. L'installer da consegnare e' qui:
echo.
echo      installer_output\%NOME_INSTALLER%.exe
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

REM Cerca ISCC.exe di qualsiasi versione di Inno Setup (6, 7, ...), nelle cartelle
REM di installazione consuete e nel PATH. Se ce n'e' piu' d'una prende la piu' recente.
:trovainno
set ISCC=
for %%d in ("%ProgramFiles(x86)%" "%ProgramFiles%" "%LOCALAPPDATA%\Programs") do (
  if not defined ISCC for /d %%v in ("%%~d\Inno Setup*") do if exist "%%~v\ISCC.exe" set "ISCC=%%~v\ISCC.exe"
)
if not defined ISCC for /f "delims=" %%p in ('where ISCC.exe 2^>nul') do if not defined ISCC set "ISCC=%%p"
exit /b 0

:senzainno
echo.
echo  Non trovo Inno Setup su questo computer e non riesco a installarlo da solo.
echo  Se e' gia' installato: ho cercato ISCC.exe in "Programmi", "Programmi (x86)"
echo  e "%LOCALAPPDATA%\Programs", nelle cartelle "Inno Setup ...". Se sta altrove,
echo  aggiungi la sua cartella al PATH oppure reinstallalo nella cartella proposta.
echo  Se non e' installato, scaricalo (e' gratuito) da:  https://jrsoftware.org/isdl.php
echo  Installalo, poi rilancia questo file.
echo.
echo  In alternativa usa costruisci_app.bat, che crea il programma
echo  senza installer (va copiata tutta la cartella dist).
echo.
pause
exit /b 1

:erroreimpostazioni
echo.
echo  Non riesco a leggere impostazioni_installer.txt (vedi il messaggio qui sopra).
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
