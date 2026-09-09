@echo off
setlocal
REM ============================================================
REM  Orario Musicale - installazione e avvio su Windows
REM  Doppio clic su questo file. La prima volta prepara tutto,
REM  le volte dopo avvia direttamente il programma.
REM ============================================================
cd /d "%~dp0"
title Orario Musicale

set PY=
where py >nul 2>nul
if not errorlevel 1 set PY=py -3
if defined PY goto pythontrovato
where python >nul 2>nul
if not errorlevel 1 set PY=python
:pythontrovato
if not defined PY goto senzapython

if not exist ".venv\Scripts\python.exe" goto prepara
goto avvia

:prepara
echo.
echo  Preparo l'ambiente: ci vogliono alcuni minuti, solo la prima volta...
echo.
%PY% -m venv .venv
if errorlevel 1 goto errore
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto errore
echo.
echo  Ambiente pronto.
echo.

:avvia
echo  Avvio del programma...
.venv\Scripts\python.exe app.py
if errorlevel 1 goto errore
exit /b 0

:senzapython
echo.
echo  Python non risulta installato su questo computer.
echo  Scaricalo da:  https://www.python.org/downloads/windows/
echo  Durante l'installazione spunta "Add python.exe to PATH".
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
