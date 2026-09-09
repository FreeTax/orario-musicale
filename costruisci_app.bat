@echo off
setlocal
REM ============================================================
REM  Orario Musicale - crea il programma con doppio clic (.exe)
REM  Da lanciare UNA VOLTA su un PC Windows. Il risultato sta in
REM  dist\Orario Musicale\Orario Musicale.exe
REM ============================================================
cd /d "%~dp0"
title Orario Musicale - costruzione

set PY=
where py >nul 2>nul
if not errorlevel 1 set PY=py -3
if defined PY goto pythontrovato
where python >nul 2>nul
if not errorlevel 1 set PY=python
:pythontrovato
if not defined PY goto senzapython

if not exist ".venv\Scripts\python.exe" goto prepara
goto costruisci

:prepara
echo  Preparo l'ambiente: alcuni minuti...
%PY% -m venv .venv
if errorlevel 1 goto errore
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto errore

:costruisci
.venv\Scripts\python.exe -m pip install -q pyinstaller
if errorlevel 1 goto errore
echo.
echo  Costruzione in corso: alcuni minuti...
echo.
.venv\Scripts\python.exe -m PyInstaller --noconfirm --windowed --name "Orario Musicale" ^
  --collect-all customtkinter --collect-all tksheet --collect-all ortools --collect-all tzdata ^
  --hidden-import orario.motore --hidden-import orario.export --hidden-import orario.export_excel ^
  --hidden-import orario.export_pdf --hidden-import orario.export_comune ^
  app.py
if errorlevel 1 goto errore

echo.
echo  Fatto. Il programma e' qui:
echo      dist\Orario Musicale\Orario Musicale.exe
echo.
echo  Puoi copiare tutta la cartella "dist\Orario Musicale" dove vuoi,
echo  anche su una chiavetta: serve tutta, non solo il file .exe.
echo  Per comodita' fai un collegamento sul desktop (tasto destro sul
echo  file .exe, "Invia a", "Desktop (crea collegamento)").
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

:errore
echo.
echo  Qualcosa non ha funzionato. Copia il messaggio qui sopra
echo  e mandalo a chi ti ha dato il programma.
echo.
pause
exit /b 1
