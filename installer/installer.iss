; Installer per Windows di "Orario Musicale" (Inno Setup 6).
; Non si lancia a mano: lo compila costruisci_installer.bat, che prima crea il programma con PyInstaller.
; Risultato: installer_output\OrarioMusicale-setup-<versione>.exe

#define NomeApp "Orario Musicale"
#define Versione "1.0"
#define EseguibileApp "Orario Musicale.exe"
#define Autore "Liceo Musicale"

[Setup]
; AppId identifica il programma per aggiornamenti e disinstallazione: non cambiarlo mai.
AppId={{1A86EF0A-48CC-414C-ACD0-995285F89076}
AppName={#NomeApp}
AppVersion={#Versione}
AppVerName={#NomeApp} {#Versione}
AppPublisher={#Autore}
DefaultDirName={autopf}\Orario Musicale
DefaultGroupName=Orario Musicale
DisableProgramGroupPage=yes
; Installazione per il solo utente corrente: nessuna richiesta di password di amministratore.
PrivilegesRequired=lowest
OutputDir=..\installer_output
OutputBaseFilename=OrarioMusicale-setup-{#Versione}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#EseguibileApp}
UninstallDisplayName={#NomeApp}
; Serve Windows 10 o successivo
MinVersion=10.0

[Languages]
Name: "italiano"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Tutto il contenuto della cartella creata da PyInstaller
Source: "..\dist\Orario Musicale\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; File di esempio e istruzioni, copiati nei Documenti dell'utente (senza sovrascrivere)
Source: "..\esempio_orario.xlsx"; DestDir: "{userdocs}\Orario Musicale"; Flags: onlyifdoesntexist skipifsourcedoesntexist
Source: "..\WINDOWS.md"; DestDir: "{app}"; DestName: "Istruzioni Windows.txt"; Flags: ignoreversion

[Icons]
Name: "{group}\{#NomeApp}"; Filename: "{app}\{#EseguibileApp}"
Name: "{group}\File di esempio"; Filename: "{userdocs}\Orario Musicale"
Name: "{group}\{cm:UninstallProgram,{#NomeApp}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#NomeApp}"; Filename: "{app}\{#EseguibileApp}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#EseguibileApp}"; Description: "{cm:LaunchProgram,{#StringChange(NomeApp, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Messages]
italiano.WelcomeLabel2=Verrà installato [name/ver] su questo computer.%n%nIl programma calcola l'orario pomeridiano del liceo musicale a partire da un file Excel con studenti, docenti e gruppi, e produce l'orario in Excel e PDF.%n%nUn file di esempio verrà messo in Documenti\Orario Musicale.
