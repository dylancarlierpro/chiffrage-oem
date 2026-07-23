; ─────────────────────────────────────────────────────────────
;  Inno Setup — Installeur Chiffrage OEM by CPI Howden
;  Placez ce fichier dans le même dossier que :
;    - dist\Chiffrage OEM.exe   (généré par PyInstaller)
;    - logo_cpi.png
;
;  Numéro de version : passé en ligne de commande via /DMyAppVersion=X.Y
;  (voir build.ps1 -Version). Par défaut "1.0" si non spécifié.
; ─────────────────────────────────────────────────────────────

#ifndef MyAppVersion
  #define MyAppVersion "1.0"
#endif

[Setup]
; Informations générales
AppName=Chiffrage OEM
AppVersion={#MyAppVersion}
AppPublisher=CPI by Howden
AppPublisherURL=https://www.howden.com
AppSupportURL=https://www.howden.com
AppUpdatesURL=https://www.howden.com

; Dossier d'installation par défaut
DefaultDirName={autopf}\CPI Howden\Chiffrage OEM
DefaultGroupName=CPI Howden\Chiffrage OEM

; Fichier de sortie (l'installeur final)
OutputDir=installeur
OutputBaseFilename=Chiffrage_OEM_Setup_v{#MyAppVersion}

; Icône de l'installeur (optionnel — remplacez par votre .ico si disponible)
; SetupIconFile=logo_cpi.ico

; Compression
Compression=lzma2/ultra64
SolidCompression=yes

; Nécessite les droits admin pour installer dans Program Files
PrivilegesRequired=admin

; Afficher un assistant d'installation classique
WizardStyle=modern

; Langue
[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

; ─────────────────────────────────────────────────────────────
;  FICHIERS À INCLURE
; ─────────────────────────────────────────────────────────────
[Files]
; L'exécutable principal (généré par PyInstaller)
Source: "dist\Chiffrage OEM.exe"; DestDir: "{app}"; Flags: ignoreversion

; Le logo (copié à côté de l'exe pour qu'il soit trouvé au lancement)
Source: "logo_cpi.png"; DestDir: "{app}"; Flags: ignoreversion

; Si vous avez déjà un chiffrage_data.json à pré-charger, décommentez :
; Source: "chiffrage_data.json"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

; ─────────────────────────────────────────────────────────────
;  RACCOURCIS
; ─────────────────────────────────────────────────────────────
[Icons]
; Raccourci dans le menu Démarrer
Name: "{group}\Chiffrage OEM"; Filename: "{app}\Chiffrage OEM.exe"

; Raccourci sur le bureau
Name: "{commondesktop}\Chiffrage OEM"; Filename: "{app}\Chiffrage OEM.exe"; Tasks: desktopicon

; ─────────────────────────────────────────────────────────────
;  TÂCHES OPTIONNELLES (proposées à l'utilisateur)
; ─────────────────────────────────────────────────────────────
[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le Bureau"; GroupDescription: "Options :"

; ─────────────────────────────────────────────────────────────
;  LANCEMENT AUTOMATIQUE APRÈS INSTALLATION
; ─────────────────────────────────────────────────────────────
[Run]
Filename: "{app}\Chiffrage OEM.exe"; Description: "Lancer Chiffrage OEM"; Flags: nowait postinstall skipifsilent

; ─────────────────────────────────────────────────────────────
;  DÉSINSTALLATION
; ─────────────────────────────────────────────────────────────
[UninstallDelete]
; Les données (remises négociées + hash admin) ne vivent PAS dans {app}
; (Program Files, non inscriptible par un utilisateur standard) mais dans le
; profil utilisateur — voir _data_dir() dans chiffrage_oem.py, qui stocke
; tout sous %APPDATA%\CPI Howden\Chiffrage OEM. Nettoyer {app} ne servait à
; rien : le vrai fichier (et ses sauvegardes .bak/.tmp) survivait à toute
; désinstallation. Commentez ces lignes si vous voulez au contraire conserver
; les données lors d'une désinstallation.
Type: files; Name: "{userappdata}\CPI Howden\Chiffrage OEM\chiffrage_data.json"
Type: files; Name: "{userappdata}\CPI Howden\Chiffrage OEM\chiffrage_data.json.tmp"
Type: files; Name: "{userappdata}\CPI Howden\Chiffrage OEM\chiffrage_crash.log"
Type: files; Name: "{userappdata}\CPI Howden\Chiffrage OEM\*.corrompu.*.bak"
Type: dirifempty; Name: "{userappdata}\CPI Howden\Chiffrage OEM"
Type: dirifempty; Name: "{userappdata}\CPI Howden"
