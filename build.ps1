# ─────────────────────────────────────────────────────────────
#  Build complet — Chiffrage OEM
#  1) Génère l'exécutable avec PyInstaller (dist\Chiffrage OEM.exe)
#  2) Compile l'installeur avec Inno Setup (installeur\Chiffrage_OEM_Setup_v<Version>.exe)
#
#  Usage : clic droit > Exécuter avec PowerShell
#          ou depuis un terminal :  .\build.ps1
#          ou avec un numéro de version :  .\build.ps1 -Version "1.1"
# ─────────────────────────────────────────────────────────────

param(
    [string]$Version = "1.0"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Find-Python {
    $candidates = @(
        "C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python312\python.exe",
        (Get-Command python -ErrorAction SilentlyContinue).Source
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c)) { return $c }
    }
    throw "Python introuvable. Installez Python 3.12+ ou ajustez le chemin dans build.ps1."
}

function Find-ISCC {
    $candidates = @(
        "C:\Users\$env:USERNAME\AppData\Local\Programs\Inno Setup 6\ISCC.exe",
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }
    throw "Inno Setup (ISCC.exe) introuvable. Installez-le : winget install JRSoftware.InnoSetup"
}

$python = Find-Python
$iscc   = Find-ISCC

Write-Host "Python  : $python" -ForegroundColor Cyan
Write-Host "ISCC    : $iscc" -ForegroundColor Cyan
Write-Host ""

# ── 1) Génération de l'icône .ico à partir du logo (si absente ou logo plus récent) ──
$logoPng = Join-Path $PSScriptRoot "logo_cpi.png"
$logoIco = Join-Path $PSScriptRoot "logo_cpi.ico"
if ((Test-Path $logoPng) -and ((-not (Test-Path $logoIco)) -or ((Get-Item $logoPng).LastWriteTime -gt (Get-Item $logoIco).LastWriteTime))) {
    Write-Host "[1/3] Génération de l'icône .ico…" -ForegroundColor Yellow
    & $python -c "from PIL import Image; img = Image.open(r'$logoPng'); img.save(r'$logoIco', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])"
    if ($LASTEXITCODE -ne 0) { throw "Échec de la génération de l'icône." }
} else {
    Write-Host "[1/3] Icône .ico déjà à jour, étape ignorée." -ForegroundColor DarkGray
}

# ── 2) Build PyInstaller ──────────────────────────────────────────────────────
Write-Host "[2/3] Génération de l'exécutable (PyInstaller)…" -ForegroundColor Yellow
& $python -m PyInstaller --onefile --windowed --name "Chiffrage OEM" --icon "logo_cpi.ico" --clean chiffrage_oem.py
if ($LASTEXITCODE -ne 0) { throw "Échec de PyInstaller." }

# Le logo doit être présent à côté de l'exe pour être trouvé au lancement
Copy-Item $logoPng (Join-Path $PSScriptRoot "dist\logo_cpi.png") -Force

# ── 3) Compilation de l'installeur (Inno Setup) ──────────────────────────────
Write-Host "[3/3] Compilation de l'installeur (Inno Setup) — version $Version…" -ForegroundColor Yellow
& $iscc "/DMyAppVersion=$Version" "installeur_chiffrage.iss"
if ($LASTEXITCODE -ne 0) { throw "Échec de la compilation Inno Setup." }

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host " Build terminé avec succès !" -ForegroundColor Green
Write-Host "   Exécutable brut : dist\Chiffrage OEM.exe" -ForegroundColor Green
Write-Host "   Installeur      : installeur\Chiffrage_OEM_Setup_v$Version.exe" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
