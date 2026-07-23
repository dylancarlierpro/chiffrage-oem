# ─────────────────────────────────────────────────────────────
#  Release complète — Chiffrage OEM
#  1) Reconstruit l'exe + l'installeur (build.ps1 -Version X.Y)
#  2) Commit + push le code sur GitHub (si des changements existent)
#  3) Crée un tag Git vX.Y + une Release GitHub avec l'installeur attaché
#
#  Usage :  .\release.ps1 -Version "1.1" -Message "Ajout du mode sombre"
# ─────────────────────────────────────────────────────────────

param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$Message
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$Repo = "dylancarlierpro/chiffrage-oem"
$Tag  = "v$Version"

function Find-Tool {
    param([string[]]$Candidates, [string]$Name)
    foreach ($c in $Candidates) {
        if ($c -and (Test-Path $c)) { return $c }
    }
    throw "$Name introuvable."
}

$git = Find-Tool @(
    "C:\Program Files\Git\cmd\git.exe",
    (Get-Command git -ErrorAction SilentlyContinue).Source
) "Git"

$gh = Find-Tool @(
    "C:\Program Files\GitHub CLI\gh.exe",
    (Get-Command gh -ErrorAction SilentlyContinue).Source
) "GitHub CLI"

Write-Host "Version : $Tag" -ForegroundColor Cyan
Write-Host "Dépôt   : $Repo" -ForegroundColor Cyan
Write-Host ""

# ── 0) Le tag ne doit pas déjà exister ────────────────────────────────────────
$existingTag = & $git tag -l $Tag
if ($existingTag) {
    throw "Le tag $Tag existe déjà. Choisissez un autre numéro de version."
}

# ── 1) Build complet (exe + installeur, nommés avec le bon numéro de version) ─
Write-Host "[1/4] Build de l'application…" -ForegroundColor Yellow
& (Join-Path $PSScriptRoot "build.ps1") -Version $Version
if ($LASTEXITCODE -ne 0) { throw "Échec du build." }

$installerPath = Join-Path $PSScriptRoot "installeur\Chiffrage_OEM_Setup_v$Version.exe"
if (-not (Test-Path $installerPath)) {
    throw "Installeur introuvable après build : $installerPath"
}

# ── 2) Commit + push du code (seulement s'il y a des changements) ────────────
Write-Host "[2/4] Commit et push du code…" -ForegroundColor Yellow
# git add -u (et non "git add .") : ne stage que les fichiers DÉJÀ suivis par
# Git. Un "git add ." pousserait aveuglément tout nouveau fichier créé à la
# racine (export de test, dump JSON renommé...) sur le dépôt distant sans
# revue. Un fichier untracked qui devrait vraiment être ajouté doit l'être
# explicitement (git add <fichier>), jamais via ce script.
& $git add -u
$untracked = & $git status --porcelain | Where-Object { $_ -match '^\?\?' }
if ($untracked) {
    Write-Host "  Fichiers non suivis ignorés par ce script (à ajouter manuellement si besoin) :" -ForegroundColor DarkYellow
    $untracked | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
}
$hasChanges = & $git status --porcelain --untracked-files=no
if ($hasChanges) {
    & $git commit -m $Message
    if ($LASTEXITCODE -ne 0) { throw "Échec du commit." }
} else {
    Write-Host "  Aucun changement de code à committer." -ForegroundColor DarkGray
}
& $git push
if ($LASTEXITCODE -ne 0) { throw "Échec du push." }

# ── 3) Tag Git ─────────────────────────────────────────────────────────────
Write-Host "[3/4] Création et push du tag $Tag…" -ForegroundColor Yellow
& $git tag $Tag
& $git push origin $Tag
if ($LASTEXITCODE -ne 0) { throw "Échec du push du tag." }

# ── 4) Release GitHub avec l'installeur attaché ───────────────────────────────
Write-Host "[4/4] Création de la Release GitHub…" -ForegroundColor Yellow
& $gh release create $Tag $installerPath `
    --repo $Repo `
    --title "Chiffrage OEM $Tag" `
    --notes $Message
if ($LASTEXITCODE -ne 0) { throw "Échec de la création de la Release." }

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host " Release $Tag publiée avec succès !" -ForegroundColor Green
Write-Host "   https://github.com/$Repo/releases/tag/$Tag" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
