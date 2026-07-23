# Sources du dossier de spécifications

`spec_fr.html` / `spec_en.html` génèrent les PDF du dossier `docs/` (`Chiffrage_OEM_-_Dossier_de_specifications.pdf` et sa version `_EN`). Les captures d'écran utilisées sont dans `screenshots/`.

## Régénérer un PDF

Avec Chrome ou Edge installé (headless print-to-pdf, sans en-tête/pied de page injecté) :

```bash
msedge --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="Chiffrage_OEM_-_Dossier_de_specifications.pdf" spec_fr.html
```

Remplacer `spec_fr.html` par `spec_en.html` pour la version anglaise. Le chemin de l'exécutable Edge sur Windows est généralement `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`.
