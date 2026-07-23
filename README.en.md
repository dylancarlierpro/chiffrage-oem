*[Lire en français](README.md)*

# Chiffrage OEM

Desktop application (PyQt6, Windows) computing OEM unit prices for industrial
parts sold to different clients, developed for **CPI Howden** (formerly CPI
Liard).

> **Portfolio note** — this repository is a public version of the project:
> the business logic is identical to the one used in production, but the
> demo data (clients, rates, discounts in `DEFAULT_DATA`) is **fictional**.
> CPI Howden's real commercial data is neither present nor accessible from
> this repository (`chiffrage_data.json` is intentionally excluded, see
> `.gitignore`).

![Chiffrage OEM — main screen](docs/src/screenshots/screenshot_0.png)

## Documentation

Full specification document (pricing logic, architecture, reliability/security
review, screenshots):
[FR](docs/Chiffrage_OEM_-_Dossier_de_specifications.pdf) ·
[EN](docs/Chiffrage_OEM_-_Dossier_de_specifications_EN.pdf).

## Business logic

We start from a reference catalog price (« PL21 ») entered by the user, then
apply successively:

1. the annual increase rates negotiated with each client (per year, per
   material — PTFE or PEEK),
2. the commercial discount negotiated with that client,

```
price_final = ceil( PL21 × ∏(1 + rate_year) × (1 + discount) )
```

The result is rounded up to the nearest whole cent.

## Features

- Material selection (PTFE/PEEK), instant calculation, calculation history
- Waterfall visualization of the effect of each year + discount
- Password-protected admin management (add/remove clients and years, edit
  rates and discounts)
- PDF export of the calculation detail, Excel export (optional, if
  `openpyxl` is installed), quick copy to clipboard
- Admin action audit log, atomic data saving, automatic recovery on a
  corrupted file

## Security

- Admin password hashed and salted (PBKDF2-HMAC-SHA256), constant-time
  comparison (`hmac.compare_digest`)
- No hardcoded default password: on first use, a password is **randomly
  generated** and shown on screen once
- Progressive slowdown after failed admin login attempts
- Any anomaly in the authentication data (manually altered file) is traced
  in the audit log

## Tech stack

- Python 3.12, PyQt6 (UI), openpyxl (Excel export, optional)
- Packaged as a Windows executable via PyInstaller (`build.ps1`)
- Installer via Inno Setup (`installeur_chiffrage.iss`)
- Automated release: build + tag + GitHub Release (`release.ps1`)

## Run the project

```bash
pip install PyQt6 openpyxl
python chiffrage_oem.py
```

On first launch, a fictional demo dataset is created and an admin password
is generated and shown on screen.

---

Created by Dylan Carlier.
