# Mbodjatech

Scripts et outils internes.

## Convertisseur PDF → Excel (`src/`)

Page statique qui convertit des PDF en classeurs Excel (.xlsx) ou CSV,
entièrement dans le navigateur. Production : https://pdf-to-excel-boost-senegal.vercel.app

### Lancer en local

```bash
cd src && python3 -m http.server 8000
# puis ouvrir http://localhost:8000
```

### Sécurité côté client

- Aucun envoi réseau : la politique CSP (`connect-src 'none'`) interdit toute
  connexion sortante, et pdf.js / ExcelJS sont hébergés dans `src/vendor/`
  (aucun CDN tiers).
- Validation des fichiers : extension ou type `.pdf`, signature `%PDF-`
  vérifiée dans les octets, 50 Mo maximum par fichier, 20 fichiers, 1000 pages.
- pdf.js est lancé avec `isEvalSupported: false` (protection CVE-2024-4367).
- Export CSV protégé contre l'injection de formules (`=`, `+`, `-`, `@`).
- En-têtes HTTP de sécurité dans `vercel.json`.

### Bibliothèques embarquées

| Bibliothèque | Version | Licence |
|---|---|---|
| pdfjs-dist (`build/pdf.min.js`, `pdf.worker.min.js`) | 3.11.174 | Apache-2.0 |
| exceljs (`dist/exceljs.min.js`) | 4.4.0 | MIT |

Fichiers copiés tels quels depuis les paquets npm officiels.
