# notificacionesCese

Python script that validates a Google Drive restaurant folder, generates QR assets and DOCX/PDF outputs, and uploads the generated files back to the same Drive folder.

This folder was prepared as a separate project based on the proven `derechosTransmision` workflow. For now, the code path is intentionally kept equivalent while the new Excel form and templates for `notificacionesCese` are developed.

## What the project expects

- A Google Drive folder containing `ANEXOA.pdf`, `ANEXOB.pdf`, `ANEXOC.pdf`, and `acta_data.xlsx`
- The annex PDFs must already be public in Drive
- Google OAuth client credentials saved locally as `src/credentials.json`

## Project structure

- `src/`: application code
- `templates/`: DOCX/PDF templates used during generation
- `runs/`: parent folder for per-execution temporary output directories

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Put your Google OAuth credentials file at `src/credentials.json`.
4. On first run, complete the Google login flow. The app will create `src/token.json` locally.

## PDF conversion

PDF export requires one of these local tools:

- Windows: `docx2pdf` plus Microsoft Word installed
- Windows, Linux, or macOS: LibreOffice with `soffice` available in `PATH`

`docx2pdf` is not listed in `requirements.txt` because the code already treats it as optional and falls back to LibreOffice when available.

If you want the Windows-first path, install it manually:

```bash
pip install docx2pdf
```

## Run

From the project root:

```bash
python src/main.py
```

Then paste the Google Drive folder link when prompted.

## Git notes

- `src/credentials.json` and `src/token.json` are intentionally ignored
- `runs/` stays in the repo only as an empty placeholder directory
- Each execution creates a unique temporary folder inside `runs/` and deletes it at the end, even on failures
- Generated files inside `runs/` should stay local and should not be committed
