# DocWise Community

Open-source OCR and RAG archive for local documents, with Arabic and English support.

DocWise Community is a local-first document intelligence app that supports Arabic/English OCR, folder indexing, Office/PDF/image files, smart filing suggestions, and advanced RAG.

## Features

- Upload PDFs, images, Word, Excel, PowerPoint, text and RTF files
- Pick local folders with a native folder chooser; watched folders rescan automatically when new files arrive, and duplicate files are detected and skipped
- Arabic + English OCR using Tesseract best-accuracy models, with parallel page OCR and live progress
- **Semantic search that understands meaning across Arabic and English** — a local multilingual embedding model (downloads automatically on first run, ~130MB, then fully offline)
- **Page viewer with search-hit highlighting**: see exactly where on the scanned page the words were found
- **Searchable PDF export**: turn any scan into a selectable, searchable PDF
- Structured field extraction (invoice number, total, currency, VAT number, IBAN, dates — Arabic-Indic digits included) and **Excel export** of the whole archive
- Ask AI with clickable citations that open the exact source page
- Optional local vision OCR via [Ollama](https://ollama.com) (auto-detected, free, private) and optional OpenAI Vision/embeddings with your own key
- OCR quality scoring, orientation auto-correction, and manual OCR correction
- Hybrid RAG: SQLite FTS5/BM25 + vector embeddings + Arabic-aware reranking
- Meta Astryx-based UI

## Install and run

Nothing to preinstall — the launcher installs everything automatically on first run (Python 3.13, Tesseract OCR, and all app dependencies). You only need internet and a few minutes the first time. Later runs start instantly.

### Windows 10/11

```powershell
git clone https://github.com/Buafra/docwise-community.git
cd docwise-community
.\start.bat
```

Without Git: download the ZIP from GitHub (**Code → Download ZIP**), extract it, and double-click `start.bat`.

Notes:
- Windows may show one admin (UAC) prompt while installing Tesseract — click **Yes**. If you decline, the app still works but OCR of scanned documents is disabled.
- If automatic Python install fails, install [Python 3.10+](https://www.python.org/downloads/) yourself (tick **"Add python.exe to PATH"**) and run `start.bat` again.

### macOS

```bash
git clone https://github.com/Buafra/docwise-community.git
cd docwise-community
./start.command
```

Without Git: download and extract the ZIP, then double-click `start.command`. If macOS blocks it ("unidentified developer"), right-click the file and choose **Open** the first time. If it opens as a text file instead of running, run in Terminal: `chmod +x start.command` then try again.

Notes:
- If Python 3.10+ is missing, the script installs it via [Homebrew](https://brew.sh) (and offers to install Homebrew first if needed — this asks for your Mac password, which is normal).
- Tesseract OCR is installed with `brew install tesseract`. Without it the app still works, but OCR of scanned documents is disabled.

### After starting

Open:

```txt
http://127.0.0.1:8120
```

Arabic and English OCR language data are already bundled in `tessdata/` — no extra downloads.

Update an existing install:

```bash
cd docwise-community
git pull
```

Then start it again (`.\start.bat` on Windows, `./start.command` on macOS).

## OCR quality

- The repo bundles Tesseract's **best-accuracy** models (`tessdata_best`) for Arabic and English — no downloads needed.
- Scanned PDF pages render at ~250 dpi and each page tries multiple language and segmentation modes; the winner is chosen by Tesseract's own word confidence, so glyph noise never beats a real reading.
- Sideways and upside-down pages (phone photos) are detected and corrected automatically.
- Best input = best output: scan at 300 dpi, dark text on light background, pages straight and uncropped.
- For the absolute best results on phone photos, handwriting, stamps, and complex invoices, set `OPENAI_API_KEY` (below) — pages Tesseract struggles with automatically fall back to OpenAI Vision OCR.

## Optional OpenAI features

Set your own key before running.

Windows:

```bat
set OPENAI_API_KEY=your_key_here
```

macOS:

```bash
export OPENAI_API_KEY=your_key_here
```

or add it locally to `start.bat` / `start.command`. Do not commit secrets.

## Rebuild UI

```bat
rebuild-astryx-webapp.bat
```

## License

AGPL-3.0-or-later. See `LICENSE`.

## Pro / Enterprise

Commercial features such as hosted SaaS, license server, customer portal, Stripe billing, team workspaces, managed AI credits, installers, cloud backup/sync, support and SLA belong in DocWise Pro/Enterprise.
