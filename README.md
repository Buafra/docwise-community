# DocWise Community

Open-source Arabic-first OCR and RAG archive for local documents.

DocWise Community is a local-first document intelligence app that supports Arabic/English OCR, folder indexing, Office/PDF/image files, smart filing suggestions, and advanced RAG.

## Features

- Upload PDFs, images, Word, Excel, PowerPoint, text and RTF files
- Add local folders and index recursively
- Arabic + English OCR using Tesseract
- Optional OpenAI Vision OCR with your own `OPENAI_API_KEY`
- OCR quality scoring and manual OCR correction
- Smart filing suggestions and safe copy-to-archive
- Advanced hybrid RAG: SQLite FTS5/BM25 + vector embeddings
- Optional OpenAI embeddings, GPT reranking and answer verification with your own key
- Arabic search normalization
- RAG evaluation cases
- Meta Astryx-based UI

## Requirements

- Windows 10/11
- [Python 3.10+](https://www.python.org/downloads/) — during install, tick **"Add python.exe to PATH"**
- [Git](https://git-scm.com/downloads) (only needed for the git clone install)
- Optional, for OCR of scanned documents and images: [Tesseract OCR for Windows](https://github.com/UB-Mannheim/tesseract/wiki) — the app runs without it, but OCR is disabled. Arabic and English language data are already bundled in `tessdata/`.

## Install and run (Windows)

```powershell
git clone https://github.com/Buafra/docwise-community.git
cd docwise-community
.\start.bat
```

Then open:

```txt
http://127.0.0.1:8120
```

On first run, `start.bat` automatically creates a virtual environment in `.venv` and installs all dependencies (needs internet). Later runs start instantly. No `pip` or `venv` commands needed.

Alternative without Git: download the ZIP from GitHub (**Code → Download ZIP**), extract it, and double-click `start.bat`.

Update an existing install:

```powershell
cd docwise-community
git pull
.\start.bat
```

## Optional OpenAI features

Set your own key before running:

```bat
set OPENAI_API_KEY=your_key_here
```

or add it locally to `start.bat`. Do not commit secrets.

## Rebuild UI

```bat
rebuild-astryx-webapp.bat
```

## License

AGPL-3.0-or-later. See `LICENSE`.

## Pro / Enterprise

Commercial features such as hosted SaaS, license server, customer portal, Stripe billing, team workspaces, managed AI credits, installers, cloud backup/sync, support and SLA belong in DocWise Pro/Enterprise.
