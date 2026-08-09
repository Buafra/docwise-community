# DocWise Community

Open-source OCR and RAG archive for local documents, with Arabic and English support.

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
