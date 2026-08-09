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

## Run

```bat
start.bat
```

Then open:

```txt
http://127.0.0.1:8120
```

## Community install command line

Fresh Windows install:

```powershell
git clone https://github.com/Buafra/docwise-community.git
cd docwise-community
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\start.bat
```

Update an existing install:

```powershell
cd docwise-community
git pull
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
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
