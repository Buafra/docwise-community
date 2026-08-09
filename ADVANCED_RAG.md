# DocWise Advanced RAG

DocWise now uses a hybrid RAG pipeline instead of only simple keyword matching.

## Pipeline

1. Extract text/OCR from documents.
2. Create structure-aware chunks by page, paragraphs, lines, tables, and slides.
3. Store chunks in SQLite.
4. Build SQLite FTS5 index for BM25-style keyword retrieval.
5. Build vector embeddings for semantic retrieval.
6. Merge FTS + vector candidates.
7. Rerank using Arabic-aware query terms, metadata, document type, amount/date intent, and source signals.
8. Send only the top cited chunks to the answer model if `OPENAI_API_KEY` is set.

## Chunking

The app uses page-preserving chunks around 700 approximate tokens with around 90 token overlap. It tries to keep paragraphs, table rows, Excel sheet rows, and PowerPoint slide text together.

## Indexes

- `documents`: document metadata and full extracted text
- `chunks`: page-aware RAG chunks
- `chunk_fts`: SQLite FTS5 full-text index
- `chunk_embeddings`: vector embeddings per chunk

## Embeddings

If `OPENAI_API_KEY` is set, DocWise uses OpenAI embeddings by default:

`text-embedding-3-small`

You can change it in `start.bat`:

```bat
set "DOCWISE_EMBED_MODEL=text-embedding-3-small"
```

If no API key is set, DocWise uses a local deterministic multilingual hash vector fallback. This is useful offline but less semantic than real embeddings.

## Retrieval

For each question:

- Arabic/English query terms are normalized and expanded.
- FTS5 retrieves exact matches.
- Vector search retrieves semantic matches.
- Results are merged and reranked.
- Source document/page/snippet are returned.

## Query planning

DocWise detects simple intents such as:

- invoice questions
- contract questions
- ID/passport questions
- amount questions
- date/expiry questions
- latest/recent questions

Example:

`كم قيمة آخر فاتورة؟`

Planner hints:

- document type: invoice
- field: amount
- sort: latest

## Rebuild indexes

Open Settings and click:

`Rebuild RAG Indexes`

Or call:

`POST /api/rebuild-rag`

## Added next-phase features

- OCR quality labels/scores (`empty`, `poor`, `medium`, `good`)
- automatic OpenAI Vision fallback for poor OCR when `OPENAI_API_KEY` is configured
- per-document **OpenAI Vision OCR** button for images/PDFs
- structured field extraction stored in `documents.fields`
- GPT reranking when `OPENAI_API_KEY` is configured (`DOCWISE_GPT_RERANK=1`)
- answer verification when `OPENAI_API_KEY` is configured (`DOCWISE_VERIFY=1`)
- RAG evaluation cases and scoring via Settings or `/api/evaluate-rag`

## Next phases still possible

- true table cell citations
- bounding boxes for PDF/image citations
- dedicated local reranker model
- larger evaluation set and trend history
