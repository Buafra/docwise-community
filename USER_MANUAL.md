# DocWise Archive AI — User Manual

DocWise Archive AI is a local document archive for Arabic/English OCR, folder indexing, smart filing, search, and RAG question answering.

## 1. Start the app

1. Open:

   `C:\Path\To\docwise-community`

2. Double-click:

   `start.bat`

3. In your browser open:

   `http://127.0.0.1:8120`

Keep the black terminal window open while using the app.

## 2. Check OCR status

Go to **Settings**.

You should see:

- Tesseract local OCR: Available
- Arabic OCR: Available
- Languages: `ara, eng, osd`

If Arabic OCR is missing, restart using `start.bat`.

## 3. Upload documents

1. Go to **Upload / Folders**.
2. Under **Upload files**, choose PDFs, images, or text files.
3. Click **Upload + Index**.
4. Wait for indexing to finish.

Supported files:

- PDF
- PNG/JPG/JPEG/WebP/BMP/TIF/TIFF
- Word: DOCX/DOCM/DOTX, legacy DOC best-effort if Microsoft Office is installed
- Excel: XLSX/XLSM/XLTX, legacy XLS best-effort if Microsoft Office is installed
- PowerPoint: PPTX/PPTM/PPSX, legacy PPT best-effort if Microsoft Office is installed
- TXT/MD/CSV/JSON/RTF

## 4. Add a folder for indexing

1. Go to **Upload / Folders**.
2. Paste a folder path, for example:

   `C:\Path\To\Documents\Scans`

3. Keep these checked if you want automatic indexing:

   - **Recursive**: scan subfolders
   - **Watch folder**: keep checking for new files
   - **Scan now**: index immediately

4. Click **Add Folder**.

DocWise does not move your original files. It only indexes them.

## 5. Search documents

1. Go to **Library**.
2. Search in Arabic or English.

Examples:

- `فاتورة الكهرباء`
- `عقد الإيجار`
- `passport expiry`
- `utility invoice`
- `245 AED`

Arabic search supports normalization, so different forms like `أ / إ / آ` are handled better.

## 6. Ask AI / RAG questions

1. Go to **Ask AI**.
2. Ask a question in Arabic or English.

DocWise uses hybrid RAG: structure-aware chunks, SQLite FTS5 keyword search, vector embeddings, and metadata-aware reranking. See `ADVANCED_RAG.md` for details.

Examples:

- `كم قيمة آخر فاتورة؟`
- `وين عقد الإيجار؟`
- `ما تاريخ انتهاء الهوية؟`
- `Find all utility invoices`

If `OPENAI_API_KEY` is not set, DocWise uses extractive local answers. If OpenAI is configured, it gives better AI answers with source citations.

## 7. View a document

Click any document card in **Library**.

You will see:

- file preview
- summary
- smart filing suggestion
- extracted OCR text
- buttons to open, reindex, or copy to archive

## 8. Smart filing

Click **Copy to Smart Filing** in the document viewer.

DocWise copies the file into:

`C:\Path\To\docwise-community\data\archive\TYPE\YEAR\`

Example:

`data\archive\invoice\2026\2026-08-08_utility_invoice_245.50.pdf`

Original files stay untouched.

## 9. Fix PDFs showing ??????

Some Arabic PDFs contain broken embedded text. They may show extracted text as `??????`.

DocWise now detects garbled PDF text and OCR-renders the PDF page instead.

If an old uploaded PDF still shows `??????`:

1. Open the document from **Library**.
2. Click **Reindex**.
3. Wait for OCR to finish.

If the PDF is image-heavy, OCR can take longer.

## 10. Delete documents

### Delete one document

Open **Library**, click the document, then choose one of two buttons:

- **Remove from Index**: removes from DocWise search/RAG only. The original file stays on disk.
- **Delete File + Index**: deletes the original file from disk and removes it from DocWise. This is permanent and asks for confirmation twice.

### Multi-delete

In **Library**, tick the checkbox on multiple document cards, then choose:

- **Remove Selected from Index**
- **Delete Selected App Files + Index**

For safety, bulk file delete only deletes app-managed uploaded/archive files. External watched-folder originals are skipped.

### Start fresh

Click **Start Fresh: Delete All App Documents** in **Library**. This clears all DocWise document records and deletes uploaded/archive files managed by DocWise.

## 11. Recommended workflow

1. Add your main scan/document folders.
2. Let DocWise index everything.
3. Search and review documents marked `needs_ocr` or `error`.
4. Reindex poor OCR documents.
5. Use **Ask AI** for questions.
6. Use **Copy to Smart Filing** only after reviewing suggestions.
