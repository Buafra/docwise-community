import base64
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

try:
    import fitz  # PyMuPDF
    FITZ_IMPORT_ERROR = None
except Exception as _fitz_exc:
    fitz = None
    FITZ_IMPORT_ERROR = str(_fitz_exc)

try:
    from PIL import Image
except Exception:
    Image = None

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

try:
    from docx import Document
except Exception:
    Document = None

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

try:
    from pptx import Presentation
except Exception:
    Presentation = None

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
UPLOADS = DATA / "uploads"
ARCHIVE = DATA / "archive"
DB_PATH = DATA / "docwise.sqlite3"
STATIC = ROOT / "static_astryx" if (ROOT / "static_astryx").exists() else ROOT / "static"

for folder in (DATA, UPLOADS, ARCHIVE):
    folder.mkdir(parents=True, exist_ok=True)

SUPPORTED = {
    ".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff",
    ".txt", ".md", ".csv", ".json", ".rtf",
    ".docx", ".docm", ".dotx", ".doc",
    ".xlsx", ".xlsm", ".xltx", ".xls",
    ".pptx", ".pptm", ".ppsx", ".ppt",
}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
TEXT_EXTS = {".txt", ".md", ".csv", ".json"}
WORD_EXTS = {".docx", ".docm", ".dotx"}
EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx"}
POWERPOINT_EXTS = {".pptx", ".pptm", ".ppsx"}
LEGACY_OFFICE_EXTS = {".doc", ".xls", ".ppt"}
ARABIC_RE = re.compile(r"[\u0600-\u06FF]")
TASHKEEL_RE = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED]")

app = FastAPI(title="DocWise Archive AI", version="0.1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
DB_LOCK = threading.Lock()
SCAN_STATE = {"running": False, "last": None, "message": "Ready", "indexed": 0, "errors": 0}


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with DB_LOCK, db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                path TEXT NOT NULL UNIQUE,
                original_name TEXT,
                source_type TEXT NOT NULL DEFAULT 'folder',
                file_ext TEXT,
                size INTEGER DEFAULT 0,
                mtime REAL DEFAULT 0,
                sha256 TEXT,
                title TEXT,
                doc_type TEXT DEFAULT 'general',
                language TEXT DEFAULT 'unknown',
                date_guess TEXT,
                company TEXT,
                amount TEXT,
                tags TEXT DEFAULT '[]',
                summary TEXT,
                fields TEXT DEFAULT '{}',
                text TEXT DEFAULT '',
                normalized_text TEXT DEFAULT '',
                status TEXT DEFAULT 'new',
                ocr_engine TEXT DEFAULT 'none',
                ocr_quality TEXT DEFAULT 'unknown',
                ocr_score REAL DEFAULT 0,
                error TEXT,
                created_at TEXT,
                updated_at TEXT
            );
            CREATE TABLE IF NOT EXISTS chunks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                page INTEGER DEFAULT 1,
                chunk_index INTEGER DEFAULT 0,
                text TEXT NOT NULL,
                normalized_text TEXT NOT NULL,
                token_count INTEGER DEFAULT 0,
                FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS chunk_embeddings (
                chunk_id INTEGER PRIMARY KEY,
                document_id INTEGER NOT NULL,
                model TEXT NOT NULL,
                dims INTEGER NOT NULL,
                embedding TEXT NOT NULL,
                created_at TEXT,
                FOREIGN KEY(chunk_id) REFERENCES chunks(id) ON DELETE CASCADE
            );
            CREATE VIRTUAL TABLE IF NOT EXISTS chunk_fts USING fts5(
                chunk_id UNINDEXED,
                document_id UNINDEXED,
                title,
                text,
                normalized_text,
                tokenize='unicode61'
            );
            CREATE TABLE IF NOT EXISTS folders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                path TEXT NOT NULL UNIQUE,
                recursive INTEGER DEFAULT 1,
                watch INTEGER DEFAULT 1,
                last_scan TEXT,
                created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS eval_cases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                question TEXT NOT NULL,
                expected TEXT DEFAULT '',
                must_cite TEXT DEFAULT '',
                created_at TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_documents_type ON documents(doc_type);
            CREATE INDEX IF NOT EXISTS idx_documents_lang ON documents(language);
            CREATE INDEX IF NOT EXISTS idx_chunks_document ON chunks(document_id);
            CREATE INDEX IF NOT EXISTS idx_embeddings_document ON chunk_embeddings(document_id);
            """
        )
        for stmt in (
            "ALTER TABLE chunks ADD COLUMN token_count INTEGER DEFAULT 0",
            "ALTER TABLE documents ADD COLUMN fields TEXT DEFAULT '{}'",
            "ALTER TABLE documents ADD COLUMN ocr_quality TEXT DEFAULT 'unknown'",
            "ALTER TABLE documents ADD COLUMN ocr_score REAL DEFAULT 0",
        ):
            try:
                conn.execute(stmt)
            except sqlite3.OperationalError:
                pass



init_db()


def normalize_arabic(text: str) -> str:
    text = text or ""
    text = TASHKEEL_RE.sub("", text)
    text = re.sub("[إأآٱ]", "ا", text)
    text = text.replace("ى", "ي").replace("ؤ", "و").replace("ئ", "ي")
    text = text.replace("ة", "ه")
    text = re.sub(r"[؟،؛ـ]", " ", text)
    text = text.lower()
    text = re.sub(r"[^\w\u0600-\u06FF]+", " ", text, flags=re.UNICODE)
    return re.sub(r"\s+", " ", text).strip()


def safe_json_load(value: str, default):
    try:
        return json.loads(value or "")
    except Exception:
        return default


def contains_keyword(normalized_text: str, keyword: str) -> bool:
    kw = normalize_arabic(keyword)
    if not kw:
        return False
    padded = f" {normalized_text} "
    if " " in kw:
        return f" {kw} " in padded
    if re.fullmatch(r"[a-z0-9]+", kw):
        return f" {kw} " in padded
    # Arabic classification keywords must match word-like tokens, not substrings.
    # This prevents false matches like عقد inside المعقد, or طبي inside unrelated words.
    words = set(normalized_text.split())
    variants = {kw, f"ال{kw}", f"و{kw}", f"بال{kw}", f"لل{kw}"}
    for word in words:
        if word in variants:
            return True
        stripped = word
        for prefix in ("وال", "بال", "لل", "ال", "و", "ب", "ل"):
            if stripped.startswith(prefix) and len(stripped) > len(prefix) + 1:
                stripped = stripped[len(prefix):]
                break
        if stripped == kw:
            return True
    return False


def approx_token_count(text: str) -> int:
    # Good enough for chunk sizing/index stats without adding tokenizer dependencies.
    words = re.findall(r"[\w\u0600-\u06FF]+", text or "", flags=re.UNICODE)
    return max(1, int(len(words) * 1.25)) if words else 0


def cosine_similarity(a: list[float], b: list[float]) -> float:
    if not a or not b or len(a) != len(b):
        return 0.0
    dot = sum(x * y for x, y in zip(a, b))
    na = sum(x * x for x in a) ** 0.5
    nb = sum(y * y for y in b) ** 0.5
    return dot / (na * nb) if na and nb else 0.0


def local_embedding(text: str, dims: int = 384) -> list[float]:
    """Deterministic multilingual lexical vector fallback when no embedding API is configured."""
    vec = [0.0] * dims
    norm = normalize_arabic(text)
    tokens = [t for t in norm.split() if len(t) > 1]
    features = []
    for tok in tokens:
        features.append(tok)
        if len(tok) >= 4:
            features.extend(tok[i:i+3] for i in range(len(tok) - 2))
    for feat in features:
        h = int(hashlib.sha256(feat.encode("utf-8")).hexdigest()[:12], 16)
        idx = h % dims
        sign = 1.0 if (h >> 3) & 1 else -1.0
        vec[idx] += sign
    length = sum(v * v for v in vec) ** 0.5
    return [v / length for v in vec] if length else vec


def embed_text(text: str) -> tuple[list[float], str]:
    mode = os.environ.get("DOCWISE_EMBEDDINGS", "auto").lower()
    if mode != "local" and os.environ.get("OPENAI_API_KEY") and OpenAI:
        try:
            client = OpenAI()
            model = os.environ.get("DOCWISE_EMBED_MODEL", "text-embedding-3-small")
            resp = client.embeddings.create(model=model, input=(text or "")[:8000])
            return list(resp.data[0].embedding), model
        except Exception:
            if mode == "openai":
                raise
    return local_embedding(text), "local-hash-multilingual-v1"


def fts_query_from_terms(terms: list[str]) -> str:
    clean = []
    for term in terms[:12]:
        term = normalize_arabic(term)
        if len(term) < 2:
            continue
        term = re.sub(r'[^\w\u0600-\u06FF]+', ' ', term, flags=re.UNICODE).strip()
        if term:
            clean.append('"' + term.replace('"', '') + '"')
    return " OR ".join(dict.fromkeys(clean))


def query_plan(question: str) -> dict:
    q = normalize_arabic(question)
    plan = {"doc_type": None, "field": None, "sort": None, "terms": query_terms(question)}
    doc_rules = {
        "invoice": ["فاتوره", "الفاتوره", "bill", "invoice", "كهرباء", "utility"],
        "contract": ["عقد", "ايجار", "contract", "lease", "rent"],
        "id": ["هويه", "جواز", "passport", "emirates id"],
        "receipt": ["ايصال", "receipt", "payment"],
        "bank": ["بنك", "bank", "iban", "statement"],
        "medical": ["طبي", "مستشفي", "medical", "hospital", "clinic"],
    }
    for doc_type, kws in doc_rules.items():
        if any(contains_keyword(q, kw) for kw in kws):
            plan["doc_type"] = doc_type
            break
    if any(contains_keyword(q, kw) for kw in ["كم", "قيمه", "مبلغ", "total", "amount", "price"]):
        plan["field"] = "amount"
    if any(contains_keyword(q, kw) for kw in ["تاريخ", "متي", "date", "expiry", "انتهاء", "due"]):
        plan["field"] = plan["field"] or "date"
    if any(contains_keyword(q, kw) for kw in ["اخر", "احدث", "latest", "newest", "recent"]):
        plan["sort"] = "latest"
    return plan


def delete_chunk_indexes_for_doc(conn: sqlite3.Connection, doc_id: int) -> None:
    chunk_ids = [r["id"] for r in conn.execute("SELECT id FROM chunks WHERE document_id=?", (doc_id,)).fetchall()]
    if chunk_ids:
        placeholders = ",".join("?" for _ in chunk_ids)
        conn.execute(f"DELETE FROM chunk_embeddings WHERE chunk_id IN ({placeholders})", chunk_ids)
    conn.execute("DELETE FROM chunk_fts WHERE document_id=?", (doc_id,))


def add_chunk_indexes(conn: sqlite3.Connection, chunk_id: int, doc_id: int, title: str, text: str, normalized_text: str) -> None:
    conn.execute(
        "INSERT INTO chunk_fts(chunk_id, document_id, title, text, normalized_text) VALUES(?,?,?,?,?)",
        (chunk_id, doc_id, title or "", text or "", normalized_text or ""),
    )
    try:
        emb, model = embed_text(f"{title}\n{text}")
        conn.execute(
            "INSERT OR REPLACE INTO chunk_embeddings(chunk_id, document_id, model, dims, embedding, created_at) VALUES(?,?,?,?,?,?)",
            (chunk_id, doc_id, model, len(emb), json.dumps(emb), now_iso()),
        )
    except Exception as exc:
        # FTS still works if embeddings fail.
        conn.execute(
            "INSERT OR REPLACE INTO chunk_embeddings(chunk_id, document_id, model, dims, embedding, created_at) VALUES(?,?,?,?,?,?)",
            (chunk_id, doc_id, f"embedding-error:{str(exc)[:80]}", 0, "[]", now_iso()),
        )


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def has_arabic(text: str) -> bool:
    return bool(ARABIC_RE.search(text or ""))


def text_is_garbled(text: str) -> bool:
    """Detect PDF text extraction that is technically present but useless, e.g. Arabic rendered as ?????."""
    clean = re.sub(r"\s+", "", text or "")
    if not clean:
        return True
    bad_chars = clean.count("?") + clean.count("�") + clean.count("□") + clean.count("■")
    weird_chars = sum(1 for ch in clean if not re.match(r"[A-Za-z0-9\u0600-\u06FF\s.,:;!؟،؛()\[\]{}_/\\+\-=٪%$€£¥@#&*'\"|<>\n\r\t-]", ch))
    greek_garbage = len(re.findall(r"[\u0370-\u03FF\u1F00-\u1FFF]", clean))
    if len(clean) >= 8 and bad_chars / max(1, len(clean)) > 0.22:
        return True
    if greek_garbage >= 2:
        return True
    if len(clean) >= 20 and weird_chars / max(1, len(clean)) > 0.08:
        return True
    if re.fullmatch(r"[?\W_\d]+", clean) and clean.count("?") >= 4:
        return True
    return False


def should_ocr_pdf_page(text: str) -> bool:
    stripped = (text or "").strip()
    return len(stripped) < 25 or text_is_garbled(stripped)


def tesseract_cmd() -> Optional[str]:
    env = os.environ.get("TESSERACT_CMD")
    if env and Path(env).exists():
        return env
    found = shutil.which("tesseract")
    if found:
        return found
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


def tessdata_dir() -> Optional[str]:
    env = os.environ.get("TESSDATA_PREFIX")
    if env and Path(env).exists():
        return env
    local = ROOT / "tessdata"
    if local.exists():
        return str(local)
    cmd = tesseract_cmd()
    if cmd:
        installed = Path(cmd).parent / "tessdata"
        if installed.exists():
            return str(installed)
    return None


def available_languages() -> list[str]:
    folder = tessdata_dir()
    if not folder:
        return []
    return sorted([p.stem for p in Path(folder).glob("*.traineddata")])


def available_ocr() -> dict:
    cmd = tesseract_cmd()
    langs = available_languages()
    return {
        "tesseract": bool(cmd and pytesseract and Image),
        "tesseract_cmd": cmd,
        "tessdata_dir": tessdata_dir(),
        "languages": langs,
        "arabic": "ara" in langs,
        "english": "eng" in langs,
        "office": {"word": bool(Document), "excel": bool(load_workbook), "powerpoint": bool(Presentation)},
        "embedding_mode": os.environ.get("DOCWISE_EMBEDDINGS", "auto"),
        "embedding_model": os.environ.get("DOCWISE_EMBED_MODEL", "text-embedding-3-small") if os.environ.get("OPENAI_API_KEY") else "local-hash-multilingual-v1",
        "openai_vision": bool(os.environ.get("OPENAI_API_KEY") and OpenAI),
        "openai_text": bool(os.environ.get("OPENAI_API_KEY") and OpenAI),
    }


def ocr_text_score(text: str) -> int:
    clean = (text or "").strip()
    if not clean:
        return 0
    arabic = len(re.findall(r"[\u0600-\u06FF]", clean))
    latin = len(re.findall(r"[A-Za-z]", clean))
    digits = len(re.findall(r"\d", clean))
    words = len(re.findall(r"[\w\u0600-\u06FF]{2,}", clean, flags=re.UNICODE))
    bad = clean.count("?") + clean.count("�") + clean.count("□")
    weird = sum(1 for ch in clean if not re.match(r"[A-Za-z0-9\u0600-\u06FF\s.,:;!؟،؛()\[\]{}_/\\+\-=٪%$€£¥@#&*'\"|<>\n\r\t-]", ch))
    # Arabic and Latin weighted equally: overweighting Arabic makes the
    # Arabic-only pass win on English documents with garbage glyph output.
    return arabic * 2 + latin * 2 + digits + words * 3 - bad * 10 - weird * 2


def ocr_quality(text: str) -> tuple[str, float]:
    clean = (text or "").strip()
    if not clean:
        return "empty", 0.0
    score = ocr_text_score(clean)
    chars = max(1, len(clean))
    bad = clean.count("?") + clean.count("�") + clean.count("□")
    bad_ratio = bad / chars
    word_count = len(re.findall(r"[\w\u0600-\u06FF]{2,}", clean, flags=re.UNICODE))
    density = score / chars
    if bad_ratio > 0.08 or word_count < 3:
        return "poor", round(max(0.0, density), 3)
    if density < 0.55:
        return "poor", round(density, 3)
    if density < 1.15:
        return "medium", round(density, 3)
    return "good", round(density, 3)


def should_try_vision_fallback(text: str, engine: str = "") -> bool:
    quality, _ = ocr_quality(text)
    if quality in ("empty", "poor"):
        return True
    # Arabic Tesseract with diacritics often looks plausible but wrong; allow opt-in automatic cloud fallback.
    return bool(os.environ.get("DOCWISE_VISION_FALLBACK", "1") == "1" and "tesseract" in engine and has_arabic(text) and quality == "medium")


def ocr_image_variants(img):
    try:
        from PIL import ImageOps, ImageFilter
        variants = [("orig", img)]
        gray = ImageOps.grayscale(img)
        variants.append(("gray", gray))
        scale = 2 if max(img.size) < 2200 else 1
        if scale > 1:
            big = gray.resize((gray.width * scale, gray.height * scale))
            variants.append(("gray2x", big))
            variants.append(("sharp2x", big.filter(ImageFilter.SHARPEN)))
            variants.append(("threshold2x", big.point(lambda x: 255 if x > 180 else 0)))
        return variants
    except Exception:
        return [("orig", img)]


def ocr_image_tesseract(path: Path) -> tuple[str, str]:
    cmd = tesseract_cmd()
    if not (cmd and pytesseract and Image):
        return "", "tesseract-not-installed"
    pytesseract.pytesseract.tesseract_cmd = cmd
    img = Image.open(path)
    errors = []
    langs = set(available_languages())
    attempts = []
    if {"ara", "eng"}.issubset(langs):
        attempts.extend(["ara+eng", "eng+ara"])
    if "ara" in langs:
        attempts.append("ara")
    if "eng" in langs:
        attempts.append("eng")
    attempts = attempts or ["eng"]
    # Pass tessdata location via env, not --tessdata-dir: an unquoted config
    # argument breaks on paths containing spaces and kills every OCR attempt.
    td = tessdata_dir()
    if td:
        os.environ["TESSDATA_PREFIX"] = td
    psm_modes = [3, 4, 6, 11, 12]

    def mean_confidence(image, lang, psm) -> float:
        """Tesseract's own per-word confidence: real text ~70-95, glyph noise
        that merely counts like text (e.g. a sideways Arabic page) ~20-45."""
        try:
            data = pytesseract.image_to_data(
                image, lang=lang, config=f"--psm {psm}", output_type=pytesseract.Output.DICT
            )
            confs = []
            for c, w in zip(data.get("conf", []), data.get("text", [])):
                try:
                    c = int(float(c))
                except (TypeError, ValueError):
                    continue
                if c >= 0 and str(w).strip():
                    confs.append(c)
            return (sum(confs) / len(confs)) if confs else 0.0
        except Exception:
            return -1.0

    def run_grid(image):
        top = {"text": "", "score": 0, "engine": "", "conf": -1.0}
        for variant_name, variant in ocr_image_variants(image):
            for lang in attempts:
                for psm in psm_modes:
                    try:
                        txt = pytesseract.image_to_string(variant, lang=lang, config=f"--psm {psm}")
                        score = ocr_text_score(txt)
                        if score > top["score"]:
                            top = {"text": txt, "score": score, "engine": f"tesseract:{lang}:psm{psm}:{variant_name}", "conf": -1.0}
                            # A clean, high-confidence read ends the search early;
                            # otherwise every page runs the whole ~100-pass grid.
                            if ocr_quality(txt)[0] == "good":
                                conf = mean_confidence(variant, lang, psm)
                                top["conf"] = conf
                                if conf < 0 or conf >= 60:
                                    return top
                    except Exception as exc:
                        errors.append(f"{lang}/psm{psm}/{variant_name}: {exc}")
        return top

    best = run_grid(img)

    # Sideways or upside-down scans (common with phone photos): if the straight
    # read failed or came back as low-confidence noise, ask Tesseract's
    # orientation detector and retry once on the corrected rotation.
    needs_rescue = ocr_quality(best["text"])[0] in ("empty", "poor") or (0 <= best["conf"] < 50)
    if needs_rescue and "osd" in langs:
        try:
            osd = pytesseract.image_to_osd(img)
            m = re.search(r"Rotate:\s*(\d+)", osd)
            angle = int(m.group(1)) if m else 0
            if angle:
                rotated = img.rotate(-angle, expand=True)
                r = run_grid(rotated)
                better_conf = r["conf"] >= 0 and r["conf"] > max(best["conf"], 0.0) + 10
                if better_conf or (r["conf"] < 0 and r["score"] > best["score"]):
                    r["engine"] += f":rot{angle}"
                    best = r
        except Exception as exc:
            errors.append(f"osd: {exc}")

    if best["text"].strip():
        return best["text"], f"{best['engine']}:conf{int(best['conf'])}" if best["conf"] >= 0 else best["engine"]
    return "", "; ".join(errors) or "tesseract-no-text"


def ocr_image_openai(path: Path) -> tuple[str, str]:
    if not (os.environ.get("OPENAI_API_KEY") and OpenAI):
        return "", "openai-not-configured"
    try:
        client = OpenAI()
        mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
        data = base64.b64encode(path.read_bytes()).decode("ascii")
        prompt = (
            "Extract all visible text from this document image. Support Arabic and English. "
            "Preserve line breaks, numbers, dates, totals, and names. Return only extracted text."
        )
        resp = client.chat.completions.create(
            model=os.environ.get("DOCWISE_VISION_MODEL", "gpt-4o-mini"),
            messages=[{"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{data}"}},
            ]}],
            temperature=0,
        )
        return resp.choices[0].message.content or "", "openai-vision"
    except Exception as exc:
        return "", f"openai-error: {exc}"


def ocr_image(path: Path) -> tuple[str, str]:
    engine = os.environ.get("DOCWISE_OCR", "auto").lower()
    if engine in ("openai", "vision"):
        txt, used = ocr_image_openai(path)
        if txt.strip():
            return txt, used
    txt, used = ocr_image_tesseract(path)
    if txt.strip() and not should_try_vision_fallback(txt, used):
        return txt, used
    txt2, used2 = ocr_image_openai(path)
    if txt2.strip() and ocr_text_score(txt2) >= ocr_text_score(txt):
        return txt2, f"{used2}:fallback-from-{used}"
    if txt.strip():
        return txt, used
    return "", f"{used}; {used2}"


def extract_text_from_pdf(path: Path) -> tuple[list[dict], str, Optional[str]]:
    pages = []
    engines = set()
    errors = []
    if not fitz:
        detail = FITZ_IMPORT_ERROR or "PyMuPDF is not installed"
        return [], "none", (
            f"PyMuPDF cannot load: {detail}. On Windows this is usually fixed by "
            "installing the Microsoft Visual C++ runtime "
            "(https://aka.ms/vs/17/release/vc_redist.x64.exe) or re-running setup.ps1."
        )
    try:
        doc = fitz.open(path)
        max_ocr_pages = int(os.environ.get("DOCWISE_MAX_OCR_PDF_PAGES", "20"))
        for i, page in enumerate(doc, start=1):
            text = page.get_text("text") or ""
            engine = "pdf-text"
            if should_ocr_pdf_page(text) and i <= max_ocr_pages:
                try:
                    # 3.5x ~= 250 dpi: measurably better Arabic diacritics and
                    # small print than the old 2.5x, at no real speed cost.
                    pix = page.get_pixmap(matrix=fitz.Matrix(3.5, 3.5), alpha=False)
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                        tmp_path = Path(tmp.name)
                    pix.save(str(tmp_path))
                    ocr_text, ocr_engine = ocr_image(tmp_path)
                    tmp_path.unlink(missing_ok=True)
                    if ocr_text.strip() and not text_is_garbled(ocr_text):
                        text = ocr_text
                        engine = f"pdf-render+{ocr_engine}"
                    elif ocr_text.strip() and len(ocr_text.strip()) > len(text.strip()):
                        text = ocr_text
                        engine = f"pdf-render+{ocr_engine}"
                    else:
                        errors.append(f"page {i}: {ocr_engine}")
                except Exception as exc:
                    errors.append(f"page {i} OCR failed: {exc}")
            engines.add(engine)
            pages.append({"page": i, "text": text})
        return pages, ",".join(sorted(engines)) or "pdf", "; ".join(errors) if errors else None
    except Exception as exc:
        return [], "pdf-error", str(exc)


def strip_rtf(text: str) -> str:
    text = re.sub(r"\\'[0-9a-fA-F]{2}", " ", text)
    text = re.sub(r"\\[a-zA-Z]+-?\d* ?", " ", text)
    text = text.replace("{", " ").replace("}", " ")
    return re.sub(r"\s+", " ", text).strip()


def extract_word_text(path: Path) -> tuple[list[dict], str, Optional[str]]:
    if not Document:
        return [], "docx-error", "python-docx is not installed"
    try:
        doc = Document(str(path))
        parts = []
        for p in doc.paragraphs:
            if p.text and p.text.strip():
                parts.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells if cell.text and cell.text.strip()]
                if cells:
                    parts.append(" | ".join(cells))
        return [{"page": 1, "text": "\n".join(parts)}], "docx-text", None
    except Exception as exc:
        return [], "docx-error", str(exc)


def extract_excel_text(path: Path) -> tuple[list[dict], str, Optional[str]]:
    if not load_workbook:
        return [], "xlsx-error", "openpyxl is not installed"
    try:
        wb = load_workbook(str(path), read_only=True, data_only=True)
        pages = []
        for i, ws in enumerate(wb.worksheets, start=1):
            lines = [f"Sheet: {ws.title}"]
            for row in ws.iter_rows(values_only=True):
                values = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if values:
                    lines.append(" | ".join(values))
            pages.append({"page": i, "text": "\n".join(lines)})
        wb.close()
        return pages, "xlsx-text", None
    except Exception as exc:
        return [], "xlsx-error", str(exc)


def extract_powerpoint_text(path: Path) -> tuple[list[dict], str, Optional[str]]:
    if not Presentation:
        return [], "pptx-error", "python-pptx is not installed"
    try:
        prs = Presentation(str(path))
        pages = []
        for i, slide in enumerate(prs.slides, start=1):
            parts = [f"Slide {i}"]
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text and shape.text.strip():
                    parts.append(shape.text.strip())
                if getattr(shape, "has_table", False):
                    for row in shape.table.rows:
                        cells = [cell.text.strip() for cell in row.cells if cell.text and cell.text.strip()]
                        if cells:
                            parts.append(" | ".join(cells))
            pages.append({"page": i, "text": "\n".join(parts)})
        return pages, "pptx-text", None
    except Exception as exc:
        return [], "pptx-error", str(exc)


def extract_legacy_office_text(path: Path) -> tuple[list[dict], str, Optional[str]]:
    """Best-effort extraction for old .doc/.xls/.ppt using installed Microsoft Office COM."""
    try:
        import win32com.client  # type: ignore
        ext = path.suffix.lower()
        if ext == ".doc":
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            text = doc.Content.Text
            doc.Close(False)
            word.Quit()
            return [{"page": 1, "text": text}], "word-com-text", None
        if ext == ".xls":
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(str(path.resolve()), ReadOnly=True)
            pages = []
            for i, ws in enumerate(wb.Worksheets, start=1):
                used = ws.UsedRange.Value
                lines = [f"Sheet: {ws.Name}"]
                if isinstance(used, tuple):
                    for row in used:
                        if not isinstance(row, tuple):
                            row = (row,)
                        values = [str(v).strip() for v in row if v is not None and str(v).strip()]
                        if values:
                            lines.append(" | ".join(values))
                elif used is not None:
                    lines.append(str(used))
                pages.append({"page": i, "text": "\n".join(lines)})
            wb.Close(False)
            excel.Quit()
            return pages, "excel-com-text", None
        if ext == ".ppt":
            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            pres = ppt.Presentations.Open(str(path.resolve()), WithWindow=False)
            pages = []
            for i, slide in enumerate(pres.Slides, start=1):
                parts = [f"Slide {i}"]
                for shape in slide.Shapes:
                    try:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            parts.append(shape.TextFrame.TextRange.Text)
                    except Exception:
                        pass
                pages.append({"page": i, "text": "\n".join(parts)})
            pres.Close()
            ppt.Quit()
            return pages, "powerpoint-com-text", None
    except Exception as exc:
        return [], "legacy-office-error", f"Legacy Office extraction failed. Install Microsoft Office or convert to docx/xlsx/pptx. Details: {exc}"
    return [], "legacy-office-error", "Unsupported legacy Office file"


def extract_text(path: Path) -> tuple[list[dict], str, Optional[str]]:
    ext = path.suffix.lower()
    if ext in TEXT_EXTS:
        for enc in ("utf-8", "utf-8-sig", "cp1256", "latin-1"):
            try:
                return [{"page": 1, "text": path.read_text(encoding=enc, errors="ignore")}], f"text:{enc}", None
            except Exception:
                pass
        return [], "text-error", "Could not decode text file"
    if ext == ".rtf":
        for enc in ("utf-8", "utf-8-sig", "cp1256", "latin-1"):
            try:
                return [{"page": 1, "text": strip_rtf(path.read_text(encoding=enc, errors="ignore"))}], f"rtf:{enc}", None
            except Exception:
                pass
        return [], "rtf-error", "Could not decode RTF file"
    if ext == ".pdf":
        return extract_text_from_pdf(path)
    if ext in IMAGE_EXTS:
        text, engine = ocr_image(path)
        return [{"page": 1, "text": text}], engine, None if text.strip() else engine
    if ext in WORD_EXTS:
        return extract_word_text(path)
    if ext in EXCEL_EXTS:
        return extract_excel_text(path)
    if ext in POWERPOINT_EXTS:
        return extract_powerpoint_text(path)
    if ext in LEGACY_OFFICE_EXTS:
        return extract_legacy_office_text(path)
    return [], "unsupported", "Unsupported file type"


def guess_metadata(path: Path, text: str) -> dict:
    lowered = normalize_arabic(text + " " + path.name)
    raw = text or ""
    lang = "Arabic + English" if has_arabic(raw) and re.search(r"[A-Za-z]", raw) else "Arabic" if has_arabic(raw) else "English" if re.search(r"[A-Za-z]", raw) else "unknown"

    rules = [
        ("news", ["news", "article", "مقال", "خبر", "اخبار", "كشفت", "اعلنت", "تقول", "هاتف", "شركة", "تقنيه"]),
        ("invoice", ["invoice", "فاتوره", "bill", "amount due", "total", "ضريبه", "vat", "aed", "درهم"]),
        ("contract", ["contract", "agreement", "عقد", "اتفاقيه", "lease", "rent", "ايجار"]),
        ("id", ["passport", "emirates id", "national id", "identity card", "جواز", "هويه", "بطاقه", "اقامه"]),
        ("receipt", ["receipt", "ايصال", "paid", "payment", "دفع", "مدفوع"]),
        ("bank", ["bank statement", "iban", "swift", "bank", "بنك", "كشف حساب", "ايبان"]),
        ("medical", ["medical", "hospital", "clinic", "doctor", "patient", "طبي", "مستشفي", "عياده", "مريض"]),
        ("certificate", ["certificate", "شهاده", "degree", "diploma"]),
        ("legal", ["court", "legal", "law", "محكمه", "قانون", "دعوي"]),
    ]
    doc_type = "general"
    for typ, kws in rules:
        if any(contains_keyword(lowered, k) for k in kws):
            doc_type = typ
            break

    date_guess = None
    date_patterns = [
        r"\b(20\d{2}[-/.]\d{1,2}[-/.]\d{1,2})\b",
        r"\b(\d{1,2}[-/.]\d{1,2}[-/.]20\d{2})\b",
        r"\b(\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+20\d{2})\b",
    ]
    for pat in date_patterns:
        m = re.search(pat, raw, re.I)
        if m:
            date_guess = m.group(1)
            break

    amount = None
    amount_patterns = [r"(?:AED|د\.إ|درهم)\s*([0-9,]+(?:\.\d{1,2})?)", r"([0-9,]+(?:\.\d{1,2})?)\s*(?:AED|د\.إ|درهم)"]
    for pat in amount_patterns:
        matches = re.findall(pat, raw, re.I)
        if matches:
            amount = matches[-1]
            break

    company = None
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    for ln in lines[:12]:
        if 2 <= len(ln) <= 80 and not text_is_garbled(ln) and not re.search(r"^[\d\W_]+$", ln):
            company = ln[:80]
            break

    title = path.stem.replace("_", " ").replace("-", " ").strip().title()
    tags = [doc_type, lang.lower().split()[0]]
    if amount:
        tags.append("amount")
    if date_guess:
        tags.append("date")

    summary = summarize_text(text, doc_type, lang, date_guess, amount, company)
    return {
        "title": title,
        "doc_type": doc_type,
        "language": lang,
        "date_guess": date_guess,
        "company": company,
        "amount": amount,
        "tags": sorted(set(tags)),
        "summary": summary,
    }


def summarize_text(text: str, doc_type: str, lang: str, date_guess: Optional[str], amount: Optional[str], company: Optional[str]) -> str:
    clean = re.sub(r"\s+", " ", text or "").strip()
    if not clean:
        return "No text extracted yet. Install Tesseract with Arabic language data or set OPENAI_API_KEY for image OCR."
    prefix = f"{doc_type.title()} document"
    if company:
        prefix += f" from {company}"
    bits = []
    if date_guess:
        bits.append(f"date: {date_guess}")
    if amount:
        bits.append(f"amount: {amount}")
    if lang != "unknown":
        bits.append(f"language: {lang}")
    meta = f" ({', '.join(bits)})" if bits else ""
    return f"{prefix}{meta}. {clean[:420]}{'...' if len(clean) > 420 else ''}"


def extract_structured_fields(text: str, meta: dict) -> dict:
    """Document-type fields used before/alongside RAG answers."""
    raw = text or ""
    fields = {
        "doc_type": meta.get("doc_type"),
        "title": meta.get("title"),
        "company": meta.get("company"),
        "date": meta.get("date_guess"),
        "amount": meta.get("amount"),
    }
    # Common fields
    m = re.search(r"(?:invoice\s*(?:no|number|#)|رقم\s*الفاتوره|فاتوره\s*رقم)\s*[:#-]?\s*([A-Z0-9\-/]+)", raw, re.I)
    if m:
        fields["invoice_number"] = m.group(1)
    m = re.search(r"(?:due\s*date|expiry\s*date|تاريخ\s*(?:الاستحقاق|الانتهاء|انتهاء))\s*[:#-]?\s*([0-9A-Za-z\-/\.\s]+)", raw, re.I)
    if m:
        fields["due_or_expiry_date"] = m.group(1).strip()[:60]
    amounts = re.findall(r"(?:AED|د\.إ|درهم)\s*([0-9,]+(?:\.\d{1,2})?)|([0-9,]+(?:\.\d{1,2})?)\s*(?:AED|د\.إ|درهم)", raw, re.I)
    clean_amounts = [a or b for a, b in amounts if (a or b)]
    if clean_amounts:
        fields["amounts_found"] = clean_amounts[-8:]
        fields["amount"] = fields.get("amount") or clean_amounts[-1]
    if meta.get("doc_type") == "contract":
        rent = re.search(r"(?:rent|annual rent|الايجار|قيمة الايجار)\s*[:#-]?\s*([0-9,]+(?:\.\d{1,2})?\s*(?:AED|درهم|د\.إ)?)", raw, re.I)
        if rent:
            fields["rent_amount"] = rent.group(1)
    if meta.get("doc_type") in ("id", "legal", "contract"):
        names = re.findall(r"(?:name|الاسم)\s*[:#-]?\s*([^\n]{3,80})", raw, re.I)
        if names:
            fields["names_found"] = [n.strip() for n in names[:6]]
    # Optional schema extraction with OpenAI for better invoices/contracts/IDs.
    if os.environ.get("OPENAI_API_KEY") and OpenAI and os.environ.get("DOCWISE_FIELD_AI", "1") == "1" and raw.strip():
        try:
            client = OpenAI()
            schema_prompt = (
                "Extract structured fields from this document as compact JSON only. "
                "Support Arabic/English. Include only fields clearly present. "
                "Useful keys: vendor, customer, invoice_number, date, due_date, total, currency, tax, parties, start_date, end_date, id_number, expiry_date, category."
            )
            resp = client.chat.completions.create(
                model=os.environ.get("DOCWISE_CHAT_MODEL", "gpt-4o-mini"),
                messages=[{"role": "user", "content": f"{schema_prompt}\n\nDocument type: {meta.get('doc_type')}\nText:\n{raw[:6000]}"}],
                temperature=0,
            )
            content = resp.choices[0].message.content or "{}"
            content = re.sub(r"^```json|```$", "", content.strip(), flags=re.I).strip()
            ai_fields = json.loads(content)
            if isinstance(ai_fields, dict):
                fields["ai"] = ai_fields
        except Exception as exc:
            fields["ai_error"] = str(exc)[:160]
    return {k: v for k, v in fields.items() if v not in (None, "", [], {})}


def chunk_pages(pages: list[dict], max_tokens: int = 700, overlap_tokens: int = 90) -> list[dict]:
    """Structure-aware-ish chunking: paragraphs/lines first, token window fallback, page citations preserved."""
    chunks = []
    for page in pages:
        raw = (page.get("text") or "").strip()
        if not raw:
            continue
        # Keep table rows/slide lines together when possible; don't flatten everything immediately.
        blocks = [b.strip() for b in re.split(r"\n\s*\n|(?<=\.)\s+(?=[A-Z\u0600-\u06FF])", raw) if b.strip()]
        if len(blocks) <= 1:
            blocks = [b.strip() for b in raw.splitlines() if b.strip()] or [raw]
        idx = 0
        current = []
        current_tokens = 0
        for block in blocks:
            bt = approx_token_count(block)
            if current and current_tokens + bt > max_tokens:
                text = "\n".join(current).strip()
                chunks.append({"page": page.get("page", 1), "chunk_index": idx, "text": text, "token_count": approx_token_count(text)})
                idx += 1
                # Lightweight overlap by carrying trailing lines.
                carry = []
                carry_tokens = 0
                for prev in reversed(current):
                    pt = approx_token_count(prev)
                    if carry_tokens + pt > overlap_tokens:
                        break
                    carry.insert(0, prev)
                    carry_tokens += pt
                current = carry
                current_tokens = carry_tokens
            if bt > max_tokens:
                words = block.split()
                step = max(1, int(max_tokens - overlap_tokens))
                for start in range(0, len(words), step):
                    part = " ".join(words[start:start + max_tokens]).strip()
                    if part:
                        chunks.append({"page": page.get("page", 1), "chunk_index": idx, "text": part, "token_count": approx_token_count(part)})
                        idx += 1
                current = []
                current_tokens = 0
            else:
                current.append(block)
                current_tokens += bt
        if current:
            text = "\n".join(current).strip()
            chunks.append({"page": page.get("page", 1), "chunk_index": idx, "text": text, "token_count": approx_token_count(text)})
    return chunks


def index_file(path: Path, source_type: str = "folder", force: bool = False) -> dict:
    path = path.resolve()
    if path.suffix.lower() not in SUPPORTED or not path.is_file():
        return {"status": "skipped", "path": str(path), "reason": "unsupported"}
    stat = path.stat()
    sha = file_hash(path)

    with DB_LOCK, db() as conn:
        existing = conn.execute("SELECT id, sha256, mtime FROM documents WHERE path=?", (str(path),)).fetchone()
        if existing and not force and existing["sha256"] == sha and abs((existing["mtime"] or 0) - stat.st_mtime) < 0.001:
            return {"status": "unchanged", "id": existing["id"], "path": str(path)}

    pages, engine, err = extract_text(path)
    full_text = "\n\n".join([p.get("text") or "" for p in pages]).strip()
    meta = guess_metadata(path, full_text)
    fields = extract_structured_fields(full_text, meta)
    quality, quality_score = ocr_quality(full_text)
    normalized = normalize_arabic(full_text + " " + meta["title"] + " " + " ".join(meta["tags"]) + " " + json.dumps(fields, ensure_ascii=False))
    status = "indexed" if full_text else "needs_ocr"
    if err and not full_text:
        status = "error" if "Unsupported" in err else "needs_ocr"

    with DB_LOCK, db() as conn:
        cur = conn.execute(
            """
            INSERT INTO documents(path, original_name, source_type, file_ext, size, mtime, sha256, title, doc_type, language,
                date_guess, company, amount, tags, summary, fields, text, normalized_text, status, ocr_engine, ocr_quality, ocr_score, error, created_at, updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(path) DO UPDATE SET
                original_name=excluded.original_name, source_type=excluded.source_type, file_ext=excluded.file_ext,
                size=excluded.size, mtime=excluded.mtime, sha256=excluded.sha256, title=excluded.title,
                doc_type=excluded.doc_type, language=excluded.language, date_guess=excluded.date_guess,
                company=excluded.company, amount=excluded.amount, tags=excluded.tags, summary=excluded.summary, fields=excluded.fields,
                text=excluded.text, normalized_text=excluded.normalized_text, status=excluded.status,
                ocr_engine=excluded.ocr_engine, ocr_quality=excluded.ocr_quality, ocr_score=excluded.ocr_score, error=excluded.error, updated_at=excluded.updated_at
            """,
            (
                str(path), path.name, source_type, path.suffix.lower(), stat.st_size, stat.st_mtime, sha,
                meta["title"], meta["doc_type"], meta["language"], meta["date_guess"], meta["company"], meta["amount"],
                json.dumps(meta["tags"], ensure_ascii=False), meta["summary"], json.dumps(fields, ensure_ascii=False),
                full_text, normalized, status, engine, quality, quality_score, err, now_iso(), now_iso(),
            ),
        )
        doc_id_row = conn.execute("SELECT id FROM documents WHERE path=?", (str(path),)).fetchone()
        doc_id = doc_id_row["id"]
        delete_chunk_indexes_for_doc(conn, doc_id)
        conn.execute("DELETE FROM chunks WHERE document_id=?", (doc_id,))
        for ch in chunk_pages(pages):
            norm_ch = normalize_arabic(ch["text"])
            cur = conn.execute(
                "INSERT INTO chunks(document_id, page, chunk_index, text, normalized_text, token_count) VALUES(?,?,?,?,?,?)",
                (doc_id, ch["page"], ch["chunk_index"], ch["text"], norm_ch, ch.get("token_count") or approx_token_count(ch["text"])),
            )
            add_chunk_indexes(conn, cur.lastrowid, doc_id, meta["title"], ch["text"], norm_ch)
        return {"status": status, "id": doc_id, "path": str(path), "engine": engine, "error": err}


def row_to_doc(row: sqlite3.Row, include_text: bool = False) -> dict:
    d = dict(row)
    d["tags"] = safe_json_load(d.get("tags"), [])
    d["fields"] = safe_json_load(d.get("fields"), {})
    if not include_text:
        d.pop("text", None)
        d.pop("normalized_text", None)
    return d


def scan_folder(path: Path, recursive: bool = True, force: bool = False) -> dict:
    path = path.expanduser().resolve()
    if not path.exists() or not path.is_dir():
        raise HTTPException(status_code=400, detail="Folder does not exist")
    pattern = "**/*" if recursive else "*"
    files = [p for p in path.glob(pattern) if p.is_file() and p.suffix.lower() in SUPPORTED]
    result = {"folder": str(path), "found": len(files), "indexed": 0, "unchanged": 0, "skipped": 0, "errors": 0, "items": []}
    SCAN_STATE.update({"running": True, "message": f"Scanning {path}", "indexed": 0, "errors": 0})
    try:
        for p in files:
            try:
                item = index_file(p, "folder", force=force)
                result["items"].append(item)
                if item["status"] in ("indexed", "needs_ocr", "error"):
                    result["indexed"] += 1
                elif item["status"] == "unchanged":
                    result["unchanged"] += 1
                else:
                    result["skipped"] += 1
                if item["status"] == "error":
                    result["errors"] += 1
            except Exception as exc:
                result["errors"] += 1
                result["items"].append({"status": "error", "path": str(p), "error": str(exc)})
            SCAN_STATE.update({"indexed": result["indexed"], "errors": result["errors"]})
    finally:
        SCAN_STATE.update({"running": False, "last": now_iso(), "message": "Scan complete"})
    with DB_LOCK, db() as conn:
        conn.execute("UPDATE folders SET last_scan=? WHERE path=?", (now_iso(), str(path)))
    return result


def watcher_loop():
    while True:
        time.sleep(int(os.environ.get("DOCWISE_WATCH_INTERVAL", "45")))
        try:
            with DB_LOCK, db() as conn:
                folders = conn.execute("SELECT path, recursive FROM folders WHERE watch=1").fetchall()
            if SCAN_STATE["running"]:
                continue
            for f in folders:
                try:
                    scan_folder(Path(f["path"]), bool(f["recursive"]), force=False)
                except Exception as exc:
                    SCAN_STATE.update({"message": f"Watch scan error: {exc}", "errors": SCAN_STATE.get("errors", 0) + 1})
        except Exception:
            pass


threading.Thread(target=watcher_loop, daemon=True).start()


class FolderRequest(BaseModel):
    path: str
    recursive: bool = True
    watch: bool = True
    scan_now: bool = True


class SearchRequest(BaseModel):
    q: str = ""
    doc_type: str = "all"
    language: str = "all"
    limit: int = 50


class AskRequest(BaseModel):
    question: str
    limit: int = 6
    use_ai: bool = True


class OrganizeRequest(BaseModel):
    document_id: int
    mode: str = "copy"  # copy or move


class BulkDeleteRequest(BaseModel):
    ids: list[int] = []
    all: bool = False
    delete_files: bool = False
    app_files_only: bool = True


class TextUpdateRequest(BaseModel):
    text: str


class EvalCaseRequest(BaseModel):
    question: str
    expected: str = ""
    must_cite: str = ""


@app.get("/api/status")
def status():
    with DB_LOCK, db() as conn:
        total = conn.execute("SELECT COUNT(*) c FROM documents").fetchone()["c"]
        needs = conn.execute("SELECT COUNT(*) c FROM documents WHERE status!='indexed'").fetchone()["c"]
        folders = conn.execute("SELECT COUNT(*) c FROM folders").fetchone()["c"]
        chunks = conn.execute("SELECT COUNT(*) c FROM chunks").fetchone()["c"]
        embeddings = conn.execute("SELECT COUNT(*) c FROM chunk_embeddings WHERE dims>0").fetchone()["c"]
        fts_rows = conn.execute("SELECT COUNT(*) c FROM chunk_fts").fetchone()["c"]
        quality_rows = conn.execute("SELECT ocr_quality, COUNT(*) c FROM documents GROUP BY ocr_quality").fetchall()
        ocr_quality_counts = {r["ocr_quality"] or "unknown": r["c"] for r in quality_rows}
    return {"ok": True, "documents": total, "needs_review": needs, "folders": folders, "chunks": chunks, "embeddings": embeddings, "fts_rows": fts_rows, "ocr_quality_counts": ocr_quality_counts, "scan": SCAN_STATE, "ocr": available_ocr()}


@app.post("/api/upload")
async def upload(files: list[UploadFile] = File(...)):
    results = []
    for f in files:
        ext = Path(f.filename or "file").suffix.lower()
        if ext not in SUPPORTED:
            results.append({"filename": f.filename, "status": "skipped", "reason": "unsupported"})
            continue
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        safe = re.sub(r"[^A-Za-z0-9_.\-\u0600-\u06FF]+", "_", f.filename or f"upload{ext}")
        dest = UPLOADS / f"{stamp}_{safe}"
        with dest.open("wb") as out:
            shutil.copyfileobj(f.file, out)
        try:
            results.append(index_file(dest, "upload", force=True))
        except Exception as exc:
            results.append({"filename": f.filename, "status": "error", "error": str(exc)})
    return {"results": results}


@app.post("/api/folders")
def add_folder(req: FolderRequest):
    folder = Path(req.path).expanduser().resolve()
    if not folder.exists() or not folder.is_dir():
        raise HTTPException(status_code=400, detail="Folder does not exist")
    with DB_LOCK, db() as conn:
        conn.execute(
            "INSERT INTO folders(path, recursive, watch, created_at) VALUES(?,?,?,?) ON CONFLICT(path) DO UPDATE SET recursive=excluded.recursive, watch=excluded.watch",
            (str(folder), int(req.recursive), int(req.watch), now_iso()),
        )
    result = None
    if req.scan_now:
        result = scan_folder(folder, req.recursive, force=False)
    return {"folder": str(folder), "scan": result}


@app.get("/api/folders")
def list_folders():
    with DB_LOCK, db() as conn:
        rows = conn.execute("SELECT * FROM folders ORDER BY created_at DESC").fetchall()
    return {"folders": [dict(r) for r in rows]}


@app.delete("/api/folders/{folder_id}")
def delete_folder(folder_id: int):
    with DB_LOCK, db() as conn:
        conn.execute("DELETE FROM folders WHERE id=?", (folder_id,))
    return {"ok": True}


@app.post("/api/scan")
def manual_scan(req: FolderRequest):
    return scan_folder(Path(req.path), recursive=req.recursive, force=True)


@app.get("/api/documents")
def documents(limit: int = 80, offset: int = 0):
    with DB_LOCK, db() as conn:
        rows = conn.execute("SELECT * FROM documents ORDER BY updated_at DESC LIMIT ? OFFSET ?", (limit, offset)).fetchall()
    return {"documents": [row_to_doc(r) for r in rows]}


@app.get("/api/documents/{doc_id}")
def document(doc_id: int):
    with DB_LOCK, db() as conn:
        row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
        chunks = conn.execute("SELECT page, chunk_index, text FROM chunks WHERE document_id=? ORDER BY page, chunk_index", (doc_id,)).fetchall()
    if not row:
        raise HTTPException(status_code=404, detail="Document not found")
    d = row_to_doc(row, include_text=True)
    d["chunks"] = [dict(c) for c in chunks]
    d["suggestion"] = filing_suggestion(d)
    return d


def is_app_managed_file(path: Path) -> bool:
    try:
        resolved = path.resolve()
        return resolved.is_relative_to(UPLOADS.resolve()) or resolved.is_relative_to(ARCHIVE.resolve())
    except Exception:
        return False


def delete_docs_by_ids(ids: list[int], delete_files: bool = False, app_files_only: bool = True) -> dict:
    ids = sorted(set(int(i) for i in ids if int(i) > 0))
    if not ids:
        return {"ok": True, "deleted_count": 0, "file_deleted_count": 0, "items": []}
    placeholders = ",".join("?" for _ in ids)
    items = []
    file_deleted_count = 0
    with DB_LOCK, db() as conn:
        rows = conn.execute(f"SELECT id, path FROM documents WHERE id IN ({placeholders})", ids).fetchall()
        for row in rows:
            path = Path(row["path"])
            item = {"id": row["id"], "path": str(path), "index_deleted": True, "file_deleted": False, "file_skipped": False, "file_error": None}
            if delete_files:
                if app_files_only and not is_app_managed_file(path):
                    item["file_skipped"] = True
                    item["file_error"] = "Skipped external file because app_files_only=true"
                else:
                    try:
                        if path.exists() and path.is_file():
                            path.unlink()
                            item["file_deleted"] = True
                            file_deleted_count += 1
                    except Exception as exc:
                        item["file_error"] = str(exc)
            items.append(item)
        for doc_id in ids:
            delete_chunk_indexes_for_doc(conn, doc_id)
        conn.execute(f"DELETE FROM chunks WHERE document_id IN ({placeholders})", ids)
        conn.execute(f"DELETE FROM documents WHERE id IN ({placeholders})", ids)
    return {"ok": True, "deleted_count": len(items), "file_deleted_count": file_deleted_count, "items": items}


@app.put("/api/documents/{doc_id}/text")
def update_document_text(doc_id: int, req: TextUpdateRequest):
    """Save manually corrected OCR text and rebuild search/RAG chunks."""
    text = req.text or ""
    with DB_LOCK, db() as conn:
        row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Document not found")
        path = Path(row["path"])
        meta = guess_metadata(path, text)
        fields = extract_structured_fields(text, meta)
        quality, quality_score = ocr_quality(text)
        normalized = normalize_arabic(text + " " + meta["title"] + " " + " ".join(meta["tags"]) + " " + json.dumps(fields, ensure_ascii=False))
        status = "indexed" if text.strip() else "needs_ocr"
        conn.execute(
            """
            UPDATE documents SET title=?, doc_type=?, language=?, date_guess=?, company=?, amount=?, tags=?, summary=?, fields=?,
                text=?, normalized_text=?, status=?, ocr_engine=?, ocr_quality=?, ocr_score=?, error=?, updated_at=? WHERE id=?
            """,
            (
                meta["title"], meta["doc_type"], meta["language"], meta["date_guess"], meta["company"], meta["amount"],
                json.dumps(meta["tags"], ensure_ascii=False), meta["summary"], json.dumps(fields, ensure_ascii=False), text, normalized, status,
                "manual-correction", quality, quality_score, None, now_iso(), doc_id,
            ),
        )
        delete_chunk_indexes_for_doc(conn, doc_id)
        conn.execute("DELETE FROM chunks WHERE document_id=?", (doc_id,))
        for ch in chunk_pages([{"page": 1, "text": text}]):
            norm_ch = normalize_arabic(ch["text"])
            cur = conn.execute(
                "INSERT INTO chunks(document_id, page, chunk_index, text, normalized_text, token_count) VALUES(?,?,?,?,?,?)",
                (doc_id, ch["page"], ch["chunk_index"], ch["text"], norm_ch, ch.get("token_count") or approx_token_count(ch["text"])),
            )
            add_chunk_indexes(conn, cur.lastrowid, doc_id, meta["title"], ch["text"], norm_ch)
    return {"ok": True, "id": doc_id, "status": status}


@app.post("/api/documents/{doc_id}/vision-ocr")
def force_vision_ocr(doc_id: int):
    """Force OpenAI Vision OCR for one image/PDF and rebuild metadata/RAG indexes."""
    if not (os.environ.get("OPENAI_API_KEY") and OpenAI):
        raise HTTPException(status_code=400, detail="OPENAI_API_KEY is not configured")
    with DB_LOCK, db() as conn:
        row = conn.execute("SELECT path FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Document not found")
    path = Path(row["path"])
    if not path.exists():
        raise HTTPException(status_code=404, detail="Original file not found")
    ext = path.suffix.lower()
    pages = []
    engine = "openai-vision"
    errors = []
    if ext in IMAGE_EXTS:
        text, used = ocr_image_openai(path)
        engine = used
        pages = [{"page": 1, "text": text}]
        if not text.strip():
            errors.append(used)
    elif ext == ".pdf":
        if not fitz:
            raise HTTPException(status_code=400, detail="PyMuPDF is not installed")
        doc = fitz.open(path)
        max_pages = int(os.environ.get("DOCWISE_MAX_OCR_PDF_PAGES", "20"))
        for i, page in enumerate(doc, start=1):
            if i > max_pages:
                break
            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp_path = Path(tmp.name)
            pix.save(str(tmp_path))
            text, used = ocr_image_openai(tmp_path)
            tmp_path.unlink(missing_ok=True)
            pages.append({"page": i, "text": text})
            engine = f"pdf-render+{used}"
            if not text.strip():
                errors.append(f"page {i}: {used}")
    else:
        raise HTTPException(status_code=400, detail="Vision OCR supports images and PDFs only")

    full_text = "\n\n".join(p.get("text") or "" for p in pages).strip()
    meta = guess_metadata(path, full_text)
    fields = extract_structured_fields(full_text, meta)
    quality, quality_score = ocr_quality(full_text)
    normalized = normalize_arabic(full_text + " " + meta["title"] + " " + " ".join(meta["tags"]) + " " + json.dumps(fields, ensure_ascii=False))
    status = "indexed" if full_text else "needs_ocr"
    err = "; ".join(errors) if errors else None
    with DB_LOCK, db() as conn:
        conn.execute(
            """
            UPDATE documents SET title=?, doc_type=?, language=?, date_guess=?, company=?, amount=?, tags=?, summary=?, fields=?,
                text=?, normalized_text=?, status=?, ocr_engine=?, ocr_quality=?, ocr_score=?, error=?, updated_at=? WHERE id=?
            """,
            (meta["title"], meta["doc_type"], meta["language"], meta["date_guess"], meta["company"], meta["amount"],
             json.dumps(meta["tags"], ensure_ascii=False), meta["summary"], json.dumps(fields, ensure_ascii=False),
             full_text, normalized, status, engine, quality, quality_score, err, now_iso(), doc_id),
        )
        delete_chunk_indexes_for_doc(conn, doc_id)
        conn.execute("DELETE FROM chunks WHERE document_id=?", (doc_id,))
        for ch in chunk_pages(pages):
            norm_ch = normalize_arabic(ch["text"])
            cur = conn.execute(
                "INSERT INTO chunks(document_id, page, chunk_index, text, normalized_text, token_count) VALUES(?,?,?,?,?,?)",
                (doc_id, ch["page"], ch["chunk_index"], ch["text"], norm_ch, ch.get("token_count") or approx_token_count(ch["text"])),
            )
            add_chunk_indexes(conn, cur.lastrowid, doc_id, meta["title"], ch["text"], norm_ch)
    return {"ok": True, "id": doc_id, "status": status, "engine": engine, "ocr_quality": quality, "ocr_score": quality_score, "error": err}


@app.delete("/api/documents/{doc_id}")
def delete_document(doc_id: int):
    """Remove only the DocWise index record. The original file stays untouched."""
    result = delete_docs_by_ids([doc_id], delete_files=False)
    result["deleted"] = "index_only"
    return result


@app.delete("/api/documents/{doc_id}/file")
def delete_document_and_file(doc_id: int):
    """Delete the original file from disk, then remove the DocWise index record."""
    result = delete_docs_by_ids([doc_id], delete_files=True, app_files_only=False)
    result["deleted"] = "file_and_index"
    return result


@app.post("/api/documents/bulk-delete")
def bulk_delete_documents(req: BulkDeleteRequest):
    with DB_LOCK, db() as conn:
        if req.all:
            ids = [r["id"] for r in conn.execute("SELECT id FROM documents").fetchall()]
        else:
            ids = req.ids
    return delete_docs_by_ids(ids, delete_files=req.delete_files, app_files_only=req.app_files_only)


@app.post("/api/reindex/{doc_id}")
def reindex(doc_id: int):
    with DB_LOCK, db() as conn:
        row = conn.execute("SELECT path FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Document not found")
    return index_file(Path(row["path"]), force=True)


@app.post("/api/rebuild-rag")
def rebuild_rag_indexes():
    """Rebuild FTS/vector indexes for already extracted documents without re-OCRing files."""
    rebuilt = 0
    with DB_LOCK, db() as conn:
        docs = conn.execute("SELECT id, title FROM documents WHERE status='indexed'").fetchall()
        for doc in docs:
            delete_chunk_indexes_for_doc(conn, doc["id"])
            chunks = conn.execute("SELECT * FROM chunks WHERE document_id=? ORDER BY page, chunk_index", (doc["id"],)).fetchall()
            for ch in chunks:
                add_chunk_indexes(conn, ch["id"], doc["id"], doc["title"], ch["text"], ch["normalized_text"])
                rebuilt += 1
    return {"ok": True, "rebuilt_chunks": rebuilt}


@app.post("/api/search")
def search(req: SearchRequest):
    tokens = query_terms(req.q)
    if req.q.strip():
        retrieved = hybrid_retrieve(req.q, limit=max(30, req.limit * 3), doc_type=None if req.doc_type == "all" else req.doc_type)
        doc_map = {}
        for ch in retrieved:
            doc_id = ch["document_id"]
            if doc_id not in doc_map or ch["score"] > doc_map[doc_id]["score"]:
                doc_map[doc_id] = ch
        ids = list(doc_map.keys())[: max(1, min(req.limit, 100))]
        if not ids:
            return {"results": []}
        placeholders = ",".join("?" for _ in ids)
        with DB_LOCK, db() as conn:
            rows = conn.execute(f"SELECT * FROM documents WHERE id IN ({placeholders})", ids).fetchall()
        by_id = {r["id"]: row_to_doc(r, include_text=True) for r in rows}
        results = []
        for doc_id in ids:
            d = by_id.get(doc_id)
            if not d:
                continue
            if req.language != "all" and req.language not in (d.get("language") or ""):
                continue
            best = doc_map[doc_id]
            d["score"] = round(best["score"], 4)
            d["snippet"] = make_snippet(best.get("text") or d.get("summary") or "", tokens)
            d["retrieval"] = best.get("retrieval", "hybrid")
            d.pop("text", None)
            d.pop("normalized_text", None)
            results.append(d)
        return {"results": results}

    with DB_LOCK, db() as conn:
        sql = "SELECT * FROM documents WHERE 1=1"
        params = []
        if req.doc_type != "all":
            sql += " AND doc_type=?"
            params.append(req.doc_type)
        if req.language != "all":
            sql += " AND language LIKE ?"
            params.append(f"%{req.language}%")
        rows = conn.execute(sql + " ORDER BY updated_at DESC LIMIT ?", params + [max(1, min(req.limit, 100))]).fetchall()
    return {"results": [row_to_doc(r) for r in rows]}


def make_snippet(text: str, tokens: list[str]) -> str:
    clean = re.sub(r"\s+", " ", text or "").strip()
    if not clean:
        return ""
    norm = normalize_arabic(clean)
    idx = 0
    for t in tokens:
        pos = norm.find(t)
        if pos >= 0:
            idx = max(0, min(len(clean), pos) - 100)
            break
    return clean[idx:idx + 360] + ("..." if len(clean) > idx + 360 else "")


def query_terms(text: str) -> list[str]:
    base = [t for t in normalize_arabic(text).split() if len(t) > 1]
    synonyms = {
        "فاتوره": ["invoice", "bill", "total", "amount", "aed", "درهم", "اجمالي", "مبلغ"],
        "الفاتوره": ["فاتوره", "invoice", "bill", "total", "amount", "aed", "درهم", "اجمالي", "مبلغ"],
        "قيمه": ["amount", "total", "مبلغ", "اجمالي", "aed", "درهم"],
        "المبلغ": ["amount", "total", "مبلغ", "اجمالي", "aed", "درهم"],
        "كم": ["amount", "total", "مبلغ", "اجمالي"],
        "عقد": ["contract", "agreement", "lease", "rent", "ايجار"],
        "هويه": ["identity", "passport", "emirates", "id", "expiry", "انتهاء"],
        "تاريخ": ["date", "expiry", "due", "تاريخ"],
    }
    out = []
    for t in base:
        out.append(t)
        for prefix in ("ال", "و", "ب", "بال", "لل"):
            if t.startswith(prefix) and len(t) > len(prefix) + 1:
                out.append(t[len(prefix):])
        out.extend(synonyms.get(t, []))
    seen = set()
    uniq = []
    for t in out:
        nt = normalize_arabic(t)
        if nt and nt not in seen:
            seen.add(nt)
            uniq.append(nt)
    return uniq


def chunk_score(question: str, chunk_norm: str) -> int:
    tokens = query_terms(question)
    if not tokens:
        return 0
    words = set(chunk_norm.split())
    score = 0
    for t in tokens:
        if t in words:
            score += 4
        if t in chunk_norm:
            score += chunk_norm.count(t)
        if len(t) > 3 and any(w.startswith(t) or t.startswith(w) for w in words if len(w) > 3):
            score += 1
    return score


def rerank_candidate(candidate: dict, question: str, plan: dict) -> float:
    score = float(candidate.get("score", 0.0))
    text = candidate.get("text") or ""
    norm = candidate.get("normalized_text") or normalize_arabic(text)
    score += min(3.0, chunk_score(question, norm) * 0.08)
    if plan.get("doc_type") and candidate.get("doc_type") == plan["doc_type"]:
        score += 1.2
    if plan.get("field") == "amount" and re.search(r"(?:AED|د\.إ|درهم|total|amount|اجمالي|مبلغ|\d+[,.]?\d*)", text, re.I):
        score += 0.9
    if plan.get("field") == "date" and re.search(r"20\d{2}|\d{1,2}[-/.]\d{1,2}|expiry|due|انتهاء|تاريخ", text, re.I):
        score += 0.7
    if plan.get("sort") == "latest" and candidate.get("updated_at"):
        score += 0.15
    return score


def gpt_rerank(question: str, candidates: list[dict], limit: int) -> list[dict]:
    if not (os.environ.get("OPENAI_API_KEY") and OpenAI) or os.environ.get("DOCWISE_GPT_RERANK", "1") != "1":
        return candidates[:limit]
    try:
        client = OpenAI()
        packed = []
        for i, c in enumerate(candidates[:20], start=1):
            packed.append({"n": i, "title": c.get("title"), "type": c.get("doc_type"), "page": c.get("page"), "text": (c.get("text") or "")[:900]})
        resp = client.chat.completions.create(
            model=os.environ.get("DOCWISE_CHAT_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": "Rerank document chunks for answering the user question. Return JSON only: {\"order\":[candidate_numbers_best_first]}. Prefer exact supporting evidence."},
                {"role": "user", "content": json.dumps({"question": question, "candidates": packed}, ensure_ascii=False)},
            ],
            temperature=0,
        )
        content = (resp.choices[0].message.content or "{}").strip()
        content = re.sub(r"^```json|```$", "", content, flags=re.I).strip()
        order = json.loads(content).get("order", [])
        by_num = {i: c for i, c in enumerate(candidates[:20], start=1)}
        reranked = [by_num[n] for n in order if n in by_num]
        seen = {id(c) for c in reranked}
        reranked.extend([c for c in candidates if id(c) not in seen])
        for c in reranked[:limit]:
            c["retrieval"] = (c.get("retrieval") or "hybrid") + "+gpt-rerank"
        return reranked[:limit]
    except Exception:
        return candidates[:limit]


def verify_answer(question: str, answer: str, sources: list[dict]) -> dict:
    if not (os.environ.get("OPENAI_API_KEY") and OpenAI) or os.environ.get("DOCWISE_VERIFY", "1") != "1":
        return {"enabled": False}
    try:
        client = OpenAI()
        resp = client.chat.completions.create(
            model=os.environ.get("DOCWISE_CHAT_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": "Verify if the answer is fully supported by the sources. Return JSON only with keys supported:boolean, confidence:0-1, reason:string."},
                {"role": "user", "content": json.dumps({"question": question, "answer": answer, "sources": sources}, ensure_ascii=False)},
            ],
            temperature=0,
        )
        content = (resp.choices[0].message.content or "{}").strip()
        content = re.sub(r"^```json|```$", "", content, flags=re.I).strip()
        data = json.loads(content)
        data["enabled"] = True
        return data
    except Exception as exc:
        return {"enabled": True, "error": str(exc)[:160]}


def hybrid_retrieve(question: str, limit: int = 8, doc_type: Optional[str] = None) -> list[dict]:
    """Hybrid RAG retrieval: SQLite FTS5/BM25 + vector embeddings + Arabic-aware lexical reranking."""
    plan = query_plan(question)
    if doc_type:
        plan["doc_type"] = doc_type
    terms = plan.get("terms") or query_terms(question)
    fts_query = fts_query_from_terms(terms)
    merged: dict[int, dict] = {}

    with DB_LOCK, db() as conn:
        params_base = []
        doc_filter = ""
        if plan.get("doc_type"):
            doc_filter = " AND documents.doc_type=?"
            params_base.append(plan["doc_type"])

        if fts_query:
            try:
                rows = conn.execute(
                    f"""
                    SELECT chunks.*, documents.title, documents.path, documents.doc_type, documents.updated_at,
                           documents.date_guess, documents.amount, documents.company, bm25(chunk_fts) AS bm25_rank
                    FROM chunk_fts
                    JOIN chunks ON chunks.id=chunk_fts.chunk_id
                    JOIN documents ON documents.id=chunks.document_id
                    WHERE chunk_fts MATCH ? AND documents.status='indexed'{doc_filter}
                    ORDER BY bm25_rank LIMIT 60
                    """,
                    [fts_query] + params_base,
                ).fetchall()
                for r in rows:
                    d = dict(r)
                    # FTS5 bm25 is usually negative; more negative is better.
                    d["score"] = 2.0 + max(0.0, -float(d.get("bm25_rank") or 0.0))
                    d["retrieval"] = "fts5-bm25"
                    merged[d["id"]] = d
            except Exception:
                pass

        # Vector retrieval. Uses OpenAI embeddings when configured; local deterministic vector otherwise.
        try:
            q_emb, q_model = embed_text(question)
            rows = conn.execute(
                f"""
                SELECT chunks.*, documents.title, documents.path, documents.doc_type, documents.updated_at,
                       documents.date_guess, documents.amount, documents.company, chunk_embeddings.embedding,
                       chunk_embeddings.dims, chunk_embeddings.model
                FROM chunk_embeddings
                JOIN chunks ON chunks.id=chunk_embeddings.chunk_id
                JOIN documents ON documents.id=chunks.document_id
                WHERE chunk_embeddings.dims>0 AND documents.status='indexed'{doc_filter}
                """,
                params_base,
            ).fetchall()
            vector_hits = []
            for r in rows:
                try:
                    emb = json.loads(r["embedding"])
                    sim = cosine_similarity(q_emb, emb)
                    min_sim = 0.18 if str(r["model"]).startswith("local-hash") else 0.12
                    if sim > min_sim:
                        d = dict(r)
                        d.pop("embedding", None)
                        d["score"] = 1.5 + sim * 3.0
                        d["retrieval"] = f"vector:{r['model']}"
                        vector_hits.append(d)
                except Exception:
                    continue
            vector_hits.sort(key=lambda x: x["score"], reverse=True)
            for d in vector_hits[:60]:
                if d["id"] in merged:
                    merged[d["id"]]["score"] += d["score"]
                    merged[d["id"]]["retrieval"] += "+vector"
                else:
                    merged[d["id"]] = d
        except Exception:
            pass

        # Safety fallback: lexical scan if indexes are empty/stale.
        if not merged:
            rows = conn.execute(
                f"""
                SELECT chunks.*, documents.title, documents.path, documents.doc_type, documents.updated_at,
                       documents.date_guess, documents.amount, documents.company
                FROM chunks JOIN documents ON documents.id=chunks.document_id
                WHERE documents.status='indexed'{doc_filter}
                """,
                params_base,
            ).fetchall()
            for r in rows:
                s = chunk_score(question, r["normalized_text"])
                if s > 0:
                    d = dict(r)
                    d["score"] = s / 5.0
                    d["retrieval"] = "lexical-fallback"
                    merged[d["id"]] = d

    ranked = []
    for d in merged.values():
        d["score"] = rerank_candidate(d, question, plan)
        ranked.append(d)
    ranked.sort(key=lambda x: x["score"], reverse=True)
    return gpt_rerank(question, ranked, max(1, min(limit, 20)))


@app.get("/api/eval-cases")
def list_eval_cases():
    with DB_LOCK, db() as conn:
        rows = conn.execute("SELECT * FROM eval_cases ORDER BY id DESC").fetchall()
    return {"cases": [dict(r) for r in rows]}


@app.post("/api/eval-cases")
def add_eval_case(req: EvalCaseRequest):
    with DB_LOCK, db() as conn:
        cur = conn.execute(
            "INSERT INTO eval_cases(question, expected, must_cite, created_at) VALUES(?,?,?,?)",
            (req.question, req.expected, req.must_cite, now_iso()),
        )
    return {"ok": True, "id": cur.lastrowid}


@app.delete("/api/eval-cases/{case_id}")
def delete_eval_case(case_id: int):
    with DB_LOCK, db() as conn:
        conn.execute("DELETE FROM eval_cases WHERE id=?", (case_id,))
    return {"ok": True}


@app.post("/api/evaluate-rag")
def evaluate_rag():
    with DB_LOCK, db() as conn:
        cases = conn.execute("SELECT * FROM eval_cases ORDER BY id").fetchall()
    results = []
    for c in cases:
        answer_data = ask(AskRequest(question=c["question"], use_ai=bool(os.environ.get("OPENAI_API_KEY"))))
        answer = answer_data.get("answer", "")
        sources = answer_data.get("sources", [])
        expected_ok = not c["expected"] or normalize_arabic(c["expected"]) in normalize_arabic(answer)
        cite_ok = not c["must_cite"] or any(normalize_arabic(c["must_cite"]) in normalize_arabic(s.get("title", "") + " " + s.get("path", "")) for s in sources)
        results.append({
            "id": c["id"], "question": c["question"], "expected": c["expected"], "must_cite": c["must_cite"],
            "answer": answer, "sources": sources, "expected_ok": expected_ok, "cite_ok": cite_ok,
            "pass": expected_ok and cite_ok,
        })
    passed = sum(1 for r in results if r["pass"])
    return {"ok": True, "total": len(results), "passed": passed, "score": round(passed / len(results), 3) if results else None, "results": results}


@app.post("/api/ask")
def ask(req: AskRequest):
    top = hybrid_retrieve(req.question, limit=max(1, min(req.limit, 12)))
    if not top:
        return {"answer": "No matching indexed text found yet. Add folders/uploads and make sure OCR is working.", "sources": [], "mode": "hybrid-empty"}

    sources = [{"document_id": t["document_id"], "title": t["title"], "page": t["page"], "path": t["path"], "snippet": t["text"][:650], "score": round(t.get("score", 0), 4), "retrieval": t.get("retrieval", "hybrid")} for t in top]
    if req.use_ai and os.environ.get("OPENAI_API_KEY") and OpenAI:
        try:
            client = OpenAI()
            context = "\n\n".join([f"[Source {i+1}: {t['title']} page {t['page']}]\n{t['text']}" for i, t in enumerate(top)])
            lang_hint = "Answer in Arabic if the question is Arabic; otherwise answer in the user's language."
            resp = client.chat.completions.create(
                model=os.environ.get("DOCWISE_CHAT_MODEL", "gpt-4o-mini"),
                messages=[
                    {"role": "system", "content": f"You are a strict document RAG assistant. Use only the provided sources. Cite sources like [Source 1]. If the sources do not clearly support the answer, say you could not find it in the indexed documents. Prefer exact amounts, dates, names, and quotes from the sources. {lang_hint}"},
                    {"role": "user", "content": f"Question: {req.question}\n\nSources:\n{context}"},
                ],
                temperature=0.2,
            )
            answer = resp.choices[0].message.content or ""
            verification = verify_answer(req.question, answer, sources)
            return {"answer": answer, "sources": sources, "mode": "ai", "verification": verification}
        except Exception as exc:
            return {"answer": fallback_answer(req.question, top) + f"\n\nAI error: {exc}", "sources": sources, "mode": "extractive"}
    return {"answer": fallback_answer(req.question, top), "sources": sources, "mode": "extractive"}


def fallback_answer(question: str, top: list[dict]) -> str:
    arabic = has_arabic(question)
    if arabic:
        intro = "أقرب النتائج من الأرشيف:"
        lines = [intro]
        for i, t in enumerate(top[:4], start=1):
            lines.append(f"[{i}] {t['title']} - صفحة {t['page']}: {t['text'][:360]}...")
        return "\n".join(lines)
    lines = ["Closest matches from your archive:"]
    for i, t in enumerate(top[:4], start=1):
        lines.append(f"[{i}] {t['title']} - page {t['page']}: {t['text'][:360]}...")
    return "\n".join(lines)


def filing_suggestion(d: dict) -> dict:
    doc_type = d.get("doc_type") or "general"
    date = d.get("date_guess") or datetime.now().strftime("%Y-%m-%d")
    year = re.search(r"20\d{2}", date)
    year_s = year.group(0) if year else datetime.now().strftime("%Y")
    company = d.get("company") or d.get("title") or "document"
    company = normalize_filename(company)[:35] or "document"
    amount = normalize_filename(d.get("amount") or "")
    ext = d.get("file_ext") or Path(d.get("path", "file.pdf")).suffix
    name_parts = [date.replace("/", "-").replace(".", "-"), doc_type, company]
    if amount:
        name_parts.append(amount)
    filename = normalize_filename("_".join(name_parts))[:120] + ext
    target = ARCHIVE / doc_type / year_s / filename
    return {"folder": str((ARCHIVE / doc_type / year_s).resolve()), "filename": filename, "target_path": str(target.resolve())}


def normalize_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9\u0600-\u06FF_.-]+", "_", value or "")
    value = re.sub(r"_+", "_", value).strip("_.-")
    return value or "document"


@app.post("/api/organize")
def organize(req: OrganizeRequest):
    with DB_LOCK, db() as conn:
        row = conn.execute("SELECT * FROM documents WHERE id=?", (req.document_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Document not found")
    d = row_to_doc(row, include_text=True)
    source = Path(d["path"])
    if not source.exists():
        raise HTTPException(status_code=404, detail="Original file missing")
    suggestion = filing_suggestion(d)
    target = Path(suggestion["target_path"])
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target = target.with_name(f"{target.stem}_{int(time.time())}{target.suffix}")
    if req.mode == "move":
        shutil.move(str(source), str(target))
        new_path = target
    else:
        shutil.copy2(source, target)
        new_path = target
    indexed = index_file(new_path, "archive", force=True)
    return {"ok": True, "mode": req.mode, "target": str(target), "indexed": indexed}


@app.post("/api/open/{doc_id}")
def open_file(doc_id: int):
    with DB_LOCK, db() as conn:
        row = conn.execute("SELECT path FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Document not found")
    path = Path(row["path"])
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    try:
        if hasattr(os, "startfile"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
        return {"ok": True}
    except Exception as exc:
        return {"ok": False, "error": str(exc), "path": str(path)}


@app.get("/api/file/{doc_id}")
def get_file(doc_id: int):
    with DB_LOCK, db() as conn:
        row = conn.execute("SELECT path FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Document not found")
    path = Path(row["path"])
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path)


app.mount("/", StaticFiles(directory=STATIC, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8120, reload=False)
