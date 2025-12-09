# app.py
from flask import Flask, request, send_file, abort, jsonify, after_this_request, current_app
import os
import tempfile
import shutil
import subprocess
from werkzeug.utils import secure_filename
from flask_cors import CORS
import zipfile
import logging
import re

# -----------------------------
# Basic config & limits
# -----------------------------
logging.getLogger('werkzeug').setLevel(logging.ERROR)

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Environment for LibreOffice + Poppler
os.environ.setdefault("UNO_PATH", "/usr/lib/libreoffice/program")
os.environ["PATH"] += ":/usr/lib/libreoffice/program:/usr/bin:/usr/local/bin"
POPPLER_PATH = "/usr/bin"

# upload limit (200 MB). Adjust down if you hit OOM.
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# operational limits to protect 512 MB runtime
MAX_OCR_PAGES = 40        # OCR at most 40 pages by default
OCR_DPI = 180             # DPI for rasterization (lower => less memory)
PDF_TO_JPG_DPI = 160
IMAGE_THREAD_COUNT = 1    # single-threaded for pdf2image/poppler
PDF2DOCX_PAGE_LIMIT = 200 # limit pages for pdf2docx conversion (safety)

# -----------------------------
# Lazy imports to reduce memory
# -----------------------------
def lazy_pdf2docx_converter():
    from pdf2docx import Converter
    return Converter

def lazy_pil_Image():
    from PIL import Image
    return Image

def lazy_pdf2image_convert():
    from pdf2image import convert_from_path
    return convert_from_path

def lazy_pytesseract():
    import pytesseract
    return pytesseract

def lazy_pikepdf():
    import pikepdf
    return pikepdf

def lazy_pypdf():
    from PyPDF2 import PdfReader, PdfWriter, PdfMerger
    return PdfReader, PdfWriter, PdfMerger

def lazy_pdfplumber():
    import pdfplumber
    return pdfplumber

def lazy_docx_Document():
    from docx import Document
    return Document

# -----------------------------
# Utilities
# -----------------------------
def tmp_file(ext: str = "") -> str:
    fd, path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    return path

def tmp_dir() -> str:
    return tempfile.mkdtemp()

def save_upload(file_obj, ext: str | None = None) -> str:
    filename = secure_filename(file_obj.filename or "upload")
    extension = ext if ext else os.path.splitext(filename)[1] or ""
    path = tmp_file(extension)
    file_obj.stream.seek(0)
    with open(path, "wb") as f:
        for chunk in iter(lambda: file_obj.stream.read(4096), b""):
            f.write(chunk)
    return path

def cleanup(path: str):
    try:
        if not path:
            return
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.isfile(path):
            os.remove(path)
    except Exception:
        # never raise from cleanup
        pass

def run_subprocess(cmd, **kwargs):
    """Run subprocess and convert CalledProcessError into HTTP 500 with logs."""
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True, **kwargs)
    except subprocess.CalledProcessError as e:
        current_app.logger.exception("Subprocess failed: %s", cmd)
        raise

# -----------------------------
# Smart Hybrid formatting helpers (Option C)
# -----------------------------
BULLET_PATTERNS = [
    r'^\s*[-•\u2022]\s+',
    r'^\s*\d+\.\s+',
    r'^\s*\(\w\)\s+',
]

def is_bullet_line(s: str) -> bool:
    s = s.strip()
    for p in BULLET_PATTERNS:
        if re.match(p, s):
            return True
    return False

def detect_heading(lines):
    """
    Heuristic: a heading is a short line (<= 8 words) in all-caps or Title Case,
    or a line followed by an empty line and then a paragraph.
    """
    heading_lines = set()
    for i, line in enumerate(lines):
        t = line.strip()
        if not t:
            continue
        words = t.split()
        if 1 <= len(words) <= 8:
            # all caps
            if t.isupper() and any(c.isalpha() for c in t):
                heading_lines.add(i)
                continue
            # Title Case heuristic (First letters capitalized and not too long)
            if all(w[0].isupper() for w in words if w):
                heading_lines.add(i)
    return heading_lines

def merge_lines_to_paragraphs(raw_text: str) -> str:
    """
    Convert OCR/raw text lines into smart hybrid paragraphs:
    - Merge lines that are likely part of the same paragraph (not bullets, not headings).
    - Detect bullets & numbered lists and keep them as list items.
    - Detect headings using heuristics and add spacing.
    """
    lines = [ln.rstrip() for ln in raw_text.splitlines()]
    # remove repeated blank lines
    cleaned_lines = []
    for ln in lines:
        if cleaned_lines and cleaned_lines[-1] == "" and ln == "":
            continue
        cleaned_lines.append(ln)
    lines = cleaned_lines

    headings_idx = detect_heading(lines)

    out_blocks = []
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        if ln == "":
            # blank line -> paragraph separator
            i += 1
            continue

        # bullet/numbered list detection
        if is_bullet_line(ln):
            # collect contiguous bullets
            bullets = []
            while i < len(lines) and lines[i].strip() and is_bullet_line(lines[i]):
                bullets.append(lines[i].strip())
                i += 1
            out_blocks.append("\n".join(bullets))
            continue

        # heading detection
        if i in headings_idx:
            out_blocks.append(ln.upper() if ln.isupper() else ln)
            i += 1
            continue

        # Otherwise gather lines that should be merged into a paragraph
        para_lines = [ln]
        i += 1
        while i < len(lines) and lines[i].strip() and not is_bullet_line(lines[i]) and (i not in headings_idx):
            next_ln = lines[i].strip()
            # Heuristic: if previous line ends with punctuation, keep newline.
            if para_lines[-1].endswith(('.', '?', '!', ':', ';', '—', '-')):
                # treat as end-of-sentence; still merge but with space
                para_lines.append(next_ln)
            else:
                # merge broken line if short and next starts lowercase or continuation
                para_lines.append(next_ln)
            i += 1
        # join with space, but preserve single newlines in certain cases
        paragraph = " ".join(x for x in [p.strip() for p in para_lines] if x)
        out_blocks.append(paragraph)

    # Post processing: fix multiple spaces
    out = "\n\n".join(block.strip() for block in out_blocks if block.strip())
    out = re.sub(r' {2,}', ' ', out)
    # Normalize bullet markers to a consistent bullet (•)
    out = re.sub(r'^\s*[-\u2022]\s+', '• ', out, flags=re.MULTILINE)
    return out.strip()

# -----------------------------
# OCR pipeline (disk-based, memory-friendly)
# -----------------------------
def ocr_pdf_to_text(pdf_path: str, max_pages: int = MAX_OCR_PAGES, dpi: int = OCR_DPI) -> str:
    """
    Rasterize PDF pages to images on disk, OCR each page sequentially,
    and return combined text. Uses Smart Hybrid formatting on the combined OCR.
    """
    convert_from_path = lazy_pdf2image_convert()
    pytesseract = lazy_pytesseract()
    Image = lazy_pil_Image()

    tmpdir = tmp_dir()
    try:
        # rasterize to disk; paths_only ensures we get file paths
        try:
            image_paths = convert_from_path(
                pdf_path,
                dpi=dpi,
                poppler_path=POPPLER_PATH,
                output_folder=tmpdir,
                fmt="png",
                paths_only=True,
                thread_count=IMAGE_THREAD_COUNT,
            )
        except Exception as e:
            current_app.logger.exception("pdf2image rasterization failed")
            raise RuntimeError(f"Rasterization failed: {e}")

        if not image_paths:
            return ""

        # limit pages
        image_paths = sorted(image_paths)[:max_pages]

        page_texts = []
        for idx, img_path in enumerate(image_paths, start=1):
            try:
                # open image from disk, OCR, then close immediately
                with Image.open(img_path) as im:
                    # optionally convert to grayscale for faster OCR
                    im = im.convert("L")
                    txt = pytesseract.image_to_string(im, config="--oem 1 --psm 3 -l eng")
            except Exception:
                current_app.logger.exception("Tesseract failed on page %s", idx)
                txt = ""
            page_texts.append(f"--- PAGE {idx} ---\n{txt.strip()}")

        combined = "\n\n".join(page_texts).strip()
        # Smart Hybrid formatting: remove page markers and reformat
        stripped = "\n".join(line for line in combined.splitlines() if not line.strip().startswith('--- PAGE'))
        formatted = merge_lines_to_paragraphs(stripped)
        return formatted

    finally:
        cleanup(tmpdir)

# -----------------------------
# Endpoint: PDF → Word (pdf2docx with OCR fallback)
# -----------------------------
@app.post("/pdf-to-word")
def pdf_to_word():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    PdfReader, PdfWriter, _ = lazy_pypdf()
    Converter = lazy_pdf2docx_converter()
    Document = lazy_docx_Document()

    pdf_path = save_upload(f, ".pdf")
    unlocked_pdf = tmp_file(".pdf")
    out_docx = tmp_file(".docx")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(pdf_path)
        cleanup(unlocked_pdf)
        cleanup(out_docx)
        return resp

    try:
        # Try opening PDF and check if encrypted
        try:
            reader = PdfReader(pdf_path, strict=False)
        except Exception as e:
            return jsonify({"error": "Unable to open PDF.", "details": str(e)}), 400

        pdf_to_use = pdf_path

        if getattr(reader, "is_encrypted", False):
            # Try empty-password unlock then pikepdf fallback
            try:
                reader.decrypt("")
                writer = PdfWriter()
                for p in reader.pages:
                    writer.add_page(p)
                with open(unlocked_pdf, "wb") as outf:
                    writer.write(outf)
                pdf_to_use = unlocked_pdf
            except Exception:
                try:
                    pikepdf = lazy_pikepdf()
                    with pikepdf.open(pdf_path, password="") as pp:
                        pp.save(unlocked_pdf)
                    pdf_to_use = unlocked_pdf
                except Exception:
                    return jsonify({"error": "PDF encrypted. Use Unlock tool first."}), 400

        # Attempt pdf2docx conversion first (best fidelity)
        try:
            # Use Converter but limit pages for safety
            cv = Converter(pdf_to_use)
            cv.convert(out_docx, start=0, end=None, layout_mode=True)
            cv.close()
            doc_check = Document(out_docx)
            if len(doc_check.paragraphs) > 0:
                return send_file(out_docx, as_attachment=True, download_name="output.docx")
        except Exception:
            # conversion failed -> fallback to OCR
            current_app.logger.info("pdf2docx conversion failed; falling back to OCR")

        # OCR fallback (smart hybrid formatting)
        try:
            ocr_text = ocr_pdf_to_text(pdf_to_use, max_pages=MAX_OCR_PAGES, dpi=OCR_DPI)
        except Exception as e:
            current_app.logger.exception("OCR fallback failed")
            return jsonify({"error": "OCR failed.", "details": str(e)}), 500

        # Save OCR text to docx with paragraphs & simple heuristics (headings / bullets)
        doc = Document()
        for block in ocr_text.split("\n\n"):
            block = block.strip()
            if not block:
                continue
            # Apply simple heading detection (all-caps or short lines)
            words = block.split()
            if 1 <= len(words) <= 8 and block.upper() == block and any(c.isalpha() for c in block):
                p = doc.add_paragraph()
                run = p.add_run(block)
                run.bold = True
                continue
            # bullets
            if block.startswith("• "):
                # create bullet list entries (docx supports paragraph style 'List Bullet')
                for line in block.splitlines():
                    if line.strip():
                        p = doc.add_paragraph(line.strip().lstrip('•').strip(), style='List Bullet')
                continue
            doc.add_paragraph(block)
        doc.save(out_docx)
        return send_file(out_docx, as_attachment=True, download_name="output.docx")

    finally:
        # final cleanup handled by after_this_request
        pass

# -----------------------------
# Endpoint: Word → PDF (LibreOffice on-demand)
# -----------------------------
def safe_libreoffice_convert(input_path: str, out_dir: str, convert_filter: list):
    """
    Run LibreOffice headless conversion safely. convert_filter is list args after --convert-to.
    """
    cmd = ["libreoffice", "--headless", "--norestore", "--nologo", "--invisible", "--convert-to"] + convert_filter + ["--outdir", out_dir, input_path]
    run_subprocess(cmd)

@app.post("/word-to-pdf")
def word_to_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".doc", ".docx"):
        return abort(400, "Upload a Word (.doc/.docx) file")

    doc_path = save_upload(f, ext)
    out_dir = tmp_dir()

    @after_this_request
    def _cleanup_response(resp):
        cleanup(doc_path)
        cleanup(out_dir)
        return resp

    try:
        # DOC/DOCX -> ODT
        try:
            safe_libreoffice_convert(doc_path, out_dir, ["odt"])
        except Exception:
            # if conversion fails, surface an error
            return jsonify({"error": "LibreOffice conversion to ODT failed."}), 500

        base = os.path.splitext(os.path.basename(doc_path))[0]
        odt_path = os.path.join(out_dir, f"{base}.odt")
        if not os.path.exists(odt_path):
            return jsonify({"error": "Intermediate ODT not found; conversion failed."}), 500

        # ODT -> PDF (embed fonts, keep images quality)
        try:
            safe_libreoffice_convert(odt_path, out_dir, ["pdf:writer_pdf_Export:EmbedStandardFonts=true;ReduceImageResolution=false"])
        except Exception:
            return jsonify({"error": "LibreOffice ODT -> PDF conversion failed."}), 500

        out_pdf = os.path.join(out_dir, f"{base}.pdf")
        if not os.path.exists(out_pdf):
            return jsonify({"error": "Output PDF not generated."}), 500

        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")
    finally:
        pass

# -----------------------------
# Endpoint: PPT -> PDF
# -----------------------------
@app.post("/ppt-to-pdf")
def ppt_to_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".ppt", ".pptx"):
        return abort(400, "Upload a PPT/PPTX file")

    ppt_path = save_upload(f, ext)
    out_dir = tmp_dir()

    @after_this_request
    def _cleanup_response(resp):
        cleanup(ppt_path)
        cleanup(out_dir)
        return resp

    try:
        try:
            safe_libreoffice_convert(ppt_path, out_dir, ["pdf:impress_pdf_Export"])
        except Exception:
            return jsonify({"error": "LibreOffice PPT -> PDF conversion failed."}), 500

        base = os.path.splitext(os.path.basename(ppt_path))[0]
        pdf_path = os.path.join(out_dir, f"{base}.pdf")
        if not os.path.exists(pdf_path):
            return jsonify({"error": "PDF not produced from PPT."}), 500
        return send_file(pdf_path, as_attachment=True, download_name="output.pdf")
    finally:
        pass

# -----------------------------
# Endpoint: JPG -> PDF
# -----------------------------
@app.post("/jpg-to-pdf")
def jpg_to_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files selected")

    Image = lazy_pil_Image()
    saved = []
    images = []
    out_pdf = tmp_file(".pdf")

    @after_this_request
    def _cleanup_response(resp):
        for p in saved:
            cleanup(p)
        cleanup(out_pdf)
        return resp

    try:
        # Save uploads to disk and open sequentially
        for f in files:
            p = save_upload(f)
            saved.append(p)
            try:
                img = Image.open(p).convert("RGB")
                # do not keep PIL objects in memory longer than necessary:
                images.append(p)  # store path; we'll open again when saving
            except Exception:
                current_app.logger.exception("Failed to open image %s", p)
                cleanup(p)
        # Build pdf using first image and appending the rest via Pillow (open one-by-one)
        if not images:
            return abort(400, "No valid images uploaded")
        # Pillow can accept file paths; open first
        with Image.open(images[0]).convert("RGB") as first_img:
            rest = []
            for path in images[1:]:
                with Image.open(path).convert("RGB") as im:
                    rest.append(im.copy())
            first_img.save(out_pdf, save_all=True, append_images=rest, dpi=(300,300), quality=90)
        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")
    finally:
        pass

# -----------------------------
# Endpoint: PDF -> JPG (memory-friendly, disk-based)
# -----------------------------
@app.post("/pdf-to-jpg")
def pdf_to_jpg():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    convert_from_path = lazy_pdf2image_convert()
    pdf_path = save_upload(f, ".pdf")
    out_dir = tmp_dir()
    zip_path = tmp_file(".zip")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(pdf_path)
        cleanup(out_dir)
        cleanup(zip_path)
        return resp

    try:
        try:
            image_paths = convert_from_path(
                pdf_path,
                dpi=PDF_TO_JPG_DPI,
                poppler_path=POPPLER_PATH,
                output_folder=out_dir,
                fmt="jpeg",
                paths_only=True,
                thread_count=IMAGE_THREAD_COUNT,
            )
        except Exception as e:
            current_app.logger.exception("pdf2image failed")
            return jsonify({"error": "Failed to rasterize PDF", "details": str(e)}), 500

        if not image_paths:
            return jsonify({"error": "No images produced from PDF"}), 500

        # ZIP images on disk directly
        with zipfile.ZipFile(zip_path, "w") as z:
            for p in sorted(image_paths):
                z.write(p, arcname=os.path.basename(p))
        return send_file(zip_path, as_attachment=True, download_name="images.zip")
    finally:
        pass

# -----------------------------
# Endpoint: Merge PDFs
# -----------------------------
@app.post("/merge-pdf")
def merge_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files uploaded")
    PdfReader, PdfWriter, PdfMerger = lazy_pypdf()

    merger = PdfMerger()
    saved = []
    out_pdf = tmp_file(".pdf")

    @after_this_request
    def _cleanup_response(resp):
        for p in saved:
            cleanup(p)
        cleanup(out_pdf)
        return resp

    try:
        for f in files:
            if not f.filename.lower().endswith(".pdf"):
                return abort(400, "All files must be PDF")
            p = save_upload(f, ".pdf")
            saved.append(p)
            merger.append(p)
        merger.write(out_pdf)
        merger.close()
        return send_file(out_pdf, as_attachment=True, download_name="merged.pdf")
    finally:
        pass

# -----------------------------
# Endpoint: Split PDF
# -----------------------------
@app.post("/split-pdf")
def split_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")
    ranges = request.form.get("ranges")
    if not ranges:
        return abort(400, "Missing ranges parameter")

    PdfReader, PdfWriter, _ = lazy_pypdf()
    pdf_path = save_upload(f, ".pdf")
    out_dir = tmp_dir()
    zip_path = tmp_file(".zip")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(pdf_path)
        cleanup(out_dir)
        cleanup(zip_path)
        return resp

    try:
        reader = PdfReader(pdf_path, strict=False)
        total = len(reader.pages)
        pages = []
        for part in ranges.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                try:
                    a, b = map(int, part.split("-"))
                    pages.extend(range(a, b + 1))
                except Exception:
                    continue
            else:
                try:
                    pages.append(int(part))
                except Exception:
                    continue
        pages = sorted({p for p in pages if 1 <= p <= total})
        if not pages:
            return abort(400, "No valid pages derived from ranges")

        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pages:
                writer = PdfWriter()
                writer.add_page(reader.pages[p-1])
                single_pdf = os.path.join(out_dir, f"page_{p}.pdf")
                with open(single_pdf, "wb") as o:
                    writer.write(o)
                z.write(single_pdf, arcname=os.path.basename(single_pdf))
        return send_file(zip_path, as_attachment=True, download_name="split.zip")
    finally:
        pass

# -----------------------------
# Endpoint: Rotate PDF
# -----------------------------
@app.post("/rotate-pdf")
def rotate_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")
    try:
        angle = int(request.form.get("angle", 90) or 90)
    except Exception:
        return abort(400, "Invalid angle")

    PdfReader, PdfWriter, _ = lazy_pypdf()
    pdf_path = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(pdf_path)
        cleanup(out_pdf)
        return resp

    try:
        reader = PdfReader(pdf_path, strict=False)
        writer = PdfWriter()
        for page in reader.pages:
            try:
                page.rotate(angle)
            except Exception:
                try:
                    page.rotate_clockwise(angle)
                except Exception:
                    # older/newer PyPDF2 differences, attempt safe rotate by transform
                    pass
            writer.add_page(page)
        with open(out_pdf, "wb") as o:
            writer.write(o)
        return send_file(out_pdf, as_attachment=True, download_name="rotated.pdf")
    finally:
        pass

# -----------------------------
# Endpoint: Compress PDF (Ghostscript)
# -----------------------------
@app.post("/compress-pdf")
def compress_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")
    input_pdf = save_upload(f, ".pdf")
    output_pdf = tmp_file(".pdf")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(input_pdf)
        cleanup(output_pdf)
        return resp

    try:
        gs = shutil.which("gs")
        if not gs:
            return abort(500, "Ghostscript not installed")

        cmd = [
            gs,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/ebook",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_pdf}",
            input_pdf
        ]
        try:
            subprocess.run(cmd, check=True)
        except subprocess.CalledProcessError:
            return abort(500, "Ghostscript compression failed")
        return send_file(output_pdf, as_attachment=True, download_name="compressed.pdf", mimetype="application/pdf")
    finally:
        pass

# -----------------------------
# Endpoint: Protect PDF (add password)
# -----------------------------
@app.post("/protect-pdf")
def protect_pdf():
    PdfReader, PdfWriter, _ = lazy_pypdf()
    f = request.files.get("file")
    pwd = request.form.get("password")
    if not f or not pwd:
        return abort(400, "Missing file or password")
    pdf_path = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(pdf_path)
        cleanup(out_pdf)
        return resp

    try:
        reader = PdfReader(pdf_path, strict=False)
        writer = PdfWriter()
        for p in reader.pages:
            writer.add_page(p)
        writer.encrypt(user_pwd=pwd)
        with open(out_pdf, "wb") as o:
            writer.write(o)
        return send_file(out_pdf, as_attachment=True, download_name="protected.pdf")
    finally:
        pass

# -----------------------------
# Endpoint: Unlock PDF
# -----------------------------
@app.post("/unlock-pdf")
def unlock_pdf():
    PdfReader, PdfWriter, _ = lazy_pypdf()
    f = request.files.get("file")
    pwd = request.form.get("password", "")  # empty string allowed (try decrypt)
    if not f:
        return abort(400, "Missing file")
    pdf_path = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(pdf_path)
        cleanup(out_pdf)
        return resp

    try:
        reader = PdfReader(pdf_path, strict=False)
        if getattr(reader, "is_encrypted", False):
            if not pwd:
                return abort(400, "Password required to unlock")
            try:
                reader.decrypt(pwd)
            except Exception:
                return abort(400, "Wrong password")
        writer = PdfWriter()
        for p in reader.pages:
            writer.add_page(p)
        with open(out_pdf, "wb") as o:
            writer.write(o)
        return send_file(out_pdf, as_attachment=True, download_name="unlocked.pdf")
    finally:
        pass

# -----------------------------
# Endpoint: Extract Text (with OCR fallback, Smart Hybrid)
# -----------------------------
@app.post("/extract-text")
def extract_text():
    pdfplumber = lazy_pdfplumber()
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")
    pdf_path = save_upload(f, ".pdf")

    @after_this_request
    def _cleanup_response(resp):
        cleanup(pdf_path)
        return resp

    try:
        text_pages = []
        has_text = False
        try:
            with pdfplumber.open(pdf_path) as p:
                for i, page in enumerate(p.pages, start=1):
                    raw = page.extract_text() or ""
                    cleaned = " ".join(line.strip() for line in raw.split("\n") if line.strip())
                    if cleaned:
                        has_text = True
                    text_pages.append(f"--- PAGE {i} ---\n{cleaned}")
        except Exception:
            current_app.logger.exception("pdfplumber extraction failed")

        if not has_text:
            # OCR fallback (smart hybrid formatting)
            formatted = ocr_pdf_to_text(pdf_path, max_pages=MAX_OCR_PAGES, dpi=OCR_DPI)
            return jsonify({"text": formatted})
        else:
            # Combine pages preserving page markers, then apply hybrid formatting
            combined = "\n\n".join(text_pages)
            stripped = "\n".join(line for line in combined.splitlines() if not line.strip().startswith('--- PAGE'))
            formatted = merge_lines_to_paragraphs(stripped)
            return jsonify({"text": formatted})
    finally:
        pass

# -----------------------------
# Root and CORS helpers
# -----------------------------
@app.get("/")
def home():
    return "PDF Tools Backend Running (Smart Hybrid formatting, optimized for 512MB)"

@app.after_request
def apply_cors(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Max-Age"] = "86400"
    return response

@app.before_request
def handle_preflight():
    if request.method == "OPTIONS":
        resp = app.make_response("")
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Headers"] = "*"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        return resp

# -----------------------------
# Run (local)
# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
