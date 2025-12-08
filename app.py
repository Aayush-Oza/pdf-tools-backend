from flask import Flask, request, send_file, abort, jsonify, after_this_request, current_app
import os
import tempfile
import shutil
import subprocess
from werkzeug.utils import secure_filename
from flask_cors import CORS
import zipfile
import logging

# Lighten logs
logging.getLogger('werkzeug').setLevel(logging.ERROR)

# --------------------------------------------------------
# APP + BASIC CONFIG
# --------------------------------------------------------
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# LibreOffice / Poppler environment
os.environ.setdefault("UNO_PATH", "/usr/lib/libreoffice/program")
os.environ["PATH"] += ":/usr/lib/libreoffice/program:/usr/bin:/usr/local/bin"

POPPLER_PATH = "/usr/bin"

# 200 MB max upload (you can reduce this if needed to save memory)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# --------------------------------------------------------
# LAZY IMPORT HELPERS (to avoid loading heavy libs until needed)
# --------------------------------------------------------
def get_pdf2docx_converter():
    from pdf2docx import Converter
    return Converter

def get_pil_image():
    from PIL import Image
    return Image

def get_convert_from_path():
    from pdf2image import convert_from_path
    return convert_from_path

def get_pytesseract():
    import pytesseract
    return pytesseract

def get_pikepdf():
    import pikepdf
    return pikepdf

def get_pypdf():
    from PyPDF2 import PdfReader, PdfWriter, PdfMerger
    return PdfReader, PdfWriter, PdfMerger

def get_pdfplumber():
    import pdfplumber
    return pdfplumber

def get_docx_document():
    from docx import Document
    return Document

# --------------------------------------------------------
# UTILS
# --------------------------------------------------------
def tmp_file(ext: str = "") -> str:
    fd, path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    return path


def save_upload(file, ext: str | None = None) -> str:
    """Save uploaded file object to a temp file and return its path."""
    filename = secure_filename(file.filename)
    extension = ext if ext else os.path.splitext(filename)[1]
    path = tmp_file(extension)

    file.stream.seek(0)
    with open(path, "wb") as f:
        for chunk in iter(lambda: file.stream.read(4096), b""):
            f.write(chunk)

    return path


def cleanup(path: str):
    """Safely delete file or directory if it exists."""
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def ocr_pdf_to_text(pdf_path: str, dpi: int = 200) -> str:
    """
    OCR an image-based PDF using Tesseract.
    Memory-friendly: render pages to disk, not all in RAM.
    """
    convert_from_path = get_convert_from_path()
    pytesseract = get_pytesseract()
    Image = get_pil_image()

    temp_dir = tempfile.mkdtemp()
    try:
        try:
            # paths_only=True → get file paths, not PIL objects
            image_paths = convert_from_path(
                pdf_path,
                dpi=dpi,
                poppler_path=POPPLER_PATH,
                output_folder=temp_dir,
                fmt="png",
                paths_only=True,
                thread_count=2,
            )
        except Exception as e:
            current_app.logger.exception("OCR pdf2image failed")
            raise RuntimeError(f"OCR rasterization failed: {e}")

        texts = []
        # Process each page one-by-one to keep memory lower
        for i, img_path in enumerate(sorted(image_paths), start=1):
            try:
                with Image.open(img_path) as img:
                    text = pytesseract.image_to_string(
                        img,
                        config="--oem 1 --psm 3 -l eng"
                    )
            except Exception:
                current_app.logger.exception("Tesseract OCR failed on page %s", i)
                text = ""
            texts.append(f"--- PAGE {i} ---\n{text.strip()}")

        return "\n\n".join(texts).strip()

    finally:
        cleanup(temp_dir)

# --------------------------------------------------------
# 1 - PDF → Word  (with OCR fallback)
# --------------------------------------------------------
@app.post("/pdf-to-word")
def pdf_to_word():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    PdfReader, PdfWriter, _ = get_pypdf()
    Converter = get_pdf2docx_converter()
    Document = get_docx_document()

    pdf = save_upload(f, ".pdf")
    unlocked_pdf = tmp_file(".pdf")
    out_docx = tmp_file(".docx")

    @after_this_request
    def _cleanup_response(response):
        cleanup(pdf)
        cleanup(unlocked_pdf)
        cleanup(out_docx)
        return response

    try:
        # Try opening PDF
        try:
            reader = PdfReader(pdf, strict=False)
        except Exception as e:
            return jsonify({
                "error": "Unable to open PDF (possibly corrupted).",
                "details": str(e)
            }), 400

        pdf_to_use = pdf

        # Handle encrypted PDF
        if getattr(reader, "is_encrypted", False):
            try:
                try:
                    reader.decrypt("")
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    with open(unlocked_pdf, "wb") as f2:
                        writer.write(f2)
                    pdf_to_use = unlocked_pdf
                except Exception:
                    pikepdf = get_pikepdf()
                    try:
                        with pikepdf.open(pdf, password="") as pp:
                            pp.save(unlocked_pdf)
                        pdf_to_use = unlocked_pdf
                    except pikepdf._qpdf.PasswordError:
                        return jsonify({
                            "error": "PDF is password protected. Use Unlock tool first."
                        }), 400
            except Exception as e:
                return jsonify({
                    "error": "Failed processing encrypted PDF",
                    "details": str(e)
                }), 400

        # Try normal PDF → Word via pdf2docx
        try:
            cv = Converter(pdf_to_use)
            cv.convert(out_docx, start=0, end=None, layout_mode=True)
            cv.close()

            doc_check = Document(out_docx)
            if len(doc_check.paragraphs) > 0:
                return send_file(out_docx, as_attachment=True, download_name="output.docx")
        except Exception:
            # fallback to OCR later
            pass

        # Fallback to OCR if pdf2docx failed
        try:
            ocr_text = ocr_pdf_to_text(pdf_to_use)
        except Exception as e:
            return jsonify({
                "error": "OCR failed. PDF may be damaged.",
                "details": str(e)
            }), 500

        doc = Document()
        for block in ocr_text.split("\n\n"):
            block = block.strip()
            if block:
                doc.add_paragraph(block)
        doc.save(out_docx)

        return send_file(out_docx, as_attachment=True, download_name="output.docx")

    finally:
        pass

# --------------------------------------------------------
# 2 - Word → PDF (LibreOffice on-demand)
# --------------------------------------------------------
@app.post("/word-to-pdf")
def word_to_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in [".doc", ".docx"]:
        return abort(400, "Upload a Word file")

    doc = save_upload(f, ext)
    out_dir = tempfile.mkdtemp()

    try:
        # Step 1: DOC/DOCX → ODT
        subprocess.run(
            [
                "libreoffice", "--headless",
                "--convert-to", "odt",
                "--outdir", out_dir, doc
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True
        )

        base = os.path.splitext(os.path.basename(doc))[0]
        odt_path = os.path.join(out_dir, f"{base}.odt")

        # Step 2: ODT → high-quality PDF
        subprocess.run(
            [
                "libreoffice", "--headless",
                "--convert-to",
                "pdf:writer_pdf_Export:EmbedStandardFonts=true;ReduceImageResolution=false",
                "--outdir", out_dir, odt_path
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True
        )

        out_pdf = os.path.join(out_dir, f"{base}.pdf")
        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")

    finally:
        cleanup(doc)
        cleanup(out_dir)

# --------------------------------------------------------
# 3 - PPT → PDF
# --------------------------------------------------------
@app.post("/ppt-to-pdf")
def ppt_to_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in [".ppt", ".pptx"]:
        return abort(400, "Upload a PPT/PPTX file")

    ppt_path = save_upload(f, ext)
    out_dir = tempfile.mkdtemp()

    try:
        subprocess.run(
            [
                "libreoffice", "--headless",
                "--convert-to", "pdf:impress_pdf_Export",
                "--outdir", out_dir,
                ppt_path
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True
        )

        base = os.path.splitext(os.path.basename(ppt_path))[0]
        pdf_path = os.path.join(out_dir, f"{base}.pdf")

        if not os.path.exists(pdf_path):
            return abort(500, "Failed to export PDF from PPT.")

        return send_file(pdf_path, as_attachment=True, download_name="output.pdf")

    finally:
        cleanup(ppt_path)
        cleanup(out_dir)

# --------------------------------------------------------
# 4 - JPG → PDF
# --------------------------------------------------------
@app.post("/jpg-to-pdf")
def jpg_to_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files selected")

    Image = get_pil_image()

    images = []
    saved = []
    out_pdf = None

    try:
        for f in files:
            p = save_upload(f)
            saved.append(p)

            img = Image.open(p).convert("RGB")
            img.info["dpi"] = (300, 300)
            images.append(img)

        out_pdf = tmp_file(".pdf")

        images[0].save(
            out_pdf,
            save_all=True,
            append_images=images[1:],
            quality=100,
            dpi=(300, 300)
        )

        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")

    finally:
        for p in saved:
            cleanup(p)
        if out_pdf:
            cleanup(out_pdf)

# --------------------------------------------------------
# 5 - PDF → JPG (memory-friendly)
# --------------------------------------------------------
@app.post("/pdf-to-jpg")
def pdf_to_jpg():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    convert_from_path = get_convert_from_path()

    pdf = save_upload(f, ".pdf")
    out_dir = tempfile.mkdtemp()
    zip_path = tmp_file(".zip")

    try:
        # paths_only=True to avoid keeping all PIL images in RAM
        image_paths = convert_from_path(
            pdf,
            dpi=180,
            poppler_path=POPPLER_PATH,
            output_folder=out_dir,
            fmt="jpeg",
            paths_only=True,
            thread_count=2,
        )

        with zipfile.ZipFile(zip_path, "w") as z:
            for img_path in sorted(image_paths):
                arcname = os.path.basename(img_path)
                z.write(img_path, arcname=arcname)

        return send_file(zip_path, as_attachment=True, download_name="images.zip")

    finally:
        cleanup(pdf)
        cleanup(out_dir)
        cleanup(zip_path)

# --------------------------------------------------------
# 6 - Merge PDF
# --------------------------------------------------------
@app.post("/merge-pdf")
def merge_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files")

    PdfReader, PdfWriter, PdfMerger = get_pypdf()

    merger = PdfMerger()
    saved = []
    out_pdf = None

    try:
        for f in files:
            if not f.filename.lower().endswith(".pdf"):
                return abort(400, "All files must be PDF")

            p = save_upload(f, ".pdf")
            saved.append(p)
            merger.append(p)

        out_pdf = tmp_file(".pdf")
        merger.write(out_pdf)
        merger.close()

        return send_file(out_pdf, as_attachment=True, download_name="merged.pdf")

    finally:
        for p in saved:
            cleanup(p)
        if out_pdf:
            cleanup(out_pdf)

# --------------------------------------------------------
# 7 - Split PDF
# --------------------------------------------------------
@app.post("/split-pdf")
def split_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    PdfReader, PdfWriter, _ = get_pypdf()

    ranges = request.form.get("ranges")
    if not ranges:
        return abort(400, "Missing ranges")

    pdf = save_upload(f, ".pdf")
    out_dir = tempfile.mkdtemp()
    zip_path = None

    try:
        reader = PdfReader(pdf, strict=False)
        total = len(reader.pages)

        pages = []
        for part in ranges.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a, b = map(int, part.split("-"))
                pages.extend(range(a, b + 1))
            else:
                pages.append(int(part))

        pages = sorted({p for p in pages if 1 <= p <= total})
        if not pages:
            return abort(400, "No valid page numbers in range")

        zip_path = tmp_file(".zip")

        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pages:
                writer = PdfWriter()
                writer.add_page(reader.pages[p - 1])

                single_pdf = os.path.join(out_dir, f"page_{p}.pdf")
                with open(single_pdf, "wb") as o:
                    writer.write(o)

                z.write(single_pdf, arcname=f"page_{p}.pdf")

        return send_file(zip_path, as_attachment=True, download_name="split.zip")

    finally:
        cleanup(pdf)
        cleanup(out_dir)
        if zip_path:
            cleanup(zip_path)

# --------------------------------------------------------
# 8 - Rotate PDF
# --------------------------------------------------------
@app.post("/rotate-pdf")
def rotate_pdf():
    PdfReader, PdfWriter, _ = get_pypdf()

    f = request.files.get("file")
    angle = int(request.form.get("angle", 90) or 90)

    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        reader = PdfReader(pdf, strict=False)
        writer = PdfWriter()

        for page in reader.pages:
            try:
                page.rotate(angle)
            except Exception:
                page.rotate_clockwise(angle)
            writer.add_page(page)

        with open(out_pdf, "wb") as o:
            writer.write(o)

        return send_file(out_pdf, as_attachment=True, download_name="rotated.pdf")

    finally:
        cleanup(pdf)
        cleanup(out_pdf)

# --------------------------------------------------------
# 9 - Compress PDF (Ghostscript)
# --------------------------------------------------------
@app.post("/compress-pdf")
def compress_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    input_pdf = save_upload(f, ".pdf")
    output_pdf = tmp_file(".pdf")

    try:
        gs = shutil.which("gs")
        if not gs:
            return abort(500, "Ghostscript is missing on server")

        cmd = [
            gs,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/ebook",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_pdf}",
            input_pdf,
        ]

        subprocess.run(cmd, check=True)

        return send_file(
            output_pdf,
            as_attachment=True,
            download_name="compressed.pdf",
            mimetype="application/pdf"
        )

    except subprocess.CalledProcessError:
        return abort(500, "Compression failed")

    finally:
        cleanup(input_pdf)
        cleanup(output_pdf)

# --------------------------------------------------------
# 10 - Protect PDF
# --------------------------------------------------------
@app.post("/protect-pdf")
def protect_pdf():
    PdfReader, PdfWriter, _ = get_pypdf()

    f = request.files.get("file")
    pwd = request.form.get("password")

    if not f or not pwd:
        return abort(400, "Missing file or password")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        reader = PdfReader(pdf, strict=False)
        writer = PdfWriter()

        for p in reader.pages:
            writer.add_page(p)

        writer.encrypt(user_pwd=pwd)

        with open(out_pdf, "wb") as o:
            writer.write(o)

        return send_file(out_pdf, as_attachment=True, download_name="protected.pdf")

    finally:
        cleanup(pdf)
        cleanup(out_pdf)

# --------------------------------------------------------
# 11 - Unlock PDF
# --------------------------------------------------------
@app.post("/unlock-pdf")
def unlock_pdf():
    PdfReader, PdfWriter, _ = get_pypdf()

    f = request.files.get("file")
    pwd = request.form.get("password", "")

    if not f:
        return abort(400, "Missing file")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        reader = PdfReader(pdf, strict=False)

        if reader.is_encrypted:
            if not pwd:
                return abort(400, "Password required")
            reader.decrypt(pwd)

        writer = PdfWriter()
        for p in reader.pages:
            writer.add_page(p)

        with open(out_pdf, "wb") as o:
            writer.write(o)

        return send_file(out_pdf, as_attachment=True, download_name="unlocked.pdf")

    finally:
        cleanup(pdf)
        cleanup(out_pdf)

# --------------------------------------------------------
# 12 - Extract Text (with OCR fallback)
# --------------------------------------------------------
@app.post("/extract-text")
def extract_text():
    pdfplumber = get_pdfplumber()

    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")

    try:
        text_pages = []
        has_text = False

        # Normal text extraction
        try:
            with pdfplumber.open(pdf) as p:
                for i, page in enumerate(p.pages, start=1):
                    raw = page.extract_text() or ""
                    cleaned = " ".join(
                        line.strip()
                        for line in raw.split("\n")
                        if line.strip()
                    )
                    if cleaned:
                        has_text = True
                    text_pages.append(f"--- PAGE {i} ---\n{cleaned}\n")
        except Exception:
            pass

        # OCR fallback
        if not has_text:
            ocr_raw = ocr_pdf_to_text(pdf)

            ocr_raw = "\n".join(
                line for line in ocr_raw.split("\n")
                if not line.strip().startswith("--- PAGE")
            )
            paragraphs = [p.strip() for p in ocr_raw.split("\n\n") if p.strip()]
            formatted = "\n\n".join(paragraphs)

            return jsonify({"text": formatted})

        formatted = "\n\n".join(text_pages).strip()
        return jsonify({"text": formatted})

    finally:
        cleanup(pdf)

# --------------------------------------------------------
# ROOT + CORS
# --------------------------------------------------------
@app.get("/")
def home():
    return "PDF Tools Backend Running (Optimized for Render Free Tier)"

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

# --------------------------------------------------------
# START SERVER (LOCAL ONLY)
# --------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
