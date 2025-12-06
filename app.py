from flask import Flask, request, send_file, abort, jsonify, after_this_request, current_app
import os
import tempfile
import shutil
import subprocess
from werkzeug.utils import secure_filename
from flask_cors import CORS

from pdf2docx import Converter
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
import pikepdf
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import pdfplumber
import zipfile

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Poppler + PATH fix for Render
os.environ["PATH"] += ":/usr/bin:/usr/local/bin"
POPPLER_PATH = "/usr/bin"

# 200 MB max
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


# --------------------------------------------------------
# Utility helpers
# --------------------------------------------------------
def tmp_file(ext: str = "") -> str:
    """Create a temp file path with given extension."""
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
        f.write(file.stream.read())

    return path


def cleanup(path: str):
    """Safely delete file or directory if it exists."""
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.isfile(path):
            os.remove(path)
    except Exception:
        # Do not crash app just because cleanup fails
        pass


def ocr_pdf_to_text(pdf_path: str, dpi: int = 300) -> str:
    """
    OCR an image-based PDF using Tesseract via pdf2image.
    Returns plain text with simple page separators.
    """
    try:
        pages = convert_from_path(pdf_path, dpi=dpi, poppler_path=POPPLER_PATH)
    except Exception as e:
        current_app.logger.exception("OCR pdf2image failed")
        raise RuntimeError(f"OCR rasterization failed: {e}")

    texts = []
    for i, page in enumerate(pages, start=1):
        try:
            text = pytesseract.image_to_string(page)
        except Exception as e:
            current_app.logger.exception("Tesseract OCR failed on page %s", i)
            text = ""
        texts.append(f"--- PAGE {i} ---\n{text.strip()}")

    return "\n\n".join(texts).strip()


# --------------------------------------------------------
# 1 - PDF → Word  (with OCR fallback)
# --------------------------------------------------------
@app.post("/pdf-to-word")
def pdf_to_word():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

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
        # ---------------------------------------------------
        # TRY OPENING PDF (check encryption)
        # ---------------------------------------------------
        try:
            reader = PdfReader(pdf)
        except Exception as e:
            return jsonify({
                "error": "Unable to open PDF (possibly corrupted).",
                "details": str(e)
            }), 400

        pdf_to_use = pdf

        # ---------------------------------------------------
        # HANDLE ENCRYPTED PDF
        # ---------------------------------------------------
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
                    try:
                        import pikepdf
                        with pikepdf.open(pdf, password="") as pp:
                            pp.save(unlocked_pdf)
                        pdf_to_use = unlocked_pdf
                    except pikepdf._qpdf.PasswordError:
                        return jsonify({"error": "PDF is password protected. Use Unlock tool first."}), 400
            except Exception as e:
                return jsonify({"error": "Failed processing encrypted PDF", "details": str(e)}), 400

        # ---------------------------------------------------
        # FIRST TRY NORMAL PDF → WORD (pdf2docx)
        # ---------------------------------------------------
        from docx import Document

        try:
            cv = Converter(pdf_to_use)
            cv.convert(out_docx, start=0, end=None, layout_mode=True)
            cv.close()

            # Check if conversion produced text (not empty)
            doc_check = Document(out_docx)
            if len(doc_check.paragraphs) > 0:
                return send_file(out_docx, as_attachment=True, download_name="output.docx")

        except Exception:
            pass  # pdf2docx failed → fallback to OCR

        # ---------------------------------------------------
        # FALLBACK TO OCR IF pdf2docx FAILED
        # ---------------------------------------------------
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
# 2 - Word → PDF (high quality, embedded fonts)
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
            ["libreoffice", "--headless",
             "--convert-to", "odt",
             "--outdir", out_dir, doc],
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
# PERFECT + COMPRESSED PPT → PDF  (Slide → Image → PDF)
# --------------------------------------------------------
@app.post("/ppt-to-pdf-perfect")
def ppt_to_pdf_perfect():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in [".ppt", ".pptx"]:
        return abort(400, "Upload a PPT/PPTX file")

    ppt = save_upload(f, ext)
    temp_dir = tempfile.mkdtemp()
    out_pdf = tmp_file(".pdf")

    try:
        # ----------------------------------------------------
        # 1. Export each slide as a PNG using LibreOffice
        # ----------------------------------------------------
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "png",
                "--outdir", temp_dir,
                ppt
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True
        )

        # Get PNG slide images
        imgs = sorted(
            os.path.join(temp_dir, name)
            for name in os.listdir(temp_dir)
            if name.lower().endswith(".png")
        )

        if not imgs:
            return abort(500, "Failed to export slides.")

        # ----------------------------------------------------
        # 2. Convert PNG → compressed JPEG (smaller PDF size)
        # ----------------------------------------------------
        slide_images = []
        from io import BytesIO

        for img_path in imgs:
            img = Image.open(img_path).convert("RGB")

            # Reduce slide resolution a bit to reduce PDF size
            img = img.resize(
                (int(img.width * 0.85), int(img.height * 0.85)),
                Image.LANCZOS
            )

            # Compress to JPEG buffer
            buf = BytesIO()
            img.save(buf, format="JPEG", quality=85)   # balanced quality
            buf.seek(0)

            # Reload as Pillow image for PDF creation
            compressed = Image.open(buf)
            slide_images.append(compressed)

        # ----------------------------------------------------
        # 3. Save all slides into a single PDF
        # ----------------------------------------------------
        slide_images[0].save(
            out_pdf,
            "PDF",
            resolution=150,             # lower DPI → smaller size
            save_all=True,
            append_images=slide_images[1:],
        )

        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")

    finally:
        cleanup(ppt)
        cleanup(temp_dir) 


# --------------------------------------------------------
# 4 - JPG → PDF (300 DPI, max quality)
# --------------------------------------------------------
@app.post("/jpg-to-pdf")
def jpg_to_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files selected")

    images = []
    saved = []
    out_pdf = None

    try:
        for f in files:
            p = save_upload(f)
            saved.append(p)

            img = Image.open(p).convert("RGB")
            img = img.resize((img.width, img.height), resample=Image.LANCZOS)  # high quality resample
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
# 5 - PDF → JPG  (high quality, vertical merge)
# --------------------------------------------------------
@app.post("/pdf-to-jpg")
def pdf_to_jpg():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    pdf = save_upload(f, ".pdf")
    out_jpg = None

    try:
        # Slightly higher DPI for better clarity
        pages = convert_from_path(pdf, dpi=350, poppler_path=POPPLER_PATH)

        if not pages:
            return abort(400, "Failed to read PDF")

        imgs = [p.convert("RGB") for p in pages]
        total_height = sum(img.height for img in imgs)
        max_width = max(img.width for img in imgs)

        merged = Image.new("RGB", (max_width, total_height), "white")
        y = 0
        for img in imgs:
            merged.paste(img, (0, y))
            y += img.height

        out_jpg = tmp_file(".jpg")
        merged.save(out_jpg, "JPEG", quality=90)

        return send_file(
            out_jpg,
            as_attachment=True,
            download_name="output.jpg",
            mimetype="image/jpeg"
        )

    finally:
        cleanup(pdf)
        if out_jpg:
            cleanup(out_jpg)


# --------------------------------------------------------
# 6 - Merge PDF
# --------------------------------------------------------
@app.post("/merge-pdf")
def merge_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files")

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

    ranges = request.form.get("ranges")
    if not ranges:
        return abort(400, "Missing ranges")

    pdf = save_upload(f, ".pdf")
    out_dir = tempfile.mkdtemp()
    zip_path = None

    try:
        reader = PdfReader(pdf)
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

        # Only valid page numbers
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
    f = request.files.get("file")
    angle = int(request.form.get("angle", 90) or 90)

    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        reader = PdfReader(pdf)
        writer = PdfWriter()

        for page in reader.pages:
            # PyPDF2 new versions: rotate() is deprecated, use rotate_clockwise
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
    f = request.files.get("file")
    pwd = request.form.get("password")

    if not f or not pwd:
        return abort(400, "Missing file or password")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        reader = PdfReader(pdf)
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
    f = request.files.get("file")
    pwd = request.form.get("password", "")

    if not f:
        return abort(400, "Missing file")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        reader = PdfReader(pdf)

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
    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")

    try:
        text_pages = []
        has_text = False

        # ---------- Improve Normal Text Extraction ----------
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
        except:
            pass

        # ---------- OCR Fallback (formatted paragraphs) ----------
        if not has_text:
            ocr_raw = ocr_pdf_to_text(pdf)
            
            # Remove page markers like "--- PAGE X ---"
            ocr_raw = "\n".join(
                line for line in ocr_raw.split("\n")
                if not line.strip().startswith("--- PAGE")
            )           
            paragraphs = [p.strip() for p in ocr_raw.split("\n\n") if p.strip()]
            formatted = "\n\n".join(paragraphs)

            return jsonify({"text": formatted})

        # normal text result
        formatted = "\n\n".join(text_pages).strip()
        return jsonify({"text": formatted})

    finally:
        cleanup(pdf)


# --------------------------------------------------------
# Test
# --------------------------------------------------------
@app.get("/")
def home():
    return "PDF Tools Backend Running"


if __name__ == "__main__":
    app.run(debug=True)
