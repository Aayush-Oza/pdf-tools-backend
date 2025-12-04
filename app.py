from flask import Flask, request, send_file, abort, jsonify
import os
import tempfile
import shutil
import subprocess
from werkzeug.utils import secure_filename
from flask_cors import CORS
from flask import after_this_request, current_app
from pdf2docx import Converter
from PIL import Image
from pdf2image import convert_from_path
import pikepdf
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import pdfplumber
import zipfile
import shutil

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Poppler + PATH fix for Render
os.environ["PATH"] += ":/usr/bin:/usr/local/bin"
POPPLER_PATH = "/usr/bin"

app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


# --------------------------------------------------------
# Utility helpers
# --------------------------------------------------------
def tmp_file(ext=""):
    fd, path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    return path


def save_upload(file, ext=None):
    filename = secure_filename(file.filename)
    extension = ext if ext else os.path.splitext(filename)[1]
    path = tmp_file(extension)

    file.stream.seek(0)
    with open(path, "wb") as f:
        f.write(file.stream.read())

    return path


def cleanup(path):
    try:
        if os.path.isdir(path):
            shutil.rmtree(path)
        elif os.path.isfile(path):
            os.remove(path)
    except:
        pass


# --------------------------------------------------------
# 1 - PDF → Word
# --------------------------------------------------------
@app.post("/pdf-to-word")
def pdf_to_word():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    pdf = save_upload(f, ".pdf")
    unlocked_pdf = tmp_file(".pdf")
    out_docx = tmp_file(".docx")

    # register cleanup of generated files AFTER response is created
    @after_this_request
    def _cleanup_response(response):
        try:
            cleanup(pdf)
            cleanup(unlocked_pdf)
            cleanup(out_docx)
        except Exception as e:
            current_app.logger.exception("Cleanup failed: %s", e)
        return response

    try:
        # -------------------------
        # Try to open and detect encryption
        # -------------------------
        try:
            reader = PdfReader(pdf)
        except Exception as e:
            current_app.logger.exception("PdfReader failed to open PDF")
            return jsonify({"error": "Unable to open PDF (possibly corrupted).", "details": str(e)}), 400

        pdf_to_use = pdf

        # If encrypted -> attempt to remove owner password with pikepdf (stronger)
        if getattr(reader, "is_encrypted", False):
            try:
                # Try trivial decrypt via PyPDF2 first (empty password)
                try:
                    reader.decrypt("")
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    with open(unlocked_pdf, "wb") as f2:
                        writer.write(f2)
                    pdf_to_use = unlocked_pdf
                except Exception:
                    # fallback to pikepdf (can remove owner-passwords)
                    try:
                        with pikepdf.open(pdf, password="") as pp:
                            pp.save(unlocked_pdf)
                        pdf_to_use = unlocked_pdf
                    except pikepdf._qpdf.PasswordError:
                        return jsonify({"error": "PDF is password protected. Use Unlock tool first."}), 400
                    except Exception as e:
                        current_app.logger.exception("pikepdf unlock failed")
                        return jsonify({"error": "Failed to unlock PDF", "details": str(e)}), 400
            except Exception as e:
                current_app.logger.exception("Encryption handling failed")
                return jsonify({"error": "Failed processing encrypted PDF", "details": str(e)}), 400

        # -------------------------
        # Quick check: is PDF image-only (scanned)?
        # -------------------------
        try:
            text_found = False
            with pdfplumber.open(pdf_to_use) as p:
                # check first 2 pages for visible text
                for page in p.pages[:2]:
                    txt = (page.extract_text() or "").strip()
                    if txt:
                        text_found = True
                        break
        except Exception as e:
            current_app.logger.exception("pdfplumber check failed; continuing to conversion")

        if not text_found:
            # it's likely scanned or image only — pdf2docx doesn't do OCR
            return jsonify({
                "error": "PDF appears to be image/scanned (no selectable text).",
                "suggestion": "Use an OCR step first (Tesseract) or convert PDF->JPG and run OCR."
            }), 400

        # -------------------------
        # Convert PDF -> DOCX
        # -------------------------
        try:
            cv = Converter(pdf_to_use)
            # use convert with explicit start/end to avoid hidden failures
            cv.convert(out_docx, start=0, end=None)
            cv.close()
        except Exception as e:
            # log stacktrace and return message to client
            current_app.logger.exception("pdf2docx conversion failed")
            return jsonify({"error": "Conversion failed", "details": str(e)}), 500

        # -------------------------
        # Send file (after_this_request will cleanup)
        # -------------------------
        return send_file(out_docx, as_attachment=True, download_name="output.docx")

    finally:
        # do not remove output files here — cleanup happens in after_this_request
        pass

# --------------------------------------------------------
# 2 - Word → PDF
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
        # STEP 1 → Convert DOCX → ODT
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

        # STEP 2 → Convert ODT → PDF (HIGH QUALITY)
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
        return abort(400, "No file")

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in [".ppt", ".pptx"]:
        return abort(400, "Upload a PPT/PPTX file")

    ppt = save_upload(f, ext)
    out_dir = tempfile.mkdtemp()

    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf",
             "--outdir", out_dir, ppt],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True
        )

        base = os.path.splitext(os.path.basename(ppt))[0]
        out_pdf = os.path.join(out_dir, f"{base}.pdf")

        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")
    finally:
        cleanup(ppt)
        cleanup(out_dir)


# --------------------------------------------------------
# 4 - JPG → PDF
# --------------------------------------------------------
@app.post("/jpg-to-pdf")
def jpg_to_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files selected")

    images = []
    saved = []

    try:
        for f in files:
            p = save_upload(f)
            saved.append(p)

            img = Image.open(p).convert("RGB")
            # Increase DPI & preserve clarity
            img.info["dpi"] = (300, 300)
            images.append(img)

        out_pdf = tmp_file(".pdf")

        images[0].save(
            out_pdf,
            save_all=True,
            append_images=images[1:],
            quality=100,         # max quality
            dpi=(300, 300)       # print quality
        )

        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")

    finally:
        for p in saved:
            cleanup(p)



# --------------------------------------------------------
# 5 - PDF → JPG
# --------------------------------------------------------
@app.post("/pdf-to-jpg")
def pdf_to_jpg():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    pdf = save_upload(f, ".pdf")

    try:
        pages = convert_from_path(pdf, dpi=200, poppler_path=POPPLER_PATH)

        if not pages:
            return abort(400, "Failed to read PDF")

        # Merge all pages vertically
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

        return send_file(out_jpg, as_attachment=True,
                         download_name="output.jpg", mimetype="image/jpeg")

    finally:
        cleanup(pdf)
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

    try:
        reader = PdfReader(pdf)
        total = len(reader.pages)

        pages = []
        for part in ranges.split(","):
            if "-" in part:
                a, b = map(int, part.split("-"))
                pages.extend(range(a, b + 1))
            else:
                pages.append(int(part))

        pages = [p for p in pages if 1 <= p <= total]

        zip_path = tmp_file(".zip")

        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pages:
                writer = PdfWriter()
                writer.add_page(reader.pages[p - 1])

                out_pdf = os.path.join(out_dir, f"page_{p}.pdf")
                with open(out_pdf, "wb") as o:
                    writer.write(o)

                z.write(out_pdf, arcname=f"page_{p}.pdf")

        return send_file(zip_path, as_attachment=True, download_name="split.zip")

    finally:
        cleanup(pdf)
        cleanup(out_dir)


# --------------------------------------------------------
# 8 - Rotate PDF
# --------------------------------------------------------
@app.post("/rotate-pdf")
def rotate_pdf():
    f = request.files.get("file")
    angle = int(request.form.get("angle", 90))

    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        reader = PdfReader(pdf)
        writer = PdfWriter()

        for page in reader.pages:
            page.rotate(angle)
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

        return send_file(output_pdf, as_attachment=True,
                         download_name="compressed.pdf", mimetype="application/pdf")

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
# 12 - Extract Text
# --------------------------------------------------------
@app.post("/extract-text")
def extract_text():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")

    try:
        text = []
        with pdfplumber.open(pdf) as p:
            for page in p.pages:
                text.append(page.extract_text() or "")

        return jsonify({"text": "\n\n--- PAGE BREAK ---\n\n".join(text)})

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
