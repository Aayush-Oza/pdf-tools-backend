from flask import Flask, request, send_file, abort, jsonify
import os
import tempfile
import shutil
import subprocess
from werkzeug.utils import secure_filename
from flask_cors import CORS

from pdf2docx import Converter
from PIL import Image
from pdf2image import convert_from_path
import pikepdf
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import pdfplumber

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Allow poppler in Docker
os.environ["PATH"] += ":/usr/bin"

app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB


# --------------------------------------------------------
# Utility helpers
# --------------------------------------------------------
def tmp_file(ext=""):
    fd, path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    return path


def save_upload(file, ext=None):
    """ SAFE write — prevents corrupt PDF uploads """
    filename = secure_filename(file.filename)
    extension = ext if ext else os.path.splitext(filename)[1]
    path = tmp_file(extension)

    with open(path, "wb") as f_out:
        f_out.write(file.read())
    try:
        file.seek(0)
    except:
        pass

    return path


def cleanup(path):
    try:
        if os.path.isdir(path):
            shutil.rmtree(path)
        else:
            os.remove(path)
    except:
        pass


# --------------------------------------------------------
# 1 — PDF → Word
# --------------------------------------------------------
@app.post("/pdf-to-word")
def pdf_to_word():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")
    out_docx = tmp_file(".docx")

    try:
        cv = Converter(pdf)
        cv.convert(out_docx)
        cv.close()
        return send_file(out_docx, as_attachment=True, download_name="output.docx")
    finally:
        cleanup(pdf)
        cleanup(out_docx)


# --------------------------------------------------------
# 2 — Word → PDF
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
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, doc],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True
        )

        base = os.path.splitext(os.path.basename(doc))[0]
        out_pdf = os.path.join(out_dir, base + ".pdf")

        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")

    finally:
        cleanup(doc)
        cleanup(out_dir)


# --------------------------------------------------------
# 3 — PPT → PDF
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
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, ppt],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True
        )

        base = os.path.splitext(os.path.basename(ppt))[0]
        out_pdf = os.path.join(out_dir, base + ".pdf")

        return send_file(out_pdf, as_attachment=True, download_name="output.pdf")

    finally:
        cleanup(ppt)
        cleanup(out_dir)


# --------------------------------------------------------
# 4 — JPG → PDF
# --------------------------------------------------------
@app.post("/jpg-to-pdf")
def jpg_to_pdf():
    files = request.files.getlist("files")
    if not files:
        return abort(400, "No files selected")

    images = []
    temp_paths = []

    try:
        for f in files:
            p = save_upload(f)
            temp_paths.append(p)
            img = Image.open(p).convert("RGB")
            images.append(img)

        out = tmp_file(".pdf")
        images[0].save(out, save_all=True, append_images=images[1:])

        return send_file(out, as_attachment=True, download_name="output.pdf")

    finally:
        for p in temp_paths:
            cleanup(p)


# --------------------------------------------------------
# 5 — PDF → JPG  (single merged JPG)
# --------------------------------------------------------
@app.post("/pdf-to-jpg")
def pdf_to_jpg():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file uploaded")

    pdf = save_upload(f, ".pdf")

    try:
        pages = convert_from_path(pdf, dpi=200)

        if not pages:
            return abort(400, "Could not read PDF pages")

        imgs = [p.convert("RGB") for p in pages]

        total_height = sum(img.height for img in imgs)
        max_width = max(img.width for img in imgs)

        merged = Image.new("RGB", (max_width, total_height), (255, 255, 255))

        y = 0
        for img in imgs:
            merged.paste(img, (0, y))
            y += img.height

        out_jpg = tmp_file(".jpg")
        merged.save(out_jpg, "JPEG", quality=90)

        return send_file(out_jpg, as_attachment=True, download_name="output.jpg")

    finally:
        cleanup(pdf)


# --------------------------------------------------------
# 6 — Merge PDFs (best + correct way)
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

        out = tmp_file(".pdf")
        merger.write(out)
        merger.close()

        return send_file(out, as_attachment=True, download_name="merged.pdf")

    finally:
        for p in saved:
            cleanup(p)
        cleanup(out)


# --------------------------------------------------------
# 7 — Split PDF
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
                a, b = part.split("-")
                pages.extend(range(int(a), int(b) + 1))
            else:
                pages.append(int(part))

        pages = [p for p in pages if 1 <= p <= total]

        zip_path = tmp_file(".zip")
        import zipfile
        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pages:
                w = PdfWriter()
                w.add_page(reader.pages[p - 1])

                out_pdf = os.path.join(out_dir, f"page_{p}.pdf")
                with open(out_pdf, "wb") as o:
                    w.write(o)

                z.write(out_pdf, arcname=f"page_{p}.pdf")

        return send_file(zip_path, as_attachment=True, download_name="split.zip")

    finally:
        cleanup(pdf)
        cleanup(out_dir)


# --------------------------------------------------------
# 8 — Rotate PDF
# --------------------------------------------------------
@app.post("/rotate-pdf")
def rotate_pdf():
    f = request.files.get("file")
    angle = int(request.form.get("angle", "90"))

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
# 9 — Compress PDF
# --------------------------------------------------------
@app.post("/compress-pdf")
def compress_pdf():
    f = request.files.get("file")
    if not f:
        return abort(400, "No file")

    pdf = save_upload(f, ".pdf")
    out_pdf = tmp_file(".pdf")

    try:
        try:
            p = pikepdf.open(pdf, allow_overwriting_input=True)
        except Exception as e:
            return abort(400, f"Invalid or damaged PDF: {str(e)}")

        p.save(out_pdf, optimize_streams=True, linearize=True)
        p.close()

        return send_file(out_pdf, as_attachment=True, download_name="compressed.pdf")

    finally:
        cleanup(pdf)
        cleanup(out_pdf)


# --------------------------------------------------------
# 10 — Protect PDF
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
# 11 — Unlock PDF
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
# 12 — Extract Text
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
                t = page.extract_text() or ""
                text.append(t)

        full_text = "\n\n--- PAGE BREAK ---\n\n".join(text)
        return jsonify({"text": full_text})

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
