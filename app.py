import re
import tempfile
import zipfile
from pathlib import Path
from typing import Tuple

import fitz  # PyMuPDF
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

# Maximum upload size: 100 MB
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024


# ----------------------------
# Helper Functions
# ----------------------------
def sanitize_filename(name: str) -> str:
    """Remove invalid filename characters."""
    name = str(name).strip()
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name.rstrip(". ") or "output"


def parse_page_range(page_range_value) -> Tuple[int, int]:
    """
    Supported formats:
    - 3
    - 1 to 2
    - 1-2
    - 1,2
    """
    text = str(page_range_value).strip().lower()
    text = text.replace("–", "-").replace("—", "-")

    numbers = [int(x) for x in re.findall(r"\d+", text)]

    if not numbers:
        raise ValueError(f"Invalid page range: {page_range_value}")

    if len(numbers) == 1:
        return numbers[0], numbers[0]

    return numbers[0], numbers[1]


def save_pdf(doc: fitz.Document, start: int, end: int, output_path: Path) -> None:
    """Save selected PDF page range as a new PDF."""
    new_doc = fitz.open()
    new_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
    new_doc.save(output_path)
    new_doc.close()


def save_jpg(doc: fitz.Document, start: int, end: int, output_base_path: Path) -> None:
    """
    Save one or multiple pages as JPG.
    If single page -> output_name.jpg
    If multiple pages -> output_name_1.jpg, output_name_2.jpg, ...
    """
    page_count = end - start + 1

    for i, page_no in enumerate(range(start, end + 1), start=1):
        page = doc.load_page(page_no - 1)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Better quality

        if page_count == 1:
            jpg_path = output_base_path.with_suffix(".jpg")
        else:
            jpg_path = output_base_path.parent / f"{output_base_path.stem}_{i}.jpg"

        pix.save(jpg_path)


def make_zip(folder_path: Path, zip_path: Path) -> None:
    """Create ZIP file from output folder."""
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file_path in folder_path.rglob("*"):
            if file_path.is_file():
                zipf.write(file_path, file_path.relative_to(folder_path))


# ----------------------------
# Routes
# ----------------------------
@app.route("/")
def home():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    pdf_file = request.files.get("pdf")

    if not pdf_file or not pdf_file.filename:
        return render_template("index.html", error="કૃપા કરીને PDF ફાઇલ અપલોડ કરો.")

    # Read editable table data from webpage
    page_ranges = request.form.getlist("page_range[]")
    output_formats = request.form.getlist("output_format[]")
    output_names = request.form.getlist("output_name[]")

    if not page_ranges or not output_formats or not output_names:
        return render_template("index.html", error="કૃપા કરીને ટેબલમાં જરૂરી માહિતી भरो.")

    if not (len(page_ranges) == len(output_formats) == len(output_names)):
        return render_template("index.html", error="ટેબલમાં આપેલ માહિતી યોગ્ય નથી.")

    # Create temp working folder
    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_split_tool_"))
    pdf_path = temp_dir / sanitize_filename(pdf_file.filename)
    output_dir = temp_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Save uploaded PDF
    pdf_file.save(pdf_path)

    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)

        files_created = 0

        for page_range, file_format, output_name in zip(page_ranges, output_formats, output_names):
            page_range = str(page_range).strip()
            file_format = str(file_format).strip().lower()
            output_name = str(output_name).strip()

            if not page_range or not file_format or not output_name:
                continue

            try:
                start_page, end_page = parse_page_range(page_range)
            except ValueError:
                continue

            # Skip invalid page ranges
            if start_page < 1 or end_page > total_pages or start_page > end_page:
                continue

            safe_output_name = sanitize_filename(output_name)
            output_base = output_dir / safe_output_name

            if file_format == "pdf":
                save_pdf(doc, start_page, end_page, output_base.with_suffix(".pdf"))
                files_created += 1

            elif file_format in {"jpg", "jpeg"}:
                save_jpg(doc, start_page, end_page, output_base)
                files_created += 1

            else:
                # Unsupported format
                continue

        doc.close()

        if files_created == 0:
            return render_template(
                "index.html",
                error="કોઈ માન્ય output તૈયાર થયું નથી. કૃપા કરીને page range અને format ચકાસો."
            )

        zip_path = temp_dir / "result.zip"
        make_zip(output_dir, zip_path)

        return send_file(
            zip_path,
            as_attachment=True,
            download_name="result.zip",
            mimetype="application/zip"
        )

    except Exception as e:
        return render_template("index.html", error=f"ભૂલ આવી: {str(e)}")


# ----------------------------
# Run App
# ----------------------------
if __name__ == "__main__":
    app.run(debug=True)
