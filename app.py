"""
app.py
------
Flask web front-end for pdf_to_excel.py.
Accepts a single PDF upload and returns the generated .xlsx file.

Usage:
    pip install flask pdfplumber openpyxl
    python app.py
Then open http://127.0.0.1:5000 in your browser.
"""

import io
import tempfile
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify

from pdf_to_excel import parse_pdf, write_excel

app = Flask(__name__)
# 16 MB upload limit — adjust as needed
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "pdf_file" not in request.files:
        return jsonify({"error": "No file part in request."}), 400

    file = request.files["pdf_file"]

    if file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are accepted."}), 400

    # Save upload to a temp file so pdfplumber can open it by path
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
        file.save(tmp_pdf.name)
        tmp_pdf_path = tmp_pdf.name

    try:
        data = parse_pdf(tmp_pdf_path)
    except Exception as exc:
        return jsonify({"error": f"Failed to parse PDF: {exc}"}), 500
    finally:
        Path(tmp_pdf_path).unlink(missing_ok=True)

    # Write Excel to an in-memory buffer
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xlsx:
        tmp_xlsx_path = tmp_xlsx.name

    try:
        write_excel(data, tmp_xlsx_path)
        buffer = io.BytesIO(Path(tmp_xlsx_path).read_bytes())
    finally:
        Path(tmp_xlsx_path).unlink(missing_ok=True)

    buffer.seek(0)
    stem = Path(file.filename).stem
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{stem}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
