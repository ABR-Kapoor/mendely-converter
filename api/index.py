from flask import Flask, jsonify, request, send_file, render_template_string
from io import BytesIO

from mendeley_core import analyze_uploaded_documents, convert_docx_to_mendeley

app = Flask(__name__)


INDEX_HTML = """
<!doctype html>
<title>Mendeley Converter</title>
<h1>Mendeley Converter (Vercel)</h1>
<form action="/convert" method="post" enctype="multipart/form-data">
  <label>Word (.docx): <input type="file" name="docx" accept=".docx" required></label><br><br>
  <label>BibTeX (.bib): <input type="file" name="bib" accept=".bib" required></label><br><br>
  <button type="submit">Convert</button>
</form>
"""


@app.get("/")
def home():
    return render_template_string(INDEX_HTML)


@app.post("/convert")
def convert():
    if "docx" not in request.files or "bib" not in request.files:
        return jsonify({"error": "Missing files"}), 400

    docx_file = request.files["docx"]
    bib_file = request.files["bib"]

    # Prepare file-like objects for the core module
    docx_bytes = docx_file.read()
    bib_bytes = bib_file.read()

    analysis = analyze_uploaded_documents(io.BytesIO(docx_bytes), io.BytesIO(bib_bytes))

    if analysis.get("unresolved"):
        return jsonify({"error": "unresolved_citations", "unresolved": analysis["unresolved"]}), 400

    try:
        out_bytes = convert_docx_to_mendeley(docx_bytes, analysis["resolved_map"])
    except Exception as exc:
        return jsonify({"error": "conversion_failed", "message": str(exc)}), 500

    return send_file(BytesIO(out_bytes), as_attachment=True, download_name="paper_mendeley.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.get("/health")
def health():
    return jsonify({"status": "healthy"})


if __name__ == "__main__":
    app.run(debug=True, port=8000)
