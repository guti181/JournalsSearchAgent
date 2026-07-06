
from __future__ import annotations

import secrets
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, send_file

from generator import generate_outputs

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/generate")
def generate():
    keren_file = request.files.get("keren")
    agreement_file = request.files.get("agreement")
    include_cleaned = request.form.get("include_cleaned_excels") == "on"

    if not keren_file or not agreement_file:
        return render_template("index.html", error="Please upload both Excel files."), 400

    if not keren_file.filename.lower().endswith(".xlsx") or not agreement_file.filename.lower().endswith(".xlsx"):
        return render_template("index.html", error="Only .xlsx files are supported."), 400

    token = secrets.token_hex(4)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    keren_path = UPLOAD_DIR / f"keren_raw_{ts}_{token}.xlsx"
    agreement_path = UPLOAD_DIR / f"agreement_raw_{ts}_{token}.xlsx"
    out_dir = OUTPUT_DIR / f"run_{ts}_{token}"
    out_dir.mkdir(parents=True, exist_ok=True)

    keren_file.save(keren_path)
    agreement_file.save(agreement_path)

    try:
        result = generate_outputs(str(keren_path), str(agreement_path), str(out_dir), include_cleaned_excels=include_cleaned)
    except Exception as exc:
        return render_template("index.html", error=f"Generation failed: {exc}"), 500

    if not include_cleaned:
        return send_file(result["html_path"], as_attachment=True, download_name=Path(result["html_path"]).name)

    zip_path = OUTPUT_DIR / f"generated_bundle_{ts}_{token}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for key in ["html_path", "cleaned_keren_xlsx_path", "cleaned_agreement_xlsx_path", "processed_keren_xlsx_path", "processed_agreement_xlsx_path"]:
            file_path = Path(result[key])
            zf.write(file_path, file_path.name)

    return send_file(zip_path, as_attachment=True, download_name="generated_bundle.zip")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)
