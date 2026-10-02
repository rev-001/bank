"""
Flask Web Application & API Server for the Offline PDF-to-Excel Agent.
Provides an interactive dashboard, live data editor, agent chat console,
and instant Excel export.
"""

import os
import sys
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_from_directory, send_file
from werkzeug.utils import secure_filename

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from agent.core_agent import PDFToExcelAgent

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["UPLOAD_FOLDER"] = str(_PROJECT_ROOT / "uploads")
app.config["OUTPUT_FOLDER"] = str(_PROJECT_ROOT / "outputs")
app.config["SAMPLES_FOLDER"] = str(_PROJECT_ROOT / "samples")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB max

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

agent = PDFToExcelAgent()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/samples", methods=["GET"])
def get_samples():
    samples_dir = Path(app.config["SAMPLES_FOLDER"])
    if not samples_dir.exists():
        return jsonify([])
    samples = []
    for f in sorted(samples_dir.glob("*.pdf")):
        samples.append({
            "name": f.name,
            "path": str(f),
            "size": f.stat().st_size
        })
    return jsonify(samples)


@app.route("/api/extract", methods=["POST"])
def extract_pdf():
    pdf_path = None
    
    # Check if uploaded file or sample path
    if "file" in request.files:
        file = request.files["file"]
        if file.filename != "":
            filename = secure_filename(file.filename)
            pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(pdf_path)
    elif request.is_json and "sample_path" in request.json:
        pdf_path = request.json["sample_path"]

    if not pdf_path or not os.path.exists(pdf_path):
        return jsonify({"error": "No valid PDF file provided"}), 400

    chat_instruction = request.json.get("chat_instruction") if request.is_json else None
    template_name = request.json.get("template_name") if request.is_json else None

    try:
        result = agent.process_pdf(
            pdf_path=pdf_path,
            template_name=template_name,
            chat_instruction=chat_instruction
        )
        master_name = Path(result.get("master_excel_path") or result["excel_path"]).name
        result["download_url"] = f"/api/download/{master_name}"
        result["excel_path"] = result.get("master_excel_path") or result["excel_path"]
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/chat", methods=["POST"])
def agent_chat():
    data = request.json or {}
    prompt = data.get("prompt", "").strip()
    pdf_path = data.get("pdf_path")
    
    if not prompt:
        return jsonify({"error": "Empty prompt"}), 400

    # Learn the rule
    learn_res = agent.learn_from_chat(prompt)

    # If an active PDF is currently being worked on, re-process it with the new knowledge
    reprocessed = None
    if pdf_path and os.path.exists(pdf_path):
        try:
            reprocessed = agent.process_pdf(pdf_path=pdf_path)
            excel_filename = Path(reprocessed.get("master_excel_path") or reprocessed["excel_path"]).name
            reprocessed["download_url"] = f"/api/download/{excel_filename}"
        except Exception:
            pass

    return jsonify({
        "learning": learn_res,
        "reprocessed": reprocessed
    })


@app.route("/api/save-template", methods=["POST"])
def save_template():
    data = request.json or {}
    template_name = data.get("template_name")
    pdf_path = data.get("pdf_path")
    column_mappings = data.get("column_mappings", {})

    if not template_name or not pdf_path or not os.path.exists(pdf_path):
        return jsonify({"error": "Missing template_name or valid pdf_path"}), 400

    try:
        t_id = agent.save_current_as_template(template_name, pdf_path, column_mappings)
        return jsonify({"status": "success", "template_id": t_id, "message": f"Template '{template_name}' saved."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/export-custom", methods=["POST"])
def export_custom():
    """Allows user to submit manually edited table rows/headers and export custom Excel."""
    data = request.json or {}
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    key_values = data.get("key_values", {})
    filename = data.get("filename", "custom_export")
    is_bank_statement = bool(data.get("is_bank_statement"))
    
    table_data = {
        "headers": headers,
        "rows": rows,
        "column_types": ["text"] * len(headers)
    }

    # Re-infer column types
    table_data = agent.schema_matcher.standardize_table(table_data)
    if is_bank_statement:
        table_data["headers"] = [
            "Txn Date",
            "Value Date",
            "Particulars",
            "Ref. No.",
            "Debit",
            "Credit",
            "Balance",
        ]
        bank_column_types = {
            "txn date": "date",
            "value date": "date",
            "particulars": "text",
            "ref. no.": "text",
            "debit": "currency",
            "credit": "currency",
            "balance": "currency",
        }
        table_data["column_types"] = [
            bank_column_types.get(str(header).strip().lower(), "text")
            for header in table_data["headers"]
        ]
    
    val_result = agent.validator.validate_and_reconcile(table_data, key_values)

    if is_bank_statement:
        output_name = f"{secure_filename(Path(filename).stem) or 'bank_statement'}_extracted.xlsx"
        output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_name)
        agent.exporter.export(
            output_path=output_path,
            table_data=val_result["table"],
            key_values=key_values,
            validation_info=val_result,
            metadata={
                "filename": filename,
                "template_name": "User-Edited Bank Statement",
                "accuracy_score": 1.0,
                "currency": "INR",
                "is_bank_statement": True,
            },
        )
        return jsonify({
            "status": "success",
            "excel_path": output_path,
            "download_url": f"/api/download/{output_name}",
        })

    master_path = os.path.join(app.config["OUTPUT_FOLDER"], "all_transactions.xlsx")
    agent.exporter.merge_into_master(
        master_path=master_path,
        table_data=val_result["table"],
        key_values=key_values,
        validation_info=val_result,
        metadata={"filename": filename, "template_name": "User-Edited Custom", "accuracy_score": 1.0}
    )

    return jsonify({
        "status": "success",
        "excel_path": master_path,
        "download_url": "/api/download/all_transactions.xlsx"
    })


@app.route("/api/history", methods=["GET"])
def get_history():
    return jsonify(agent.get_history(limit=50))


@app.route("/api/templates", methods=["GET"])
def get_templates():
    return jsonify(agent.get_templates())


@app.route("/api/download/<filename>", methods=["GET"])
def download_file(filename):
    out_dir = app.config["OUTPUT_FOLDER"]
    return send_from_directory(out_dir, filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
