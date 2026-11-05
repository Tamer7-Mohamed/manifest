
from flask import Flask, request, jsonify, send_file
import os
import sys
import tempfile
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import manifest_to_excel_hapag_egypt as parser

app = Flask(__name__)


@app.route("/", methods=["GET"])
def home():
    return jsonify({"status": "API is working"})


@app.route("/<path:path>", methods=["POST"])
@app.route("/", methods=["POST"])
def process_manifest(path=""):

    if "file" not in request.files:
        return jsonify({"error": "لم يتم إرسال ملف PDF"}), 400

    uploaded_file = request.files["file"]

    if not uploaded_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "يجب رفع ملف PDF فقط"}), 400

    temp_pdf = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".pdf", delete=False
        ) as tmp:
            uploaded_file.save(tmp)
            temp_pdf = tmp.name

        total_pages = parser.open_pdf(temp_pdf)

        batch_size = 15
        all_parts = []

        for start in range(0, total_pages, batch_size):
            end = min(start + batch_size, total_pages)
            result = parser.process(start, end)
            all_parts.append(result)

        parser.finish(__import__("json").dumps(all_parts))

        output_path = "/tmp/out.xlsx"

        if not os.path.isfile(output_path):
            return jsonify({"error": "لم يتم إنشاء ملف Excel"}), 500

        return send_file(
            output_path,
            as_attachment=True,
            download_name="Hapag_Lloyd_Egypt_Checked.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "error": "حدث خطأ أثناء تحليل المانيفست",
            "details": str(e)
        }), 500

    finally:
        try:
            if parser._pdf is not None:
                parser._pdf.close()
                parser._pdf = None
        except Exception:
            pass

        if temp_pdf and os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except Exception:
                pass
