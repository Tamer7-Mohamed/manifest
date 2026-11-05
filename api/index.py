
from flask import Flask, request, jsonify, send_file
import os
import sys
import json
import tempfile
import traceback

sys.path.append(os.path.dirname(os.path.dirname(__file__)))

import manifest_to_excel_hapag_egypt as parser

app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "online",
        "message": "Hapag-Lloyd Egypt Manifest Checker"
    })

@app.route("/api/process", methods=["POST"])
def process_manifest():
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

        # فتح ملف PDF
        total_pages = parser.open_pdf(temp_pdf)

        # معالجة الصفحات على دفعات لتقليل استهلاك الذاكرة
        batch_size = 30
        all_parts = []

        for start in range(0, total_pages, batch_size):
            end = min(start + batch_size, total_pages)
            result = parser.process(start, end)
            all_parts.append(result)

        # إنشاء ملف Excel النهائي
        parser.finish(json.dumps(all_parts))

        output_path = "/tmp/out.xlsx"

        if not os.path.exists(output_path):
            return jsonify({
                "error": "لم يتم إنشاء ملف Excel"
            }), 500

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
        if temp_pdf and os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except Exception:
                pass
