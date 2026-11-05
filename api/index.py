
from flask import Flask, request, jsonify
import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(__file__)))

app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "online",
        "message": "Hapag-Lloyd Egypt Manifest Checker API"
    })

@app.route("/api", methods=["GET"])
def api_status():
    return jsonify({"status": "API is running"})
