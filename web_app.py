"""
Web interface for AI Small Business Operations Agent.
Run with: python web_app.py
Open http://localhost:5000 in your browser.
"""

import os
import sys
import json
from datetime import datetime, timezone
from flask import Flask, render_template, request, jsonify, send_file

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from agent import AIBusinessAgent

app = Flask(__name__)
agent = AIBusinessAgent()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/process", methods=["POST"])
def process_order():
    data = request.json
    message = data.get("message", "")
    customer = data.get("customer", "Customer")

    if not message:
        return jsonify({"error": "No message provided"}), 400

    result = agent.process_order(message, customer)
    return jsonify(result)


@app.route("/api/products")
def get_products():
    products = agent.sheets.read_products() if agent.sheets else []
    return jsonify(products)


@app.route("/api/orders")
def get_orders():
    orders = getattr(agent.sheets, "orders", []) if agent.sheets else []
    return jsonify(orders)


@app.route("/api/logs")
def get_logs():
    logs = getattr(agent.sheets, "logs", []) if agent.sheets else []
    return jsonify(logs)


@app.route("/api/download-excel")
def download_excel():
    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventory.xlsx")
    if os.path.exists(excel_path):
        return send_file(excel_path, as_attachment=True, download_name="inventory.xlsx")
    return jsonify({"error": "Excel file not found"}), 404


@app.route("/api/sheet-info")
def sheet_info():
    sheet_id = os.environ.get("GOOGLE_SHEET_ID", "1Q3aTYv-3WpF2pX1QEFlsbhUbSMuis6XI6DrjQQnGR5o")
    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventory.xlsx")
    return jsonify({
        "excel_local_path": excel_path,
        "excel_exists": os.path.exists(excel_path),
        "google_sheet_url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
        "google_sheet_id": sheet_id
    })


@app.route("/api/health")
def health():
    return jsonify({
        "status": "ok",
        "deepseek_configured": bool(os.environ.get("DEEPSEEK_API_KEY")),
        "telegram_configured": bool(os.environ.get("TELEGRAM_BOT_TOKEN")),
        "sheets_configured": bool(os.environ.get("GOOGLE_SHEET_ID") or agent.sheets),
        "excel_persisted": os.path.exists(os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventory.xlsx")),
    })


if __name__ == "__main__":
    app.run(debug=True, port=5000)
