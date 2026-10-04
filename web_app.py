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
app.config["TEMPLATES_AUTO_RELOAD"] = True
agent = AIBusinessAgent()


@app.after_request
def add_header(response):
    """Disable caching for all API endpoints so UI always displays real-time stock."""
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/process", methods=["POST"])
def process_order():
    data = request.json or {}
    message = data.get("message", "")
    customer = data.get("customer", "Customer")

    if not message:
        return jsonify({"error": "No message provided"}), 400

    result = agent.process_order(message, customer)
    return jsonify(result)


@app.route("/api/pos-checkout", methods=["POST"])
def pos_checkout():
    """Direct POS cart checkout supporting multiple line items and accurate stock deduction."""
    data = request.json or {}
    items = data.get("items", [])
    customer = data.get("customer", "Walk-In Customer")

    if not items:
        return jsonify({"error": "No cart items provided"}), 400

    import uuid
    from datetime import datetime, timezone
    from agent import (
        TELEGRAM_CUSTOMER_CHAT_ID,
        TELEGRAM_OWNER_CHAT_ID,
        AUTO_APPROVE_MAX_VALUE,
    )

    order_id = f"ORD-{uuid.uuid4().hex[:8].upper()}"
    timestamp = datetime.now(timezone.utc).isoformat()

    total_value = sum(float(i.get("price", 0)) * int(i.get("qty", 1)) for i in items)
    inventory = agent.sheets.read_products() if agent.sheets else []
    inv_map = {item.get("SKU"): item for item in inventory}

    # Rule 1: High-value escalation (>= $100)
    if total_value >= AUTO_APPROVE_MAX_VALUE:
        items_summary = "\n".join(
            f"  • {i.get('qty')}x {i.get('product')} ({i.get('variant')}) - ${float(i.get('price', 0)) * int(i.get('qty', 1)):.2f}"
            for i in items
        )
        msg = (
            f"⚠️ <b>Order Escalation Required</b>\n\n"
            f"Order ID: {order_id}\n"
            f"Customer: {customer}\n"
            f"Total Value: ${total_value:.2f} (Threshold: ${AUTO_APPROVE_MAX_VALUE})\n\n"
            f"Items:\n{items_summary}\n\n"
            f"Reason: Order value exceeds ${AUTO_APPROVE_MAX_VALUE}. Manager override required before stock deduction."
        )
        if TELEGRAM_OWNER_CHAT_ID:
            agent.telegram.send_message(TELEGRAM_OWNER_CHAT_ID, msg)

        if agent.sheets:
            agent.sheets.append_log({
                "log_id": f"LOG-{uuid.uuid4().hex[:8].upper()}",
                "timestamp": timestamp,
                "action": "owner_escalation",
                "order_id": order_id,
                "details": f"High value cart order (${total_value:.2f} >= ${AUTO_APPROVE_MAX_VALUE})",
                "customer_notified": "pending",
                "owner_notified": "yes",
            })

        return jsonify({
            "status": "escalated",
            "order_id": order_id,
            "total": round(total_value, 2),
            "reason": f"Order total ${total_value:.2f} exceeds ${AUTO_APPROVE_MAX_VALUE} threshold."
        })

    # Rule 2: Stock sufficiency check
    for i in items:
        sku = i.get("sku")
        qty = int(i.get("qty", 1))
        inv_item = inv_map.get(sku)
        if not inv_item or inv_item.get("Stock", 0) < qty:
            # Insufficient stock -> suggest alternative branch
            return jsonify({
                "status": "alternative_suggested",
                "order_id": order_id,
                "total": round(total_value, 2),
                "reason": f"Insufficient shelf stock for {i.get('product')}. Please offer backstock alternatives."
            })

    # Rule 3: Auto-approve & fulfill
    for i in items:
        sku = i.get("sku")
        qty = int(i.get("qty", 1))
        inv_item = inv_map.get(sku)
        new_stock = inv_item.get("Stock", 0) - qty
        if agent.sheets:
            agent.sheets.update_stock(sku, new_stock)
            inv_item["Stock"] = new_stock

        if agent.sheets:
            agent.sheets.append_order({
                "order_id": order_id,
                "customer": customer,
                "product": i.get("product"),
                "variant": i.get("variant", ""),
                "quantity": qty,
                "total_value": float(i.get("price", 0)) * qty,
                "status": "approved",
                "timestamp": timestamp,
                "notes": "POS Auto-fulfilled (<$100)",
            })

    items_receipt = "\n".join(
        f"  • {i.get('qty')}x {i.get('product')} ({i.get('variant')}) - ${float(i.get('price', 0)) * int(i.get('qty', 1)):.2f}"
        for i in items
    )
    receipt_msg = (
        f"✅ <b>Order Approved (Elora POS)</b>\n\n"
        f"Order: {order_id}\n"
        f"Customer: {customer}\n"
        f"Items:\n{items_receipt}\n\n"
        f"<b>Total Paid: ${total_value:.2f}</b>\n\n"
        f"Shelf stock deducted and registered in database."
    )
    if TELEGRAM_CUSTOMER_CHAT_ID:
        agent.telegram.send_message(TELEGRAM_CUSTOMER_CHAT_ID, receipt_msg)

    if agent.sheets:
        agent.sheets.append_log({
            "log_id": f"LOG-{uuid.uuid4().hex[:8].upper()}",
            "timestamp": timestamp,
            "action": "order_approved",
            "order_id": order_id,
            "details": f"POS cart checkout fulfilled (${total_value:.2f})",
            "customer_notified": "yes",
            "owner_notified": "no",
        })

    return jsonify({
        "status": "approved",
        "order_id": order_id,
        "total": round(total_value, 2),
        "items_count": len(items)
    })


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
