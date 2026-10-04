"""
Script to capture implementation screenshots and diagrams for the assignment report.
Uses Playwright to capture Web UI execution and Matplotlib to generate high-res diagrams.
"""
import os
import sys
import time
import threading

# Add project root to path
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Ensure screenshots dir exists
output_dir = os.path.join(PROJECT_ROOT, "screenshots")
os.makedirs(output_dir, exist_ok=True)


def generate_workflow_diagram():
    """Generate high-resolution visual workflow diagram with accurate multi-path execution."""
    fig, ax = plt.subplots(figsize=(12, 10), dpi=200)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 13)
    ax.axis('off')

    # Colors
    c_blue = '#2563eb'
    c_green = '#059669'
    c_amber = '#d97706'
    c_purple = '#7c3aed'
    c_dark = '#0f172a'
    c_bg = '#f8fafc'

    # Title
    ax.text(6, 12.5, "AI Small Business Operations Agent - Workflow Architecture", 
            ha='center', va='center', fontsize=14, fontweight='bold', color=c_dark)
    ax.text(6, 12.1, "Autonomous Order Processing, Multi-Path AI Reasoning & Branch-Specific Execution", 
            ha='center', va='center', fontsize=10, color='#64748b')

    # Top pipeline steps
    steps = [
        (6, 11.2, "Step 1: Order Intake", "Receives customer request via Web POS Terminal, Webhook, or CLI", c_blue),
        (6, 10.1, "Step 2: AI Intent Extraction (Decision Point 1)", "DeepSeek LLM parses product, variant, quantity, and intent into structured JSON", c_purple),
        (6, 9.0, "Step 3: Read Live Inventory", "Queries Google Sheets / Persistent Excel (inventory.xlsx) for real-time stock", c_blue),
        (6, 7.9, "Step 4: AI Fulfillment & Action Reasoning (Decision Points 2 & 3)", "DeepSeek LLM evaluates stock levels, pricing thresholds ($100 rule), and product validity", c_purple),
    ]

    for x, y, title, subtitle, col in steps:
        box = patches.FancyBboxPatch((x-4.5, y-0.42), 9.0, 0.84, boxstyle="round,pad=0.1",
                                     fc=c_bg, ec=col, lw=2)
        ax.add_patch(box)
        ax.text(x, y+0.12, title, ha='center', va='center', fontsize=10, fontweight='bold', color=col)
        ax.text(x, y-0.18, subtitle, ha='center', va='center', fontsize=8, color='#475569')

    # Arrows between top steps
    for y in [10.7, 9.6, 8.5]:
        ax.annotate('', xy=(6, y-0.15), xytext=(6, y+0.35),
                    arrowprops=dict(arrowstyle="->", lw=2, color='#64748b'))

    # Step 5 Header: 5-Way Autonomous Router
    ax.annotate('', xy=(6, 6.95), xytext=(6, 7.45),
                arrowprops=dict(arrowstyle="->", lw=2, color='#64748b'))

    router_box = patches.FancyBboxPatch((1.5, 6.4), 9.0, 0.55, boxstyle="round,pad=0.08",
                                        fc='#e0e7ff', ec='#4338ca', lw=2)
    ax.add_patch(router_box)
    ax.text(6, 6.67, "Step 5: Autonomous Workflow Routing (5 Mutually Exclusive Execution Paths)", 
            ha='center', va='center', fontsize=10, fontweight='bold', color='#312e81')

    # 5 Branch columns
    branch_data = [
        (1.4, "APPROVE", "Auto-fulfill\nStock OK\nValue < $100", c_green, [
            "• Deduct inventory",
            "• Record in Orders tab",
            "• Customer Telegram alert",
            "• Append to Logs tab",
        ]),
        (3.7, "ESCALATE", "Owner Review\nValue ≥ $100\nThreshold triggered", c_amber, [
            "• Stock UNTOUCHED",
            "• Owner Telegram alert",
            "• Order marked PENDING",
            "• Append to Logs tab",
        ]),
        (6.0, "SUGGEST ALT", "Stock Deficit\nAlternative items\navailable", c_purple, [
            "• Stock UNTOUCHED",
            "• Query substitute items",
            "• Proposal via Telegram",
            "• Append to Logs tab",
        ]),
        (8.3, "CLARIFY", "Ambiguous order\nMissing variant\nUnknown SKU", c_blue, [
            "• Stock UNTOUCHED",
            "• Prompt clarification",
            "• Await customer input",
            "• Append to Logs tab",
        ]),
        (10.6, "REJECT", "Cannot Fulfill\nZero inventory\nNo substitutes", '#dc2626', [
            "• Stock UNTOUCHED",
            "• Send decline notice",
            "• Polite explanation",
            "• Append to Logs tab",
        ]),
    ]

    for x_center, b_title, b_condition, col, actions in branch_data:
        # Branch header card
        header_box = patches.FancyBboxPatch((x_center-1.05, 5.0), 2.1, 1.1, boxstyle="round,pad=0.08",
                                            fc='#ffffff', ec=col, lw=2)
        ax.add_patch(header_box)
        ax.text(x_center, 5.75, b_title, ha='center', va='center', fontsize=9, fontweight='bold', color=col)
        ax.text(x_center, 5.3, b_condition, ha='center', va='center', fontsize=7, color='#334155', linespacing=1.15)

        # Arrow from router to branch header
        ax.annotate('', xy=(x_center, 6.15), xytext=(6, 6.4),
                    arrowprops=dict(arrowstyle="->", lw=1.5, color=col))

        # Branch execution details card
        actions_box = patches.FancyBboxPatch((x_center-1.05, 2.3), 2.1, 2.45, boxstyle="round,pad=0.08",
                                             fc='#ffffff', ec=col, lw=1.4, ls='--')
        ax.add_patch(actions_box)
        ax.text(x_center, 4.45, "Execution Actions:", ha='center', va='center', fontsize=7.5, fontweight='bold', color=col)
        
        y_act = 4.0
        for act in actions:
            ax.text(x_center-0.95, y_act, act, ha='left', va='center', fontsize=6.8, color='#1e293b')
            y_act -= 0.45

        # Arrow from branch header down to execution card
        ax.annotate('', xy=(x_center, 4.8), xytext=(x_center, 5.0),
                    arrowprops=dict(arrowstyle="->", lw=1.5, color=col))

        # Arrow from execution card down to completion node
        ax.annotate('', xy=(6, 1.25), xytext=(x_center, 2.25),
                    arrowprops=dict(arrowstyle="->", lw=1.2, color='#94a3b8'))

    # Final Step: Unified Return
    final_box = patches.FancyBboxPatch((1.5, 0.4), 9.0, 0.8, boxstyle="round,pad=0.1",
                                       fc='#0f172a', ec='#334155', lw=1.5)
    ax.add_patch(final_box)
    ax.text(6, 0.95, "Step 10: Deliver Structured Response to Calling Interface", 
            ha='center', va='center', fontsize=9.5, fontweight='bold', color='#ffffff')
    ax.text(6, 0.62, "Formats receipt / status notice, returns JSON to POS Web Terminal, Webhook, or Telegram", 
            ha='center', va='center', fontsize=7.8, color='#94a3b8')

    plt.tight_layout()
    diagram_path = os.path.join(output_dir, "00_workflow_diagram.png")
    plt.savefig(diagram_path, bbox_inches='tight')
    plt.close()
    print(f"Generated workflow diagram: {diagram_path}")


def generate_architecture_diagram():
    """Generate high-resolution visual system architecture diagram."""
    fig, ax = plt.subplots(figsize=(10, 6.5), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 8)
    ax.axis('off')

    c_blue = '#1d4ed8'
    c_dark = '#0f172a'
    c_card = '#f8fafc'

    # Title
    ax.text(5, 7.6, "AI Small Business Operations Agent - System Architecture", 
            ha='center', va='center', fontsize=14, fontweight='bold', color=c_dark)
    ax.text(5, 7.2, "Python Agent Core, DeepSeek LLM, Google Sheets & Telegram Bot Integration", 
            ha='center', va='center', fontsize=10, color='#64748b')

    # Input Clients Layer
    ax.text(1.8, 6.4, "Input Channels", fontsize=11, fontweight='bold', color=c_blue)
    inputs = [
        (1.8, 5.5, "Customer Chat / Web UI\n(Flask App / Browser)"),
        (1.8, 4.3, "Webhook Trigger\n(n8n REST API Endpoint)"),
        (1.8, 3.1, "CLI Terminal\n(Interactive Operator Mode)"),
    ]
    for x, y, text in inputs:
        box = patches.FancyBboxPatch((x-1.5, y-0.45), 3.0, 0.9, boxstyle="round,pad=0.08",
                                     fc='#eff6ff', ec=c_blue, lw=1.5)
        ax.add_patch(box)
        ax.text(x, y, text, ha='center', va='center', fontsize=8, fontweight='bold', color='#1e3a8a')

    # Central Agent Core Layer
    ax.text(5.0, 6.4, "AI Agent Orchestration Core", fontsize=11, fontweight='bold', color='#7c3aed')
    core_box = patches.FancyBboxPatch((3.7, 2.4), 2.6, 3.8, boxstyle="round,pad=0.1",
                                     fc='#faf5ff', ec='#9333ea', lw=2)
    ax.add_patch(core_box)
    ax.text(5.0, 5.8, "AIBusinessAgent\n(agent.py)", ha='center', va='center', fontsize=9.5, fontweight='bold', color='#581c87')
    ax.text(5.0, 5.0, "• 10-Step Pipeline\n• 3 AI Decision Points\n• Business Rule Engine\n• 5-Way Autonomous Router",
            ha='center', va='center', fontsize=7.5, color='#3b0764')
    ax.text(5.0, 3.8, "Resilient Data Layer\n(Google Sheets Client\n+ In-Memory Mock Fallback)",
            ha='center', va='center', fontsize=7.5, fontweight='bold', color='#6b21a8')
    ax.text(5.0, 2.8, "Telegram Client\n(Async Notifications)",
            ha='center', va='center', fontsize=7.5, fontweight='bold', color='#6b21a8')

    # External Services Layer
    ax.text(8.2, 6.4, "External Services (4 Tools)", fontsize=11, fontweight='bold', color='#059669')
    externals = [
        (8.2, 5.5, "DeepSeek LLM API\n(deepseek-chat V3 Model)\nReasoning & Decision Engine", '#2563eb'),
        (8.2, 4.2, "Google Sheets API\n(gspread / OAuth2 / Service Account)\nProducts, Orders, Customers, Logs", '#059669'),
        (8.2, 2.9, "Telegram Bot API\n(@business0psBot)\nInstant Alerts to Customer & Owner", '#0284c7'),
    ]
    for x, y, text, col in externals:
        box = patches.FancyBboxPatch((x-1.5, y-0.45), 3.0, 0.9, boxstyle="round,pad=0.08",
                                     fc='#f0fdf4', ec=col, lw=1.5)
        ax.add_patch(box)
        ax.text(x, y, text, ha='center', va='center', fontsize=7.5, fontweight='bold', color=col)

    # Connecting Arrows
    for _, y, _ in inputs:
        ax.annotate('', xy=(3.6, y), xytext=(3.3, y),
                    arrowprops=dict(arrowstyle="->", lw=1.8, color='#94a3b8'))

    for _, y, _, _ in externals:
        ax.annotate('', xy=(6.7, y), xytext=(6.35, y),
                    arrowprops=dict(arrowstyle="<->", lw=1.8, color='#94a3b8'))

    plt.tight_layout()
    arch_path = os.path.join(output_dir, "00_architecture_diagram.png")
    plt.savefig(arch_path, bbox_inches='tight')
    plt.close()
    print(f"Generated architecture diagram: {arch_path}")


def capture_web_ui():
    """Capture UI states with Playwright (connects to existing server or launches one)."""
    import urllib.request
    from playwright.sync_api import sync_playwright

    server_running = False
    try:
        urllib.request.urlopen("http://127.0.0.1:5000/api/health", timeout=1)
        server_running = True
        print("Connected to existing web server on port 5000")
    except Exception:
        pass

    if not server_running:
        from web_app import app
        server = threading.Thread(target=lambda: app.run(port=5000, debug=False, use_reloader=False))
        server.daemon = True
        server.start()
        time.sleep(1.5)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1280, "height": 720})
        page.goto("http://127.0.0.1:5000")
        page.wait_for_selector(".product-card", timeout=10000)
        time.sleep(1)

        # 1. Main POS Terminal Interface (Catalog & Empty Cart)
        p1 = os.path.join(output_dir, "01_web_ui_dashboard.png")
        page.screenshot(path=p1)
        print(f"Captured: {p1}")

        # 2. Scenario 1 - Auto-approve
        page.click(".preset-chip.approve")
        page.wait_for_selector("dialog#decision-dialog[open]", timeout=15000)
        time.sleep(1)
        p2 = os.path.join(output_dir, "02_scenario_auto_approve.png")
        page.screenshot(path=p2)
        print(f"Captured: {p2}")
        page.click(".modal-close-btn")
        time.sleep(0.5)

        # 3. Scenario 2 - Escalate to Owner
        page.click(".preset-chip.escalate")
        page.wait_for_selector("dialog#decision-dialog[open]", timeout=15000)
        time.sleep(1)
        p3 = os.path.join(output_dir, "03_scenario_escalate.png")
        page.screenshot(path=p3)
        print(f"Captured: {p3}")
        page.click(".modal-close-btn")
        time.sleep(0.5)

        # 4. Scenario 3 - Suggest Alternative
        page.click(".preset-chip.suggest")
        page.wait_for_selector("dialog#decision-dialog[open]", timeout=15000)
        time.sleep(1)
        p4 = os.path.join(output_dir, "04_scenario_alternative.png")
        page.screenshot(path=p4)
        print(f"Captured: {p4}")
        page.click(".modal-close-btn")
        time.sleep(0.5)

        # 5. Scenario 4 - Clarification
        page.click(".preset-chip.clarify")
        page.wait_for_selector("dialog#decision-dialog[open]", timeout=15000)
        time.sleep(1)
        p5 = os.path.join(output_dir, "05_scenario_clarify.png")
        page.screenshot(path=p5)
        print(f"Captured: {p5}")
        page.click(".modal-close-btn")
        time.sleep(0.5)

        browser.close()


if __name__ == "__main__":
    generate_workflow_diagram()
    generate_architecture_diagram()
    try:
        capture_web_ui()
    except Exception as e:
        print(f"Web UI capture error: {e}")
