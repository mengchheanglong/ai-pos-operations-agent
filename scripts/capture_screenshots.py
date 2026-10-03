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
    """Generate high-resolution visual workflow diagram."""
    fig, ax = plt.subplots(figsize=(10, 8), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 12)
    ax.axis('off')

    # Colors
    c_blue = '#2563eb'
    c_green = '#059669'
    c_amber = '#d97706'
    c_purple = '#7c3aed'
    c_dark = '#1e293b'
    c_bg = '#f8fafc'

    # Title
    ax.text(5, 11.5, "AI Small Business Operations Agent - Workflow Diagram", 
            ha='center', va='center', fontsize=14, fontweight='bold', color=c_dark)
    ax.text(5, 11.1, "10-Step Autonomous Order Fulfillment & Multi-Path Routing", 
            ha='center', va='center', fontsize=10, color='#64748b')

    # Nodes
    steps = [
        (5, 10.2, "1. Customer Order Request", "Webhook / Web Interface / CLI entry point", c_blue),
        (5, 9.0, "2. AI Intent Extraction (Decision 1)", "DeepSeek LLM extracts product, variant, quantity, intent", c_purple),
        (5, 7.8, "3. Read Inventory", "Google Sheets API / In-Memory Mock Inventory", c_blue),
        (5, 6.6, "4. AI Fulfillment & Action Decision (Decisions 2 & 3)", "DeepSeek LLM applies business rules based on stock & value", c_purple),
    ]

    for x, y, title, subtitle, col in steps:
        box = patches.FancyBboxPatch((x-3.2, y-0.45), 6.4, 0.9, boxstyle="round,pad=0.1",
                                     fc=c_bg, ec=col, lw=2)
        ax.add_patch(box)
        ax.text(x, y+0.12, title, ha='center', va='center', fontsize=10, fontweight='bold', color=col)
        ax.text(x, y-0.18, subtitle, ha='center', va='center', fontsize=7.5, color='#475569')

    # Arrows for first 4 steps
    for y in [9.7, 8.5, 7.3]:
        ax.annotate('', xy=(5, y-0.2), xytext=(5, y+0.3),
                    arrowprops=dict(arrowstyle="->", lw=2, color='#64748b'))

    # Step 5: Autonomous 5-Way Branching
    ax.annotate('', xy=(5, 5.6), xytext=(5, 6.1),
                arrowprops=dict(arrowstyle="->", lw=2, color='#64748b'))

    ax.text(5, 5.4, "5. Autonomous Workflow Routing (5 Possible Paths)", 
            ha='center', va='center', fontsize=10, fontweight='bold', color=c_dark)

    branches = [
        (1.2, 4.2, "Approve", "Auto-fulfill\nValue < $100\nStock OK", c_green),
        (3.1, 4.2, "Escalate", "Owner Review\nValue >= $100\nOwner Alert", c_amber),
        (5.0, 4.2, "Suggest Alt", "Out of Stock\nOffer Sub.\nVariants", c_purple),
        (6.9, 4.2, "Clarify", "Incomplete Info\nor Unknown SKU", c_blue),
        (8.8, 4.2, "Reject", "Cannot Fulfill\nNo Stock/Alt\nAvailable", '#dc2626'),
    ]

    for x, y, title, desc, col in branches:
        # Branch box
        box = patches.FancyBboxPatch((x-0.85, y-0.6), 1.7, 1.2, boxstyle="round,pad=0.08",
                                     fc='#ffffff', ec=col, lw=1.8)
        ax.add_patch(box)
        ax.text(x, y+0.3, title, ha='center', va='center', fontsize=8.5, fontweight='bold', color=col)
        ax.text(x, y-0.15, desc, ha='center', va='center', fontsize=6.8, color='#334155', linespacing=1.2)

        # Connector arrow from routing
        ax.annotate('', xy=(x, y+0.65), xytext=(5, 5.2),
                    arrowprops=dict(arrowstyle="->", lw=1.2, color=col))

    # Downstream execution steps (Steps 6 to 10)
    exec_steps = [
        (5, 2.2, "6. Execute Action (Update Stock in Google Sheets)"),
        (5, 1.5, "7. Create Order Record (Orders Sheet) & Notify Customer (Telegram)"),
        (5, 0.8, "8. Notify Owner if Escalated (Telegram) | 9. Append Audit Log (Logs Sheet)"),
        (5, 0.1, "10. Return Final Structured Result to Caller"),
    ]

    for x, y, text in exec_steps:
        box = patches.FancyBboxPatch((x-4.2, y-0.25), 8.4, 0.5, boxstyle="round,pad=0.06",
                                     fc='#f1f5f9', ec='#cbd5e1', lw=1)
        ax.add_patch(box)
        ax.text(x, y, text, ha='center', va='center', fontsize=8, fontweight='bold', color=c_dark)

    # Connector down
    ax.annotate('', xy=(5, 2.5), xytext=(5, 3.4),
                arrowprops=dict(arrowstyle="->", lw=2, color='#64748b'))

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
    """Run Flask server in thread and capture UI states with Playwright."""
    from playwright.sync_api import sync_playwright
    from web_app import app

    server = threading.Thread(target=lambda: app.run(port=5000, debug=False, use_reloader=False))
    server.daemon = True
    server.start()
    time.sleep(1.5)  # Wait for server

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1000, "height": 800})
        page.goto("http://127.0.0.1:5000")
        page.wait_for_timeout(1000)

        # 1. Main Dashboard
        p1 = os.path.join(output_dir, "01_web_ui_dashboard.png")
        page.screenshot(path=p1)
        print(f"Captured: {p1}")

        # 2. Scenario 1 - Auto-approve
        page.fill("#message", "I want 2 black Basic Shirts")
        page.click("#submit-btn")
        page.wait_for_selector("#result.show", timeout=15000)
        time.sleep(1)
        p2 = os.path.join(output_dir, "02_scenario_auto_approve.png")
        page.screenshot(path=p2)
        print(f"Captured: {p2}")

        # 3. Scenario 2 - Escalate to Owner
        page.fill("#message", "I need 12 black Basic Shirts")
        page.click("#submit-btn")
        page.wait_for_selector("#result.show", timeout=15000)
        time.sleep(1)
        p3 = os.path.join(output_dir, "03_scenario_escalate.png")
        page.screenshot(path=p3)
        print(f"Captured: {p3}")

        # 4. Scenario 3 - Suggest Alternative
        page.fill("#message", "I want 25 black Basic Shirts")
        page.click("#submit-btn")
        page.wait_for_selector("#result.show", timeout=15000)
        time.sleep(1)
        p4 = os.path.join(output_dir, "04_scenario_alternative.png")
        page.screenshot(path=p4)
        print(f"Captured: {p4}")

        # 5. Scenario 4 - Clarification
        page.fill("#message", "Do you have blue jeans?")
        page.click("#submit-btn")
        page.wait_for_selector("#result.show", timeout=15000)
        time.sleep(1)
        p5 = os.path.join(output_dir, "05_scenario_clarify.png")
        page.screenshot(path=p5)
        print(f"Captured: {p5}")

        browser.close()


if __name__ == "__main__":
    generate_workflow_diagram()
    generate_architecture_diagram()
    try:
        capture_web_ui()
    except Exception as e:
        print(f"Web UI capture error: {e}")
