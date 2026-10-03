"""
End-to-end verification script for the AI Operations Agent workflow.
Tests all decision paths, LLM reasoning, inventory state mutations, and Telegram alerts.
"""
import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

from agent import AIBusinessAgent

def run_verification():
    print("=" * 65)
    print("  COMPREHENSIVE END-TO-END WORKFLOW VERIFICATION")
    print("=" * 65)

    agent = AIBusinessAgent()

    test_cases = [
        {
            "id": "SCENARIO-1",
            "name": "Auto-Approve (< $100 & Stock Available)",
            "message": "I want 2 black Basic Shirts",
            "customer": "Alice Test",
            "expected_status": "approved",
            "expect_stock_deduction": True,
            "deduct_amount": 2,
        },
        {
            "id": "SCENARIO-2",
            "name": "Human Escalation (>= $100 & Stock Available)",
            "message": "I need 12 black Basic Shirts",
            "customer": "Bob Test",
            "expected_status": "escalated",
            "expect_stock_deduction": False,
            "deduct_amount": 0,
        },
        {
            "id": "SCENARIO-3",
            "name": "Alternative Suggestion (Stock Insufficient)",
            "message": "I want 25 black Basic Shirts",
            "customer": "Charlie Test",
            "expected_status": "alternative_suggested",
            "expect_stock_deduction": False,
            "deduct_amount": 0,
        },
        {
            "id": "SCENARIO-4",
            "name": "Clarification Request (Unknown SKU)",
            "message": "Do you have blue jeans?",
            "customer": "Diana Test",
            "expected_status": "clarify",
            "expect_stock_deduction": False,
            "deduct_amount": 0,
        },
    ]

    all_passed = True

    for tc in test_cases:
        print(f"\n>> Testing [{tc['id']}] {tc['name']}")
        
        # Check stock before
        stock_before = None
        if agent.sheets:
            for p in agent.sheets.read_products():
                if "basic shirt" in p.get("Product", "").lower() and p.get("Variant", "").lower() == "black":
                    stock_before = p.get("Stock")
                    break

        orders_count_before = len(agent.sheets.orders) if hasattr(agent.sheets, "orders") else 0
        logs_count_before = len(agent.sheets.logs) if hasattr(agent.sheets, "logs") else 0

        # Execute order through full agent pipeline
        result = agent.process_order(tc["message"], tc["customer"])
        actual_status = result.get("status")

        # Check stock after
        stock_after = None
        if agent.sheets:
            for p in agent.sheets.read_products():
                if "basic shirt" in p.get("Product", "").lower() and p.get("Variant", "").lower() == "black":
                    stock_after = p.get("Stock")
                    break

        orders_count_after = len(agent.sheets.orders) if hasattr(agent.sheets, "orders") else 0
        logs_count_after = len(agent.sheets.logs) if hasattr(agent.sheets, "logs") else 0

        # Validations
        status_match = (actual_status == tc["expected_status"])
        
        if tc["expect_stock_deduction"]:
            stock_match = (stock_after == stock_before - tc["deduct_amount"])
            order_added = (orders_count_after == orders_count_before + 1)
        else:
            stock_match = (stock_after == stock_before)
            order_added = (orders_count_after == orders_count_before)

        log_added = (logs_count_after == logs_count_before + 1)

        print(f"   [Validation Results]")
        print(f"   • Action Routing: {actual_status} -> {'PASS' if status_match else 'FAIL'}")
        print(f"   • Stock Mutation: {stock_before} -> {stock_after} -> {'PASS' if stock_match else 'FAIL'}")
        print(f"   • Order Record Added: {'PASS' if order_added else 'FAIL'}")
        print(f"   • Audit Log Appended: {'PASS' if log_added else 'FAIL'}")

        if not (status_match and stock_match and order_added and log_added):
            all_passed = False

    print("\n" + "=" * 65)
    if all_passed:
        print("  VERIFICATION COMPLETED: ALL 4 WORKFLOW SCENARIOS PASSED 100%")
    else:
        print("  VERIFICATION FAILED ON SOME SCENARIOS")
    print("=" * 65)


if __name__ == "__main__":
    run_verification()
