import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------
# 1. Generate realistic sample Tally Daybook Excel (.xlsx)
# ---------------------------------------------------------
wb = openpyxl.Workbook()

# Sheet 1: Tally_Daybook_FY2526
ws = wb.active
ws.title = "Tally_Daybook_FY2526"

headers = [
    "Date", "Voucher_No", "Voucher_Type", "Ledger_Head", "Party_Name",
    "Party_GSTIN", "Payment_Mode", "Base_Amount", "CGST", "SGST",
    "IGST", "TDS_Section", "TDS_Deducted", "Net_Payable", "Supporting_Doc", "Narration"
]

sample_rows = [
    ["2025-05-12", "PUR-001", "Purchase", "Computer & Peripherals", "Infiniti Retail Ltd (Croma)", "27AABCI1234D1Z2", "NEFT", 95000, 0, 0, 0, "NA", 0, 112100, "Tax Invoice (Billed to Rajesh Sharma - Director)", "MacBook Air purchased; ITC claimed in GSTR-3B"],
    ["2025-06-18", "CSH-014", "Payment", "Repairs & Maintenance", "Ramesh Sutar Carpentry", "URD", "CASH", 14500, 0, 0, 0, "NA", 0, 14500, "Handwritten Cash Receipt", "Single day cash payment for cabin partition repair"],
    ["2025-08-04", "PUR-042", "Purchase", "Office Supplies", "Shreeji Enterprises (Andheri)", "27AAECS9988K1Z1", "NEFT", 20000, 0, 0, 3600, "NA", 0, 23600, "Tax Invoice #SE/25-26/108", "Intra-state Mumbai supply wrongly billed with 18% IGST"],
    ["2025-09-15", "PUR-067", "Journal", "Postage & Courier", "Blue Dart Express Ltd", "27AAACB0446L1ZS", "NEFT", 2000, 180, 180, 0, "NA", 0, 2360, "Tax Invoice + Signed Dispatch Log", "Monthly client document courier charges"],
    ["2025-10-22", "CSH-089", "Payment", "Printing & Stationery", "Petty Cash - Staff Slip", "URD", "CASH", 8400, 0, 0, 0, "NA", 0, 8400, "Internal Petty Cash Slip Only", "Letterhead printing; vendor bill misplaced by office boy"],
    ["2025-11-30", "JRN-112", "Journal", "Legal & Professional Fees", "M/s Mehta & Co. Advocates", "27AAFFM4321P1Z9", "NEFT", 65000, 0, 0, 0, "194J", 0, 65000, "Invoice #MC/25-26/044", "Annual ROC & contract advisory retainer credited without TDS"],
    ["2025-12-31", "JRN-145", "Journal", "Office Rent - Nariman Point", "Nariman Point Properties LLP", "27AABFN7766R1Z4", "NEFT", 65000, 5850, 5850, 0, "194I", 0, 76700, "Rent Invoice Dec 2025", "Monthly office rent > Rs.50,000/month without 194I TDS"],
    ["2026-01-31", "JRN-160", "Journal", "Godown Rent - Bhiwandi", "Shree Logistics Warehousing", "URD", "NEFT", 35000, 0, 0, 0, "194I", 0, 35000, "Rent Agreement + Receipt", "Monthly godown rent Rs.35,000 (Compliant under FA 2025 >50k/mo limit)"],
    ["2026-02-14", "JRN-188", "Journal", "AMC & Repairs Contract", "CoolAir HVAC Services Pvt Ltd", "27AAACC8877H1Z3", "CHEQUE", 42000, 3780, 3780, 0, "194C", 0, 49560, "Tax Invoice #CA/882", "Single HVAC AMC contract bill > Rs.30,000 without 2% TDS u/s 194C"]
]

ws.append(headers)
for r in sample_rows:
    ws.append(r)

# Professional Styling for Excel
header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
warn_fill = PatternFill(start_color="FFF7ED", end_color="FFF7ED", fill_type="solid")
thin_border = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1")
)

for col_idx, cell in enumerate(ws[1], 1):
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

for row_idx in range(2, ws.max_row + 1):
    is_even = (row_idx % 2 == 0)
    for col_idx in range(1, ws.max_column + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        c.border = thin_border
        c.font = Font(name="Calibri", size=10)
        if is_even:
            c.fill = zebra_fill
        if col_idx in [8, 9, 10, 11, 13, 14]:
            c.number_format = '#,##0'
            c.alignment = Alignment(horizontal="right")

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

col_widths = [12, 12, 13, 26, 28, 18, 13, 14, 11, 11, 11, 12, 14, 14, 34, 45]
for idx, width in enumerate(col_widths, 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = width

# Sheet 2: FY25_26_Statutory_Rules_Ref
ws_rules = wb.create_sheet("FY25_26_Audit_Rules_Ref")
ws_rules.append(["Statutory Area", "Section / Standard", "FY 2025-26 Threshold / Rule", "Audit Impact / Disallowance"])
rules_data = [
    ["Cash Expenditure", "Sec 40A(3) IT Act", "Exceeding Rs. 10,000 in a single day to single person (Rs. 35,000 for GTA)", "100% disallowance of expense in business income computation"],
    ["TDS - Professional / Technical", "Sec 194J (FA 2025)", "Aggregate > Rs. 50,000 p.a. (10% Professional / 2% Technical)", "30% expense disallowance u/s 40(a)(ia) + Interest u/s 201(1A)"],
    ["TDS - Rent", "Sec 194I (FA 2025)", "Exceeding Rs. 50,000 per month or part of a month (@ 10% Land/Bldg, 2% P&M)", "30% expense disallowance u/s 40(a)(ia); replaces old Rs. 2.40L annual limit"],
    ["TDS - Contractors", "Sec 194C", "Single payment > Rs. 30,000 OR Aggregate > Rs. 1,00,000 p.a. (1% Ind/HUF, 2% Others)", "30% expense disallowance u/s 40(a)(ia)"],
    ["GST Input Tax Credit", "Sec 16(2) CGST Act", "Valid Tax Invoice with Company GSTIN, receipt of goods, GSTR-2B reflection", "100% ITC reversal + Interest u/s 50 if billed in personal name"],
    ["Place of Supply (GST)", "Sec 8 & 12 IGST Act", "Same State Code (e.g. 27 -> 27 Maharashtra) requires CGST + SGST", "Wrong tax head (IGST claimed instead of CGST+SGST) requires refund/re-payment"],
    ["Audit Evidence", "SA 500", "External third-party invoice carries higher reliability than internal cash slip", "Qualification / Management Representation / Vouching exception"]
]
for r in rules_data:
    ws_rules.append(r)

for cell in ws_rules[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.border = thin_border

for row_idx in range(2, ws_rules.max_row + 1):
    for col_idx in range(1, 5):
        c = ws_rules.cell(row=row_idx, column=col_idx)
        c.border = thin_border
        c.alignment = Alignment(wrap_text=True, vertical="top")

for idx, w in enumerate([24, 22, 46, 48], 1):
    ws_rules.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w

excel_filename = "apex_tally_daybook_fy25_26.xlsx"
wb.save(excel_filename)

# ---------------------------------------------------------
# 2. Generate Complete Python (NiceGUI + FastAPI + Pandas) App
# ---------------------------------------------------------
python_app_code = '''"""
The First-Year Audit Trail: Python Edition (NiceGUI + FastAPI + Pandas)
=======================================================================
A full-stack pure-Python Statutory Audit Training Simulator for CA Article Trainees.

Key Capabilities Unlocked over Static HTML:
1. Tally Excel (.xlsx / .csv) Ingestion via Pandas (`ingest_tally_daybook`)
2. Automated Statutory Rule & Exception Scanner (FY 2025-26 Finance Act 2025 compliant)
3. Randomized Exception Mutator (`generate_randomized_batch`) so every trainee gets unique numbers
4. Tactile NiceGUI Web Interface (Tailwind CSS + Vue/Quasar under the hood, 100% Python)
5. Exportable Excel Working Paper Memo (`export_audit_observations_xlsx`) and REST API endpoints

Quick Start:
------------
1. Install dependencies:
   pip install nicegui pandas openpyxl

2. Run the simulator:
   python audit_trail_nicegui_app.py

3. Open http://localhost:8080 in your browser (FastAPI Swagger docs also live at http://localhost:8080/docs)
"""

from __future__ import annotations
import io
import random
from dataclasses import dataclass, asdict
from datetime import date, timedelta
from typing import List, Dict, Optional

import pandas as pd
from fastapi import UploadFile, File
from nicegui import ui, app

# ==============================================================================
# 1. STATUTORY AUDIT ENGINE (FY 2025-26 / AY 2026-27 RULES)
# ==============================================================================

COMPANY_NAME = "Apex Trading Pvt Ltd"
COMPANY_GSTIN = "27AABCA1234F1Z5"
COMPANY_STATE_CODE = "27"  # Maharashtra


@dataclass
class AuditVoucher:
    voucher_id: str
    voucher_date: str
    ledger_head: str
    party_name: str
    party_gstin: str
    payment_mode: str
    base_amount: float
    cgst: float
    sgst: float
    igst: float
    tds_section: str
    tds_deducted: float
    supporting_doc: str
    narration: str
    expected_action: str       # "VERIFY" or "QUERY"
    expected_rule_code: str    # e.g., "SEC_40A3", "SEC_16_2", "INTRA_IGST", "SEC_194J", "SEC_194I", "SEC_194C", "SA_500", "CLEAN"
    senior_hint: str
    statutory_explanation: str


RULE_OPTIONS = {
    "CLEAN": "Clean & Compliant — Valid supporting, GST & TDS compliance verified",
    "SEC_16_2": "Sec 16(2) CGST Act — Invoice in Director's personal name / Company GSTIN missing",
    "SEC_40A3": "Sec 40A(3) IT Act — Cash expenditure > ₹10,000 in a single day (100% Disallowance)",
    "INTRA_IGST": "Sec 8/12 IGST Act — Intra-state supply (27→27) wrongly charged IGST instead of CGST+SGST",
    "SA_500": "SA 500 Audit Evidence — Missing external third-party invoice (internal slip only)",
    "SEC_194J": "Sec 194J Default — Professional fees > ₹50,000 p.a. without 10% TDS (30% disallowance u/s 40(a)(ia))",
    "SEC_194I": "Sec 194I Default — Monthly Rent > ₹50,000/month (FA 2025 limit) without 10% TDS",
    "SEC_194C": "Sec 194C Default — Single contract bill > ₹30,000 without TDS (1%/2%)",
}


def format_inr(amount: float) -> str:
    """Format number in Indian Rupee numbering system (e.g., ₹1,45,000)."""
    amt = int(round(amount))
    s = str(abs(amt))
    if len(s) > 3:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.append(rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.append(rest)
        formatted = ",".join(reversed(parts)) + "," + last3
    else:
        formatted = s
    prefix = "-₹" if amt < 0 else "₹"
    return f"{prefix}{formatted}"


def auto_classify_tally_row(row: pd.Series) -> AuditVoucher:
    """
    Inspects a single row from a Tally Daybook Excel/CSV export and automatically
    determines whether it violates Indian Statutory Audit rules for FY 2025-26.
    """
    v_id = str(row.get("Voucher_No", "V-000"))
    v_date = str(row.get("Date", "2025-10-01"))[:10]
    ledger = str(row.get("Ledger_Head", ""))
    party = str(row.get("Party_Name", ""))
    gstin = str(row.get("Party_GSTIN", "URD")).strip().upper()
    mode = str(row.get("Payment_Mode", "NEFT")).strip().upper()
    base = float(row.get("Base_Amount", 0) or 0)
    cgst = float(row.get("CGST", 0) or 0)
    sgst = float(row.get("SGST", 0) or 0)
    igst = float(row.get("IGST", 0) or 0)
    tds_sec = str(row.get("TDS_Section", "NA")).strip().upper()
    tds_amt = float(row.get("TDS_Deducted", 0) or 0)
    doc = str(row.get("Supporting_Doc", "Tax Invoice"))
    narration = str(row.get("Narration", ""))

    # Default clean state
    action = "VERIFY"
    code = "CLEAN"
    hint = "Verify Buyer GSTIN, tax head split, payment mode, TDS threshold (FA 2025), and external supporting."
    explanation = "Compliant transaction. External supporting document and statutory compliances are in order."

    # Rule 1: Personal Name / B2C ITC Claimed (Sec 16(2))
    if "director" in doc.lower() or "personal" in doc.lower() or "rajesh sharma" in doc.lower():
        action, code = "QUERY", "SEC_16_2"
        hint = "Check whose name and GSTIN appear in the 'Billed To' block before allowing Input Tax Credit."
        explanation = "Violation of Sec 16(2) CGST Act: Invoice is in Director's personal name without Company GSTIN. ITC must be reversed."

    # Rule 2: Missing external third-party invoice (SA 500)
    elif "internal" in doc.lower() or "petty cash slip only" in doc.lower() or "misplaced" in narration.lower():
        action, code = "QUERY", "SA_500"
        hint = "Does an internal slip signed by staff constitute sufficient appropriate external audit evidence?"
        explanation = "Insufficient Audit Evidence (SA 500): Expense exceeding petty norms lacks external vendor bill."

    # Rule 3: Cash Payment > Rs. 10,000 in a single day (Sec 40A(3))
    elif mode == "CASH" and base > 10000:
        action, code = "QUERY", "SEC_40A3"
        hint = "Check the payment mode and compare the amount against Section 40A(3) single-day cash limits."
        explanation = f"Sec 40A(3) Violation: Cash payment of {format_inr(base)} exceeds ₹10,000 in a single day. 100% disallowed in tax computation."

    # Rule 4: Intra-state supply charging IGST instead of CGST + SGST
    elif gstin.startswith(COMPANY_STATE_CODE) and igst > 0 and (cgst == 0 and sgst == 0):
        action, code = "QUERY", "INTRA_IGST"
        hint = f"Compare Vendor State Code ({gstin[:2]}) with Company State Code ({COMPANY_STATE_CODE}). Which GST head applies?"
        explanation = "Wrong GST Tax Head (Sec 8/12 IGST Act): Both supplier and recipient are in Maharashtra (27), so CGST+SGST applies, not IGST."

    # Rule 5: TDS u/s 194J (FY 2025-26 threshold > Rs. 50,000)
    elif (tds_sec == "194J" or "professional" in ledger.lower() or "legal" in ledger.lower()) and base > 50000 and tds_amt == 0:
        action, code = "QUERY", "SEC_194J"
        hint = "Finance Act 2025 raised Sec 194J limit to ₹50,000. Check whether this bill crosses ₹50,000 and if TDS was deducted."
        explanation = f"Sec 194J Default: Professional fee of {format_inr(base)} exceeds ₹50,000 threshold with ₹0 TDS. Triggers 30% disallowance ({format_inr(base * 0.30)}) u/s 40(a)(ia)."

    # Rule 6: TDS u/s 194I Rent (FY 2025-26 threshold > Rs. 50,000 per month!)
    elif (tds_sec == "194I" or "rent" in ledger.lower()):
        if base > 50000 and tds_amt == 0:
            action, code = "QUERY", "SEC_194I"
            hint = "Under Finance Act 2025, Sec 194I applies when monthly rent exceeds ₹50,000/month."
            explanation = f"Sec 194I Default: Monthly rent of {format_inr(base)} exceeds ₹50,000/month without 10% TDS. Triggers 30% disallowance u/s 40(a)(ia)."
        else:
            action, code = "VERIFY", "CLEAN"
            hint = "Careful! Check the Finance Act 2025 amendment to Sec 194I (replaced ₹2.40L/yr with >₹50,000/month)."
            explanation = f"Budget 2025 Trap Avoided! Monthly rent of {format_inr(base)} is ≤ ₹50,000/month, so NO TDS is required u/s 194I in FY 2025-26."

    # Rule 7: TDS u/s 194C Contractors (Single > Rs. 30,000)
    elif (tds_sec == "194C" or "contract" in ledger.lower() or "amc" in ledger.lower()) and base > 30000 and tds_amt == 0:
        action, code = "QUERY", "SEC_194C"
        hint = "Check single-invoice threshold under Section 194C (₹30,000 single / ₹1,00,000 aggregate)."
        explanation = f"Sec 194C Default: Single contract bill of {format_inr(base)} exceeds ₹30,000 without TDS. 30% disallowance risk u/s 40(a)(ia)."

    return AuditVoucher(
        voucher_id=v_id,
        voucher_date=v_date,
        ledger_head=ledger,
        party_name=party,
        party_gstin=gstin,
        payment_mode=mode,
        base_amount=base,
        cgst=cgst,
        sgst=sgst,
        igst=igst,
        tds_section=tds_sec,
        tds_deducted=tds_amt,
        supporting_doc=doc,
        narration=narration,
        expected_action=action,
        expected_rule_code=code,
        senior_hint=hint,
        statutory_explanation=explanation,
    )


def generate_randomized_batch(seed: Optional[int] = None) -> List[AuditVoucher]:
    """
    Dynamically generates a fresh randomized audit batch so two trainees sitting
    next to each other cannot copy voucher numbers or exact Rupee figures.
    """
    rng = random.Random(seed)
    cash_amt = rng.choice([12800, 14500, 16200, 18900])
    laptop_amt = rng.choice([78000, 86000, 95000, 105000])
    igst_base = rng.choice([18000, 24000, 32000])
    prof_fee = rng.choice([60000, 65000, 75000, 85000])
    high_rent = rng.choice([58000, 65000, 72000])
    safe_rent = rng.choice([32000, 35000, 42000, 48000])  # <= 50,000 FA 2025 compliant trap!
    amc_bill = rng.choice([36000, 42000, 48500])

    raw_records = [
        {"Date": "2025-05-14", "Voucher_No": f"PUR-{rng.randint(101,199)}", "Ledger_Head": "Computers & Laptops",
         "Party_Name": "Infiniti Retail Ltd (Croma)", "Party_GSTIN": "27AABCI1234D1Z2", "Payment_Mode": "NEFT",
         "Base_Amount": laptop_amt, "CGST": 0, "SGST": 0, "IGST": 0, "TDS_Section": "NA", "TDS_Deducted": 0,
         "Supporting_Doc": "Tax Invoice (Billed to Director Personal Name)", "Narration": "Laptop purchased; ITC claimed in books"},

        {"Date": "2025-06-21", "Voucher_No": f"CSH-{rng.randint(201,299)}", "Ledger_Head": "Repairs & Maintenance",
         "Party_Name": "Sutar Woodworks", "Party_GSTIN": "URD", "Payment_Mode": "CASH",
         "Base_Amount": cash_amt, "CGST": 0, "SGST": 0, "IGST": 0, "TDS_Section": "NA", "TDS_Deducted": 0,
         "Supporting_Doc": "Cash Voucher + Handwritten Slip", "Narration": "Single-day cash payment for office desk repair"},

        {"Date": "2025-08-09", "Voucher_No": f"PUR-{rng.randint(301,399)}", "Ledger_Head": "Office Consumables",
         "Party_Name": "Shreeji Stationers (Andheri, Mumbai)", "Party_GSTIN": "27AAECS9988K1Z1", "Payment_Mode": "NEFT",
         "Base_Amount": igst_base, "CGST": 0, "SGST": 0, "IGST": round(igst_base * 0.18, 2), "TDS_Section": "NA", "TDS_Deducted": 0,
         "Supporting_Doc": "Vendor Tax Invoice", "Narration": "Local Mumbai purchase billed with 18% IGST"},

        {"Date": "2025-09-17", "Voucher_No": f"PUR-{rng.randint(401,499)}", "Ledger_Head": "Postage & Courier",
         "Party_Name": "Blue Dart Express Ltd", "Party_GSTIN": "27AAACB0446L1ZS", "Payment_Mode": "NEFT",
         "Base_Amount": 2500, "CGST": 225, "SGST": 225, "IGST": 0, "TDS_Section": "NA", "TDS_Deducted": 0,
         "Supporting_Doc": "Valid B2B Tax Invoice + Dispatch Register", "Narration": "Client audit report dispatches"},

        {"Date": "2025-11-11", "Voucher_No": f"CSH-{rng.randint(501,599)}", "Ledger_Head": "Printing & Stationery",
         "Party_Name": "Petty Cash Reimbursement", "Party_GSTIN": "URD", "Payment_Mode": "CASH",
         "Base_Amount": 7800, "CGST": 0, "SGST": 0, "IGST": 0, "TDS_Section": "NA", "TDS_Deducted": 0,
         "Supporting_Doc": "Internal Petty Cash Slip Only", "Narration": "Brochure printing; external vendor bill misplaced"},

        {"Date": "2025-12-19", "Voucher_No": f"JRN-{rng.randint(601,699)}", "Ledger_Head": "Legal & Professional Fees",
         "Party_Name": "M/s Mehta & Associates", "Party_GSTIN": "27AAFFM4321P1Z9", "Payment_Mode": "NEFT",
         "Base_Amount": prof_fee, "CGST": 0, "SGST": 0, "IGST": 0, "TDS_Section": "194J", "TDS_Deducted": 0,
         "Supporting_Doc": "Professional Fee Bill", "Narration": "Retainer consultancy fee credited without TDS"},

        {"Date": "2026-01-31", "Voucher_No": f"JRN-{rng.randint(701,799)}", "Ledger_Head": "Office Rent - Nariman Point",
         "Party_Name": "Nariman Chambers LLP", "Party_GSTIN": "27AABFN7766R1Z4", "Payment_Mode": "NEFT",
         "Base_Amount": high_rent, "CGST": round(high_rent * 0.09, 2), "SGST": round(high_rent * 0.09, 2), "IGST": 0,
         "TDS_Section": "194I", "TDS_Deducted": 0, "Supporting_Doc": "Monthly Rent Invoice", "Narration": "Monthly office rent credited"},

        {"Date": "2026-02-28", "Voucher_No": f"JRN-{rng.randint(801,899)}", "Ledger_Head": "Godown Rent - Bhiwandi",
         "Party_Name": "Bhiwandi Warehousing Co.", "Party_GSTIN": "URD", "Payment_Mode": "NEFT",
         "Base_Amount": safe_rent, "CGST": 0, "SGST": 0, "IGST": 0, "TDS_Section": "194I", "TDS_Deducted": 0,
         "Supporting_Doc": "Registered Leave & License + Receipt", "Narration": "Monthly warehouse rent (FA 2025 threshold check)"},

        {"Date": "2026-03-15", "Voucher_No": f"JRN-{rng.randint(901,999)}", "Ledger_Head": "AMC Contractor Charges",
         "Party_Name": "CoolAir HVAC Services Pvt Ltd", "Party_GSTIN": "27AAACC8877H1Z3", "Payment_Mode": "CHEQUE",
         "Base_Amount": amc_bill, "CGST": round(amc_bill * 0.09, 2), "SGST": round(amc_bill * 0.09, 2), "IGST": 0,
         "TDS_Section": "194C", "TDS_Deducted": 0, "Supporting_Doc": "AMC Tax Invoice", "Narration": "Single annual AC maintenance bill"},
    ]

    df = pd.DataFrame(raw_records)
    return [auto_classify_tally_row(row) for _, row in df.iterrows()]


def build_working_paper_excel(results: List[Dict]) -> bytes:
    """Exports trainee decisions & raised audit queries into a formatted .xlsx Working Paper."""
    output = io.BytesIO()
    df_out = pd.DataFrame(results)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="WP_Audit_Observations")
    return output.getvalue()


# ==============================================================================
# 2. FASTAPI REST ENDPOINT (FOR FIRM LMS / EXTERNAL INTEGRATIONS)
# ==============================================================================

@app.post("/api/scan-tally-excel")
async def api_scan_tally_excel(file: UploadFile = File(...)):
    """Upload a Tally Daybook .xlsx/.csv and get instant JSON statutory audit flags."""
    contents = await file.read()
    df = pd.read_csv(io.BytesIO(contents)) if file.filename.endswith(".csv") else pd.read_excel(io.BytesIO(contents))
    vouchers = [asdict(auto_classify_tally_row(row)) for _, row in df.iterrows()]
    return {
        "company": COMPANY_NAME,
        "financial_year": "2025-26",
        "total_vouchers_scanned": len(vouchers),
        "exceptions_flagged": sum(1 for v in vouchers if v["expected_action"] == "QUERY"),
        "vouchers": vouchers,
    }


# ==============================================================================
# 3. INTERACTIVE NICEGUI FRONTEND ("TACTILE AUDIT WORKSPACE")
# ==============================================================================

@ui.page("/")
def audit_simulator_page():
    ui.colors(primary="#1E3A8A", secondary="#D97706", positive="#15803D", warning="#EA580C", negative="#DC2626")

    state = {
        "trainee_name": "Article Trainee",
        "vouchers": generate_randomized_batch(),
        "idx": 0,
        "decisions": {},  # voucher_id -> dict
        "selected_rule": "CLEAN",
    }

    # --- Top HUD Header ---
    with ui.header().classes("bg-slate-900 text-white px-6 py-3 flex items-center justify-between shadow-lg border-b border-slate-700"):
        with ui.row().classes("items-center gap-3"):
            ui.icon("verified_user", size="28px").classes("text-amber-400")
            with ui.column().classes("gap-0"):
                ui.label("THE FIRST-YEAR AUDIT TRAIL — PYTHON ENGINE").classes("font-bold tracking-wide text-base")
                ui.label(f"Client: {COMPANY_NAME} | GSTIN: {COMPANY_GSTIN} | FY 2025-26").classes("text-xs text-slate-400 font-mono")

        with ui.row().classes("items-center gap-3"):
            progress_badge = ui.badge("W/P 1 / 9", color="amber-7").classes("px-3 py-1 font-mono text-xs")
            score_badge = ui.badge("Accuracy: 100%", color="green-7").classes("px-3 py-1 font-mono text-xs")

            async def handle_excel_upload(e):
                raw = e.content.read()
                df = pd.read_csv(io.BytesIO(raw)) if e.name.endswith(".csv") else pd.read_excel(io.BytesIO(raw))
                state["vouchers"] = [auto_classify_tally_row(row) for _, row in df.iterrows()]
                state["idx"] = 0
                state["decisions"] = {}
                ui.notify(f"Loaded {len(state['vouchers'])} vouchers from {e.name}!", type="positive")
                render_workspace.refresh()

            ui.upload(label="Upload Tally .xlsx", auto_upload=True, on_upload=handle_excel_upload) \\
                .props("accept=.xlsx,.csv flat dense color=amber").classes("max-w-[180px] text-xs")

            def regenerate_batch():
                state["vouchers"] = generate_randomized_batch()
                state["idx"] = 0
                state["decisions"] = {}
                ui.notify("Generated fresh randomized audit figures for new trainee!", type="info")
                render_workspace.refresh()

            ui.button("New Random Batch", icon="shuffle", on_click=regenerate_batch).props("outline dense color=white").classes("px-3 text-xs")

    # --- Main Dual-Pane Audit Desk ---
    @ui.refreshable
    def render_workspace():
        vouchers: List[AuditVoucher] = state["vouchers"]
        total = len(vouchers)
        done_count = len(state["decisions"])
        correct_count = sum(1 for d in state["decisions"].values() if d["is_correct"])
        acc = round((correct_count / done_count) * 100) if done_count else 100

        progress_badge.set_text(f"Audited: {done_count} / {total}")
        score_badge.set_text(f"Partner Confidence: {acc}%")

        if state["idx"] >= total:
            # --- LEVEL COMPLETE: WORKING PAPER SUMMARY & EXCEL EXPORT ---
            with ui.column().classes("w-full max-w-6xl mx-auto p-6 gap-4"):
                with ui.card().classes("w-full p-6 bg-white border-t-4 border-blue-900 shadow-md"):
                    with ui.row().classes("w-full justify-between items-center mb-4"):
                        with ui.column().classes("gap-1"):
                            ui.label("DRAFT STATUTORY AUDIT OBSERVATION SCHEDULE (W/P REF: D-1)").classes("text-xl font-bold text-slate-900")
                            ui.label(f"Entity: {COMPANY_NAME} | Financial Year: 2025-26 (AY 2026-27) | Score: {correct_count}/{total} ({acc}%)").classes("text-sm text-slate-600")

                        def download_wp():
                            xlsx_bytes = build_working_paper_excel(list(state["decisions"].values()))
                            ui.download(xlsx_bytes, "Apex_Audit_Working_Paper_FY25_26.xlsx")

                        ui.button("Export Working Paper (.xlsx)", icon="download", on_click=download_wp).classes("bg-blue-900 text-white")

                    columns = [
                        {"name": "voucher_id", "label": "Voucher #", "field": "voucher_id", "align": "left"},
                        {"name": "ledger_head", "label": "Ledger Account", "field": "ledger_head", "align": "left"},
                        {"name": "amount", "label": "Base Amount", "field": "amount", "align": "right"},
                        {"name": "trainee_action", "label": "Your Stamp", "field": "trainee_action", "align": "center"},
                        {"name": "status", "label": "Review Result", "field": "status", "align": "center"},
                        {"name": "statutory_note", "label": "Statutory Provision & Impact", "field": "statutory_note", "align": "left"},
                    ]
                    ui.table(columns=columns, rows=list(state["decisions"].values()), row_key="voucher_id").classes("w-full")
            return

        v = vouchers[state["idx"]]
        existing_decision = state["decisions"].get(v.voucher_id)

        with ui.row().classes("w-full max-w-7xl mx-auto p-6 gap-6 items-start no-wrap"):
            # LEFT STAGE (60%): Realistic Ivory Tax Voucher / Ledger Evidence
            with ui.card().classes("w-7/12 bg-[#FCFBF7] border border-stone-300 shadow-xl p-6 relative rounded-md"):
                with ui.row().classes("w-full justify-between items-start border-b-2 border-slate-800 pb-3 mb-4"):
                    with ui.column().classes("gap-0"):
                        ui.label("EVIDENCE VOUCHER / LEDGER EXTRACT").classes("text-xs font-mono tracking-widest text-slate-500")
                        ui.label(f"{v.party_name}").classes("text-xl font-bold text-slate-900")
                        ui.label(f"Supplier GSTIN: {v.party_gstin}").classes("text-xs font-mono text-slate-700")
                    with ui.column().classes("items-end gap-1"):
                        ui.badge(f"Voucher: {v.voucher_id}", color="blue-10").classes("font-mono px-2 py-1")
                        ui.label(f"Date: {v.voucher_date}").classes("text-xs font-mono text-slate-600")

                # Stamp Overlay if already audited
                if existing_decision:
                    stamp_color = "border-green-700 text-green-700" if existing_decision["trainee_action"] == "VERIFY" else "border-orange-600 text-orange-600"
                    stamp_text = "✓ VERIFIED & TICKED" if existing_decision["trainee_action"] == "VERIFY" else "? AUDIT QUERY RAISED"
                    ui.label(stamp_text).classes(f"border-4 {stamp_color} font-mono font-extrabold px-4 py-1 rounded text-lg rotate-[-3deg] self-end mb-2")

                with ui.grid(columns=2).classes("w-full gap-3 bg-slate-100 p-4 rounded border border-slate-200 mb-4 text-sm"):
                    ui.label(f"Account Head: {v.ledger_head}").classes("font-semibold text-slate-800")
                    ui.label(f"Payment Mode: {v.payment_mode}").classes("font-mono font-bold text-slate-800")
                    ui.label(f"Billed To GSTIN: {COMPANY_GSTIN if v.expected_rule_code != 'SEC_16_2' else 'MISSING (Personal Name)'}") \\
                        .classes("font-mono text-xs text-slate-700")
                    ui.label(f"Supporting Attached: {v.supporting_doc}").classes("text-xs font-medium text-amber-900")

                # Financial Breakdown Table
                total_val = v.base_amount + v.cgst + v.sgst + v.igst - v.tds_deducted
                with ui.column().classes("w-full font-mono text-sm border border-slate-300 rounded overflow-hidden mb-4"):
                    with ui.row().classes("w-full justify-between bg-slate-800 text-white px-4 py-2 font-bold"):
                        ui.label("Particulars / Tax Component")
                        ui.label("Amount (INR)")
                    for label_txt, val in [
                        ("Taxable / Base Value", v.base_amount),
                        ("CGST (9%)", v.cgst),
                        ("SGST (9%)", v.sgst),
                        ("IGST (18%)", v.igst),
                        (f"Less: TDS Deducted ({v.tds_section})", -v.tds_deducted),
                    ]:
                        with ui.row().classes("w-full justify-between px-4 py-1.5 border-b border-slate-200"):
                            ui.label(label_txt)
                            ui.label(format_inr(val))
                    with ui.row().classes("w-full justify-between px-4 py-2 bg-amber-50 font-bold text-slate-900"):
                        ui.label("Net Recorded / Paid")
                        ui.label(format_inr(total_val))

                ui.label(f"Narration: {v.narration}").classes("text-xs italic text-slate-600")

            # RIGHT STAGE (40%): Audit Action Console & Senior's Notebook
            with ui.column().classes("w-5/12 gap-4"):
                with ui.card().classes("w-full p-5 bg-white shadow-md border border-slate-200"):
                    ui.label("AUDIT WORKPAPER DECISION CONSOLE").classes("text-xs font-mono font-bold tracking-wider text-blue-900 mb-2")

                    rule_select = ui.select(
                        options=RULE_OPTIONS,
                        value=state["selected_rule"],
                        label="Select Statutory Basis / Audit Observation",
                    ).classes("w-full mb-4")

                    def submit_audit(action_chosen: str):
                        chosen_rule = rule_select.value
                        if action_chosen == "VERIFY":
                            is_ok = (v.expected_action == "VERIFY")
                        else:
                            is_ok = (v.expected_action == "QUERY" and chosen_rule == v.expected_rule_code)

                        state["decisions"][v.voucher_id] = {
                            "voucher_id": v.voucher_id,
                            "ledger_head": v.ledger_head,
                            "amount": format_inr(v.base_amount),
                            "trainee_action": action_chosen,
                            "is_correct": is_ok,
                            "status": "ACCURATE" if is_ok else "REVIEW NEEDED",
                            "statutory_note": v.statutory_explanation,
                        }
                        if is_ok:
                            ui.notify("Spot on! Statutory treatment matches FY 2025-26 audit standards.", type="positive")
                        else:
                            ui.notify(f"Senior Review Note: {v.statutory_explanation}", type="warning", timeout=6000)
                        render_workspace.refresh()

                    with ui.row().classes("w-full gap-3"):
                        ui.button("STAMP VERIFIED (✓)", on_click=lambda: submit_audit("VERIFY")) \\
                            .classes("flex-1 bg-green-700 text-white font-bold py-2")
                        ui.button("RAISE QUERY (?)", on_click=lambda: submit_audit("QUERY")) \\
                            .classes("flex-1 bg-orange-600 text-white font-bold py-2")

                    if existing_decision:
                        bg = "bg-green-50 border-green-500" if existing_decision["is_correct"] else "bg-amber-50 border-amber-500"
                        with ui.card().classes(f"w-full mt-4 p-3 border-l-4 {bg}"):
                            ui.label(existing_decision["status"]).classes("font-bold text-xs")
                            ui.label(v.statutory_explanation).classes("text-xs text-slate-700")

                    with ui.row().classes("w-full justify-between mt-4 pt-3 border-t border-slate-200"):
                        ui.button("← Previous", on_click=lambda: (state.update(idx=max(0, state["idx"] - 1)), render_workspace.refresh())).props("flat dense")
                        ui.button("Next Voucher →", on_click=lambda: (state.update(idx=state["idx"] + 1), render_workspace.refresh())).props(" color=primary dense").classes("px-4")

                # Collapsible Senior's Field Notebook & FA 2025 Quick Card
                with ui.expansion("Senior's Field Notebook & FY 2025-26 Thresholds", icon="menu_book", value=True).classes("w-full bg-amber-50 border border-amber-200 rounded shadow-sm"):
                    with ui.column().classes("p-3 gap-2 text-xs text-slate-800"):
                        ui.label(f"Hint for {v.voucher_id}: {v.senior_hint}").classes("font-medium text-amber-950")
                        ui.separator()
                        ui.label("FY 2025-26 Quick Statutory Thresholds (Finance Act 2025):").classes("font-bold text-blue-950")
                        ui.label("• Sec 40A(3) Cash Limit: > ₹10,000 in a single day (100% disallowed)")
                        ui.label("• Sec 194J (Prof/Tech): > ₹50,000 p.a. (10% / 2%)")
                        ui.label("• Sec 194I (Rent): > ₹50,000 per month (replaces ₹2.40L annual limit!)")
                        ui.label("• Sec 194C (Contract): > ₹30,000 single bill OR > ₹1,00,000 aggregate")

    render_workspace()


if __name__ in {"__main__", "__mp_main__"}:
    ui.run(title="First-Year Audit Trail (Python Engine)", port=8080, reload=False)
'''

py_filename = "audit_trail_nicegui_app.py"
with open(py_filename, "w", encoding="utf-8") as f:
    f.write(python_app_code)

print("Generated files successfully:", excel_filename, py_filename)
