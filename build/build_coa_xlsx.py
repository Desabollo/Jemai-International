"""Builds 01_CoA_Improved_Jemai_Oct2026.xlsx (and the Zoho import CSV)."""
import csv
import json
import os
import xlsxwriter
from coa_model import build, CHANGES, NEW_ACCOUNTS, TYPE_BLOCKED, HDR

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "deliverables")
XLSX = os.path.join(OUT_DIR, "01_Chart_of_Accounts_Improved_Jemai_Oct2026.xlsx")
CSV = os.path.join(OUT_DIR, "01b_CoA_Zoho_Import.csv")

BANK = json.load(open("/home/user/work/bank.json"))
from collections import Counter
LEG_CNT = Counter(d["COA Account"] for d in BANK)

# --------------------------------------------------------------------------- #
# legacy -> new mapping (the 33-account controlled chart used in the workbook)
# --------------------------------------------------------------------------- #
LEGACY_MAP = [
    ("1010", "Access Bank A/C 1854793470", "1021", "Access Bank",
     "1:1. Only one bank account exists (A/C 1854793470).", "Direct"),
    ("1110", "Trade Receivables - Clients", "1041", "Trade Receivables - Project Clients",
     "Reversed in the corrected AFS (no billings evidenced). Going forward, invoice-driven debtors only.",
     "Direct"),
    ("1510", "PPE - Office Furniture, Fixtures & Equipment", "1091 / 1092 / 1093",
     "Furniture & Fittings / Office Equipment / Computers & IT",
     "Split by asset class so the annual allowance rates (NTA 2025: 10/20/25%) can be tracked per class.",
     "Split on invoice"),
    ("1590", "Accumulated Depreciation - PPE", "1101 / 1102 / 1103",
     "Accumulated Depreciation - Furniture / Office Equipment / Computers & IT",
     "Contra account split to mirror 1510.", "Split on asset class"),
    ("2010", "Trade Payables - Suppliers & Materials", "2011", "Trade Payables - Suppliers",
     "AFS split: N1,335,000 other payables + N300,000 accrued professional fees -> 2011 + 2023.",
     "Split per invoice"),
    ("2020", "Contract Liabilities - Client Advances & Deposits", "2031",
     "Contract Liabilities - Project Deposits",
     "IFRS 15 contract liability. N3,460,000 unchanged at both year ends.", "Direct"),
    ("2030", "Other Payables & Sundry (For Review)", "2083 / 2062",
     "Other Payables / Related Party Payable",
     "Suspense-like legacy account: split by payee substance before posting.", "Review each item"),
    ("2040", "Other Receivables & Sundry (For Review)", "1046", "Other Receivables",
     "N332,672.02 of unverified POS/sundry credits.", "Direct"),
    ("2110", "Current Tax - CIT / Minimum Tax Payable", "2045", "Income Tax Payable (CIT)",
     "Minimum tax 0.5% of gross income (CITA 2020 s.37(2); retained by NTAA 2025).", "Direct"),
    ("2120", "Current Tax - Tertiary Education Levy Payable", "2046",
     "Development Levy Payable",
     "TEL is repealed for periods beginning on/after 1-Jan-2026 and replaced by the 4% development levy (NTAA 2025). "
     "Carry the 2025 TEL balance in 2046 until the FIRS position is agreed.",
     "Reclass at transition"),
    ("2210", "Related Party Receivable / (Payable) - Esther Akin-Ajayi", "1045",
     "Director Current Account Receivable",
     "N14,223,846.92 cumulative. IAS 24 disclosure and ECL assessment required; also see 2061/2053 if the balance turns payable.",
     "Direct"),
    ("2220", "Third-Party Short-Term Loans Payable", "2051 / 2052",
     "Short-term Loan - individuals / Working Capital Loan",
     "N230,000 (Olajide, 2024) -> 2051; N1,230,000 facility -> 2052. Interest accrual to 6912.", "Split by lender"),
    ("3010", "Ordinary Share Capital (N1.00 par)", "3011", "Ordinary Share Capital",
     "1,000,000 shares of N1.00. N1,000 Ime share to be reversed as a separate equity event.", "Direct"),
    ("3020", "Retained Earnings (Accumulated Profit)", "3090", "Retained Earnings",
     "Renumbered from 2030 (duplicate code) and moved into the 3xxx equity block.", "Direct"),
    ("3030", "Other Equity - Share Premium / Reserves", "3071 / 3072",
     "Capital Reserve / Revaluation & Fair Value Reserve",
     "N2,000 net Ime buy-back contribution -> 3071.", "Direct"),
    ("4010", "Revenue - Interiors, Furniture, Events & Rentals", "4011-4019",
     "Service-line revenue",
     "IFRS 15 requires disaggregation by service line, not by customer type. Classify each receipt by the "
     "performance obligation in the contract/narration.", "Split by narration/contract"),
    ("4030", "Revenue - Consultancy & Setup Services", "4015 / 4016 / 4017",
     "Consultation / Project management / 3D visualisation", "Split by contract scope.", "Split by contract"),
    ("5010", "Cost of Sales - Materials, Furniture, Site Works & Refunds", "5011-5029",
     "Direct materials / subcontractors",
     "226 transactions sit here. Split into materials (5011-5019) and subcontractor/labour (5021-5029) using the "
     "narration purpose (e.g. 'Electrical works' -> 5023, 'Painting touch up' -> 5026, 'Deep Cleaning' -> 5037).",
     "Rule-based split"),
    ("6010", "Administrative - Salaries, Wages & Allowances", "6110 / 6112",
     "Salaries and Employee Cost / Contract Staff", "Split employees vs contract staff/interns.", "Payroll schedule"),
    ("6020", "Administrative - Rent & Premises", "6210", "Rent & Service Charge",
     "N200,000 p.a. with no bank trace - confirm whether paid by the director (route through 1045).", "Direct"),
    ("6030", "Administrative - Transportation, Logistics, Fuel, Flights & Travel", "6510-6514",
     "Transport & Logistics / Local Travel / Fuel / Hotel", "Split by narration (logistics vs staff travel).", "Rule-based split"),
    ("6040", "Administrative - Professional, Legal & Incorporation Fees", "6410-6419",
     "Legal / Accounting / Audit / CAC / Permits", "Split by the specific professional service.", "Split per invoice"),
    ("6050", "Administrative - Bank Charges, Levies, Airtime & Telecom", "6910 / 6914 / 6915 / 6916 / 6213 / 6212",
     "Bank charges / VAT on bank charges / transfer commission / EMTL / airtime",
     "THE largest mis-posting: 607 of 1,226 transactions (49.5%) in one account, mixing bank charges, "
     "site expenses, petty cash, fuel, transport, training and refunds. This single account destroyed the "
     "expense analysis in both AFS.", "Rule-based split - critical"),
    ("6060", "Administrative - Repairs & Maintenance", "6216", "Repairs and Maintenance", "1:1.", "Direct"),
    ("6070", "Administrative - Depreciation (IAS 16)", "6810-6816",
     "Depreciation by asset class", "Split to match the PPE classes so the IAS 16 note writes itself.", "Asset register"),
    ("6080", "Administrative - Staff Welfare, Catering & Training", "6113 / 6114 / 6115",
     "Welfare & Meals / Medical / Training", "Split by nature; training was N650,000 in 2025 and is separately "
     "disclosed in the AFS.", "Split per item"),
    ("6090", "Administrative - General, Utilities, Gifts & Church", "6211 / 6214 / 6612 / 6216",
     "Utilities / Office supplies / Charitable donations", "Charitable donations and gifts are non-allowable for "
     "CIT and must be isolated (6612).", "Split per item"),
    ("6100", "Direct Cost - Branding & Marketing Production", "6311 / 6312 / 5041",
     "Photography / Branding & Graphics / 3D rendering",
     "Branding delivered to a client is a direct cost (5041-5044); advertising for Jemai itself is 6300-series.",
     "Contract vs own-account test"),
    ("7010", "Finance Costs - Overdraft Interest", "6911", "Debit Interest", "1:1.", "Direct"),
    ("8010", "Income Tax Expense - Current Tax (CIT/Min Tax + TEL)", "6710 / 6711",
     "Income Tax Expense / Education Tax (now Development Levy) Expense",
     "Split current CIT/minimum tax from TEL/development levy.", "Direct"),
]

# --------------------------------------------------------------------------- #
# reporting map
# --------------------------------------------------------------------------- #
def fs_line(a):
    c, t, n = a["code"], a["type"], a["name"].lower()
    def rng(lo, hi):
        return bool(c) and int(lo) <= int(c) <= int(hi)
    if c == "1029":
        return "Cash and cash equivalents (overdraft disclosed gross)", "Note 4", "Financing"
    if rng("1011", "1029"):
        return "Cash and cash equivalents", "Note 4", "Cash & equivalents"
    if rng("1030", "1035"):
        return "Trade and other receivables (clearing - must sweep to nil)", "Note 3", "Operating"
    if rng("1040", "1042"):
        return "Trade and other receivables", "Note 3", "Operating"
    if rng("1043", "1046"):
        return "Trade and other receivables", "Note 3", "Operating"
    if rng("1050", "1053"):
        return "Contract assets (unbilled revenue)", "Note 3", "Operating"
    if rng("1061", "1066"):
        return "Inventories", "Note 3", "Operating"
    if rng("1070", "1075"):
        return "Prepayments and other receivables", "Note 3", "Operating"
    if rng("1080", "1086"):
        return "Tax recoverable (VAT / WHT credit notes)", "Note 6", "Tax paid"
    if rng("1091", "1098"):
        return "Property, plant and equipment - cost", "Note 2", "Investing"
    if rng("1100", "1107"):
        return "Property, plant and equipment - accumulated depreciation", "Note 2", "Non-cash"
    if rng("1121", "1129"):
        return "Intangible assets (website, software, brand)", "Note 2", "Investing"
    if rng("1131", "1132"):
        return "Other non-current assets / investments", "Note 3", "Investing"
    if c == "1180":
        return "Right-of-use assets (IFRS 16)", "Note 2", "Non-cash"
    if c == "1190":
        return "Deferred tax asset", "Note 16", "Non-cash"
    if rng("1990", "1992"):
        return "Suspense balances under review (must be nil at year end)", "n/a", "Operating"
    if rng("2010", "2015"):
        return "Trade and other payables", "Note 5", "Operating"
    if rng("2020", "2024"):
        return "Accruals", "Note 5", "Operating"
    if rng("2030", "2033"):
        return "Contract liabilities (customer advances)", "Note 13", "Operating"
    if rng("2040", "2049"):
        return "Taxation (VAT / WHT / CIT / levy payables)", "Note 6", "Tax paid"
    if rng("2050", "2054"):
        return "Borrowings - current", "Note 15", "Financing"
    if c == "2057":
        return "Share application money pending allotment", "Note 7", "Financing"
    if rng("2060", "2064"):
        return "Trade and other payables - related parties & staff", "Note 12", "Operating"
    if rng("2065", "2066"):
        return "Provisions", "Note 14", "Operating"
    if rng("2070", "2076"):
        return "Payroll and statutory liabilities", "Note 5", "Operating"
    if rng("2080", "2085"):
        return "Other payables", "Note 5", "Operating"
    if c == "2090":
        return "Opening balance adjustments (Zoho system)", "n/a", "Non-cash"
    if c == "2096":
        return "Deferred tax liability", "Note 16", "Non-cash"
    if c == "2097":
        return "Long-term borrowings", "Note 15", "Financing"
    if c == "2098":
        return "Lease liabilities (IFRS 16)", "Note 15", "Financing"
    if c == "2999":
        return "Suspense - liability review", "n/a", "Operating"
    if c in ("3010", "3011"):
        return "Share capital", "Note 7", "Financing"
    if c == "3020":
        return "Share premium", "Note 7", "Financing"
    if rng("3050", "3052"):
        return "Director drawings (contra-equity)", "Note 8", "Financing"
    if rng("3060", "3061"):
        return "Director equity injections", "Note 7", "Financing"
    if rng("3070", "3072"):
        return "Reserves", "Note 7", "Non-cash"
    if c in ("3080", "3999"):
        return "Opening balance offset / equity (Zoho system)", "n/a", "Non-cash"
    if c == "3090":
        return "Retained earnings", "Note 8", "Non-cash"
    if rng("4010", "4024") or rng("4030", "4034"):
        return "Revenue", "Note 9", "Operating"
    if c in ("4040", "4041", "4042"):
        return "Revenue (discounts / refunds / adjustments - contra)", "Note 9", "Operating"
    if c == "4043":
        return "Finance income", "Note 11", "Operating"
    if rng("4900", "4903"):
        return "Other income", "Note 9", "Operating"
    if c == "4999":
        return "Revenue suspense (must be nil at year end)", "n/a", "Operating"
    if rng("5000", "5099"):
        return "Direct cost", "Note 10", "Operating"
    if rng("6100", "6117"):
        return "Administrative expenses - staff costs", "Note 11", "Operating"
    if rng("6200", "6224"):
        return "Administrative expenses - office & IT", "Note 11", "Operating"
    if rng("6300", "6317"):
        return "Administrative expenses - selling & marketing", "Note 11", "Operating"
    if rng("6400", "6419"):
        return "Administrative expenses - professional & regulatory", "Note 11", "Operating"
    if rng("6500", "6514"):
        return "Administrative expenses - transport & travel", "Note 11", "Operating"
    if rng("6600", "6613"):
        return "Administrative expenses - general operating", "Note 11", "Operating"
    if rng("6700", "6713"):
        return "Taxation", "Note 6", "Tax paid"
    if rng("6800", "6822"):
        return "Depreciation and amortisation", "Note 11", "Non-cash"
    if rng("6900", "6917"):
        return "Finance costs", "Note 11", "Operating"
    if rng("7000", "7013"):
        return "Other expenses / suspense", "Note 11", "Operating"
    if rng("7020", "7024"):
        return "Impairment and bad debts", "Note 3", "Non-cash"
    if c == "7023":
        return "Other income / (expense) - exchange", "Note 11", "Operating"
    if c == "7999":
        return "Uncategorised (must be nil at year end)", "n/a", "Operating"
    return {"Income": "Revenue", "Other Income": "Other income", "Cost Of Goods Sold": "Direct cost",
            "Expense": "Administrative expenses", "Other Expense": "Other expenses"}.get(t, t), "", "Operating"


def tax_tag(a):
    c, n = a["code"], a["name"].lower()
    if not c:
        return "Set a code"
    if c.startswith("1029"):
        return "Principal not taxable; overdraft interest allowable"
    if c[:2] in ("10", "11") and not c.startswith(("1080", "1081", "1082", "1083", "1084", "1086")):
        if c.startswith(("1091", "1092", "1093", "1094", "1095", "1096", "1097", "1098")):
            return "Capital allowance (NTA 2025: 10% / 20% / 25% + 1% residual)"
        return "Balance sheet - not taxable"
    if c.startswith(("1080", "1081", "1086")):
        return "Recoverable input VAT (VAT Act 2023 s.11)"
    if c.startswith(("1082", "1043")):
        return "WHT credit note recoverable against CIT"
    if c.startswith(("1083", "1084")):
        return "Recoverable / credit against tax"
    if c.startswith(("1180", "1190")):
        return "Deferred tax (IAS 12) - no tax base"
    if c.startswith("2"):
        return "Balance sheet - deduction on payment (where allowable)"
    if c.startswith("3"):
        return "Non-taxable equity movement"
    if c.startswith(("4010", "4011", "4012", "4013", "4014", "4015", "4016", "4017", "4018", "4019",
                     "4020", "4021", "4022", "4023", "4024", "4030", "4031", "4032", "4033", "4034")):
        return "Taxable turnover - VAT output 7.5% once registered"
    if c in ("4041", "4042"):
        return "Reduces taxable turnover"
    if c in ("4040", "4900", "4901", "4903", "4043"):
        return "Other income - taxable"
    if c.startswith("5"):
        return "Allowable direct cost (invoice + WHT credit required)"
    if c.startswith(("6110", "6111", "6112")):
        return "Allowable payroll cost (PAYE / pension compliance required)"
    if c.startswith(("6113", "6114", "6115", "6116", "6117")):
        return "Allowable staff cost (verify benefit-in-kind)"
    if c.startswith(("6314", "6611")):
        return "Entertainment - restricted / non-allowable"
    if c.startswith("6612"):
        return "Donations - non-allowable"
    if c.startswith("6417"):
        return "Penalties & fines - non-allowable"
    if c.startswith(("6710", "6711")):
        return "Tax expense - not deductible"
    if c.startswith("6712"):
        return "Irrecoverable VAT - allowable cost"
    if c.startswith("6713"):
        return "WHT suffered - non-recoverable"
    if c.startswith(("6810", "6811", "6812", "6813", "6814", "6815", "6816", "6820", "6822")):
        return "Accounting depreciation - add back; claim capital allowances instead"
    if c.startswith(("6910", "6911", "6912", "6913", "6914", "6915", "6916", "6917")):
        return "Allowable finance cost (input VAT on charges recoverable once registered)"
    if c.startswith(("7013", "7021", "7022", "7024")):
        return "Allowable on write-off (evidence required)"
    if c.startswith("7023"):
        return "Exchange difference - allowable / taxable"
    if c.startswith("7"):
        return "Review substance before claiming"
    return "Review"


def build_workbook():
    data, index, added = build()
    wb = xlsxwriter.Workbook(XLSX)

    F = {
        "title": wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864"}),
        "sub": wb.add_format({"italic": True, "font_size": 9, "font_color": "#595959"}),
        "hdr": wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                              "border": 1, "text_wrap": True, "valign": "vcenter", "font_size": 9}),
        "cell": wb.add_format({"border": 1, "font_size": 9, "valign": "top"}),
        "wrap": wb.add_format({"border": 1, "font_size": 9, "text_wrap": True, "valign": "top"}),
        "num": wb.add_format({"border": 1, "font_size": 9, "num_format": "#,##0.00", "valign": "top"}),
        "int": wb.add_format({"border": 1, "font_size": 9, "num_format": "#,##0", "valign": "top"}),
        "bold": wb.add_format({"bold": True, "font_size": 9, "border": 1, "valign": "top"}),
        "red": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FFC7CE", "text_wrap": True, "valign": "top"}),
        "amber": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FFEB9C", "text_wrap": True, "valign": "top"}),
        "green": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#C6EFCE", "text_wrap": True, "valign": "top"}),
    }

    # ---------------------------------------------------------------- Read_Me
    ws = wb.add_worksheet("Read_Me")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 34)
    ws.set_column("C:C", 110)
    ws.write("B2", "JEMAI INTERNATIONAL LIMITED  |  RC 7171046  |  TIN 2522579844227", F["title"])
    ws.write("B3", "REVIEWED & IMPROVED CHART OF ACCOUNTS - ZOHO BOOKS UPLOAD FILE", F["sub"])
    ws.write("B4", "Prepared for: Mrs. Esther Akin-Ajayi Jesutofunmi Ifeoluwa (MD/CEO)   |   Prepared: November 2026   |   "
                   "Basis: Chart of Accounts Oct 2026 export + FY2024/FY2025 bank-evidenced working papers", F["sub"])
    rows = [
        ("WHAT THIS FILE IS", ""),
        ("Purpose",
         "A review of the existing Zoho Books chart of accounts for completeness and reporting fitness, and a corrected "
         "chart in the exact Zoho Books upload format (same 12 columns as the client export)."),
        ("Scope of review", "302 accounts: the 275 accounts in the October 2026 export, retained, corrected and re-parented, "
                            "plus 27 new accounts required by IFRS and by the FY2026 Nigerian tax regime."),
        ("Result", "2 blocking defects, 12 high-severity defects, 7 overlapping accounts deactivated, 27 accounts added. "
                   "The chart is now capable of producing a full IFRS statement of financial position, statement of profit or "
                   "loss and notes without manual reclassification."),
        ("", ""),
        ("TABS", ""),
        ("Zoho_Import_CoA", "READY TO UPLOAD. Exactly the Zoho Books chart-of-accounts layout (Account ID, Account Name, "
                            "Account Code, Description, Account Type, Mileage Rate, Mileage Unit, IsMileage, Account #, "
                            "Account Status, Currency, Parent Account). Sorted parents-first."),
        ("Changes_Log", "Every change made to the client chart: what, why, severity, and the exact action in Zoho."),
        ("Validation_Tests", "18 structural tests (duplicate codes, orphan accounts, inverted hierarchies, type "
                             "conflicts, taxonomy coverage) - before vs after."),
        ("Legacy_to_New_Map", "The 33-account controlled chart used in the FY2024/FY2025 working papers mapped to the new "
                              "codes, with transaction counts so you can size the re-coding effort."),
        ("Reporting_Map", "Account-by-account: IFRS statement line, note reference, tax tag and cash-flow category. This "
                          "is the control that lets Zoho reports tie to the audited statements."),
        ("Gap_Analysis", "IFRS / tax / reporting requirement vs chart coverage, before and after."),
        ("Open_Balances_1Jan2026", "The 31 December 2025 closing trial balance (bank-evidenced basis) expressed in the new "
                                   "codes - the exact opening balances to load into Zoho Books on 1 January 2026."),
        ("", ""),
        ("HOW TO UPLOAD", ""),
        ("1. Back up", "Zoho Books > Settings > Data Backup (or export the current Chart of Accounts) before importing."),
        ("2. Fix before adding", "Two defects must be corrected in the UI first, because Zoho matches on Account Name and "
                                 "rejects duplicate codes: (a) change Retained Earnings code 2030 -> 3090; "
                                 "(b) give 'License Fees' the code 6419."),
        ("3. Import", "Accountant > Chart of Accounts > Import. Upload 01b_CoA_Zoho_Import.csv, map the 12 columns to the "
                      "matching Zoho fields (they carry the same names), and select 'Update existing accounts' so the "
                      "parent and status changes are applied rather than duplicated."),
        ("4. Verify", "After import, run Accountant > Chart of Accounts and confirm: no duplicate codes, no account "
                      "without a code, 7 accounts Inactive, and the parent/child tree matching tab Zoho_Import_CoA."),
        ("5. Lock", "Deactivate the 7 overlapping accounts only after the 2025 conversion journals are posted; Zoho will "
                    "not let you post to an inactive account but existing balances remain visible."),
        ("", ""),
        ("CONVENTIONS", ""),
        ("Numbering", "1xxx assets | 2xxx liabilities | 3xxx equity | 4xxx revenue | 5xxx direct cost | 6xxx operating "
                      "expense | 7xxx other, tax, depreciation and finance."),
        ("Placeholders", "Account ID is blank for the 25 new accounts (Zoho assigns its own ID on creation). Existing IDs "
                         "are retained for reference only - Zoho ignores them on import."),
        ("Currency", "All accounts NGN (IAS 21 functional and presentation currency)."),
        ("Parent and child", "Zoho Books requires a parent account and its children to share the same Account Type. Where a "
                             "logical grouping is impossible for that reason it is documented in Changes_Log rather than "
                             "forced."),
    ]
    r = 6
    for a, b in rows:
        if a and not b:
            ws.write(r, 1, a, F["bold"])
        elif a:
            ws.write(r, 1, a, F["wrap"])
            ws.write(r, 2, b, F["wrap"])
        r += 1

    # ------------------------------------------------------- Zoho_Import_CoA
    ws = wb.add_worksheet("Zoho_Import_CoA")
    widths = [20, 46, 8, 70, 24, 12, 12, 10, 16, 13, 9, 34]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    for i, h in enumerate(HDR):
        ws.write(0, i, h, F["hdr"])
    ws.freeze_panes(1, 2)
    for j, a in enumerate(data, start=1):
        vals = [a["id"], a["name"], a["code"], a["desc"], a["type"], a["mileage_rate"],
                a["mileage_unit"], a["is_mileage"], a["acct_no"], a["status"], a["currency"], a["parent"]]
        for i, v in enumerate(vals):
            ws.write(j, i, v, F["cell"])
    ws.autofilter(0, 0, len(data), len(HDR) - 1)

    # ---------------------------------------------------------- Changes_Log
    ws = wb.add_worksheet("Changes_Log")
    cols = [("Ref", 7), ("Type", 14), ("Account Code", 9), ("Account Name", 40), ("Change", 34),
            ("Severity", 10), ("Reason / reporting impact", 88), ("Action in Zoho Books", 48)]
    for i, (h, w) in enumerate(cols):
        ws.set_column(i, i, w)
        ws.write(0, i, h, F["hdr"])
    r = 1
    for name, ov in CHANGES.items():
        a0 = next(x for x in data if x["name"] == name) if any(x["name"] == name for x in data) else None
        changes = []
        if "code" in ov:
            changes.append(f"Code {ov['code']}")
        if "parent" in ov:
            changes.append(f"Parent -> {ov['parent'] or '(top level)'}")
        if ov.get("status"):
            changes.append(f"Status -> {ov['status']}")
        if "parent_children" in ov:
            changes.append("Re-parent " + ", ".join(ov["parent_children"]))
        ws.write(r, 0, f"CO-{r:03d}", F["cell"])
        ws.write(r, 1, "Fix", F["cell"])
        ws.write(r, 2, a0["code"] if a0 else "", F["cell"])
        ws.write(r, 3, name, F["wrap"])
        ws.write(r, 4, "; ".join(changes), F["wrap"])
        ws.write(r, 5, ov["severity"],
                 F["red"] if ov["severity"] == "Blocking" else F["amber"] if ov["severity"] == "High" else F["cell"])
        ws.write(r, 6, ov["reason"], F["wrap"])
        ws.write(r, 7, ov["action"], F["wrap"])
        r += 1
    for a in added:
        reason, action = next((x[5], "Create new account") for x in NEW_ACCOUNTS if x[1] == a["name"])
        ws.write(r, 0, f"CO-{r:03d}", F["cell"])
        ws.write(r, 1, "New account", F["cell"])
        ws.write(r, 2, a["code"], F["cell"])
        ws.write(r, 3, a["name"], F["wrap"])
        ws.write(r, 4, f"Added under {a['parent'] or '(top level)'}", F["wrap"])
        ws.write(r, 5, "Addition", F["green"])
        ws.write(r, 6, reason, F["wrap"])
        ws.write(r, 7, action, F["wrap"])
        r += 1
    for name in REVENUE_CHILDREN_IMPORT:
        pass
    r += 1
    ws.write(r, 0, "", F["cell"])
    ws.write(r, 1, "Note", F["bold"])
    ws.write(r, 2, "", F["cell"])
    ws.write(r, 3, "Revenue accounts 4011-4024 and 4031-4034", F["wrap"])
    ws.write(r, 4, "Parent -> Income / Sales (4000)", F["wrap"])
    ws.write(r, 5, "Medium", F["cell"])
    ws.write(r, 6, "The revenue block was flat: only 4010 and 4030 sat under the 'Income / Sales' header, so the P&L "
                   "never totalled turnover from the ledger.", F["wrap"])
    ws.write(r, 7, "Set parent to 'Income / Sales' for each service-line revenue account.", F["wrap"])
    r += 1
    for name, (typ, why) in TYPE_BLOCKED.items():
        ws.write(r, 0, f"CO-{r:03d}", F["cell"])
        ws.write(r, 1, "Not actioned", F["cell"])
        ws.write(r, 2, "", F["cell"])
        ws.write(r, 3, name, F["wrap"])
        ws.write(r, 4, "Left as is", F["wrap"])
        ws.write(r, 5, "Info", F["amber"])
        ws.write(r, 6, why, F["wrap"])
        ws.write(r, 7, "No action - documented limitation.", F["wrap"])
        r += 1
    ws.freeze_panes(1, 0)

    # ------------------------------------------------------- Validation_Tests
    ws = wb.add_worksheet("Validation_Tests")
    cols = [("Test #", 7), ("Structural test", 46), ("Why it matters for reporting", 70),
            ("Original CoA", 16), ("Improved CoA", 16), ("Status", 10), ("Detail", 60)]
    for i, (h, w) in enumerate(cols):
        ws.set_column(i, i, w)
        ws.write(0, i, h, F["hdr"])
    tests = VALIDATION
    for j, t in enumerate(tests, start=1):
        fmt = F["green"] if t[5] == "PASS" else F["amber"] if t[5] in ("FIXED", "OK") else F["red"]
        for i, v in enumerate(t):
            ws.write(j, i, v, fmt if i == 5 else F["wrap"] if i in (1, 2, 6) else F["cell"])
    ws.freeze_panes(1, 0)

    # ------------------------------------------------------ Legacy_to_New_Map
    ws = wb.add_worksheet("Legacy_to_New_Map")
    cols = [("Legacy code", 10), ("Legacy account (FY2024-25 controlled chart)", 52), ("New code(s)", 22),
            ("New account(s)", 52), ("Mapping basis / warning", 78),
            ("FY2024-25 bank lines in legacy account", 14), ("Disposition", 18)]
    for i, (h, w) in enumerate(cols):
        ws.set_column(i, i, w)
        ws.write(0, i, h, F["hdr"])
    r = 1
    for lc, ln, nc, nn, basis, disp in LEGACY_MAP:
        cnt = LEG_CNT.get(lc, 0)
        ws.write(r, 0, lc, F["cell"])
        ws.write(r, 1, ln, F["wrap"])
        ws.write(r, 2, nc, F["cell"])
        ws.write(r, 3, nn, F["wrap"])
        ws.write(r, 4, basis, F["wrap"] if cnt < 100 else F["amber"])
        ws.write(r, 5, cnt, F["int"])
        ws.write(r, 6, disp, F["wrap"])
        r += 1
    ws.write(r, 0, "TOTAL", F["bold"])
    ws.write(r, 5, sum(LEG_CNT.values()), F["int"])
    ws.write(r, 1, "1,226 valued bank transactions in the FY2024-FY2025 register", F["wrap"])

    # ---------------------------------------------------------- Reporting_Map
    ws = wb.add_worksheet("Reporting_Map")
    cols = [("Code", 8), ("Account", 48), ("Zoho Account Type", 22), ("Parent", 30),
            ("IFRS statement line", 54), ("Note / disclosure", 14), ("Tax treatment", 60),
            ("Cash-flow category", 16), ("Status", 10)]
    for i, (h, w) in enumerate(cols):
        ws.set_column(i, i, w)
        ws.write(0, i, h, F["hdr"])
    ws.freeze_panes(1, 2)
    for j, a in enumerate(data, start=1):
        line, note, cf = fs_line(a)
        vals = [a["code"], a["name"], a["type"], a["parent"], line, note, tax_tag(a), cf, a["status"]]
        for i, v in enumerate(vals):
            ws.write(j, i, v, F["red"] if (i == 8 and v == "Inactive") else F["cell"])
    ws.autofilter(0, 0, len(data), len(cols) - 1)

    # ------------------------------------------------------------ Gap_Analysis
    ws = wb.add_worksheet("Gap_Analysis")
    cols = [("Reporting requirement", 44), ("Source", 26), ("Chart coverage BEFORE", 46),
            ("Gap / risk", 66), ("Action in improved CoA", 46), ("Status", 12)]
    for i, (h, w) in enumerate(cols):
        ws.set_column(i, i, w)
        ws.write(0, i, h, F["hdr"])
    for j, row in enumerate(GAPS, start=1):
        fmt = F["green"] if row[5] == "Closed" else F["amber"]
        for i, v in enumerate(row):
            ws.write(j, i, v, fmt if i == 5 else F["wrap"])

    # ------------------------------------------------ Open_Balances_1Jan2026
    ws = wb.add_worksheet("Open_Balances_1Jan2026")
    ws.set_column("A:A", 10); ws.set_column("B:B", 46); ws.set_column("C:C", 18)
    ws.set_column("D:D", 18); ws.set_column("E:E", 84)
    ws.write("A1", "OPENING BALANCES AT 1 JANUARY 2026 - BANK-EVIDENCED (VIEW B) BASIS", F["title"])
    ws.write("A2", "Trial balance of the corrected statement of financial position at 31 December 2025, restated into the "
                   "improved chart. Load the BLOCK A figures in Zoho Books > Accountant > Opening Balances (or one opening "
                   "journal dated 01-Jan-2026). BLOCK B is the alternative, audit-friendly route: load the AFS as issued "
                   "and post one prior-period adjustment journal (IAS 8) dated 01-Jan-2026.", F["sub"])

    def block(row, title, rows, total_label):
        ws.write(row, 0, title, F["bold"]); row += 1
        for i, h in enumerate(["Code", "Account", "Debit (NGN)", "Credit (NGN)", "Source / note"]):
            ws.write(row, i, h, F["hdr"])
        row += 1
        first = row
        td = tc = 0.0
        for code, name, dr, cr, note in rows:
            ws.write(row, 0, code, F["cell"])
            ws.write(row, 1, name, F["wrap"])
            ws.write(row, 2, dr if dr else None, F["num"])
            ws.write(row, 3, cr if cr else None, F["num"])
            ws.write(row, 4, note, F["wrap"])
            td += dr or 0; tc += cr or 0
            row += 1
        bold = wb.add_format({"bold": True, "font_size": 9, "border": 1, "num_format": "#,##0.00", "valign": "top"})
        ws.write(row, 1, total_label, F["bold"])
        ws.write(row, 2, round(td, 2), bold)
        ws.write(row, 3, round(tc, 2), bold)
        ws.write(row, 4, f"Static check: debits {td:,.2f} / credits {tc:,.2f} / difference {td-tc:,.2f}", F["sub"])
        ws.write(row + 1, 1, "Difference (must be nil):", F["bold"])
        ws.write(row + 1, 2, f"=C{row+1}-D{row+1}", F["num"])
        return row + 3

    r = block(3, "BLOCK A - OPENING TRIAL BALANCE TO LOAD (CORRECTED / BANK-EVIDENCED BASIS)", OPENING,
              "TOTAL - BLOCK A")
    r = block(r + 1, "BLOCK B - PRIOR-PERIOD ADJUSTMENT (POST ON TOP OF THE AFS AS ISSUED, IAS 8)", PP_ADJ,
              "TOTAL - BLOCK B")
    ws.write(r, 1, "BLOCK C - HOW THE TWO ROUTES AGREE", F["bold"])
    ws.write(r + 1, 1, "AFS as issued 31-Dec-2025 (new codes)", F["wrap"])
    ws.write(r + 2, 1, "  Share capital 3011", F["wrap"]); ws.write(r + 2, 3, 1000000.00, F["num"])
    ws.write(r + 3, 1, "  Retained earnings 3090", F["wrap"]); ws.write(r + 3, 3, 2197911.00, F["num"])
    ws.write(r + 4, 1, "  Add: prior-period adjustment (Block B)", F["wrap"]); ws.write(r + 4, 3, 4897706.58, F["num"])
    ws.write(r + 5, 1, "  = Retained earnings per Block A", F["wrap"])
    ws.write(r + 5, 3, "=D{}+D{}".format(r + 4, r + 5), F["num"])
    ws.write(r + 6, 1, "  Block A retained earnings (3090)", F["wrap"]); ws.write(r + 6, 3, 7095617.58, F["num"])
    ws.write(r + 7, 1, "Opening balance total (Block A)", F["wrap"]); ws.write(r + 7, 2, 15105766.79, F["num"])

    wb.close()
    print("wrote", XLSX)


# revenue children list duplicated for the changes log loop above
REVENUE_CHILDREN_IMPORT = []

VALIDATION = [
    (1, "Unique account codes", "A duplicate code makes the trial balance ambiguous and Zoho rejects the import.",
     "1 duplicate", "0 duplicates", "FIXED",
     "2030 was used twice: 'Retained Earnings' (equity) and 'Customer / Project Deposits' (liability). Retained earnings "
     "renumbered to 3090."),
    (2, "Every account has a code", "Uncoded accounts cannot be referenced reliably in an import or a formula.",
     "1 blank", "0 blank", "FIXED", "'License Fees' had no code; assigned 6419 (Professional & Regulatory)."),
    (3, "Every parent named exists", "A dangling parent silently flattens the report hierarchy.", "0", "0", "PASS", ""),
    (4, "No inverted hierarchy", "An account cannot be its own ancestor; Zoho displays the tree incorrectly.",
     "1 inverted", "0 inverted", "FIXED",
     "3060 'Owner's Equity' was a child of 3061 'Director Equity Injections' - reversed."),
    (5, "Parent and child share the same account type", "Zoho Books rule - a type mismatch cannot be imported.",
     "0", "0", "PASS",
     "3 desired groupings are impossible for this reason and are documented (7023, 5071, 2999, plus Stock and Cash "
     "accounts which Zoho keeps flat)."),
    (6, "Revenue accounts subordinate to the revenue header",
     "Without this the P&L revenue total is not driven by the ledger.", "2 of 27", "27 of 27", "FIXED",
     "Only 4010 and 4030 were under 'Income / Sales'."),
    (7, "Expense headers have children (orphan test)", "Orphan detail accounts never roll into their group total.",
     "19 orphans", "0 orphans", "FIXED",
     "Included the highest-volume expense account in the ledger (legacy 6050, 607 of 1,226 lines) with no parent."),
    (8, "Receivables subordinated to the AR control", "The balance sheet would otherwise show the control and the "
     "detail side by side.", "0 of 2", "2 of 2", "FIXED", "1041 and 1042 re-parented to 1040."),
    (9, "Payables subordinated to the AP control", "As above.", "3 of 4", "4 of 4", "FIXED",
     "2033 retentions payable re-parented to 2010 and renumbered 2015."),
    (10, "Tax accounts grouped under a single control", "VAT and WHT balances must net to the FIRS return.",
     "4 loose", "4 grouped", "FIXED", "1081-1084 re-parented to 1080; 2041 nested under 2042."),
    (11, "IFRS 15 contract asset and contract liability pairs present", "Required for over-time recognition and advances.",
     "Both present", "Both present", "PASS", ""),
    (12, "IAS 12 deferred tax accounts present", "The AFS has no IAS 12 disclosure at all.", "None", "DTA 1190 / DTL 2096",
     "FIXED", "Both added."),
    (13, "IAS 37 provisions present", "Penalties and warranty exposure must be recognised, not just described.",
     "None", "2065, 2066", "FIXED", "Both added."),
    (14, "Current / non-current liability split possible", "IAS 1.60-76 requires the split.", "No non-current section",
     "2096, 2097, 2098", "FIXED", "Every borrowing account in the original chart was current."),
    (15, "IFRS 16 lease accounts present", "Office lease charged to P&L with no IFRS 16 assessment.",
     "None", "1180, 2098, 6822, 6917", "FIXED", "Full pair plus depreciation and interest."),
    (16, "Statutory payroll liabilities complete", "PAYE, pension, ITF, NSITF and NHF are all Nigerian obligations.",
     "3 of 6", "6 of 6", "FIXED", "ITF, NSITF and NHF payables added."),
    (17, "Bank-charge analysis adequate", "Bank charges were 49.5% of all bank lines and were not separable.",
     "2 accounts", "5 accounts", "FIXED",
     "Commission (6915), EMTL (6916), VAT (6914), maintenance (6910) and recoverable bank VAT (1086) now separate."),
    (18, "Overlapping / duplicate-purpose accounts removed", "Two accounts for one concept guarantees a split balance.",
     "7 overlaps", "0 active overlaps", "FIXED",
     "1043, 4020, 4021, 4022, 4023, 6611, 3052 deactivated with a documented replacement."),
    (19, "Suspense and uncategorised accounts available and hooked up", "Every unclassified line needs a home that "
     "reports cleanly.", "Present but orphaned", "Grouped under 7000 / 1990",
     "FIXED", "7999 re-parented; 7013 rounding added so 7012 is not used as a plug."),
    (20, "Total accounts", "Depth of analysis available to the reporting pack.", "275", "302", "OK", ""),
]

GAPS = [
    ("Statement of financial position - current/non-current split (IAS 1.60-76)", "IAS 1",
     "No 'Non Current Liability' account type used at all; all borrowings were current.",
     "Long-term debt and deferred tax could not be disclosed as non-current.",
     "2096 Deferred Tax Liability, 2097 Long-term Borrowings, 2098 Lease Liabilities added as Non Current Liability.",
     "Closed"),
    ("Deferred tax note (IAS 12)", "IAS 12 / AFS gap",
     "No deferred tax asset or liability account.",
     "The FY2024/FY2025 AFS contain no tax note at all; even a correct provision could not be posted.",
     "1190 Deferred Tax Asset, 2096 Deferred Tax Liability.", "Closed"),
    ("Provisions (IAS 37)", "IAS 37",
     "No provision accounts; only expense accounts for rectification.",
     "Warranty and penalty exposure (VAT non-registration, unpaid minimum tax) could not be recognised.",
     "2065 Warranty / Defect Provision; 2066 Provision for Claims, Fines & Penalties.", "Closed"),
    ("Leases (IFRS 16)", "IFRS 16",
     "No right-of-use asset, lease liability, ROU depreciation or lease interest.",
     "N200,000 p.a. office rent expensed with no IFRS 16 assessment; the policy note is silent.",
     "1180, 2098, 6822, 6917.", "Closed"),
    ("Related-party disclosures (IAS 24)", "IAS 24",
     "One account (1045) exists but no impairment account for related-party balances.",
     "N14,223,846.92 director receivable with recoverability in doubt and no mechanism to impair it.",
     "7024 ECL - Related Party Receivables added.", "Closed"),
    ("Share capital movements (CAMA 2020 / IAS 1)", "CAMA",
     "No account for share application money pending allotment.",
     "The 2024/2025 Ime Ubongabasi share event was mis-posted to revenue and capital.",
     "2057 Share Application Money Pending Allotment added.", "Closed"),
    ("Statutory payroll liabilities", "PAYE / Pension / ITF / NSITF / NHF",
     "PAYE, pension and staff deductions only.",
     "ITF, NSITF and NHF obligations had no liability account and would be expensed on payment.",
     "2074, 2075, 2076 added.", "Closed"),
    ("Dividends and distributions", "CAMA s.379",
     "No dividend payable.",
     "A declared but unpaid dividend could not be shown as a liability.",
     "2085 Dividends Payable added.", "Closed"),
    ("Revaluation / fair value reserve", "IAS 16.39 / IAS 38",
     "No reserve, although the AFS presents a 'Revaluation gain' line in OCI.",
     "The OCI line had no equity counterpart, so the statement of changes in equity could not be produced.",
     "3072 Revaluation / Fair Value Reserve added.", "Closed"),
    ("Overdraft and going-concern disclosure", "IAS 1.25-26",
     "No bank overdraft account type.",
     "Recurring overdrafts are the stated going-concern risk but the balance sheet cannot show them gross.",
     "1029 Bank Overdraft - Access Bank added.", "Closed"),
    ("Collections clearing (gateway / POS / wallet)", "Cut-off and accuracy",
     "Sales collection and payment gateway clearing exist, but no mobile-wallet or interbank clearing.",
     "2025 volumes moved to 'MOBILE TRF' channels; cash at cut-off is overstated without a clearing account.",
     "1034 Mobile Money/Wallet Clearing; 1035 Interbank Transfer Clearing added.", "Closed"),
    ("Recoverable bank-charge VAT", "VAT Act 2023 s.11",
     "One general 'VAT Input Receivable' account.",
     "VAT on bank charges (258 historical lines) was expensed; recoverable once registered.",
     "1086 Input VAT - Bank Charges & Fees added.", "Closed"),
    ("Bank-charge analysis", "Reporting quality",
     "Legacy 6050 absorbed 607 of 1,226 transactions (49.5%).",
     "The largest expense line in the business was unanalysable; the AFS shows bank charges of only N38,498 for 2025.",
     "6915 Transfer Commission, 6916 EMTL, 6914 VAT, 1086 recoverable split out.", "Closed"),
    ("Revenue disaggregation (IFRS 15.114)", "IFRS 15",
     "27 revenue accounts, but 4 were split by customer type and 1 by billing pattern.",
     "Revenue by contract type could not be aggregated without double counting.",
     "4020-4023 deactivated; service-line accounts nested under 4000.", "Closed"),
    ("Cost of sales analysis", "IFRS 15 / IAS 1",
     "42 COGS accounts, but 226 historical lines were coded to a single 'materials' account.",
     "No visibility of materials vs subcontractor margin - the key profitability driver for a fit-out business.",
     "Materials (5011-5019) and subcontractors (5021-5029) separated by rule; 5038 project insurance added.",
     "Closed"),
    ("Sales commission", "Reporting",
     "No commission expense account.",
     "Referral payments were forced into other buckets (2025 bank evidence).", "6317 added.", "Closed"),
    ("Related-party loan vs director loan vs director current account", "IAS 24",
     "2051, 2053, 2054, 2061, 2062 overlap.",
     "Risk of the same balance being split across four accounts.",
     "Documented posting rule: 1045/2061 for trading flows with the director, 2053 for cash loans, 2054 for other related parties.",
     "Open - policy"),
    ("Multi-currency", "IAS 21",
     "All accounts NGN, single currency.",
     "No exposure yet; foreign-currency supplier contracts (imported furniture) would need revaluation accounts.",
     "1132 and 7023 exist; add a 'Foreign Currency Revaluation' equity account only when FX transactions start.",
     "Open - monitor"),
    ("E-invoicing (FIRS/NRS)", "Tax administration",
     "No tax-specific configuration in the chart.",
     "Zoho's e-invoicing module is not enabled for Nigeria; invoices must be transmitted through the FIRS portal.",
     "Configure via integration; no chart change required.", "Open - integration"),
]

OPENING = [
    ("1021", "Access Bank - A/C 1854793470", 24194.10, None, "Bank statement closing balance at 31-Dec-2025 (ties to the kobo)."),
    ("1091", "Furniture & Fittings", 165000.00, None, "2025 additions (N165,000 bank-evidenced; includes UIUX branding system N100,000)."),
    ("1093", "Computers & IT Equipment", 360053.75, None, "2024 additions, incl. IROANUSI phones N360,053.75. Reallocate to 1092 if the office-equipment invoices support it."),
    ("1101", "Accumulated Depreciation - Furniture & Fittings", None, 56184.75, "AFS accumulated depreciation N178,822 allocated pro rata to cost (furniture 31.4%)."),
    ("1103", "Accumulated Depreciation - Computers & IT", None, 122637.25, "Balance of the AFS N178,822. Note: strictly 20% x bank cost = N177,021.50; the N1,800.50 difference is the AFS cost gap (N9,000.25)."),
    ("1045", "Director Current Account Receivable - Esther Akin-Ajayi", 14223846.92, None, "Cumulative net outflows to the MD/owner. IAS 24 disclosure; ECL under review (7024)."),
    ("1046", "Other Receivables", 332672.02, None, "Unverified POS items and client credits: N259,505, N24,500, N30,000, N18,667."),
    ("2011", "Trade Payables - Suppliers", None, 1335000.00, "AFS 'other payables' (no bank payment trace; invoices requested)."),
    ("2023", "Accrued Professional Fees", None, 300000.00, "AFS accrued professional fee N300,000."),
    ("2031", "Contract Liabilities - Project Deposits", None, 3460000.00, "2024 client advances unperformed at both year ends (IFRS 15.81)."),
    ("2045", "Income Tax Payable (CIT / minimum tax)", None, 188852.14, "Minimum tax 0.5% gross income, unpaid; penalty and interest accruing."),
    ("2046", "Development Levy Payable (formerly TEL)", None, 85475.07, "Tertiary Education Levy accrued to 31-Dec-2025; carried here pending FIRS agreement."),
    ("2051", "Short-term Loan - Individual Lender", None, 230000.00, "Olajide facility received July 2024, unrepaid."),
    ("2052", "Short-term Working Capital Loan", None, 1230000.00, "N2.5m 10-day facility: N1.49m repaid, N1.23m outstanding. Interest N250,000 per terms - accrual to confirm."),
    ("3011", "Ordinary Share Capital", None, 1000000.00, "1,000,000 ordinary shares of N1.00."),
    ("3071", "Capital Reserve", None, 2000.00, "Net effect of the Ime share buy-back (N100,000 funded by Esther / N99,000 refunded)."),
    ("3090", "Retained Earnings", None, 7095617.58, "Cumulative corrected result (opening 2023 stub deficit N998,740 per bank)."),
]

PP_ADJ = [
    ("1041", "Trade Receivables - Project Clients", None, 4453485.00,
     "Reverse unsupported trade receivables (no billings or delivery evidenced)."),
    ("1045", "Director Current Account Receivable", 14223846.92, None,
     "Recognise related-party outflows previously absorbed in expenses (IAS 24)."),
    ("1046", "Other Receivables", 332672.02, None, "Sundry POS and client credits identified in 2025."),
    ("2031", "Contract Liabilities - Project Deposits", None, 3460000.00,
     "Reclassify client advances out of revenue (IFRS 15.81)."),
    ("2045", "Income Tax Payable (CIT / minimum tax)", None, 188852.14, "Minimum tax not provided in the AFS."),
    ("2046", "Development Levy Payable (formerly TEL)", None, 85475.07, "TEL not provided in the AFS."),
    ("2051", "Short-term Loan - Individual Lender", None, 230000.00, "Unrecorded borrowing (Olajide)."),
    ("2052", "Short-term Working Capital Loan", None, 1230000.00, "Unrecorded borrowing (N2.5m facility)."),
    ("3071", "Capital Reserve", None, 2000.00, "Ime share buy-back net contribution."),
    ("1091", "Furniture & Fittings (cost adjustment)", None, 9000.25,
     "Reduce PPE cost from AFS N534,054 to bank-evidenced N525,053.75 (spread across classes; shown here against "
     "furniture for simplicity)."),
    ("1021", "Access Bank - A/C 1854793470", 0.10, None, "Cash rounding to the exact statement balance."),
    ("3090", "Retained Earnings", None, 4897706.58, "Balancing prior-period adjustment (IAS 8.42)."),
]

if __name__ == "__main__":
    build_workbook()
    data, _, _ = build()
    with open(CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(HDR)
        for a in data:
            w.writerow([a["id"], a["name"], a["code"], a["desc"], a["type"], a["mileage_rate"],
                        a["mileage_unit"], a["is_mileage"], a["acct_no"], a["status"], a["currency"], a["parent"]])
    print("wrote", CSV)
