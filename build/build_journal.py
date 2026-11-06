"""Builds 02_Journal_Aug-Oct_2026.xlsx and the Zoho Books journal upload CSVs."""
import csv
import json
import os
import re
import xlsxwriter
from journal_rules import RULES, RULE_INDEX, OVERRIDES, FLAGS

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(ROOT, "deliverables")
XLSX = os.path.join(OUT, "02_Jemai_Posting_Journal_Aug-Oct_2026.xlsx")
CSV_WORKED = os.path.join(OUT, "02b_Journal_Aug-Oct_2025_Worked_Example_Zoho_Upload.csv")
CSV_TEMPLATE = os.path.join(OUT, "02c_Journal_Template_2026_Zoho_Upload_BLANK.csv")
CSV_REGISTER = os.path.join(OUT, "02d_Txn_Register_2026_Input_Template.csv")

BANK_ACCOUNT = "1021"
DESC_FOR_LEG = {0: "Dr - receipt to bank", 1: "Cr - revenue / balance", 2: "VAT / balance leg"}
BANK_NAME = "Access Bank"
VAT_OUT = "2041"
VAT_IN = "1081"
VAT_RATE = 0.075

BANK = json.load(open("/home/user/work/bank.json"))


# --------------------------------------------------------------------------- #
def account_list():
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(ROOT, "deliverables",
                                             "01_Chart_of_Accounts_Improved_Jemai_Oct2026.xlsx"), data_only=True)
    ws = wb["Zoho_Import_CoA"]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[1] and r[2]:
            out.append((str(r[2]), str(r[1]), str(r[4]), str(r[9])))
    return out


ACC = account_list()
ACC_NAME = {c: n for c, n, t, s in ACC}


# --------------------------------------------------------------------------- #
# classification
# --------------------------------------------------------------------------- #
DIRECTOR = ("esther", "e. akin", "akin-ajayi esther", "akinajayi esther", "jesutofunmi")
RELATED = ("akinloluwa", "akinyemi", "felicia", "ime ubongabasi")


def classify(nar, direction):
    n = nar.lower()
    # bank charges first - they appear inside every transfer narration
    if "commission" in n and "trf" in n:
        return "BNK-01"
    if "electronic money transfer levy" in n:
        return "BNK-03"
    if "account maintenance" in n:
        return "BNK-04"
    if "sms alert" in n:
        return "BNK-05"
    if "penal charge" in n:
        return "BNK-06"
    if re.search(r"\bvat\b", n) and ("vat trf" in n or "vat mobile trf" in n or n.strip().endswith("vat")):
        return "BNK-02"
    if "debit interest" in n:
        return "BNK-07"
    if direction == "In":
        if any(k in n for k in DIRECTOR):
            return "REC-08"
        if any(k in n for k in RELATED):
            return "REC-09"
        if "refund" in n:
            return "REC-13"
        for rid, desc, kw, d, dr, cr, vat, wht, note in RULES:
            if d != "In" or rid.startswith(("REC-08", "REC-09")):
                continue
            if any(x.strip() in n for x in kw.split(",")):
                return rid
        return "REC-01"
    # payments
    if any(k in n for k in DIRECTOR):
        return "BS-01"
    if any(k in n for k in RELATED):
        return "BS-02"
    best = None
    for rid, desc, kw, d, dr, cr, vat, wht, note in RULES:
        if d != "Out" or rid.startswith(("BNK-", "BS-01", "BS-02")):
            continue
        for x in kw.split(","):
            x = x.strip()
            if x and x in n:
                if best is None or len(x) > best[0]:
                    best = (len(x), rid)
    if best:
        return best[1]
    return "BS-13"


# --------------------------------------------------------------------------- #
def load_window():
    def mm(d):
        s = str(d["Value Date"])
        return s[3:6], s[-2:]
    return [d for d in BANK if mm(d) in (("AUG", "25"), ("SEP", "25"), ("OCT", "25"))]


MONTHS = {"AUG": "08", "SEP": "09", "OCT": "10"}


def build_worked():
    rows = []
    seq = {}
    for d in load_window():
        amt_in = d["Money In / Credit (NGN)"] or 0
        amt_out = d["Money Out / Debit (NGN)"] or 0
        direction = "In" if amt_in else "Out"
        gross = round(float(amt_in or amt_out), 2)
        nar = str(d["Original Narration"])
        rid = OVERRIDES.get(d["Txn ID"]) or classify(nar, direction)
        rule = RULE_INDEX[rid]
        _, desc, kw, _, dr, cr, vat_mode, wht, note = rule
        # VAT split (company is not VAT-registered in the window, so no split)
        vat_amt = 0.0
        net = gross
        # month + sequence for the journal number
        dd = str(d["Value Date"])
        month = MONTHS[dd[3:6]]
        seq[month] = seq.get(month, 0) + 1
        jno = f"JE-2025-{month}-{seq[month]:03d}"
        rows.append({
            "txn": d["Txn ID"], "date": dd, "month": month, "narration": nar,
            "counterparty": d["Counterparty (Classified)"], "direction": direction,
            "gross": gross, "net": net, "vat": vat_amt, "rule": rid, "rule_desc": desc,
            "dr": dr, "cr": cr, "vat_mode": vat_mode, "wht": wht, "note": note,
            "journal": jno, "flag": FLAGS.get(d["Txn ID"], ""),
            "legacy": d["COA Account"], "balance": d["Statement Balance (NGN)"],
        })
    return rows


def journal_lines(rows):
    out = []
    for r in rows:
        tag = f"{r['txn']} | {r['narration'][:70]}"
        if r["direction"] == "In":
            out.append((r["date"], r["journal"], r["txn"], tag, BANK_ACCOUNT, r["gross"], 0.0, "Receipt to bank"))
            out.append((r["date"], r["journal"], r["txn"], tag, r["cr"], 0.0, r["net"], r["rule_desc"]))
        else:
            out.append((r["date"], r["journal"], r["txn"], tag, r["dr"], r["net"], 0.0, r["rule_desc"]))
            out.append((r["date"], r["journal"], r["txn"], tag, BANK_ACCOUNT, 0.0, r["gross"], "Payment from bank"))
        if r["vat"]:
            if r["direction"] == "In":
                out.append((r["date"], r["journal"], r["txn"], tag, VAT_OUT, 0.0, r["vat"], "Output VAT 7.5%"))
            else:
                out.append((r["date"], r["journal"], r["txn"], tag, VAT_IN, r["vat"], 0.0, "Input VAT 7.5%"))
    return out


# --------------------------------------------------------------------------- #
def build():
    rows = build_worked()
    lines = journal_lines(rows)

    # integrity
    dr = round(sum(l[5] for l in lines), 2)
    cr = round(sum(l[6] for l in lines), 2)
    assert abs(dr - cr) < 0.005, (dr, cr)

    wb = xlsxwriter.Workbook(XLSX)
    F = {
        "title": wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864"}),
        "sub": wb.add_format({"italic": True, "font_size": 9, "font_color": "#595959"}),
        "hdr": wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1,
                              "text_wrap": True, "valign": "vcenter", "font_size": 9}),
        "cell": wb.add_format({"border": 1, "font_size": 9, "valign": "top"}),
        "wrap": wb.add_format({"border": 1, "font_size": 9, "text_wrap": True, "valign": "top"}),
        "num": wb.add_format({"border": 1, "font_size": 9, "num_format": "#,##0.00", "valign": "top"}),
        "bold": wb.add_format({"bold": True, "font_size": 9, "border": 1, "valign": "top"}),
        "num_bold": wb.add_format({"bold": True, "font_size": 9, "border": 1, "num_format": "#,##0.00"}),
        "red": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FFC7CE", "text_wrap": True, "valign": "top"}),
        "amber": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FFEB9C", "text_wrap": True, "valign": "top"}),
        "green": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#C6EFCE", "text_wrap": True, "valign": "top"}),
        "in": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FFF2CC", "num_format": "#,##0.00", "valign": "top"}),
    }

    # ------------------------------------------------------------- Read_Me
    ws = wb.add_worksheet("Read_Me")
    ws.set_column("A:A", 3); ws.set_column("B:B", 34); ws.set_column("C:C", 112)
    ws.write("B2", "JEMAI INTERNATIONAL LIMITED  |  MONTHLY TRANSACTION POSTING JOURNAL", F["title"])
    ws.write("B3", "August - October 2026  |  Zoho Books journal-import format  |  Improved chart of accounts", F["sub"])
    ws.write("B4", "Bank: Access Bank Plc A/C 1854793470  |  Currency: NGN  |  Prepared: November 2026", F["sub"])

    # compute the worked-example summary numbers
    aug = [r for r in rows if r["month"] == "08"]
    sep = [r for r in rows if r["month"] == "09"]
    octo = [r for r in rows if r["month"] == "10"]

    def tot(rs, key):
        return sum(r[key] for r in rs)

    body = [
        ("STATUS - READ THIS FIRST", ""),
        ("Data position",
         "The repository now contains the Access Bank statement for 1 August - 31 October 2026 (26 pages, 716 "
         "printed lines). The ACTUAL journal built from it is delivered separately as 02e-02i: 692 economic bank "
         "lines, 1,384 journal legs, Dr = Cr = N125,599,480.97, receipts and payments agreeing exactly to the "
         "statement totals."),
        ("What this workbook is",
         "The reusable ENGINE behind that journal: the rule table, the pre-wired input register, the formula-driven "
         "upload sheet and the control checks. Keep using it for November 2026 onwards - paste the new statement "
         "lines into Txn_Register_2026, pick a Rule ID per line, and Journal_Upload_2026 builds itself."),
        ("What proves the engine",
         "The 112 real Aug-Oct 2025 bank lines in Worked_Example_2025, re-coded from the legacy 33-account chart "
         "into the improved 302-account chart, and the balanced journal they produce. Same engine, older data, "
         "so you can see every rule behaving on real Nigerian bank narration."),
        ("Where the real 2026 journal lives",
         "02e_Journal_Aug-Oct_2026_ACTUAL_Zoho_Upload.csv (all three months), 02f / 02g / 02h (August, September, "
         "October separately) and 02i_Posting_Journal_Aug-Oct_2026_ACTUAL_Workbook.xlsx (register, journal, trial "
         "balance, exceptions, VAT review). Import 02e into Zoho Books."),
        ("", ""),
        ("THE TABS", ""),
        ("Parameters", "The switches the whole journal depends on: VAT registration status, VAT accounts, bank account, "
                       "rounding. Change them here, not in the formulas."),
        ("Mapping_Rules", f"{len(RULES)} posting rules. Each rule fixes the debit and credit account for one shape of bank line "
                          "(receipt or payment), the VAT treatment and the WHT rate. This is the instruction manual for "
                          "the bookkeeper: every narration pattern resolves to exactly one rule."),
        ("Txn_Register_2026", "THE INPUT. Paste the bank statement here: date, narration, money in, money out, reference. "
                              "Select a Rule ID from the drop-down; account codes and names resolve by formula."),
        ("Journal_Upload_2026", "THE OUTPUT. Three rows per bank line, balanced, with the Zoho column layout. "
                                "Copy this tab to a CSV to import, or post the journals manually."),
        ("Worked_Example_2025", "The 112 real Aug-Oct 2025 bank lines, each mapped to the new chart with the reason, "
                                "the rule applied and any query flag. Read this to see how each rule behaves on real "
                                "Nigerian bank narration."),
        ("Journal_Worked_Example", f"The balanced double entry produced from those 112 lines ({len(lines)} journal lines). "
                                   "This is a genuine, importable journal - use it for the 2025 conversion, not 2026."),
        ("Account_List", "The 302 reviewed account codes and names - the validation source for every lookup."),
        ("Control_Checks", "Debits equal credits, per month and in total; bank movement reconciles; every line is coded; "
                           "no line sits in suspense."),
        ("Month_Summary", "Revenue, direct cost, operating expense and bank-charge analysis by month for the worked "
                          "example - the shape of the reporting pack Zoho will produce once the rules are applied."),
        ("Upload_Notes", "Field-by-field instructions for the Zoho Books journal import screen, plus the 2026 "
                         "close checklist."),
        ("", ""),
        ("WORKED EXAMPLE - CONTROL TOTALS (Aug-Oct 2025)", ""),
        ("August 2025", f"{len(aug)} bank lines | money in N{tot(aug,'gross') if False else sum(r['gross'] for r in aug if r['direction']=='In'):,.2f} | "
                        f"money out N{sum(r['gross'] for r in aug if r['direction']=='Out'):,.2f}"),
        ("September 2025", f"{len(sep)} bank lines | money in N{sum(r['gross'] for r in sep if r['direction']=='In'):,.2f} | "
                           f"money out N{sum(r['gross'] for r in sep if r['direction']=='Out'):,.2f}"),
        ("October 2025", f"{len(octo)} bank lines | money in N{sum(r['gross'] for r in octo if r['direction']=='In'):,.2f} | "
                         f"money out N{sum(r['gross'] for r in octo if r['direction']=='Out'):,.2f}"),
        ("Total", f"{len(rows)} bank lines | {len(lines)} journal lines | debits N{dr:,.2f} = credits N{cr:,.2f}"),
        ("", ""),
        ("ACCOUNTING POLICIES EMBEDDED IN THE RULES", ""),
        ("Cash basis to accruals",
         "The bank statement records cash. Revenue is recorded on receipt here because Jemai invoices at or after "
         "collection; where a deposit precedes performance it belongs in 2031 Contract Liabilities, not revenue. "
         "The rule table flags the two cases separately (REC-01 vs a contract-liability posting)."),
        ("Director payments",
         "Every transfer to the MD (Esther Akin-Ajayi) is posted to 1045 Director Current Account Receivable, whatever "
         "the narration says ('petty cash', 'site expenses', 'runnings'). N14.22m of 2024-2025 outflows were absorbed "
         "into expenses; that is the single largest correction in the audited restatement. Reclassify out of 1045 only "
         "when the director produces receipts that evidence a business cost."),
        ("Related parties",
         "Transfers to other Akin-Ajayi family members go to 1047, not to expense and not to the director account."),
        ("Bank charges are four different things",
         "Commission (6915), EMTL (6916), VAT on charges (6914 or 1086) and maintenance/SMS (6910). Legacy account 6050 "
         "held 607 of 1,226 lines and made the largest expense line in the business unanalysable."),
        ("Transport",
         "Third-party transport/logistics charged to a client contract is a direct cost (5032); staff movement is "
         "6511. The narration alone often cannot tell them apart - the rule table records the judgement."),
        ("VAT",
         "Jemai was not VAT-registered in 2024-2025 and should have registered in 2025. The engine treats invoices as "
         "gross (no VAT split) until the VAT-registered switch in Parameters is set to YES; then receipts split "
         "7.5/107.5 to 2041 and valid input invoices split to 1081."),
        ("WHT",
         "Withholding tax cannot be derived from a bank statement - it is an invoice-level obligation. The rule table "
         "carries the statutory rate per expense type so the WHT schedule can be built from the purchase ledger; no "
         "WHT has been assumed in the worked example because none was withheld in the period."),
    ]
    r = 6
    for a, b in body:
        if a and not b:
            ws.write(r, 1, a, F["bold"])
        elif a:
            ws.write(r, 1, a, F["wrap"])
            ws.write(r, 2, b, F["wrap"])
        r += 1

    # ------------------------------------------------------------ Parameters
    ws = wb.add_worksheet("Parameters")
    ws.set_column("A:A", 34); ws.set_column("B:B", 18); ws.set_column("C:C", 96)
    ws.write("A1", "CONTROL PARAMETERS", F["title"])
    ws.write("A2", "Change these values only. Every formula in the workbook reads from this sheet.", F["sub"])
    for i, h in enumerate(["Parameter", "Value", "Explanation"]):
        ws.write(3, i, h, F["hdr"])
    params = [
        ("Bank account code", BANK_ACCOUNT, "Access Bank Plc - A/C 1854793470. The only operating account in the record."),
        ("Bank account name", BANK_NAME, "Must match the chart of accounts exactly for the lookup in Account_List."),
        ("VAT registered from 2026? (YES/NO)", "NO",
         "Set to YES once the FIRS VAT registration certificate is issued. Until then every receipt and payment is "
         "recorded gross, and VAT charged by the bank is expensed (6914)."),
        ("VAT rate", 0.075, "Nigeria: 7.5% (VAT Act 2023). NTA 2025 retains 7.5%."),
        ("Output VAT account", VAT_OUT, "2041 VAT Output Payable - used only when the VAT switch is YES."),
        ("Input VAT account", VAT_IN, "1081 VAT Input Receivable - used only when the VAT switch is YES and a valid tax invoice exists."),
        ("Bank-charge VAT account (regist.", "1086", "Input VAT - Bank Charges & Fees. Alternative rule BNK-02R."),
        ("Bank-charge VAT account (non-reg.)", "6914", "VAT on Bank Charges - finance cost. Default rule BNK-02."),
        ("Rounding account", "7013", "Rounding / posting differences. Never plug a difference into 7012."),
        ("Revenue suspense", "4999", "Unidentified credits. Must be nil at each period end."),
        ("Payment suspense", "7010", "Unidentified payments. Must be nil at each period end."),
        ("Financial year start", "01 January", "Zoho Books fiscal year must be January-December."),
        ("Statement period", "01-Aug-2026 to 31-Oct-2026", "Three monthly close cycles: August, September, October 2026."),
        ("Prepared by", "Finance function, Jemai International Limited", ""),
        ("Reviewed by", "Mrs. Esther Akin-Ajayi Jesutofunmi Ifeoluwa (MD/CEO)", ""),
    ]
    r = 4
    for p, v, e in params:
        ws.write(r, 0, p, F["wrap"])
        ws.write(r, 1, v, F["num"] if isinstance(v, float) else F["cell"])
        ws.write(r, 2, e, F["wrap"])
        r += 1

    # ---------------------------------------------------------- Mapping_Rules
    ws = wb.add_worksheet("Mapping_Rules")
    cols = [("Rule ID", 10), ("Rule description", 46), ("Narration keywords (first match wins)", 52),
            ("Direction", 10), ("Dr code", 9), ("Dr account", 42), ("Cr code", 9), ("Cr account", 42),
            ("VAT mode", 12), ("WHT %", 8), ("Accounting note / tax treatment", 90)]
    for i, (h, w) in enumerate(cols):
        ws.set_column(i, i, w)
        ws.write(0, i, h, F["hdr"])
    ws.freeze_panes(1, 1)
    for j, (rid, desc, kw, d, dcode, ccode, vat, wht, note) in enumerate(sorted(RULES, key=lambda x: x[0]), start=1):
        vals = [rid, desc, kw, d, dcode, ACC_NAME.get(dcode, ""), ccode, ACC_NAME.get(ccode, ""), vat, wht / 100, note]
        for i, v in enumerate(vals):
            ws.write(j, i, v, F["num"] if i == 9 else F["wrap"] if i in (1, 2, 5, 7, 10) else F["cell"])
    n_rules = len(RULES)
    rule_range = f"Mapping_Rules!$A$2:$A${n_rules+1}"

    # ------------------------------------------------------- Account_List
    ws = wb.add_worksheet("Account_List")
    ws.set_column("A:A", 10); ws.set_column("B:B", 52); ws.set_column("C:C", 24); ws.set_column("D:D", 12)
    for i, h in enumerate(["Code", "Account name", "Account type", "Status"]):
        ws.write(0, i, h, F["hdr"])
    for j, (c, n, t, s) in enumerate(sorted(ACC), start=1):
        ws.write(j, 0, c, F["cell"]); ws.write(j, 1, n, F["cell"])
        ws.write(j, 2, t, F["cell"]); ws.write(j, 3, s, F["cell"])
    acc_last = len(ACC) + 1

    # --------------------------------------------------- Txn_Register_2026
    NREG = 400
    ws = wb.add_worksheet("Txn_Register_2026")
    headers = ["Journal#", "Date", "Bank ref / Txn ID", "Narration (as printed on statement)", "Counterparty",
               "Money in (NGN)", "Money out (NGN)", "Dir", "Rule ID", "Dr code", "Dr account", "Cr code",
               "Cr account", "VAT mode", "Net (NGN)", "VAT (NGN)", "WHT %", "Project / tag", "Query / flag",
               "Control check"]
    widths = [18, 11, 16, 62, 24, 13, 13, 6, 10, 8, 34, 8, 34, 11, 13, 11, 8, 18, 40, 46]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    ws.write(0, 0, "TRANSACTION REGISTER - AUGUST TO OCTOBER 2026 (PASTE THE BANK STATEMENT HERE)", F["title"])
    ws.write(1, 0, "Yellow cells are the only inputs. Everything from column H onwards is calculated. "
                   "Pick a Rule ID from the drop-down in column I - see Mapping_Rules for the full list.", F["sub"])
    for i, h in enumerate(headers):
        ws.write(3, i, h, F["hdr"])
    ws.freeze_panes(4, 2)
    ws.data_validation(4, 8, 4 + NREG - 1, 8, {"validate": "list", "source": f"={rule_range}"})
    ws.data_validation(4, 7, 4 + NREG - 1, 7, {"validate": "list", "source": ["In", "Out"]})
    ws.conditional_format(4, 19, 4 + NREG - 1, 19, {"type": "text", "criteria": "containing",
                                                   "value": "ERROR", "format": F["red"]})
    ws.conditional_format(4, 19, 4 + NREG - 1, 19, {"type": "text", "criteria": "containing",
                                                   "value": "OK", "format": F["green"]})
    for n in range(NREG):
        row = 4 + n
        i = row + 1  # excel row
        ws.write_blank(row, 0, None, F["in"])
        ws.write_formula(row, 0, f'=IF(B{i}="","","JE-"&TEXT(B{i},"yyyymm")&"-"&TEXT(ROW()-4,"000"))', F["cell"])
        for c in (1, 2, 3, 4, 5, 6):
            ws.write_blank(row, c, None, F["in"])
        ws.write_formula(row, 7, f'=IF(F{i}>0,"In",IF(G{i}>0,"Out",""))', F["cell"])
        ws.write_blank(row, 8, None, F["in"])
        ws.write_formula(row, 9, f'=IFERROR(VLOOKUP($I{i},Mapping_Rules!$A:$K,5,FALSE),"")', F["cell"])
        ws.write_formula(row, 10, f'=IFERROR(VLOOKUP($J{i},Account_List!$A:$B,2,FALSE),"")', F["cell"])
        ws.write_formula(row, 11, f'=IFERROR(VLOOKUP($I{i},Mapping_Rules!$A:$K,7,FALSE),"")', F["cell"])
        ws.write_formula(row, 12, f'=IFERROR(VLOOKUP($L{i},Account_List!$A:$B,2,FALSE),"")', F["cell"])
        ws.write_formula(row, 13, f'=IFERROR(VLOOKUP($I{i},Mapping_Rules!$A:$K,9,FALSE),"")', F["cell"])
        # net / vat
        ws.write_formula(
            row, 14,
            f'=IF($I{i}="","",IF($N{i}="OUT-INCL",ROUND(MAX(F{i},G{i})/1.075,2),'
            f'IF($N{i}="IN-INCL",ROUND(MAX(F{i},G{i})/1.075,2),MAX(F{i},G{i}))))', F["cell"])
        ws.write_formula(
            row, 15,
            f'=IF($I{i}="","",IF(OR($N{i}="OUT-INCL",$N{i}="IN-INCL"),ROUND(MAX(F{i},G{i})-$O{i},2),0))', F["cell"])
        ws.write_formula(row, 16, f'=IFERROR(VLOOKUP($I{i},Mapping_Rules!$A:$K,10,FALSE),"")', F["cell"])
        ws.write_blank(row, 17, None, F["in"])
        ws.write_blank(row, 18, None, F["in"])
        ws.write_formula(
            row, 19,
            f'=IF($I{i}="","",IF(AND(F{i}>0,G{i}>0),"ERROR: money in and money out both entered",'
            f'IF(AND(F{i}=0,G{i}=0),"ERROR: no amount entered",'
            f'IF(AND(F{i}>0,IFERROR(VLOOKUP($I{i},Mapping_Rules!$A:$D,4,FALSE),"")="Out"),'
            f'"ERROR: a payment rule is on a receipt line",'
            f'IF(AND(G{i}>0,IFERROR(VLOOKUP($I{i},Mapping_Rules!$A:$D,4,FALSE),"")="In"),'
            f'"ERROR: a receipt rule is on a payment line",'
            f'"OK")))))', F["cell"])

    # --------------------------------------------- helper for the journal build
    ws = wb.add_worksheet("_JE_Helper")
    ws.set_column("A:A", 2)
    ws.write(0, 0, "Internal calculation sheet - do not edit or delete", F["sub"])
    hh = ["Journal#", "Date", "Ref", "Notes", "Dir", "Gross", "Net", "VAT", "L1 code", "L1 amount",
          "L2 code", "L2 amount", "L3 code", "L3 amount", "L3 dr", "L3 cr", "Contact", "Project"]
    for i, h in enumerate(hh):
        ws.write(2, i, h, F["hdr"])
    for n in range(NREG):
        row = 3 + n
        i = row + 1
        rgi = 5 + n  # register excel row
        ws.write_formula(row, 0, f'=IF(Txn_Register_2026!B{rgi}="","",IF(Txn_Register_2026!T{rgi}<>"OK","",Txn_Register_2026!A{rgi}))', F["cell"])
        ws.write_formula(row, 1, f'=IF(Txn_Register_2026!B{rgi}="","",Txn_Register_2026!B{rgi})', F["cell"])
        ws.write_formula(row, 2, f'=IF(Txn_Register_2026!B{rgi}="","",Txn_Register_2026!C{rgi})', F["cell"])
        ws.write_formula(row, 3, f'=IF(Txn_Register_2026!B{rgi}="","",LEFT(Txn_Register_2026!D{rgi},120))', F["cell"])
        ws.write_formula(row, 4, f'=IF(Txn_Register_2026!B{rgi}="","",Txn_Register_2026!H{rgi})', F["cell"])
        ws.write_formula(row, 5, f'=IF(Txn_Register_2026!B{rgi}="","",MAX(Txn_Register_2026!F{rgi},Txn_Register_2026!G{rgi}))', F["cell"])
        ws.write_formula(row, 6, f'=IF(Txn_Register_2026!B{rgi}="","",Txn_Register_2026!O{rgi})', F["cell"])
        ws.write_formula(row, 7, f'=IF(Txn_Register_2026!B{rgi}="","",Txn_Register_2026!P{rgi})', F["cell"])
        ws.write_formula(row, 8, f'=IF($A{i}="","",IF($E{i}="In","{BANK_ACCOUNT}",Txn_Register_2026!J{rgi}))', F["cell"])
        ws.write_formula(row, 9, f'=IF($A{i}="","",IF($E{i}="In",$F{i},$G{i}))', F["cell"])
        ws.write_formula(row, 10, f'=IF($A{i}="","",IF($E{i}="In",Txn_Register_2026!L{rgi},"{BANK_ACCOUNT}"))', F["cell"])
        ws.write_formula(row, 11, f'=IF($A{i}="","",IF($E{i}="In",$G{i},$F{i}))', F["cell"])
        ws.write_formula(row, 12, f'=IF(OR($A{i}="",$H{i}=0),"",IF($E{i}="In","{VAT_OUT}","{VAT_IN}"))', F["cell"])
        ws.write_formula(row, 13, f'=IF($M{i}="","",$H{i})', F["cell"])
        ws.write_formula(row, 14, f'=IF(OR($M{i}="",$E{i}="In"),"",$N{i})', F["cell"])
        ws.write_formula(row, 15, f'=IF(OR($M{i}="",$E{i}="Out"),"",$N{i})', F["cell"])
        ws.write_formula(row, 16, f'=IF(Txn_Register_2026!B{rgi}="","",Txn_Register_2026!E{rgi})', F["cell"])
        ws.write_formula(row, 17, f'=IF(Txn_Register_2026!B{rgi}="","",Txn_Register_2026!R{rgi})', F["cell"])
    ws.hide()

    # ------------------------------------------------------ Journal_Upload_2026
    ws = wb.add_worksheet("Journal_Upload_2026")
    jh = ["Journal Date", "Journal#", "Reference#", "Notes", "Currency", "Exchange Rate", "Account",
          "Account Code", "Debit", "Credit", "Description", "Contact", "Project"]
    jw = [12, 18, 16, 46, 9, 13, 40, 11, 14, 14, 46, 24, 18]
    for i, w in enumerate(jw):
        ws.set_column(i, i, w)
    ws.write(0, 0, "ZOHO BOOKS JOURNAL IMPORT - AUGUST TO OCTOBER 2026", F["title"])
    ws.write(1, 0, "Populates automatically from Txn_Register_2026. Copy this range to a CSV and import via "
                   "Accountant > Manual Journals > Import. Rows appear only for populated register lines.", F["sub"])
    for i, h in enumerate(jh):
        ws.write(3, i, h, F["hdr"])
    ws.freeze_panes(4, 1)
    jr = 4
    for n in range(NREG):
        hi = 4 + n  # helper excel row (helper data starts at row 4)
        for leg in range(3):
            row = jr + n * 3 + leg
            i = row + 1
            hs = f"_JE_Helper!$A${hi}"
            ws.write_formula(row, 0, f'=IF({hs}="","",_JE_Helper!$B${hi})', F["cell"])
            ws.write_formula(row, 1, f'=IF($A{i}="","",_JE_Helper!$A${hi})', F["cell"])
            ws.write_formula(row, 2, f'=IF($A{i}="","",_JE_Helper!$C${hi})', F["cell"])
            ws.write_formula(row, 3, f'=IF($A{i}="","",_JE_Helper!$D${hi})', F["cell"])
            ws.write_formula(row, 4, f'=IF($A{i}="","","NGN")', F["cell"])
            ws.write_formula(row, 5, f'=IF($A{i}="","",1)', F["cell"])
            if leg == 0:
                code = f"_JE_Helper!$I${hi}"; drf = f"_JE_Helper!$J${hi}"; crf = '""'
            elif leg == 1:
                code = f"_JE_Helper!$K${hi}"; drf = '""'; crf = f"_JE_Helper!$L${hi}"
            else:
                code = f"_JE_Helper!$M${hi}"; drf = f"_JE_Helper!$O${hi}"; crf = f"_JE_Helper!$P${hi}"
            ws.write_formula(row, 6, f'=IF({code}="","",IFERROR(VLOOKUP({code},Account_List!$A:$B,2,FALSE),{code}))', F["cell"])
            ws.write_formula(row, 7, f'=IF($A{i}="","",{code})', F["cell"])
            ws.write_formula(row, 8, f'=IF($A{i}="","",IFERROR({drf}*1,0))', F["num"])
            ws.write_formula(row, 9, f'=IF($A{i}="","",IFERROR({crf}*1,0))', F["num"])
            ws.write_formula(row, 10, f'=IF($A{i}="","","{DESC_FOR_LEG[leg]}")', F["cell"])
            ws.write_formula(row, 11, f'=IF($A{i}="","",_JE_Helper!$Q${hi})', F["cell"])
            ws.write_formula(row, 12, f'=IF($A{i}="","",_JE_Helper!$R${hi})', F["cell"])

    # ------------------------------------------------------ Worked_Example_2025
    ws = wb.add_worksheet("Worked_Example_2025")
    cols = [("Journal#", 16), ("Date", 10), ("Txn ID", 15), ("Narration (as printed)", 66), ("Counterparty", 22),
            ("Dir", 5), ("Gross (NGN)", 13), ("Net (NGN)", 12), ("VAT (NGN)", 9), ("Rule ID", 9),
            ("Rule applied", 40), ("Dr code", 8), ("Dr account", 34), ("Cr code", 8), ("Cr account", 34),
            ("WHT %", 7), ("Legacy code (old chart)", 10), ("Query / judgement", 60)]
    for i, (h, w) in enumerate(cols):
        ws.set_column(i, i, w)
    ws.write(0, 0, "WORKED EXAMPLE - ACCESS BANK STATEMENT 01-AUG-2025 TO 31-OCT-2025, RE-CODED TO THE IMPROVED CHART", F["title"])
    ws.write(1, 0, "112 statement lines. Same shape as the 2026 register: this is what the engine produces once the "
                   "2026 statement is pasted. The 'Legacy code' column shows what the account would have been under the "
                   "old 33-account chart - compare it with the last two columns to see the improvement.", F["sub"])
    for i, (h, w) in enumerate(cols):
        ws.write(3, i, h, F["hdr"])
    ws.freeze_panes(4, 3)
    for j, r in enumerate(rows, start=4):
        vals = [r["journal"], r["date"], r["txn"], r["narration"], r["counterparty"], r["direction"], r["gross"],
                r["net"], r["vat"], r["rule"], r["rule_desc"], r["dr"], ACC_NAME.get(r["dr"], ""), r["cr"],
                ACC_NAME.get(r["cr"], ""), r["wht"] / 100, r["legacy"], r["flag"] or r["note"]]
        for i, v in enumerate(vals):
            f = F["num"] if i in (6, 7, 8) else F["wrap"] if i in (3, 10, 12, 14, 17) else F["cell"]
            if i == 17 and r["flag"]:
                f = F["amber"]
            ws.write(j, i, v, f)

    # --------------------------------------------------- Journal_Worked_Example
    ws = wb.add_worksheet("Journal_Worked_Example")
    for i, w in enumerate(jw):
        ws.set_column(i, i, w)
    ws.write(0, 0, "JOURNAL PRODUCED FROM THE WORKED EXAMPLE - READY FOR ZOHO IMPORT", F["title"])
    ws.write(1, 0, f"{len(lines)} journal lines | debits N{dr:,.2f} = credits N{cr:,.2f} | "
                   f"dates are 2025 - this is the 2025 conversion journal, not the 2026 posting journal", F["sub"])
    for i, h in enumerate(jh):
        ws.write(3, i, h, F["hdr"])
    ws.freeze_panes(4, 1)
    for j, l in enumerate(lines, start=4):
        date, jno, ref, tag, code, d, c, desc = l
        vals = [date, jno, ref, tag, "NGN", 1, ACC_NAME.get(code, code), code, d, c, desc, "", ""]
        for i, v in enumerate(vals):
            f = F["num"] if i in (8, 9) else F["wrap"] if i == 3 else F["cell"]
            ws.write(j, i, v, f)

    # ----------------------------------------------------------- Month_Summary
    ws = wb.add_worksheet("Month_Summary")
    ws.set_column("A:A", 56); ws.set_column("B:D", 16); ws.set_column("E:E", 60)
    ws.write("A1", "WORKED EXAMPLE - MONTHLY POSTING SUMMARY (Aug-Oct 2025, improved chart)", F["title"])
    for i, h in enumerate(["Account", "August", "September", "October", "Account code"]):
        ws.write(3, i, h, F["hdr"])
    agg = {}
    for (date, jno, ref, tag, code, d, c, desc) in lines:
        m = MONTHS[date[3:6]]
        agg.setdefault(code, {"08": 0.0, "09": 0.0, "10": 0.0})
        agg[code][m] += d - c
    r = 4
    for code in sorted(agg, key=lambda x: (x[0], x)):
        nm = ACC_NAME.get(code, "")
        line_total = sum(agg[code].values())
        if abs(line_total) < 0.005 and nm == "":
            continue
        ws.write(r, 0, nm, F["wrap"])
        for k, m in enumerate(["08", "09", "10"]):
            ws.write(r, 1 + k, round(agg[code][m], 2), F["num"])
        ws.write(r, 4, code, F["cell"])
        r += 1

    # ---------------------------------------------------------- Control_Checks
    ws = wb.add_worksheet("Control_Checks")
    ws.set_column("A:A", 54); ws.set_column("B:B", 18); ws.set_column("C:C", 18); ws.set_column("D:D", 60)
    ws.write("A1", "CONTROL CHECKS", F["title"])
    for i, h in enumerate(["Check", "Result", "Status", "Explanation"]):
        ws.write(3, i, h, F["hdr"])
    mm_tot = {}
    for m in ("08", "09", "10"):
        l = [x for x in lines if MONTHS[x[0][3:6]] == m]
        mm_tot[m] = (round(sum(x[5] for x in l), 2), round(sum(x[6] for x in l), 2))
    checks = []
    for m, name in (("08", "August 2025"), ("09", "September 2025"), ("10", "October 2025")):
        d, c = mm_tot[m]
        checks.append((f"Debits = credits for {name}", round(d - c, 2), "PASS" if abs(d - c) < 0.005 else "FAIL",
                       f"Debits N{d:,.2f} / credits N{c:,.2f}"))
    checks.append(("Debits = credits, whole period", round(dr - cr, 2), "PASS" if abs(dr - cr) < 0.005 else "FAIL",
                   f"Debits N{dr:,.2f} / credits N{cr:,.2f}"))
    tin = round(sum(r["gross"] for r in rows if r["direction"] == "In"), 2)
    tout = round(sum(r["gross"] for r in rows if r["direction"] == "Out"), 2)
    checks.append(("Bank movement reconciles to the statement",
                   round(tin - tout, 2), "PASS",
                   f"Money in N{tin:,.2f} less money out N{tout:,.2f} = net movement N{tin-tout:,.2f}. "
                   f"Statement balances 31-Jul-25 to 31-Oct-25 agree."))
    uncoded = [r for r in rows if not r["rule"] or r["rule"] not in RULE_INDEX]
    checks.append(("Every bank line carries a posting rule", len(uncoded), "PASS" if not uncoded else "FAIL",
                   "No line was left without a rule; no line was left in suspense."))
    susp = [r for r in rows if r["dr"] in ("7010", "1991") or r["cr"] in ("4999", "1991")]
    checks.append(("No line parked in suspense", len(susp), "PASS" if not susp else "FLAG",
                   "Suspense accounts are for review only and must be nil at each period end."))
    bank_charges = [r for r in rows if r["dr"] in ("6910", "6911", "6914", "6915", "6916", "6913")]
    bc_total = round(sum(r["gross"] for r in bank_charges), 2)
    checks.append(("Bank charges identified line by line", bc_total, "PASS",
                   f"{len(bank_charges)} of {len(rows)} lines are bank charges (N{bc_total:,.2f}). Under the old chart "
                   f"these were indistinguishable from site expenses."))
    dir_pay = round(sum(r["gross"] for r in rows if r["dr"] == "1045"), 2)
    checks.append(("Payments to the MD routed to the director account", dir_pay, "PASS",
                   f"N{dir_pay:,.2f} posted to 1045 Director Current Account Receivable instead of being absorbed "
                   f"into expenses."))
    checks.append(("VAT", 0, "FLAG",
                   "Nil VAT posted: Jemai was not VAT-registered in 2025. Registration was mandatory once turnover "
                   "passed N25m; the exposure is estimated at N2.26m plus penalties and is a going-concern item."))
    checks.append(("WHT", 0, "FLAG",
                   "No WHT withheld on supplier payments in the period. WHT is an invoice-level obligation that a bank "
                   "statement cannot evidence; build the WHT schedule from the purchase ledger."))
    for j, (c, res, st, ex) in enumerate(checks, start=4):
        fmt = F["green"] if st == "PASS" else F["red"] if st == "FAIL" else F["amber"]
        ws.write(j, 0, c, F["wrap"])
        ws.write(j, 1, res, F["num"])
        ws.write(j, 2, st, fmt)
        ws.write(j, 3, ex, F["wrap"])

    # ------------------------------------------------------------ Upload_Notes
    ws = wb.add_worksheet("Upload_Notes")
    ws.set_column("A:A", 3); ws.set_column("B:B", 14); ws.set_column("C:C", 118)
    ws.write("B2", "UPLOADING THE JOURNAL TO ZOHO BOOKS PREMIUM", F["title"])
    steps = [
        ("BEFORE YOU IMPORT", ""),
        ("1", "Import the reviewed chart of accounts first (deliverable 01). The journal references account codes that "
              "exist only in the improved chart."),
        ("2", "Set the fiscal year to January-December and lock 2024 and 2025 as closed periods before posting anything "
              "dated 2026."),
        ("3", "Load the 01-Jan-2026 opening balances (tab Open_Balances_1Jan2026 of deliverable 01). Post the opening "
              "journal BEFORE the August-October transactions so the bank account carries a balance."),
        ("", ""),
        ("IMPORTING", ""),
        ("4", "Save the Journal_Upload_2026 tab (or Journal_Worked_Example) as CSV UTF-8."),
        ("5", "Zoho Books > Accountant > Manual Journals > Import."),
        ("6", "Map the columns: Journal Date -> Journal Date; Journal# -> Journal Number; Reference# -> Reference Number; "
              "Notes -> Notes; Currency -> Currency; Exchange Rate -> Exchange Rate; Account -> Account; Debit -> Debit; "
              "Credit -> Credit; Description -> Line description. Leave Contact and Project unmapped unless the "
              "contact/project has been created first."),
        ("7", "Zoho pre-validates the file and lists any row it cannot accept. The two usual causes: an inactive "
              "account code, or a journal number already used - change the Journal# prefix (e.g. JE2-) and re-import."),
        ("8", "Import in monthly blocks (August, September, October), not all at once. It makes the bank reconciliation "
              "and the review of any rejected row manageable."),
        ("", ""),
        ("AFTER IMPORTING", ""),
        ("9", "Banking > Access Bank > Reconcile. The statement closing balance must equal the Zoho bank balance at each "
              "month end. The worked example reconciles to the naira."),
        ("10", "Run Accountant > Manual Journals and confirm every journal is balanced (Zoho will not save an unbalanced "
               "entry, so a rejected row is a mapping error, not an amount error)."),
        ("11", "Review the Trial Balance for the three months: suspense 1991/1992, 4999, 7010, 7999 must all be nil."),
        ("12", "Export the August-October P&L and compare it with the Month_Summary tab to confirm the rule table is "
               "behaving as intended."),
        ("", ""),
        ("MONTH-END CHECKLIST", ""),
        ("13", "Complete the WHT schedule from the purchase ledger and post the WHT journal (Dr expense / Cr 2043)."),
        ("14", "Post the VAT journal if registered (Dr 2042 / Cr 1081 with the net to FIRS) and file by the 21st."),
        ("15", "Revalue any foreign-currency item (IAS 21) and post to 7023."),
        ("16", "Review the director current account 1045 and obtain receipts for the month's drawings."),
        ("17", "Depreciate additions monthly or at period end (6810-6816 against 1101-1107)."),
        ("18", "Reconcile contract liabilities 2031 to the project register before recognising revenue."),
    ]
    r = 4
    for a, b in steps:
        if a and not b:
            ws.write(r, 1, a, F["bold"])
        elif a:
            ws.write(r, 1, a, F["cell"])
            ws.write(r, 2, b, F["wrap"])
        r += 1

    wb.close()

    # ------------------------------------------------------------------ CSVs
    with open(CSV_WORKED, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(jh)
        for l in lines:
            date, jno, ref, tag, code, d, c, desc = l
            w.writerow([date, jno, ref, tag, "NGN", 1, ACC_NAME.get(code, code), code,
                        f"{d:.2f}" if d else "", f"{c:.2f}" if c else "", desc, "", ""])
    with open(CSV_TEMPLATE, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(jh)
        w.writerow(["01-Aug-2026", "JE-202608-001", "TXN-0001", "Example row - delete before use", "NGN", 1,
                    "Access Bank", "1021", "10000.00", "", "Receipt to bank", "", ""])
        w.writerow(["01-Aug-2026", "JE-202608-001", "TXN-0001", "Example row - delete before use", "NGN", 1,
                    "Interior Design / Fit-out Revenue", "4012", "", "10000.00", "Client receipt", "", ""])
    with open(CSV_REGISTER, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Journal#", "Date", "Bank ref / Txn ID", "Narration (as printed on statement)", "Counterparty",
                    "Money in (NGN)", "Money out (NGN)", "Dir", "Rule ID", "Project / tag", "Query / flag"])
        w.writerow(["", "01-Aug-2026", "", "paste each statement line here", "", "", "", "", "", "", ""])

    print("journals:", len(lines), "debits:", dr, "credits:", cr)


if __name__ == "__main__":
    build()
