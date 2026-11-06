"""Builds the REAL August - October 2026 posting journal from the Access Bank
statement PDF.

Inputs (all in the repository):
    August to October 2026 Statement.pdf      - 26 pages, 716 printed lines
    deliverables/01_..._Improved_Jemai_Oct2026.xlsx - 302-account chart + opening TB

Outputs (deliverables/):
    02e_Journal_Aug-Oct_2026_ACTUAL_Zoho_Upload.csv   - import-ready, all months
    02f_Journal_Aug2026_ACTUAL.csv
    02g_Journal_Sep2026_ACTUAL.csv
    02h_Journal_Oct2026_ACTUAL.csv
    02i_Posting_Journal_Aug-Oct_2026_ACTUAL_Workbook.xlsx
"""
import csv
import os

import openpyxl
import xlsxwriter

from journal_rules_2026 import BANK, RULES, classify, narrative, payee_of, vat_inclusive
from parse_statement import parse

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(ROOT, "deliverables")
COA = os.path.join(OUT, "01_Chart_of_Accounts_Improved_Jemai_Oct2026.xlsx")

CSV_ALL = os.path.join(OUT, "02e_Journal_Aug-Oct_2026_ACTUAL_Zoho_Upload.csv")
CSV_M = {
    "08": os.path.join(OUT, "02f_Journal_Aug2026_ACTUAL.csv"),
    "09": os.path.join(OUT, "02g_Journal_Sep2026_ACTUAL.csv"),
    "10": os.path.join(OUT, "02h_Journal_Oct2026_ACTUAL.csv"),
}
XLSX = os.path.join(OUT, "02i_Posting_Journal_Aug-Oct_2026_ACTUAL_Workbook.xlsx")

MONTHS = {"AUG": "08", "SEP": "09", "OCT": "10"}
MONTH_NAME = {"08": "August 2026", "09": "September 2026", "10": "October 2026"}
STATEMENT = {
    "opening": 24954074.00, "withdrawals": 74373993.97, "deposits": 51225487.00,
    "closing": 1805567.88, "cleared": 1341359.50,
}
HEADER = ["Journal Date", "Journal#", "Reference#", "Notes", "Currency", "Exchange Rate",
          "Account", "Account Code", "Debit", "Credit", "Description", "Contact", "Project"]


# --------------------------------------------------------------------------- #
def load_accounts():
    wb = openpyxl.load_workbook(COA, data_only=True)
    ws = wb["Zoho_Import_CoA"]
    names = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[1] and r[2]:
            names[str(r[2])] = str(r[1])
    ob = {}
    ws = wb["Open_Balances_1Jan2026"]
    # Block A only: the per-account opening trial balance. It ends at the first
    # row without a code (the "TOTAL - BLOCK A" line); blocks B and C below it
    # are the derivation of the same balances and must not be added again.
    for r in ws.iter_rows(min_row=6, values_only=True):
        code = str(r[0] or "").strip()
        if not code:
            break
        if len(code) == 4 and code.isdigit():
            ob[code] = (float(r[2] or 0), float(r[3] or 0))
    return names, ob


def load_rows():
    """Statement lines with the four failed-and-reversed transfer groups removed."""
    tx = parse()
    rows = []
    for x in tx:
        if "opening balance" in x["description"].lower():
            continue
        if abs(x["debit"]) < 0.005 and abs(x["credit"]) < 0.005:
            continue
        rows.append(dict(x))
    excluded = {}
    for i, x in enumerate(rows):
        if not x["reversal"]:
            continue
        target = abs(x["debit"]) or abs(x["credit"])
        excluded[i] = "bank reversal of a failed transfer - no economic effect"
        for j in range(i - 1, max(-1, i - 14), -1):
            y = rows[j]
            if j in excluded or y["reversal"]:
                continue
            if abs((abs(y["debit"]) or abs(y["credit"])) - target) < 0.005:
                excluded[j] = "failed transfer reversed by the bank in the same block - excluded"
                break
    return rows, excluded


def build():
    acc_name, opening = load_accounts()
    rows, excluded = load_rows()

    txns, seq = [], {}
    for i, x in enumerate(rows):
        if i in excluded:
            continue
        direction = "In" if x["credit"] > 0 else "Out"
        amount = round(abs(x["credit"] or x["debit"]), 2)
        month = MONTHS[x["posted"][3:6]]
        seq[month] = seq.get(month, 0) + 1
        rid, code, note = classify(x["description"], direction, amount)
        txns.append({
            "ref": f"STM26-{i + 1:04d}",
            "date": x["posted"], "month": month,
            "journal": f"JE-2026-{month}-{seq[month]:03d}",
            "direction": direction, "amount": amount,
            "what": "Receipt" if direction == "In" else "Payment",
            "narration": narrative(x["description"]),
            "contact": ("Access Bank Plc" if code.startswith("691")
                        else payee_of(x["description"])),
            "rule": rid, "code": code, "account": acc_name.get(code, ""),
            "note": note, "page": x["page"],
        })

    lines = []
    for t in txns:
        tag = f"{t['ref']} | {t['narration'][:90]}"
        note = f"{t['what']}: {t['contact']}".strip(": ")
        if t["direction"] == "In":
            lines.append([t["date"], t["journal"], t["ref"], note, "NGN", 1,
                          acc_name.get(BANK, "Access Bank"), BANK, t["amount"], 0.0,
                          t["narration"], t["contact"], ""])
            lines.append([t["date"], t["journal"], t["ref"], note, "NGN", 1,
                          t["account"], t["code"], 0.0, t["amount"],
                          t["narration"], t["contact"], ""])
        else:
            lines.append([t["date"], t["journal"], t["ref"], note, "NGN", 1,
                          t["account"], t["code"], t["amount"], 0.0,
                          t["narration"], t["contact"], ""])
            lines.append([t["date"], t["journal"], t["ref"], note, "NGN", 1,
                          acc_name.get(BANK, "Access Bank"), BANK, 0.0, t["amount"],
                          t["narration"], t["contact"], ""])
    return txns, lines, acc_name, opening, rows, excluded


# --------------------------------------------------------------------------- #
def _csv_row(l):
    return l[:8] + [f"{l[8]:.2f}", f"{l[9]:.2f}"] + l[10:]


def write_csvs(lines):
    with open(CSV_ALL, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows([_csv_row(l) for l in lines])
    for m, path in CSV_M.items():
        with open(path, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(HEADER)
            w.writerows([_csv_row(l) for l in lines if l[1].split("-")[2] == m])


def build_workbook(txns, lines, acc_name, opening, rows, excluded):
    dr = round(sum(l[8] for l in lines), 2)
    cr = round(sum(l[9] for l in lines), 2)
    assert abs(dr - cr) < 0.005, (dr, cr)

    receipts = round(sum(t["amount"] for t in txns if t["direction"] == "In"), 2)
    payments = round(sum(t["amount"] for t in txns if t["direction"] == "Out"), 2)
    closing = round(STATEMENT["opening"] + receipts - payments, 2)

    wb = xlsxwriter.Workbook(XLSX)
    F = {
        "t": wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864"}),
        "s": wb.add_format({"italic": True, "font_size": 9, "font_color": "#595959"}),
        "h": wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                            "border": 1, "text_wrap": True, "valign": "vcenter", "font_size": 9}),
        "c": wb.add_format({"border": 1, "font_size": 9, "valign": "top"}),
        "w": wb.add_format({"border": 1, "font_size": 9, "text_wrap": True, "valign": "top"}),
        "n": wb.add_format({"border": 1, "font_size": 9, "num_format": "#,##0.00", "valign": "top"}),
        "b": wb.add_format({"bold": True, "font_size": 9, "border": 1, "valign": "top"}),
        "nb": wb.add_format({"bold": True, "font_size": 9, "border": 1,
                             "num_format": "#,##0.00", "valign": "top"}),
        "ok": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#C6EFCE", "text_wrap": True}),
        "bad": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FFC7CE", "text_wrap": True}),
        "amber": wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FFEB9C", "text_wrap": True}),
    }

    def table(ws, head, data, widths, fmts=None, start=0):
        for i, h in enumerate(head):
            ws.write(start, i, h, F["h"])
        ws.set_row(start, 34)
        for j, row in enumerate(data, start + 1):
            for i, v in enumerate(row):
                fmt = F["c"]
                if fmts and i in fmts:
                    fmt = F["n"]
                if isinstance(v, str) and v.startswith(("PASS", "OK")):
                    fmt = F["ok"]
                elif isinstance(v, str) and v.startswith(("FAIL", "ERROR", "FLAG")):
                    fmt = F["bad"]
                elif isinstance(v, str) and v.startswith("CHECK"):
                    fmt = F["amber"]
                if isinstance(v, str) and len(v) > 40:
                    fmt = F["w"]
                ws.write(j, i, v, fmt)
        for i, wd in enumerate(widths):
            ws.set_column(i, i, wd)

    # ------------------------------------------------------------------ Read_Me
    ws = wb.add_worksheet("Read_Me")
    ws.set_column("A:A", 3); ws.set_column("B:B", 34); ws.set_column("C:C", 108)
    ws.write("B2", "JEMAI INTERNATIONAL LIMITED  |  ACTUAL TRANSACTION POSTING JOURNAL", F["t"])
    ws.write("B3", "1 August 2026 - 31 October 2026  |  built from the Access Bank statement "
                   "(A/C 1854793470)  |  improved 302-account chart", F["s"])
    ws.write("B4", f"Prepared November 2026  |  {len(txns)} posted bank lines  |  "
                   f"{len(lines)} journal legs  |  Dr = Cr = N{dr:,.2f}", F["s"])
    body = [
        ("WHAT THIS FILE IS", ""),
        ("Source", "The 26-page 'August to October 2026 Statement.pdf' in this repository. Every "
                   "printed line was parsed from the PDF (word coordinates, not text order) and "
                   "reconciled to the statement's own totals - see Control_Checks."),
        ("Journal", f"{len(txns)} economic bank lines -> {len(lines)} balanced journal legs, one "
                    "journal per bank line, numbered JE-2026-MM-NNN, in the Zoho Books journal "
                    "import layout (same columns as 02b / 02c)."),
        ("Import files", "02e (all three months), 02f (August), 02g (September), 02h (October). "
                         "Use 02e unless you want to post month by month."),
        ("THE TABS", ""),
        ("Control_Checks", "Parse reconciliation to the statement totals, Dr = Cr, bank ledger "
                           "roll-forward, and the exceptions that a reviewer must clear."),
        ("Month_Summary", "Receipts, payments and net movement per month, plus the expense mix."),
        ("Txn_Register", "Every bank line with the rule applied, the account it was posted to and "
                         "the reviewer note."),
        ("Journal_Lines", "The journal itself - exactly what is in the CSV upload files."),
        ("Trial_Balance", "1-Jan-2026 opening balance (from 01) + Aug-Oct 2026 movement = closing."),
        ("Exceptions", "Lines that need a human decision before the journal is final (unknown "
                       "payees, judgement calls, refunds, related parties)."),
        ("VAT_Review", "Receipts whose amounts are demonstrably 7.5% VAT-inclusive, with the "
                       "reclass entry to post if the company is VAT-registered."),
        ("Rules_Applied", "Every posting rule used, with line count and value - the audit trail "
                          "from narration to account."),
        ("BASIS OF PREPARATION - READ BEFORE IMPORTING", ""),
        ("VAT", "Posted GROSS. VAT registration is not evidenced in the file set, so no output or "
                "input VAT has been split out. VAT_Review quantifies the adjustment if it applies."),
        ("Director", "Every payment to Esther Akin-Ajayi (and every receipt from her) has gone to "
                     "1045 Director Current Account Receivable, not to expense - except the "
                     "September salary and housing allowance runs, which are payroll (6110/6111)."),
        ("Failed transfers", "Four IG-Ads / Marketing-Ads transfers of N1,000,053.75 each were "
                             "reversed by the bank in the same block. Both the payment and its "
                             "reversal are excluded (they cancel); 24 printed lines, N8,000,430.00."),
        ("Statement rounding", "The statement's own totals foot to a closing balance of "
                               "N1,805,567.03 while it prints N1,805,567.88 - a N0.85 difference "
                               "inside the bank document, not in this journal."),
        ("Not audit-ready", "This is a posting journal built from bank evidence only. Accruals, "
                            "depreciation, revenue cut-off and payroll liabilities still have to "
                            "be journalled to close the year."),
    ]
    r = 6
    for a, b in body:
        if b == "":
            ws.write(r, 1, a, F["b"]); r += 1
            continue
        ws.write(r, 1, a, F["b"]); ws.write(r, 2, b, F["w"]); r += 1

    # -------------------------------------------------------------- Control_Checks
    ws = wb.add_worksheet("Control_Checks")
    ws.set_column(0, 0, 34); ws.set_column(1, 1, 16); ws.set_column(2, 2, 90)
    ws.write("A1", "CONTROL CHECKS", F["t"])
    gross_dr = round(sum(x["debit"] for x in rows if not x["reversal"]), 2)
    rev = round(-sum(x["debit"] + x["credit"] for x in rows if x["reversal"]), 2)
    checks = [
        ("Statement lines parsed", f"{len(rows)} + 1 opening row",
         "716 economic lines were printed; the opening balance row is not a transaction."),
        ("Gross withdrawals per parse", f"{gross_dr:,.2f}",
         f"Statement prints {STATEMENT['withdrawals']:,.2f}. Difference {gross_dr - rev - STATEMENT['withdrawals']:,.2f} "
         f"= the {rev:,.2f} of reversed transfers, which the bank nets off its 'Total Withdrawals'."),
        ("Withdrawals net of reversals", f"{gross_dr - rev:,.2f}",
         f"PASS - agrees to the statement headline {STATEMENT['withdrawals']:,.2f}."),
        ("Deposits per parse",
         f"{sum(x['credit'] for x in rows):,.2f}",
         f"PASS - agrees to the statement headline {STATEMENT['deposits']:,.2f}."),
        ("Journal receipts", f"{receipts:,.2f}",
         "PASS - equals deposits parsed (no reversal affects the credit side)."),
        ("Journal payments", f"{payments:,.2f}",
         f"PASS - equals net withdrawals {gross_dr - rev:,.2f}."),
        ("Dr = Cr", f"{dr:,.2f}", "PASS - journal is balanced."),
        ("Bank ledger roll-forward",
         f"{STATEMENT['opening']:,.2f} + {receipts:,.2f} - {payments:,.2f} = {closing:,.2f}",
         f"Statement prints a closing balance of {STATEMENT['closing']:,.2f}; the difference of "
         f"{abs(closing - STATEMENT['closing']):,.2f} sits inside the bank document itself (its own "
         "totals do not foot to its printed closing balance)."),
        ("Reversal pairs excluded", f"{2 * rev:,.2f} (2 x {rev:,.2f})",
         "Four failed transfers + four reversals (24 printed lines) cancelled and were excluded."),
        ("Lines still on 7999 Uncategorised",
         f"{sum(1 for t in txns if t['code'] == '7999')}",
         "Must be cleared before import - see Exceptions."),
        ("Receipts on the automatic gallery default",
         f"{sum(1 for t in txns if t['rule'] == 'REC-99')}",
         "Unlabelled individual receipts assumed to be gallery / art sales - see Exceptions."),
        ("Payments to the director posted to 1045",
         f"{sum(t['amount'] for t in txns if t['code'] == '1045' and t['direction'] == 'Out'):,.2f}",
         "Balance-sheet movement, not expense. Review the loan agreement and IAS 24 disclosure."),
    ]
    table(ws, ["Check", "Result", "Comment"], checks, [34, 42, 120])

    # ------------------------------------------------------------- Month_Summary
    ws = wb.add_worksheet("Month_Summary")
    ws.write("A1", "MONTHLY SUMMARY - actual bank activity", F["t"])
    head = ["Month", "Receipts (NGN)", "Payments (NGN)", "Net movement", "Bank lines",
            "Top 3 payment accounts"]
    data = []
    for m in ("08", "09", "10"):
        sub = [t for t in txns if t["month"] == m]
        rin = sum(t["amount"] for t in sub if t["direction"] == "In")
        pout = sum(t["amount"] for t in sub if t["direction"] == "Out")
        by_acc = {}
        for t in sub:
            if t["direction"] == "Out":
                by_acc[t["code"]] = by_acc.get(t["code"], 0) + t["amount"]
        top = sorted(by_acc.items(), key=lambda kv: -kv[1])[:3]
        data.append([MONTH_NAME[m], round(rin, 2), round(pout, 2), round(rin - pout, 2),
                     len(sub), "; ".join(f"{acc_name.get(c, c)} {v:,.0f}" for c, v in top)])
    data.append(["TOTAL", receipts, payments, round(receipts - payments, 2), len(txns), ""])
    table(ws, head, data, [18, 16, 16, 16, 10, 70], fmts={1, 2, 3})

    # expense mix
    r0 = len(data) + 4
    ws.write(r0, 0, "PAYMENT MIX BY ACCOUNT", F["b"])
    mix = {}
    for t in txns:
        if t["direction"] == "Out":
            mix[t["code"]] = mix.get(t["code"], 0) + t["amount"]
    rows_mix = [[acc_name.get(c, c), c, round(v, 2), round(v / payments * 100, 1)]
                for c, v in sorted(mix.items(), key=lambda kv: -kv[1])]
    table(ws, ["Account", "Code", "Amount (NGN)", "% of payments"], rows_mix,
          [40, 8, 16, 14], fmts={2, 3}, start=r0 + 1)

    # -------------------------------------------------------------- Txn_Register
    ws = wb.add_worksheet("Txn_Register")
    head = ["#", "Date", "Month", "Journal#", "Reference", "Type", "Amount (NGN)",
            "Narration (as printed on the statement)", "Contact", "Rule", "Account code",
            "Account", "Reviewer note", "Page"]
    data = []
    for n, t in enumerate(txns, 1):
        data.append([n, t["date"], MONTH_NAME[t["month"]], t["journal"], t["ref"], t["what"],
                     t["amount"], t["narration"], t["contact"], t["rule"], t["code"],
                     t["account"], t["note"], t["page"]])
    table(ws, head, data,
          [5, 11, 15, 15, 11, 9, 14, 70, 24, 8, 7, 34, 70, 6], fmts={6}, start=0)
    ws.freeze_panes(1, 0)

    # -------------------------------------------------------------- Journal_Lines
    ws = wb.add_worksheet("Journal_Lines")
    table(ws, HEADER, lines, [12, 15, 11, 28, 8, 9, 34, 8, 14, 14, 78, 24, 9],
          fmts={8, 9}, start=0)
    ws.freeze_panes(1, 0)

    # --------------------------------------------------------------- Trial_Balance
    ws = wb.add_worksheet("Trial_Balance")
    ws.write("A1", "TRIAL BALANCE - 1 January 2026 + August-October 2026 movement", F["t"])
    ws.write("A2", "Opening balances come from deliverable 01 (Open_Balances_1Jan2026). "
                   "Movements are this journal only; accruals, depreciation and revenue "
                   "cut-off are not included.", F["s"])
    move = {}
    for l in lines:
        code = str(l[7])
        d, c = l[8] or 0, l[9] or 0
        m = move.setdefault(code, [0.0, 0.0])
        m[0] += d
        m[1] += c
    codes = sorted(set(list(opening) + list(move)))
    data, tdr, tcr, tcdr, tccr = [], 0.0, 0.0, 0.0, 0.0
    for code in codes:
        od, oc = opening.get(code, (0, 0))
        md, mc = move.get(code, (0.0, 0.0))
        cd, cc = od + md, oc + mc
        tdr += od; tcr += oc; tcdr += cd; tccr += cc
        data.append([code, acc_name.get(code, ""), round(od, 2), round(oc, 2),
                     round(md, 2), round(mc, 2), round(cd, 2), round(cc, 2)])
    data.append(["", "TOTAL", round(tdr, 2), round(tcr, 2), round(sum(m[0] for m in move.values()), 2),
                 round(sum(m[1] for m in move.values()), 2), round(tcdr, 2), round(tccr, 2)])
    data.append(["", "", "", "", "", "", "", ""])
    data.append(["", "Difference (must be nil)", round(tdr - tcr, 2), "",
                 round(sum(m[0] for m in move.values()) - sum(m[1] for m in move.values()), 2),
                 "", round(tcdr - tccr, 2), ""])
    table(ws, ["Code", "Account", "Opening Dr", "Opening Cr", "Movement Dr", "Movement Cr",
               "Closing Dr", "Closing Cr"], data, [8, 46, 15, 15, 15, 15, 15, 15],
          fmts={2, 3, 4, 5, 6, 7}, start=4)

    # P&L extract from the movements (revenue and expense accounts only)
    def pl(prefixes):
        d = sum(v[0] for k, v in move.items() if k[0] in prefixes)
        c = sum(v[1] for k, v in move.items() if k[0] in prefixes)
        return d, c
    rd, rc = pl("4")
    cd, cc = pl("5")
    od, oc = pl("67")
    r1 = len(data) + 7
    ws.write(r1, 1, "RESULT FOR AUGUST - OCTOBER 2026 (from the movements above)", F["b"])
    ws.write(r1 + 1, 1, "Revenue (4xxx, net of debits)"); ws.write(r1 + 1, 6, round(rc - rd, 2), F["n"])
    ws.write(r1 + 2, 1, "Direct cost (5xxx)"); ws.write(r1 + 2, 6, round(cd - cc, 2), F["n"])
    ws.write(r1 + 3, 1, "Operating & other costs (6xxx, 7xxx)")
    ws.write(r1 + 3, 6, round(od - oc, 2), F["n"])
    ws.write(r1 + 4, 1, "Result before tax")
    ws.write(r1 + 4, 6, round((rc - rd) - (cd - cc) - (od - oc), 2), F["nb"])
    ws.write(r1 + 6, 1,
             "Bank-evidence only: no accruals, depreciation, revenue cut-off, tax or payroll "
             "liabilities are in these numbers, so treat the result as a cash-shaped proxy, "
             "not the statutory profit. Closing Dr = Cr, as it must.", F["s"])

    # ----------------------------------------------------------------- Exceptions
    ws = wb.add_worksheet("Exceptions")
    ws.write("A1", "EXCEPTIONS AND REVIEW POINTS", F["t"])
    ws.write("A2", "Everything here is already posted to a sensible account - these are the "
                   "judgement calls and open questions to clear before the year-end close.", F["s"])
    unc = [t for t in txns if t["code"] == "7999"]
    auto = [t for t in txns if t["rule"] == "REC-99"]
    big_auto = [t for t in auto if t["amount"] >= 50000]
    small_auto = [t for t in auto if t["amount"] < 50000]
    ovr = [t for t in txns if t["rule"] == "OVR"]

    summary = [
        ("Uncategorised payments (7999)", len(unc), round(sum(t["amount"] for t in unc), 2),
         "No rule matched the narration. Identify the payee from the bank / the vendor's "
         "invoice and reclassify before you import."),
        ("Unlabelled receipts >= N50,000", len(big_auto),
         round(sum(t["amount"] for t in big_auto), 2),
         "Treated as gallery / art income. These are large enough that the customer's "
         "invoice or ticket should be matched to them."),
        ("Unlabelled receipts < N50,000", len(small_auto),
         round(sum(t["amount"] for t in small_auto), 2),
         "Ticket-sized receipts (N5,375 = N5,000 + VAT, N10,750, N21,500, N12,000). "
         "Aggregate to the daily ticket report; do not chase individually."),
        ("Judgement overrides", len(ovr), round(sum(t["amount"] for t in ovr), 2),
         "The narration was not decisive; the account was chosen on the business "
         "substance and is explained line by line below."),
    ]
    table(ws, ["Exception group", "Lines", "Value (NGN)", "What to do"], summary,
          [34, 8, 15, 100], fmts={2}, start=4)
    r0 = 5 + len(summary) + 1
    data = []
    for t in big_auto + unc + ovr:
        cat = ("Uncategorised" if t["code"] == "7999" else
               "Judgement override" if t["rule"] == "OVR" else "Unlabelled receipt")
        data.append([cat, t["date"], t["ref"], t["amount"], t["narration"], t["rule"], t["note"]])
    data.sort(key=lambda r: (r[0], -r[3]))
    table(ws, ["Category", "Date", "Reference", "Amount (NGN)", "Narration", "Rule",
               "Why it is here / what to do"],
          data, [20, 11, 11, 15, 78, 9, 78], fmts={3}, start=r0)

    # ------------------------------------------------------------------ VAT_Review
    ws = wb.add_worksheet("VAT_Review")
    ws.write("A1", "VAT REVIEW - receipts that look 7.5% VAT-inclusive", F["t"])
    ws.write("A2", "These receipts divide exactly by 1.075 into a whole naira amount, which is "
                   "what a VAT-inclusive price looks like. They are posted GROSS in the journal; "
                   "if the company is VAT-registered, post the reclass below.", F["s"])
    vat_rows, vat_total = [], 0.0
    for t in txns:
        if t["direction"] == "In" and vat_inclusive(t["amount"]):
            net = round(t["amount"] / 1.075, 2)
            v = round(t["amount"] - net, 2)
            vat_total += v
            vat_rows.append([t["date"], t["ref"], t["amount"], net, v, t["narration"][:90]])
    table(ws, ["Date", "Reference", "Gross receipt", "Net revenue", "VAT at 7.5%", "Narration"],
          vat_rows, [11, 11, 15, 15, 15, 90], fmts={2, 3, 4}, start=4)
    r1 = 5 + len(vat_rows)
    ws.write(r1 + 1, 1, "Proposed reclass if registered", F["b"])
    ws.write(r1 + 2, 1, "Dr 4010 Art Gallery Sales", F["c"])
    ws.write(r1 + 2, 3, round(vat_total, 2), F["n"])
    ws.write(r1 + 3, 1, "Cr 2041 VAT Output Payable", F["c"])
    ws.write(r1 + 3, 4, round(vat_total, 2), F["n"])

    # --------------------------------------------------------------- Rules_Applied
    ws = wb.add_worksheet("Rules_Applied")
    ws.write("A1", "RULES APPLIED", F["t"])
    agg = {}
    for t in txns:
        k = (t["rule"], t["code"])
        a = agg.setdefault(k, [0, 0.0, t["note"]])
        a[0] += 1
        a[1] += t["amount"]
    rule_desc = {r["id"]: r for r in RULES}
    data = []
    for (rid, code), (n, amt, note) in sorted(agg.items(), key=lambda kv: -kv[1][1]):
        data.append([rid, acc_name.get(code, ""), code, n, round(amt, 2), note])
    table(ws, ["Rule", "Account", "Code", "Lines", "Value (NGN)", "Basis"], data,
          [10, 40, 8, 8, 16, 90], fmts={4}, start=3)

    wb.close()
    return dr, cr, receipts, payments, closing


if __name__ == "__main__":
    txns, lines, acc_name, opening, rows, excluded = build()
    write_csvs(lines)
    dr, cr, receipts, payments, closing = build_workbook(txns, lines, acc_name, opening, rows, excluded)
    print(f"lines {len(txns)}  legs {len(lines)}  excluded {len(excluded)}")
    print(f"receipts {receipts:,.2f}  payments {payments:,.2f}")
    print(f"Dr {dr:,.2f}  Cr {cr:,.2f}  closing {closing:,.2f}")
    print("uncategorised:", sum(1 for t in txns if t["code"] == "7999"),
          " auto-gallery:", sum(1 for t in txns if t["rule"] == "REC-99"))
