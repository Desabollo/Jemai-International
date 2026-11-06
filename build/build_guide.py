"""Builds 03_Zoho_Books_Premium_Setup_Guide_Jemai.docx"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(ROOT, "deliverables", "03_Zoho_Books_Premium_Setup_Guide_Jemai_International.docx")

NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xA6, 0x1B, 0x1B)

doc = Document()

# ----------------------------------------------------------------- page setup
for s in doc.sections:
    s.top_margin = Cm(2.0)
    s.bottom_margin = Cm(1.8)
    s.left_margin = Cm(2.2)
    s.right_margin = Cm(2.0)

st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(10.5)
st.paragraph_format.space_after = Pt(5)
st.paragraph_format.line_spacing = 1.12

for name, size, colour in (("Heading 1", 17, NAVY), ("Heading 2", 13.5, NAVY),
                           ("Heading 3", 11.5, NAVY)):
    h = doc.styles[name]
    h.font.name = "Calibri"
    h.font.size = Pt(size)
    h.font.color.rgb = colour
    h.font.bold = True
    h.paragraph_format.space_before = Pt(14 if name == "Heading 1" else 10)
    h.paragraph_format.space_after = Pt(5)
    h.paragraph_format.keep_with_next = True


def h1(t):
    p = doc.add_heading(t, level=1)
    return p


def h2(t):
    return doc.add_heading(t, level=2)


def h3(t):
    return doc.add_heading(t, level=3)


def p(t, bold=False, italic=False, size=None, colour=None, space=5):
    par = doc.add_paragraph()
    run = par.add_run(t)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if colour:
        run.font.color.rgb = colour
    par.paragraph_format.space_after = Pt(space)
    return par


def rich(parts, space=5):
    """parts = [(text, bold), ...]"""
    par = doc.add_paragraph()
    for text, bold in parts:
        r = par.add_run(text)
        r.bold = bold
    par.paragraph_format.space_after = Pt(space)
    return par


def bullet(t, level=0):
    par = doc.add_paragraph(t, style="List Bullet")
    par.paragraph_format.left_indent = Cm(0.63 + 0.5 * level)
    par.paragraph_format.space_after = Pt(2)
    return par


def numbered(t):
    par = doc.add_paragraph(t, style="List Number")
    par.paragraph_format.left_indent = Cm(0.63)
    par.paragraph_format.space_after = Pt(2)
    return par


def shade(cell, hexcolour):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolour)
    tcPr.append(shd)


def table(headers, rows, widths=None, fontsize=8.5, header_fill="1F3864", zebra=True):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    trPr = t.rows[0]._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)
    hdr = t.rows[0].cells
    for i, htxt in enumerate(headers):
        hdr[i].text = ""
        par = hdr[i].paragraphs[0]
        run = par.add_run(htxt)
        run.bold = True
        run.font.size = Pt(fontsize)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(hdr[i], header_fill)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ""
            par = cells[i].paragraphs[0]
            run = par.add_run(str(val))
            run.font.size = Pt(fontsize)
            par.paragraph_format.space_after = Pt(0)
            if ri % 2 == 1 and zebra:
                shade(cells[i], "F2F2F2")
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def callout(title, text, colour="FFF2CC"):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.rows[0].cells[0]
    c.text = ""
    par = c.paragraphs[0]
    r = par.add_run(title + "  ")
    r.bold = True
    r.font.size = Pt(9)
    r2 = par.add_run(text)
    r2.font.size = Pt(9)
    shade(c, colour)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def pagebreak():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# =========================================================== COVER + CONTROL
tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("\n\n\nJEMAI INTERNATIONAL LIMITED")
r.bold = True
r.font.size = Pt(26)
r.font.color.rgb = NAVY

tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("ZOHO BOOKS PREMIUM")
r.bold = True
r.font.size = Pt(30)
r.font.color.rgb = NAVY

tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("Implementation &\nFinancial Reporting Setup Guide")
r.bold = True
r.font.size = Pt(20)
r.font.color.rgb = GREY

tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("\n\nBuilt on the FY2024-FY2025 audit working papers,\nthe reviewed chart of accounts and the Access Bank statement register\n")
r.italic = True
r.font.size = Pt(11)
r.font.color.rgb = GREY

tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("RC 7171046  |  TIN 2522579844227  |  Lagos, Nigeria\nNovember 2026")
r.font.size = Pt(11)
r.bold = True

pagebreak()

h1("Document control")
table(["Item", "Detail"], [
    ["Client", "Jemai International Limited (RC 7171046, TIN 2522579844227)"],
    ["Engagement", "Design and configuration of a Zoho Books Premium accounting system capable of producing "
                   "IFRS-compliant financial statements and Nigerian tax filings"],
    ["Document", "Zoho Books Premium - Implementation & Financial Reporting Setup Guide"],
    ["Version", "1.0 - for management review and sign-off"],
    ["Date", "November 2026"],
    ["Prepared by", "Finance systems implementation (Arena.ai agent working session)"],
    ["Approved by", "Mrs. Esther Akin-Ajayi Jesutofunmi Ifeoluwa, Managing Director/CEO"],
    ["Source documents", "Chart_of_Accounts Oct 2026.xlsx (Zoho export, 275 accounts); "
                         "Jemai_International_Master_Workbook_2024-2025.xlsx (35 tabs, 1,226 bank transactions); "
                         "Jemai Int'nal 2024 AFS (signed 31-Oct-2026); Jemai Int'nal 2025 AFS (signed 31-Oct-2026)"],
    ["Companion deliverables", "01_Chart_of_Accounts_Improved_Jemai_Oct2026.xlsx (+ CSV import file)\n"
                               "02_Jemai_Posting_Journal_Aug-Oct_2026.xlsx (posting engine + 2025 worked example)\n"
                               "02e-02i: the ACTUAL August-October 2026 journal (CSV upload files + workbook)"],
], widths=[3.4, 13.2], fontsize=9)

callout("How to use this guide.",
        "Sections 1-4 are for the director: what is wrong today, what the system will do, and the decisions required "
        "before configuration starts. Sections 5-16 are the configuration manual for the implementer, written in the "
        "sequence the settings should be applied. Sections 17-20 cover go-live, controls, reporting and the risks we "
        "have not been able to close. Appendix A is the sign-off checklist.")

h2("Contents")
for t in [
    "1. Executive summary and the case for change",
    "2. The company as the system must see it",
    "3. What the historical files prove about the current bookkeeping",
    "4. Why Zoho Books Premium, and the decisions required before configuration",
    "5. Organisation profile and global settings",
    "6. Chart of accounts configuration and import",
    "7. Opening balances at 1 January 2026",
    "8. Bank accounts, feeds, rules and reconciliation",
    "9. Contacts: customers, vendors, related parties",
    "10. Projects, budgets and job costing",
    "11. Sales cycle: quotes, invoices, retentions, advances and credit notes",
    "12. Purchase cycle, expenses, staff claims and withholding tax",
    "13. Payroll and statutory deductions",
    "14. Tax configuration: VAT, WHT, CIT, development levy, deferred tax",
    "15. Workflows, approvals, roles and period locks",
    "16. Budgets, analytics and the IFRS reporting pack",
    "17. Data migration and go-live plan",
    "18. Month-end and year-end close procedures",
    "19. Controls, governance and audit readiness",
    "20. Risks, open items and assumptions",
    "Appendix A - Implementation sign-off checklist",
    "Appendix B - Tax parameters used (confirm before reliance)",
    "Appendix C - Opening trial balance at 1 January 2026",
    "Appendix D - Journal import field map",
    "Appendix E - Documents to request from the client",
]:
    par = doc.add_paragraph(t)
    par.paragraph_format.space_after = Pt(1)
    par.paragraph_format.left_indent = Cm(0.4)

pagebreak()

# =========================================================== 1 EXEC SUMMARY
h1("1. Executive summary and the case for change")

p("Jemai International Limited is a N30m-turnover interior design, furniture and fit-out business that has been "
  "operating a corporate bank account since May 2024. The company has two years of signed audited financial "
  "statements, a 1,226-line bank transaction register and a 275-account Zoho Books chart of accounts. It does not "
  "have a reliable bookkeeping system: the chart of accounts it owns is not the chart of accounts its transactions "
  "were posted to, and the reports the system must produce (an IFRS statement of financial position, an IFRS 15 "
  "revenue note and a Nigerian tax computation) cannot be generated from the way the ledger has been used.")

p("This guide sets out how to configure Zoho Books Premium so that the same information that took an audit "
  "reconstruction to assemble becomes available every month, automatically.")

h2("1.1 The five findings that drive the design")
table(["#", "Finding", "Evidence in the files", "What the system must do"], [
    ["1", "Half of all bank transactions were posted to one account",
     "607 of 1,226 lines (49.5%) sat in legacy account 6050 'Bank charges, levies, airtime & telecom', which also "
     "absorbed site expenses, fuel, transport, training and refunds",
     "A chart with four separate bank-charge accounts, plus bank rules that classify the narration pattern "
     "automatically at import"],
    ["2", "Owner payments were absorbed into expenses",
     "N5.05m (2024) and N9.60m (2025) paid to the MD were expensed; the corrected position is a N14.22m receivable "
     "from the director, undisclosed in the AFS",
     "A director current account with its own expense-claim workflow, so drawings and business costs never mix"],
    ["3", "Revenue was recognised on cash received, including advances",
     "N3.46m of client deposits sat in turnover at both year ends although the work was unperformed",
     "Project-based invoicing with a contract-liability account and a revenue-recognition control report"],
    ["4", "Tax was not provided and VAT registration was missed",
     "The AFS report nil tax; minimum tax and TEL of N235,622 were unpaid; 2025 turnover of N30.1m crossed the N25m "
     "VAT threshold while unregistered (exposure ~N2.26m)",
     "VAT 7.5% as a configured tax rate, WHT deduction on purchases, and a monthly tax dashboard"],
    ["5", "The chart of accounts had structural defects",
     "A duplicate account code (2030 used twice), an account with no code, an inverted equity hierarchy, 19 orphan "
     "accounts with no parent, no deferred tax or provision accounts and no non-current liability section",
     "The reviewed 302-account chart in deliverable 01, imported before any transaction is posted"],
], widths=[0.8, 3.9, 6.0, 5.9], fontsize=8.5)

h2("1.2 What good looks like after go-live")
for t in [
    "Every bank line is coded to a named account within 24 hours of the statement line appearing, using a bank rule "
    "that has been tested against the historical narrations.",
    "The trial balance for any month balances to the bank statement and to the project register without adjustment "
    "journals.",
    "The director's current account, contract liabilities, VAT, WHT and tax payables are visible as balances at any "
    "time - not reconstructed at year end.",
    "Monthly management accounts (P&L by project, balance sheet, cash flow) are produced by the 7th working day.",
    "The auditor receives a Zoho login with read-only access and exports the general ledger, the trial balance and "
    "the project profitability report directly from the system.",
]:
    bullet(t)

h2("1.3 What this engagement does not do")
p("It does not rebuild 2024 and 2025 in Zoho. Those years are closed and audited; correcting them would require "
  "restating signed financial statements. The historical evidence is used to (a) design the chart, (b) test the "
  "posting rules and (c) establish the 1 January 2026 opening balances. The reviewed chart and the posting engine "
  "carry the corrections forward from 2026.")

pagebreak()

# =========================================================== 2 COMPANY
h1("2. The company as the system must see it")
p("These are the settings that Zoho Books will pre-print on every invoice, every report and every tax filing. They "
  "must match the CAC and FIRS records exactly - the 2024-2025 file already contains one compliance defect caused by "
  "a mismatch (the bank mandate reads 'JEMAI INTERNATIONAL LIMTED').")

table(["Attribute", "Value to configure", "Source"], [
    ["Registered name", "JEMAI INTERNATIONAL LIMITED", "CAC certificate / AFS cover"],
    ["Registration number", "RC 7171046", "CAC"],
    ["Tax identification number", "2522579844227", "FIRS / NRS"],
    ["Date of incorporation", "9 October 2023", "CAC"],
    ["Registered office", "24, Ashaolu Street, Off Akilo Road, Ogba, Lagos State, Nigeria", "AFS / CAC"],
    ["Principal activities", "Interior design and fit-out; furniture production and supply; architectural design; "
                             "visual art and installations", "M&AA / AFS directors' report"],
    ["Directors", "Mrs. Esther Akin-Ajayi Jesutofunmi Ifeoluwa (MD/CEO) - sole director per the AFS signature block",
     "AFS page 3"],
    ["Shareholders", "Esther Akin-Ajayi - 1,000,000 ordinary shares of N1.00 (100%)", "AFS note 7"],
    ["Authorised / issued capital", "N1,000,000 divided into 1,000,000 shares of N1.00 each, fully paid", "AFS note 7"],
    ["Functional & presentation currency", "Nigerian Naira (NGN) - single currency", "IAS 21"],
    ["Financial year", "1 January to 31 December", "AFS"],
    ["Accounting framework", "Full IFRS as issued by the FRCN; CAMA 2020 disclosures", "AFS note 1"],
    ["Auditors", "Gbenga Adeyeye & Co (Chartered Accountants), 239 Egbeda-Idimu Road, Idimu, Lagos", "AFS page 3"],
    ["Bankers", "Access Bank Plc - A/C 1854793470 (current account)", "Bank statements"],
    ["VAT status", "Not registered at 31 December 2025; registration obligation arose in 2025", "VAT Act 2023 s.6"],
], widths=[3.6, 9.6, 3.4], fontsize=8.5)

callout("Configure the legal name with care.",
        "Zoho pre-fills the customer-facing name from the organisation profile. The bank mandate spelling "
        "('LIMTED') should be corrected with Access Bank before invoices and payment instructions go out under the "
        "new system, otherwise payment-matching failures will continue.", "FFC7CE")

pagebreak()

# =========================================================== 3 EVIDENCE
h1("3. What the historical files prove about the current bookkeeping")

h2("3.1 The ledger and the chart of accounts are disconnected")
p("The October 2026 chart of accounts contains 275 accounts. The transaction register for 2024-2025 uses 18 legacy "
  "codes, only one of which (1010) exists in the current chart at the same code. In practice the bookkeeping was "
  "run on a 33-account controlled chart, and the 275-account chart was created separately in Zoho. Until the "
  "transactions are mapped, every report out of Zoho will show the new accounts as unused and the P&L will not "
  "reconcile to the audited statements.")

table(["Legacy code", "What it held", "Lines", "% of all lines", "Problem"], [
    ["6050", "Bank charges, levies, airtime, telecom - and everything else that did not fit", "607", "49.5%",
     "Mixes bank commission, EMTL, VAT on charges, site expenses, fuel, transport, training and refunds"],
    ["5010", "Cost of sales - materials, furniture, site works and refunds", "226", "18.4%",
     "Materials, subcontractors and labour in one line - gross margin cannot be analysed"],
    ["2210", "Related party receivable / (payable) - Esther Akin-Ajayi", "186", "15.2%",
     "Every payment to the MD, whatever the purpose, including genuinely business-related costs"],
    ["4010", "Revenue - interiors, furniture, events and rentals", "72", "5.9%",
     "No distinction between service lines; IFRS 15 disaggregation impossible"],
    ["All others", "14 legacy accounts", "135", "11.0%", "Coarse but usable"],
], widths=[2.2, 5.8, 1.4, 1.8, 5.4], fontsize=8.5)

h2("3.2 The 24-point exception dashboard")
p("The master workbook records 2 passes, 13 failures and 9 flags against the audited statements. The failures that "
  "a better system would have prevented or detected are:")
for t in [
    "Receivables of N6.81m across the two years with no billing, delivery note or credit terms behind them (FAIL).",
    "Minimum tax and tertiary education levy of N235,622 not provided in either year (FAIL).",
    "Client advances of N3.46m recognised as revenue (FAIL - IFRS 15).",
    "No IAS 24 related-party disclosure for N5.05m and N9.60m of director payments (FAIL).",
    "Two unrecorded borrowings totalling N1.46m (FAIL - IFRS 9).",
    "No IAS 12 tax note at all (FAIL).",
    "Cash flow statement net change that does not agree to the movement in the bank balance (FAIL).",
    "VAT registration not obtained after turnover crossed N25m (FLAG - N2.26m exposure).",
    "WHT not withheld on qualifying supplier and professional payments (FLAG - up to N2.7m exposure).",
]:
    bullet(t)
p("Every one of these is a data-capture failure, not an accounting-judgement failure. Zoho Books, configured as set "
  "out in this guide, makes each of them structurally difficult: invoices create receivables, projects hold "
  "advances, vendors carry WHT, and the tax accounts are in the chart with a monthly report mapped to them.")

h2("3.3 The three-month test window used to prove the design")
p("The most recent three-month block of bank statement on file (1 August - 31 October 2025) contains 112 lines: "
  "N4,431,000 in and N4,430,917.59 out. Re-coded through the reviewed chart and the 110 posting rules, those 112 "
  "lines produce 224 balanced journal lines (N8,861,917.59 each way) and reconcile to the statement to the kobo: "
  "August net +N617.48, September +N308.29, October -N843.36. That is the proof that the design works before it is "
  "applied to 2026.")

table(["Month", "Lines", "Money in (NGN)", "Money out (NGN)", "Net (NGN)", "Statement agrees"], [
    ["August 2025", "34", "351,000.00", "350,382.52", "617.48", "Yes"],
    ["September 2025", "59", "3,510,000.00", "3,509,691.71", "308.29", "Yes"],
    ["October 2025", "19", "570,000.00", "570,843.36", "(843.36)", "Yes"],
    ["Total", "112", "4,431,000.00", "4,430,917.59", "82.41", "Yes"],
], widths=[3.4, 1.6, 3.2, 3.2, 2.6, 2.6], fontsize=8.5)

pagebreak()

# =========================================================== 4 WHY PREMIUM
h1("4. Why Zoho Books Premium, and the decisions required before configuration")

h2("4.1 The requirement set")
p("The system has to do eight things. The list is drawn from what the historical files show was missing, not from a "
  "generic feature comparison.")
table(["Requirement", "Why Jemai needs it", "Edition that carries it"], [
    ["Project accounting and job costing", "Every naira of revenue and cost in the 2024-2025 register belongs to a "
     "client project. Without projects there is no gross margin per contract - the single most important number in "
     "a fit-out business.", "Premium"],
    ["Budgets with variance", "Project quotes must be held as budgets so that cost overruns are visible while the "
     "job is running, not after the audit.", "Premium"],
    ["Approval workflows", "One person (the MD) authorises payments; the system must record who approved what, "
     "including director expense claims.", "Premium"],
    ["Custom roles and audit trail", "The auditor and the bookkeeper need different access. A single login cannot "
     "evidence segregation of duties.", "Premium"],
    ["Multiple taxes and tax groups", "7.5% VAT, WHT at several rates, development levy and minimum tax must all sit "
     "in the same ledger.", "Standard and above"],
    ["Bank feeds, rules and reconciliation", "1,226 statement lines in two years; 73 of the 112 test-window lines "
     "were bank charges. Manual coding is what failed.", "Standard and above"],
    ["Custom reports and reporting tags", "The IFRS statements need revenue by service line, cost by project and "
     "expense by nature - tags and custom reports make that a click.", "Professional and above"],
    ["Recurring invoices and payment reminders", "Retention releases, monthly retainers and rent are recurring "
     "transactions.", "Professional and above"],
], widths=[4.2, 8.4, 4.0], fontsize=8.5)

callout("Recommendation.",
        "Zoho Books Premium is the correct tier for Jemai. Standard and Professional would leave projects and budgets "
        "out, which is precisely the capability the business needs to stop losing margin on fit-out jobs. Elite adds "
        "features (advanced analytics, additional users) that a N30m-turnover company will not use in the first "
        "year. Confirm current plan inclusions, user counts and Nigerian pricing directly with Zoho before "
        "purchase - plan packaging changes and the features above should be verified against the live comparison "
        "page at the time of purchase.")

h2("4.2 Decisions required before configuration starts")
p("Configuration cannot begin until these are decided. Each carries a recommended default so that progress is not "
  "blocked.")

table(["#", "Decision", "Options", "Recommendation", "Owner"], [
    ["D1", "Opening balance basis", "(a) AFS as issued; (b) AFS plus prior-period adjustments (bank-evidenced 'View B')",
     "(b). Load the AFS as issued and post the single IAS 8 prior-period adjustment journal in deliverable 01, tab "
     "Open_Balances_1Jan2026, Block B. It is auditable, it preserves the signed statements and it moves the ledger to "
     "the position the evidence supports.", "MD + auditor"],
    ["D2", "VAT registration", "(a) Register now and disclose prior periods; (b) wait and carry the exposure",
     "(a). Turnover crossed N25m in 2025; the exposure grows every month (N50,000 first-month penalty plus N25,000 "
     "per month plus 10% interest). Register, then set the VAT switch in the posting engine to YES from the "
     "registration month.", "MD + tax adviser"],
    ["D3", "Director current account treatment", "(a) Keep as a receivable and pursue repayment; (b) convert part to "
     "salary/dividend; (c) impair",
     "(a) with a documented loan agreement and a quarterly review. N14.22m is 94% of the corrected balance sheet; "
     "without a repayment plan the going-concern note cannot be signed comfortably.", "MD"],
    ["D4", "Revenue recognition", "(a) Cash received; (b) percentage of completion; (c) on delivery/handover",
     "(c) point-in-time on handover for furniture, (b) for fit-out contracts, with advances held in contract "
     "liabilities. This is what IFRS 15 requires and what the AFS policy note already claims.", "MD + auditor"],
    ["D5", "Inventory", "(a) No inventory - direct expensing; (b) materials inventory with WIP",
     "(b) for materials held for a specific job (account 1061-1064), because the register shows deposits for "
     "furniture and curtains paid weeks before installation.", "MD"],
    ["D6", "WHT policy", "(a) Continue without deducting; (b) deduct at the statutory rate on every qualifying payment",
     "(b). Non-deduction does not save tax - it makes the expense non-allowable and exposes the company to penalty "
     "and interest as an agent of FIRS.", "MD + tax adviser"],
    ["D7", "Second bank account", "(a) Single account (current position); (b) second account for collections and a "
     "wallet for petty cash",
     "(b). A single account that also handles the director's personal flows is the root of the related-party "
     "problem.", "MD"],
    ["D8", "Go-live date and data cut-over", "(a) 1 January 2026 opening with August-October 2026 journals posted; "
     "(b) 1 August 2026",
     "(a). The opening balances are already computed at 1 January 2026; the posting journal deliverable covers "
     "August to October 2026 and January-July can be posted from the same rule set.", "MD"],
], widths=[0.9, 3.0, 4.4, 6.4, 2.0], fontsize=8)

pagebreak()

# =========================================================== 5 ORG SETTINGS
h1("5. Organisation profile and global settings")
p("Path: Settings (gear icon) > Organisation Profile, then Settings > Preferences, then Settings > General.")

h2("5.1 Organisation profile")
table(["Field", "Value"], [
    ["Organisation name / legal name", "Jemai International Limited"],
    ["Organisation type", "Business"],
    ["Industry", "Interior design / furniture / construction services"],
    ["Registered address", "24, Ashaolu Street, Off Akilo Road, Ogba, Lagos State, Nigeria"],
    ["Country", "Nigeria"],
    ["Currency", "NGN - Nigerian Naira (base currency; do not add foreign currencies until an FX transaction arises)"],
    ["Financial year", "January to December"],
    ["Tax ID / TIN", "2522579844227"],
    ["Business registration number", "RC 7171046"],
    ["Time zone", "Africa/Lagos (GMT+1)"],
    ["Date format", "dd-MMM-yyyy (this is the format of the bank statements and the working papers - consistency "
                    "prevents the day/month errors seen in the 2024 AFS notes)"],
    ["Fiscal year lock", "Lock 2024 and 2025. Set a period lock one month behind the current month as a standing rule."],
], widths=[5.4, 11.2], fontsize=8.5)

h2("5.2 Preferences that must be switched on")
table(["Setting", "Path", "Value / action", "Why"], [
    ["Projects", "Settings > Preferences > Projects", "Enable", "Job costing (section 10)"],
    ["Timesheets", "Settings > Preferences > Projects", "Enable", "Capturing staff time to projects where chargeable"],
    ["Budgets", "Settings > Preferences > Budgets", "Enable, and set the budget basis to 'Project'", "Variance against "
     "quote (section 16)"],
    ["Approvals", "Settings > Preferences > Approval", "Enable multi-level approval for bills and expenses",
     "Section 15"],
    ["Custom fields", "Settings > Preferences > Custom Fields", "See 5.3", "Related-party identification and IFRS 15"],
    ["Reporting tags", "Settings > Preferences > Reporting Tags", "Revenue line (5 values), Customer type (3), "
     "Funding source (2)", "Disaggregation without multiplying revenue accounts"],
    ["Multiple transactions numbering", "Settings > Preferences > Transaction Numbering", "Invoice: INV-YYYY-####; "
     "Bill: BILL-YYYY-####; Journal: JE-YYYY-MM-###; Credit note: CN-YYYY-####", "Journal numbering matches the "
     "posting engine"],
    ["Attachments", "Settings > Preferences > Attachments", "Allow up to 5 attachments per transaction", "Invoice "
     "and receipt evidence is the main control gap in the 2024-2025 file"],
    ["Auto-number duplicate bills", "Settings > Preferences", "On", "Prevents the same supplier invoice being entered "
     "twice"],
], widths=[3.4, 4.2, 4.6, 4.4], fontsize=8)

h2("5.3 Custom fields to create")
p("Custom fields are how a small chart of accounts stays small. Each field below replaces a family of accounts that "
  "would otherwise be created.")
table(["Module", "Field", "Type / values", "Purpose"], [
    ["Customers", "Relationship", "Dropdown: Client - corporate, Client - residential, Client - government, "
     "Director/related party, Supplier, Employee", "IAS 24 population; replaces the 'corporate/residential client "
     "revenue' accounts that were retired in the reviewed chart"],
    ["Vendors", "Vendor category", "Dropdown: Subcontractor, Materials, Professional, Landlord, Bank, Statutory, "
     "Related party", "WHT rate defaults and the supplier spend analysis"],
    ["Vendors", "WHT applicable rate", "Number (%)", "Deduction at source (section 12.4)"],
    ["Invoices", "Contract performance", "Dropdown: Advance received, Work certified, Handed over, Retainer",
     "IFRS 15 control - only 'Handed over' and 'Work certified' post to revenue without review"],
    ["Invoices", "Retention %", "Number", "Retention receivable 1042"],
    ["Expenses", "Paid on behalf of", "Dropdown: Company, Director, Related party, Client", "Stops director "
     "payments being expensed (the N14.22m error)"],
    ["Journals", "Source", "Dropdown: Bank statement, Payroll, Depreciation, Tax, Prior-year adjustment, Audit "
     "adjustment", "Reportable audit trail"],
    ["Projects", "Project code", "Text (PRJ-YYYY-###)", "Matches the client contract file"],
], widths=[2.4, 3.6, 5.6, 5.0], fontsize=8)

pagebreak()

# =========================================================== 6 COA
h1("6. Chart of accounts configuration and import")
p("Zoho Books creates a default chart on sign-up. Do not build on it: delete nothing, import deliverable 01 over "
  "the top, and deactivate what is not needed. The reviewed chart has 302 accounts, of which 7 are deliberately "
  "inactive.")

h2("6.1 What changed and why it matters")
table(["Change", "Count", "Effect on reporting"], [
    ["Duplicate account code resolved (2030 Retained Earnings -> 3090)", "1", "Zoho rejects duplicate codes at "
     "import; the trial balance sorts correctly; equity sits in the 3xxx block"],
    ["Account with a blank code given a code (License Fees -> 6419)", "1", "The account can be referenced by formulas "
     "and by the journal import"],
    ["Inverted hierarchy corrected (3060 Owner's Equity was a child of 3061)", "1", "The equity section displays "
     "correctly in the balance sheet"],
    ["Orphan accounts re-parented to their group header", "19", "P&L and balance-sheet group totals actually "
     "aggregate"],
    ["Revenue accounts nested under the revenue header", "25", "Revenue totals come from the ledger, not from a "
     "manual schedule"],
    ["Overlapping accounts deactivated", "7", "One concept, one account - no split balances"],
    ["New accounts for IFRS and Nigerian tax", "27", "Deferred tax, provisions, leases, ITF/NSITF/NHF, share "
     "application money, EMTL, transfer commission, ECL on related-party balances, air travel, related-party "
     "receivable"],
], widths=[7.6, 1.6, 7.4], fontsize=8.5)

h2("6.2 Import sequence")
p("Do this before any transaction is entered. The import is reversible (each account can be edited) but re-doing it "
  "after postings exist is painful.")
for t in [
    "Take a backup: Settings > Data Backup, and export the current chart of accounts.",
    "Fix the two blocking defects in the UI first, because Zoho matches on account name and will reject a duplicate "
    "code: change Retained Earnings from 2030 to 3090, and give 'License Fees' the code 6419.",
    "Accountant > Chart of Accounts > Import, and upload 01b_CoA_Zoho_Import.csv.",
    "Map the twelve columns to the Zoho fields of the same name: Account Name, Account Type, Account Code, "
    "Description, Parent Account, Currency, Account Number, Account Status, Mileage Rate, Mileage Unit, Is Mileage. "
    "Account ID is reference only and can be left unmapped.",
    "Choose 'Update existing accounts' when the option is offered, so the parent and status changes apply instead of "
    "creating duplicates.",
    "Verify against tab Zoho_Import_CoA: 302 accounts, zero duplicate codes, zero uncoded accounts, 7 inactive, and "
    "the tree matching parent for parent.",
    "Run Accountant > Chart of Accounts > export and compare row counts. Then open the Trial Balance and confirm it "
    "is empty with 302 accounts listed.",
]:
    numbered(t)

callout("Zoho hierarchy rule.",
        "A parent account and its children must share the same Account Type. Five logical groupings are therefore "
        "impossible and are documented rather than forced: 'Other Expense' (7023 Exchange gain or loss) cannot sit "
        "under any header; 'Purchase Discounts' (5071) is typed Expense but is economically a cost-of-sales contra; "
        "'Other Liability' (2999) cannot nest under a current-liability header; and Zoho keeps Stock and Cash "
        "accounts flat. Handle these by report grouping, not by adding accounts.")

pagebreak()

# =========================================================== 7 OPENING BALANCES
h1("7. Opening balances at 1 January 2026")

p("Opening balances are the point at which most Nigerian SME implementations fail. Zoho Books has a dedicated "
  "opening-balance function for contacts and items, but the ledger itself is opened with a single journal. Do it in "
  "this order.")

h2("7.1 The opening trial balance")
table(["Code", "Account", "Debit (NGN)", "Credit (NGN)"], [
    ["1021", "Access Bank - A/C 1854793470", "24,194.10", ""],
    ["1091", "Furniture & Fittings (cost)", "165,000.00", ""],
    ["1093", "Computers & IT Equipment (cost)", "360,053.75", ""],
    ["1101", "Accumulated depreciation - furniture", "", "56,184.75"],
    ["1103", "Accumulated depreciation - computers", "", "122,637.25"],
    ["1045", "Director current account receivable - Esther Akin-Ajayi", "14,223,846.92", ""],
    ["1046", "Other receivables", "332,672.02", ""],
    ["2011", "Trade payables - suppliers", "", "1,335,000.00"],
    ["2023", "Accrued professional fees", "", "300,000.00"],
    ["2031", "Contract liabilities - project deposits", "", "3,460,000.00"],
    ["2045", "Income tax payable (minimum tax)", "", "188,852.14"],
    ["2046", "Development levy payable (2025 TEL balance)", "", "85,475.07"],
    ["2051", "Short-term loan - individual lender", "", "230,000.00"],
    ["2052", "Short-term working capital loan", "", "1,230,000.00"],
    ["3011", "Ordinary share capital", "", "1,000,000.00"],
    ["3071", "Capital reserve", "", "2,000.00"],
    ["3090", "Retained earnings", "", "7,095,617.58"],
    ["", "TOTAL", "15,105,766.79", "15,105,766.79"],
], widths=[1.8, 9.4, 3.0, 2.4], fontsize=8.5)
p("The full working, including the alternative route (load the AFS as issued and post one IAS 8 prior-period "
  "adjustment journal of N4,897,706.58 to retained earnings), is in deliverable 01, tab Open_Balances_1Jan2026. The "
  "recommended route is the second: it preserves the signed statements as the audit trail and shows the correction "
  "as a single, explainable entry.", space=6)

h2("7.2 Entering the balances")
for t in [
    "Accountant > Manual Journal > New. Date 01-01-2026. Reference 'OB-2026'. Notes 'Opening balances per audited "
    "financial statements at 31 December 2025, restated for prior-period adjustments (IAS 8)'.",
    "Enter one line per account in the table above. If the recommended route is used, first enter the AFS-as-issued "
    "balances, then post the adjustment journal as a separate entry dated 01-01-2026 so the two are distinguishable.",
    "Do not use the Opening Balance Offset accounts (3080, 3999, 2090) for these postings - Zoho controls them, and "
    "a manual entry will break the opening-balance report.",
    "Post customer opening balances through the contact record (Customers > open the customer > Opening Balance) so "
    "that the receivable control account 1040 agrees to the subledger. The 2026 opening has no trade receivables, "
    "which is itself the corrected position.",
    "Post vendor opening balances through Vendors > Opening Balance (N1,335,000 across the suppliers in Appendix E).",
    "Check: Reports > Trial Balance at 01-01-2026 must equal the table above, and Reports > Balance Sheet must equal "
    "N14,926,944.79 of net assets (total debits less accumulated depreciation).",
]:
    numbered(t)

callout("Why the opening balance sheet is unusual.",
        "Ninety-four per cent of the opening assets are a receivable from the director. That is the corrected "
        "position after removing N6.81m of unsupported trade receivables and recognising N14.22m of owner payments as "
        "a receivable rather than an expense. If the MD instead treats part of those flows as remuneration, the "
        "opening balances must be restated before go-live - decision D3.", "FFC7CE")

pagebreak()

# =========================================================== 8 BANK
h1("8. Bank accounts, feeds, rules and reconciliation")

h2("8.1 Accounts to configure")
table(["Account", "Type", "Code", "Purpose"], [
    ["Access Bank - A/C 1854793470", "Bank", "1021", "Operating account - the only account evidenced in the record"],
    ["Bank Overdraft - Access Bank", "Bank Overdraft", "1029", "Created so an overdrawn position is disclosed gross "
     "(recurring overdrafts are the stated going-concern risk)"],
    ["Mobile Money / Wallet Clearing", "Other Current Asset", "1034", "OPay / Moniepoint / PalmPay settlements"],
    ["Payment Gateway Clearing", "Other Current Asset", "1033", "Paystack - the register shows outflow 'to Paystack "
     "checkout' and Paystack collections"],
    ["Interbank Transfer Clearing", "Other Current Asset", "1035", "Transfers in transit between the company's own "
     "accounts"],
], widths=[4.6, 3.4, 1.6, 7.0], fontsize=8.5)

h2("8.2 Feeding the statement in")
p("Zoho Books can pull Nigerian bank statements through a direct feed where the bank is supported, and can always "
  "import a CSV, OFX or QIF file. For Access Bank, confirm feed support during set-up; if it is not available, use "
  "the CSV import, which is what the posting engine is built around.")
for t in [
    "Banking > Access Bank > Import Statement. Upload the CSV in the Access Bank column order (date, narration, "
    "debit, credit, balance).",
    "Map the columns on the import screen and save the mapping as a template so that every month uses the same one.",
    "Import one month at a time. The statement must run from the first to the last day of the month and the closing "
    "balance must agree with the next month's opening balance before anything is categorised.",
    "Reconcile at import: Zoho shows the statement closing balance against the ledger balance. If they differ, stop "
    "and find the missing line before categorising anything.",
]:
    numbered(t)

h2("8.3 Bank rules - the most important configuration in this guide")
p("A bank rule matches the narration and posts the transaction to an account automatically. Because 73 of the 112 "
  "lines in the test window were bank charges, and every transfer in the register is followed by a commission line, "
  "a VAT line and periodic levy lines, rules do most of the work. Create them in this order (higher priority first) "
  "and use 'contains' criteria on the narration.")

table(["Priority", "Rule name", "Criterion (narration contains)", "Debit", "Credit", "Lines in test window"], [
    ["1", "Bank commission", "COMMISSION", "6915", "1021", "31"],
    ["2", "EMTL", "ELECTRONIC MONEY TRANSFER LEVY", "6916", "1021", "7"],
    ["3", "Account maintenance", "ACCOUNT MAINTENANCE", "6910", "1021", "3"],
    ["4", "SMS alert", "SMS ALERT", "6910", "1021", "4"],
    ["5", "VAT on bank charges", "VAT", "6914", "1021", "28"],
    ["6", "Penal charge", "PENAL CHARGE", "6910", "1021", "1"],
    ["7", "Debit interest", "DEBIT INTEREST", "6911", "1021", "1"],
    ["8", "Director payment", "ESTHER", "1045", "1021", "8"],
    ["9", "Related party payment", "AKINLOLUWA / AKINYEMI", "1047", "1021", "1"],
    ["10", "Salary", "SALARY", "6110", "1021", "2"],
    ["11", "Painting subcontractor", "PAINTING", "5026", "1021", "3"],
    ["12", "Furniture subcontractor", "FURNITURE / AGEGE", "5021", "1021", "2"],
    ["13", "Site visit / travel", "SITE VISIT / FLIGHT", "6511 / 6515", "1021", "2"],
    ["", "Remaining 98 rules", "see deliverable 02, tab Mapping_Rules", "110 rules in total", "", ""],
], widths=[1.4, 3.2, 4.4, 2.4, 1.4, 3.2], fontsize=8)

callout("Rule order matters.",
        "Every transfer narration contains the payee and the purpose, and the commission and VAT lines repeat the "
        "same narration. If the 'Director payment' rule outranks the 'Bank commission' rule, a 25-naira commission "
        "will be posted to the director's account. Commission, EMTL, maintenance, SMS and VAT must always be the "
        "highest-priority rules.", "FFC7CE")

h2("8.4 Reconciliation cadence")
for t in [
    "Daily: review and categorise each new bank line. Unreviewed lines are the raw material of the next audit "
    "adjustment.",
    "Weekly: reconcile to the bank statement. A reconciliation that takes more than 30 minutes is a signal that a "
    "rule is missing.",
    "Monthly: lock the reconciliation before the period is closed; produce the Bank Reconciliation report and file it "
    "with the month-end pack.",
    "Standing control: accounts 1033, 1034 and 1035 must return to nil at each month end. A residual balance is a "
    "timing difference that must be listed and cleared.",
]:
    bullet(t)

pagebreak()

# =========================================================== 9 CONTACTS
h1("9. Contacts: customers, vendors, related parties")

h2("9.1 Customers to create from the bank evidence")
p("The project register in the master workbook identifies the clients behind the largest collections. Create these "
  "first; the rest can be added as contracts are signed.")
table(["Customer", "Evidence (2024-2025 collections)", "Suggested project code", "Relationship tag"], [
    ["The Victoria King (TVK)", "N5,810,000", "PRJ-TVK-01", "Client - residential"],
    ["Joshua Oluwadololemi Abolade", "N3,340,000", "PRJ-JOA-01", "Client - residential"],
    ["Hassan Umar Abba", "N4,225,000", "PRJ-HUA-01", "Client - residential"],
    ["Kehinde Elizabeth Adeshakin", "N3,062,000", "PRJ-KEA-01", "Client - residential"],
    ["Ibrahim Junaidu", "N3,000,000", "PRJ-IJ-01", "Client - residential"],
    ["Ekpunobi Judith", "N2,484,000", "PRJ-EJ-01", "Client - residential"],
    ["Amaowoh Obongibiangeke", "N3,392,000", "PRJ-AO-01", "Client - residential"],
    ["Famakinwa Moyosore", "N1,540,000", "PRJ-FM-01", "Client - residential"],
    ["CIIFA SafeHouse", "N1,300,000", "PRJ-CIIFA-01", "Client - corporate"],
    ["Auxit Media", "N700,000", "PRJ-AM-01", "Client - corporate"],
    ["Venture Garden Nigeria (VigiPay)", "N380,000 (Oct-2025)", "PRJ-VGN-01", "Client - corporate"],
    ["Paystack / NIP collections (unattributed)", "N1,827,000", "PRJ-GEN-01", "Client - unidentified"],
], widths=[4.8, 4.0, 3.4, 3.6], fontsize=8)

h2("9.2 Vendors")
p("Create a vendor for every counterparty that is paid more than once, and set the vendor category so the WHT rate "
  "defaults correctly:")
for t in [
    "Subcontractors (5021-5029): Adeniyi Emmanuel Olusegun (furniture), Oluwadamilare Samuel Adeyeye (painting), "
    "Samuel Obi Chukwu (electrical), Destiny Ajiri Osaremwinda (cleaning), Bello Opeyemi Esther (mirrors), Imoikor "
    "Christiana Imade (fluted panels), David Onyekachukwuomeke (curtains), Ukaire Jidech, De Great Ify-God Venture "
    "(electrical fittings), Stollage Enterprise.",
    "Professional (641x): auditors Gbenga Adeyeye & Co, tax adviser, legal counsel, Covenant University consultancy.",
    "Statutory: FIRS/NRS, Lagos State LIRS (PAYE), pension administrator, ITF, NSITF, FMBN.",
    "Landlord: the Ogba office landlord (rent N200,000 p.a. - note it has no bank trace in either year; establish "
    "whether the MD pays it personally).",
    "Related party: Esther Akin-Ajayi (director), Emmanuel Akinloluwa Akin-Ajayi, Akinyemi Akin-Ajayi, Felicia, Ime "
    "Ubongabasi (former shareholder). Mark the relationship field; these contacts must appear in the IAS 24 "
    "disclosure, not in the ordinary vendor analysis.",
]:
    bullet(t)

h2("9.3 The related-party rule that must be enforced")
p("No payment to a director or a related party is ever posted to an expense account. It is posted to 1045 (director) "
  "or 1047 (family and associates) and cleared later by a properly documented expense claim or by repayment. This "
  "single rule prevents the largest error in the audited history: N5.05m and N9.60m of owner payments expensed as if "
  "they were the company's costs.")

pagebreak()

# =========================================================== 10 PROJECTS
h1("10. Projects, budgets and job costing")
p("This is the module that justifies the Premium tier. In a fit-out business the only number that matters monthly is "
  "gross margin per contract. Without projects, every cost sits in a company-wide pool and margin can only be "
  "calculated after the audit.")

h2("10.1 Setting up a project")
for t in [
    "Projects > New Project. Name it with the client and the scope ('TVK - Interior decoration, Ikoyi').",
    "Project code PRJ-YYYY-## in the custom field, matching the contract file.",
    "Customer: the customer created in section 9.1.",
    "Billing method: 'Fixed cost for project' for a lump-sum fit-out; 'Based on project hours' only where staff time "
    "is genuinely chargeable.",
    "Budget: enter the quoted cost by category (materials, subcontractors, site operations, design delivery) - not a "
    "single number. Category budgets are what make the variance report useful.",
    "Attach the signed contract, the quote and the client's payment schedule to the project.",
]:
    numbered(t)

h2("10.2 Tagging transactions to projects")
table(["Transaction", "How to tag", "Effect"], [
    ["Client invoice / advance", "Select project on the invoice line", "Revenue and contract liability by project"],
    ["Vendor bill for materials", "Select project on the bill line", "Direct cost 5011-5019 by project"],
    ["Subcontractor payment", "Select project", "Direct cost 5021-5029 by project"],
    ["Site transport, cleaning, fuel", "Select project", "Site operations 5030-5037 by project"],
    ["3D rendering / presentation", "Select project", "Design delivery 5041-5044 by project"],
    ["Petty cash issued to site", "Select project on the expense", "5031 by project"],
], widths=[4.4, 4.4, 7.0], fontsize=8.5)

h2("10.3 Reports to run every month")
for t in [
    "Projects > Project Summary: revenue, cost, margin and budget consumption per contract.",
    "Projects > Project Profitability: the same on a percentage basis; anything below the quoted margin must be "
    "explained by the project lead.",
    "Projects > Budget vs Actual: line-by-line consumption; investigate any category above 90% before the next "
    "purchase order.",
    "Projects > Timesheet Summary (if timesheets are used): recoverable versus non-recoverable time.",
]:
    bullet(t)

pagebreak()

# =========================================================== 11 SALES
h1("11. Sales cycle: quotes, invoices, retentions, advances and credit notes")

h2("11.1 The document flow")
table(["Step", "Zoho document", "Accounting effect", "Control"], [
    ["Quote issued", "Estimate", "None", "The estimate line items become the project budget template"],
    ["Advance / deposit received", "Retainer invoice or a customer payment on account, tagged 'Advance received'",
     "Dr 1021 / Cr 2031 Contract liabilities - project deposits", "No revenue. This is the error that put N3.46m of "
     "deposits into turnover"],
    ["Work certified / milestone", "Invoice, tagged 'Work certified'", "Dr 1041 / Cr 4012 (and 2041 VAT once "
     "registered)", "Invoice only against a signed valuation"],
    ["Handover", "Final invoice", "Dr 1041 / Cr 4012-4019", "Handover certificate attached"],
    ["Retention withheld by client", "Invoice line with retention %", "Dr 1042 Retention receivables", "Tracked "
     "separately until release"],
    ["Retention released", "Customer payment applied to 1042", "Dr 1021 / Cr 1042", "Never recognised as revenue "
     "again (retired account 4023)"],
    ["Customer refund", "Credit note", "Dr 4042 / Cr 1021 or 2064", "Requires the original invoice reference"],
], widths=[3.0, 3.4, 4.4, 5.0], fontsize=8.5)

h2("11.2 Invoices")
for t in [
    "Items: create service items per revenue line (interior fit-out, furniture supply, design fee, project management, "
    "3D visualisation, installation). Do not create an item per client.",
    "Tax: leave the invoice tax as 'None' until the VAT certificate is issued, then apply the 7.5% VAT rate. Never "
    "type a VAT amount manually.",
    "Template: include RC 7171046, TIN 2522579844227, the registered address, the bank account and the payment terms. "
    "This is what FIRS and the auditor look for first.",
    "Currency: NGN only. If an imported furniture supplier invoices in USD, record the purchase in USD on the vendor "
    "bill and let Zoho revalue it; do not hand-convert in the narration.",
    "Attachments: the signed contract or purchase order must be attached to every invoice above N500,000.",
]:
    bullet(t)

h2("11.3 Contract liabilities and revenue recognition control")
p("At each month end, run the Contract Liabilities report (the balance on 2031 by project) and compare it to the "
  "project register. A project that is complete but still carries a contract liability means revenue is understated; "
  "a project that has started but carries revenue and no cost means the reverse. Two of the thirteen audit failures "
  "in the historical file originate here.")

pagebreak()

# =========================================================== 12 PURCHASES
h1("12. Purchase cycle, expenses, staff claims and withholding tax")

h2("12.1 The flow")
table(["Step", "Document", "Accounting effect", "Note"], [
    ["Materials or subcontractor ordered", "Purchase order (project-tagged)", "None", "PO approval above N250,000"],
    ["Supplier delivers and invoices", "Vendor bill", "Dr 5011-5029 / Cr 2011", "Bill date = invoice date, not "
     "payment date (accruals)"],
    ["WHT deducted", "Bill-level WHT or a separate journal", "Dr 2011 / Cr 2043", "Section 12.4"],
    ["Bill approved", "Approval workflow", "None", "Two-step: preparer, MD"],
    ["Payment made", "Vendor payment", "Dr 2011 / Cr 1021", "Applied to the bill, not to the account"],
    ["Retention withheld from subcontractor", "Bill line", "Cr 2015 Retention payables", "Released on defect "
     "clearance"],
], widths=[3.4, 3.4, 4.4, 5.0], fontsize=8.5)

h2("12.2 The deposits problem")
p("The 2025 register shows deposits paid to furniture and curtain suppliers (N1,050,000 and N300,000 in September) "
  "weeks before delivery. If those are posted as cost on payment, the P&L shows a loss in September and a windfall "
  "at handover. Post them as supplier advances (1072) or inventory (1061-1064) and release to cost when the goods are "
  "delivered to site.")

h2("12.3 Staff and director expense claims")
p("This is where the biggest historical error is fixed. The director's spending is captured as an expense claim "
  "against her current account, with receipts attached:")
for t in [
    "Expenses > New Expense. Paid through: 'Employee' (or 'Director'), not 'Bank account'.",
    "Custom field 'Paid on behalf of' = Director. This posts Dr 1045 / Cr 1044 (staff advances) or directly to 1045.",
    "Attach the receipt. No receipt, no reclassification out of 1045.",
    "The claim is approved by a second person where one exists; where the MD is the claimant, the auditor reviews the "
    "quarterly reconciliation.",
    "Never post a director claim to a cost account directly - that was the N14.22m error.",
]:
    bullet(t)

h2("12.4 Withholding tax on purchases")
p("Where applicable, WHT is deducted at invoice level and remitted by the 21st of the following month. Zoho Books "
  "does not operate Nigeria's WHT regime as a native module in the way it does Indian TDS, so the working method is a "
  "bill-level journal:")
table(["Transaction", "Debit", "Credit", "Timing"], [
    ["Supplier bill received gross", "5011-5029 (gross)", "2011 (gross)", "Invoice date"],
    ["WHT deducted at source", "2011 (WHT amount)", "2043 WHT payable", "Invoice date"],
    ["Supplier paid net of WHT", "2011 (net amount)", "1021", "Payment date"],
    ["WHT remitted to FIRS", "2043", "1021", "By the 21st"],
    ["WHT credit note received", "1082 WHT credit receivable", "2043 or 6713", "On receipt of the credit note"],
], widths=[4.4, 3.6, 3.6, 4.2], fontsize=8.5)
callout("Confirm the rates.",
        "The rates in the posting engine (5% contracts and professional services, 10% rent, 2% goods) reflect the "
        "Deduction of Tax at Source (Withholding) Regulations as applied in the working papers. The Regulations "
        "include exemptions (small companies, low-value transactions) and the position changed with the Nigeria Tax "
        "Act 2025. Confirm the applicable rate for each vendor category with the tax adviser before the first "
        "deduction - Appendix B lists the parameters that need sign-off.")

pagebreak()

# =========================================================== 13 PAYROLL
h1("13. Payroll and statutory deductions")
p("Zoho Books does not operate Nigerian payroll legislation natively. Two workable models, in order of preference:")

h2("13.1 Model A - payroll outside Zoho, posted as a monthly journal (recommended)")
p("Use a payroll schedule (spreadsheet or Zoho People) to compute gross pay, PAYE, pension, NHF, ITF and NSITF, then "
  "post one journal per month:")
table(["Line", "Account", "Debit", "Credit"], [
    ["Gross salaries and wages", "6110 Salaries and Employee Cost", "X", ""],
    ["Employer pension contribution (10%)", "6117 Pension Expense", "Y", ""],
    ["Employer ITF (1%) and NSITF (1%)", "6115 Staff Training (or 6113)", "Z", ""],
    ["PAYE withheld", "2044 PAYE payable", "", "P"],
    ["Employee pension contribution (8%)", "2072 Pension payable", "", "Q"],
    ["NHF withheld (2.5% of basic)", "2076 NHF payable", "", "R"],
    ["Net pay", "2071 Salaries payable (or directly 1021 on payment)", "", "S"],
    ["Total", "", "X+Y+Z", "P+Q+R+S"],
], widths=[5.0, 5.6, 2.6, 2.6], fontsize=8.5)
p("Then post the remittances as separate entries: Dr 2044 / Cr 1021 (PAYE, by the 10th), Dr 2072 / Cr 1021 (pension, "
  "by the 7th working day), Dr 2074/2075/2076 / Cr 1021 (ITF, NSITF, NHF as they fall due).")

h2("13.2 Model B - Zoho Books 'Salary' journals")
p("If payroll stays simple (two employees), create a recurring journal template per employee and clone it monthly. "
  "This is quicker to run but harder to audit; the payroll schedule must still exist as the underlying record.")

callout("The 2025 salary anomaly.",
        "The AFS reports N1,181,000 of salaries and a separate N650,000 of staff training. The bank register shows "
        "salary credits to two named beneficiaries and a N650,000 gap that equals the training line exactly. Before "
        "payroll goes live, tie every salary credit in the 2025 statement to a named employee and confirm whether the "
        "N650,000 was a benefit paid through the MD's account. Payroll that is not evidenced is a non-allowable "
        "expense and a PAYE exposure.", "FFC7CE")

pagebreak()

# =========================================================== 14 TAX
h1("14. Tax configuration: VAT, WHT, CIT, development levy, deferred tax")

h2("14.1 VAT")
for t in [
    "Register with FIRS/NRS if not already done (decision D2). Zoho: Settings > Taxes > New Tax, name 'VAT 7.5%', "
    "rate 7.5%, type 'Tax on sales and purchases', linked to account 2041 VAT Output Payable.",
    "Input VAT goes to 1081 VAT Input Receivable. Bank-charge VAT goes to 1086 once registered (rule BNK-02R), and "
    "to 6914 while unregistered (rule BNK-02).",
    "Set the tax to apply to all taxable supplies. Zero-rated and exempt items (medical, exported services) get their "
    "own tax with 0%.",
    "Filing: monthly VAT return by the 21st of the following month. Zoho's Nigerian VAT return report should be "
    "checked against a manual computation in the first three months; where the report is not available for the "
    "country edition, run Reports > Tax Summary and reconcile to the 2041/1081 balances by hand.",
    "Controls: 2041 must equal 7.5/107.5 of taxable turnover; 1081 must equal input VAT evidenced by valid tax "
    "invoices only; 2042 VAT Control is the net settlement account and must clear to nil when the return is filed "
    "and paid.",
]:
    bullet(t)

h2("14.2 Company income tax, minimum tax and development levy")
table(["Obligation", "Basis", "Rate used in this setup", "Account", "Due date"], [
    ["Company income tax", "Assessable profit", "0% small-company tier (turnover and asset tests met historically) - "
     "confirm the NTA 2025 thresholds", "6710 / 2045", "Within 6 months of year end"],
    ["Minimum tax", "Gross income, where CIT computes below it", "0.5% of gross income", "6710 / 2045",
     "With the CIT return"],
    ["Development levy (replaced TEL)", "Assessable profit", "4% under NTA 2025 - confirm applicability and any "
     "small-company exemption", "6711 / 2046", "With the CIT return"],
    ["Deferred tax", "Temporary differences", "IAS 12 at the expected future rate", "1190 / 2096",
     "Year end only"],
], widths=[3.4, 3.0, 4.6, 2.2, 3.0], fontsize=8.5)
callout("Tax parameters need sign-off.",
        "The 2026 year is the first full year under the Nigeria Tax Act 2025 and the Nigeria Tax Administration Act "
        "2025. The bands, the minimum tax mechanics, the development levy rate and the small-company thresholds all "
        "changed. The rates above are the working assumptions carried from the FY2024-2025 file; each one must be "
        "confirmed by the tax adviser before the first 2026 provision is posted. They are not advice.", "FFC7CE")

h2("14.3 Tax calendar to load into Zoho (recurring reminders)")
table(["Month", "Filing / payment", "System action"], [
    ["Every month, 10th", "PAYE remittance to LIRS", "Task on the journal with account 2044"],
    ["Every month, 21st", "VAT return and payment to FIRS", "Task on the VAT report"],
    ["Every month, 21st", "WHT remittance to FIRS (deducted in the prior month)", "Task on the 2043 balance"],
    ["Every month, 7th working day", "Pension remittance to the PFA", "Task on 2072"],
    ["31 January", "Annual employer return (PAYE)", "Task"],
    ["30 April", "CIT / development levy payment for the prior year", "Task on 2045/2046"],
    ["30 April", "Audited financial statements and tax returns filed", "Task"],
    ["30 June", "CAC annual return", "Task on 6414"],
], widths=[4.0, 7.6, 5.0], fontsize=8.5)

pagebreak()

# =========================================================== 15 CONTROLS
h1("15. Workflows, approvals, roles and period locks")

h2("15.1 Approval rules to configure")
table(["Module", "Condition", "Approval", "Reason"], [
    ["Bills", "Amount > N250,000", "MD approval before payment", "Highest-value outflow control"],
    ["Bills", "Vendor category = Related party", "MD + independent review", "IAS 24 discipline"],
    ["Expenses", "Paid on behalf of = Director", "MD approval + receipt mandatory", "Prevents the N14.22m error "
     "recurring"],
    ["Expenses", "Category = Subcontractor or Professional", "MD approval", "These carry WHT obligations"],
    ["Journals", "Source = Prior-year adjustment or Audit adjustment", "MD + auditor review", "No unexplained "
     "adjustments after close"],
    ["Purchase orders", "Amount > N250,000", "MD approval", "Commits the project budget"],
    ["Credit notes", "Any", "MD approval", "Refunds were N890,161 in 2025 and were not controlled"],
], widths=[2.6, 4.6, 4.2, 5.2], fontsize=8.5)

h2("15.2 Roles")
table(["User", "Role", "Access", "Rationale"], [
    ["MD (Esther Akin-Ajayi)", "Admin", "Full, including approval authority", "Only signatory today"],
    ["Bookkeeper / accountant", "Custom role: create and edit transactions, no delete, no period unlock",
     "Day-to-day entry, bank reconciliation, reporting", "Segregation of duties within a small team"],
    ["Auditor (Gbenga Adeyeye & Co)", "Custom role: read-only + reports + export", "Review access during the audit "
     "window, removable afterwards", "Replaces the request-for-spreadsheets cycle"],
    ["Tax adviser", "Read-only, reports only", "VAT, WHT and CIT schedules", "Filing support"],
    ["Project lead (if appointed)", "Custom role: projects and timesheets only", "No bank or vendor access",
     "Cost capture at source"],
], widths=[3.6, 5.0, 4.0, 4.0], fontsize=8.5)

h2("15.3 Period locks and the closing discipline")
for t in [
    "Lock each month on the 10th working day of the following month, after the bank reconciliation, VAT, WHT and "
    "payroll entries are posted.",
    "Keep one unlocked month only. Two open months is the point at which adjustments stop being traceable.",
    "Only the MD can unlock a period, and the unlock is reported in the audit trail - export it as part of the "
    "year-end pack.",
    "Take a data backup before and after every year-end close.",
]:
    bullet(t)

pagebreak()

# =========================================================== 16 REPORTING
h1("16. Budgets, analytics and the IFRS reporting pack")

h2("16.1 The monthly pack and where each report comes from")
table(["Report", "Zoho path", "Answers"], [
    ["Statement of profit or loss", "Reports > Profit and Loss (accrual, with comparative)", "Did we make a margin?"],
    ["Statement of financial position", "Reports > Balance Sheet", "What do we own and owe?"],
    ["Cash flow (indirect)", "Reports > Cash Flow", "Where did the cash go? (the AFS cash flow did not tie - this "
     "one does)"],
    ["Revenue by service line", "Reports > Sales by Item / by Reporting Tag", "IFRS 15 disaggregation"],
    ["Direct cost and gross margin by project", "Reports > Projects > Project Profitability", "Which jobs made money"],
    ["Expense by nature", "Reports > Expense Details, grouped by account", "IFRS Note 11 administrative expenses"],
    ["Trade receivables ageing", "Reports > Receivables Aging", "Recoverability and ECL"],
    ["Trade payables ageing", "Reports > Payables Aging", "Supplier pressure and WHT completeness"],
    ["Tax summary (VAT)", "Reports > Taxes > Tax Summary", "VAT return and 2041/1081 tie-out"],
    ["Trial balance", "Reports > Trial Balance", "Feeds the audit file"],
    ["General ledger", "Reports > General Ledger", "Line-by-line audit evidence"],
    ["Audit trail", "Settings > Audit Trail / Activities", "Who changed what, when"],
    ["Budget vs actual", "Reports > Budget vs Actual", "Variance to the approved budget"],
    ["Custom report - related parties", "Reports > Custom Reports, filtered on the Relationship custom field",
     "IAS 24 disclosure"],
], widths=[4.6, 6.4, 5.6], fontsize=8.5)

h2("16.2 Mapping the ledger to the audited statements")
p("Deliverable 01, tab Reporting_Map, maps all 302 accounts to an IFRS statement line, a note reference, a tax tag "
  "and a cash-flow category. Load this mapping into the report groupings so that the P&L and balance sheet in Zoho "
  "read the same as the audited statements. Without it, every report has to be re-mapped by hand at year end - which "
  "is precisely what happened for 2024 and 2025.")

h2("16.3 Budgets")
for t in [
    "Create one annual budget for the company by account, and one budget per project by cost category.",
    "Review budget vs actual monthly with the MD; the discussion should be about the two largest adverse variances, "
    "not the whole list.",
    "Rebuild the annual budget at mid-year if the project pipeline changes by more than 20%.",
]:
    bullet(t)

pagebreak()

# =========================================================== 17 GO LIVE
h1("17. Data migration and go-live plan")

h2("17.1 Sequence and timing")
table(["Week", "Activity", "Deliverable", "Owner"], [
    ["1", "Decisions D1-D8 signed off; Zoho subscription purchased; organisation profile configured",
     "Signed decision log", "MD"],
    ["1", "Chart of accounts imported and verified", "302-account chart live", "Implementer"],
    ["2", "Contacts, projects and items created; custom fields and reporting tags configured",
     "Master data file", "Implementer + MD"],
    ["2", "Taxes, approvals, roles and workflows configured", "Configuration checklist", "Implementer"],
    ["3", "Opening balances posted at 01-Jan-2026 and agreed to the audited statements",
     "Opening trial balance signed", "Implementer + MD"],
    ["3", "Bank feed / statement import mapping tested with the August 2026 statement",
     "First reconciliation", "Bookkeeper"],
    ["4", "January-July 2026 bank transactions posted using the reviewed rules", "Ledger to 31-Jul-2026",
     "Bookkeeper"],
    ["4", "August-October 2026 posting journal imported (deliverable 02)", "Ledger to 31-Oct-2026", "Bookkeeper"],
    ["5", "Parallel run: Zoho P&L vs the manual record for August-October; investigate every difference",
     "Reconciliation memo", "Implementer + MD"],
    ["6", "Training (2 sessions), month-end close rehearsal, sign-off", "Signed go-live certificate",
     "MD + implementer"],
], widths=[1.4, 8.0, 4.6, 2.8], fontsize=8.5)

h2("17.2 What to migrate, and what not to")
table(["Migrate", "Do not migrate"], [
    ["Opening trial balance at 01-Jan-2026", "The 2024 and 2025 transaction detail - those years are closed and "
     "audited; keep the working papers as the archive"],
    ["Customer and vendor master data with opening balances", "Obsolete customers with one bank receipt and no "
     "contract"],
    ["Project master data for live contracts, with budgets", "Completed 2024-2025 projects, other than as reference "
     "notes"],
    ["Item list (service items for revenue, materials categories for cost)", "An item per stock-keeping unit - Jemai "
     "does not hold meaningful stock"],
    ["Recurring journals (payroll, depreciation, rent, subscriptions)", "Historical journals - the opening balance "
     "summarises them"],
], widths=[8.3, 8.3], fontsize=8.5)

h2("17.3 Go-live acceptance criteria")
for t in [
    "Trial balance at 01-01-2026 agrees to the audited statements (with the prior-period adjustment shown separately).",
    "Bank reconciliation for each month to 31-10-2026 shows zero unreconciled items and the statement balance.",
    "Every bank line in the period has a category; suspense accounts (1991, 1992, 4999, 7010, 7999) are nil.",
    "The P&L for August-October 2026 agrees to an independent recomputation from the bank statement.",
    "VAT, WHT, PAYE and pension balances agree to the underlying schedules.",
    "The MD can produce the monthly pack without help.",
]:
    bullet(t)

pagebreak()

# =========================================================== 18 CLOSE
h1("18. Month-end and year-end close procedures")

h2("18.1 Month-end checklist (target: working day 7)")
table(["#", "Step", "Report / evidence", "Owner"], [
    ["1", "Import and reconcile the bank statement; confirm zero unreconciled items", "Bank Reconciliation", "Bookkeeper"],
    ["2", "Categorise every line; clear the clearing accounts 1033/1034/1035", "Banking > Transactions", "Bookkeeper"],
    ["3", "Issue all client invoices for work certified; raise retainer/advance invoices for deposits received",
     "Invoices list", "Bookkeeper"],
    ["4", "Enter all vendor bills received; accrue unbilled supplier costs", "Bills list", "Bookkeeper"],
    ["5", "Post subcontractor and professional WHT journals", "WHT schedule", "Bookkeeper"],
    ["6", "Post payroll and the statutory remittances", "Payroll journal", "Bookkeeper"],
    ["7", "Review the director current account; post claims with receipts", "1045 movement schedule", "MD"],
    ["8", "Release supplier advances and inventory to cost for goods delivered", "Project cost report", "Bookkeeper"],
    ["9", "Review contract liabilities against the project register", "Contract Liabilities report", "MD"],
    ["10", "Post depreciation (monthly, or at least quarterly)", "PPE schedule", "Bookkeeper"],
    ["11", "Compute and post VAT (dr 2041/cr 1081 with the net to 2042)", "VAT return", "Bookkeeper"],
    ["12", "Run the monthly pack (P&L, balance sheet, cash flow, project profitability)", "Monthly pack", "Bookkeeper"],
    ["13", "Lock the period", "Lock confirmation", "MD"],
], widths=[0.9, 8.4, 4.4, 2.6], fontsize=8.5)

h2("18.2 Year-end additions")
for t in [
    "Expected credit loss assessment on 1045 and 1046 (IFRS 9) - the largest judgement in the accounts.",
    "Contract-by-contract revenue and cost review against IFRS 15 at the reporting date.",
    "Capital allowances computation on the PPE classes (NTA 2025 rates) and the deferred tax calculation (IAS 12).",
    "Provisions: warranty (2065) and any regulatory penalties (2066).",
    "Related-party schedule (IAS 24) from the custom field, not from a manual search of narations.",
    "Going-concern assessment, with the director's repayment plan for 1045 documented.",
    "Auditor access and the export pack: trial balance, GL, project profitability, tax schedules, reconciliations, "
    "audit trail.",
]:
    bullet(t)

pagebreak()

# =========================================================== 19 CONTROLS
h1("19. Controls, governance and audit readiness")
table(["Risk", "Control in Zoho", "Test the auditor will apply"], [
    ["Revenue overstated (advances treated as income)", "Invoice custom field 'Contract performance'; contract "
     "liability account 2031; month-end contract liability review", "Agree 2031 to the unperformed contracts in the "
     "project register"],
    ["Unsupported receivables", "Receivables arise only from invoices; no manual journal to 1041", "Sample invoices "
     "back to signed contracts and delivery evidence"],
    ["Director payments expensed", "1045 with expense-claim workflow and mandatory receipts", "Agree 1045 movement to "
     "bank lines to the MD's accounts"],
    ["Unrecorded borrowings", "All loans drawn and repaid through 2051/2052/2053 with loan agreements attached",
     "Agree loan balances to agreements and bank evidence"],
    ["Tax not provided", "Monthly tax dashboard; 2044/2045/2046/2043 balances reviewed at close", "Agree tax "
     "payables to computations and payment receipts"],
    ["Cash flow does not tie", "Cash flow report generated from the ledger; opening and closing cash agreed to the "
     "bank", "Recompute the movement in cash"],
    ["Unauthorised payments", "Approval workflow above N250,000 and for related parties", "Inspect the audit trail "
     "for approver identity"],
    ["Backdating and late adjustments", "Period locks and the 'Source' custom field on journals", "List journals "
     "posted after the lock date"],
], widths=[4.4, 6.4, 5.8], fontsize=8.5)

p("Document retention: CAMA and FIRS require accounting records to be kept for six years; the company should keep "
  "the 2024-2025 audit file permanently. Zoho retains data while the subscription is live - take a monthly backup "
  "export and store it outside the system.")

pagebreak()

# =========================================================== 20 RISKS
h1("20. Risks, open items and assumptions")

h2("20.1 Risks that could stop a clean go-live")
table(["#", "Risk", "Impact", "Mitigation", "Owner"], [
    ["R1", "The August-October 2026 journal is built from bank evidence only - no accruals, depreciation, "
     "revenue cut-off, payroll liabilities or tax provisions", "The period result in the ledger is a "
     "cash-shaped proxy and will move once the closing entries are posted", "Post the month-end accrual pack "
     "(section 18) before drawing any conclusion from the P&L", "Bookkeeper + MD"],
    ["R2", "VAT registration status and any historical exposure unresolved", "Registration changes every receipt and "
     "purchase posting; the exposure accrues monthly", "Decide D2 before the August 2026 journal is posted", "MD + "
     "tax adviser"],
    ["R3", "The director current account (N14.22m) has no repayment plan or documentation", "Going concern and ECL "
     "cannot be concluded; the balance sheet is 94% one receivable", "Sign a loan agreement and a repayment schedule; "
     "review quarterly", "MD"],
    ["R4", "Only one person operates the system", "No segregation of duties", "Approval workflows plus auditor "
     "read-only access; external review of the related-party reconciliation", "MD"],
    ["R5", "Tax rates for 2026 are not signed off", "Wrong provision, wrong returns", "Appendix B sign-off with the "
     "tax adviser before the first 2026 provision", "MD + adviser"],
    ["R6", "Supplier invoices and contracts are largely missing (N34m of cost in 2024-2025 unsupported)", "Costs are "
     "non-allowable and the audit qualification repeats", "Invoice-and-contract-first policy at bill entry; no "
     "bill without a document", "Bookkeeper"],
], widths=[0.8, 4.6, 4.0, 4.6, 2.6], fontsize=8)

h2("20.2 Assumptions in the deliverables")
for t in [
    "The functional and presentation currency is NGN and no foreign-currency transaction arose in the periods "
    "reviewed.",
    "The 1 January 2026 opening position is the corrected (bank-evidenced) statement of financial position at "
    "31 December 2025, with the prior-period adjustment shown separately.",
    "Accumulated depreciation of N178,822 is allocated across PPE classes in proportion to cost; the client may "
    "reallocate once the fixed-asset register is agreed.",
    "Bank statements are the primary evidence; where the bank and the AFS disagree, the bank prevails except for "
    "items the AFS supports with third-party evidence.",
    "No lien, charge or guarantee exists over the bank account beyond the overdraft indicated by the statements.",
]:
    bullet(t)

pagebreak()

# =========================================================== APPENDIX A
h1("Appendix A - Implementation sign-off checklist")
table(["#", "Item", "Evidence of completion", "Done (initials/date)"], [
    ["1", "Organisation profile matches CAC and FIRS exactly", "Screenshot of profile", ""],
    ["2", "Financial year Jan-Dec; 2024-2025 locked", "Preferences screenshot", ""],
    ["3", "302-account chart imported; zero duplicate codes; 7 inactive", "Chart of accounts export", ""],
    ["4", "Reporting tags and custom fields created", "Settings screenshots", ""],
    ["5", "Projects, budgets and timesheets enabled", "Preferences screenshot", ""],
    ["6", "Opening balances posted and agreed to the AFS", "Trial balance at 01-01-2026", ""],
    ["7", "Bank accounts created including the overdraft and clearing accounts", "Chart of accounts", ""],
    ["8", "110 bank rules created and tested against the historical narrations", "Rule list", ""],
    ["9", "Customers, vendors and related parties created with relationship tags", "Contact export", ""],
    ["10", "Projects created with contract budgets", "Project list", ""],
    ["11", "VAT, WHT, CIT and development levy accounts and rates configured", "Tax settings screenshot", ""],
    ["12", "Payroll journal template created and posted for one month", "Journal", ""],
    ["13", "Approval workflows live", "Workflow settings", ""],
    ["14", "Roles assigned; auditor read-only access tested", "Users list", ""],
    ["15", "January-October 2026 posted and reconciled", "Bank reconciliation reports", ""],
    ["16", "First monthly pack produced from the system", "Monthly pack PDF", ""],
    ["17", "Training delivered to MD and bookkeeper", "Attendance note", ""],
    ["18", "Backup taken and stored outside the system", "Backup file", ""],
], widths=[0.9, 7.4, 5.0, 3.3], fontsize=8.5)
p("Signed: ______________________________  (Managing Director)          Date: ____________________")
p("Signed: ______________________________  (Implementer)                 Date: ____________________")

pagebreak()

# =========================================================== APPENDIX B
h1("Appendix B - Tax parameters used (confirm before reliance)")
p("These are the parameters embedded in the posting rules and the guide. They must be signed off by the tax adviser "
  "before the first 2026 provision, because 2026 is the first year under the Nigeria Tax Act 2025 and the Nigeria Tax "
  "Administration Act 2025.")
table(["Parameter", "Value used", "Basis", "Status"], [
    ["VAT rate", "7.5%", "VAT Act 2023 (retained)", "Confirm registration effective date"],
    ["VAT filing", "Monthly, by the 21st", "VAT Act 2023", "Confirm"],
    ["VAT registration threshold", "N25m taxable turnover in any 12 months", "VAT Act 2023 s.6 (threshold may have "
     "been revised under NTA 2025)", "Confirm"],
    ["CIT - small company tier", "0%", "CITA 2020 s.40(1)(a); NTA 2025 revised the test to turnover-based bands",
     "Confirm the current tests"],
    ["CIT - standard rate", "20% / 30% depending on size band", "NTA 2025", "Confirm which band Jemai falls in"],
    ["Minimum tax", "0.5% of gross income", "CITA 2020 s.37(2) as carried into the file", "Confirm under NTAA 2025"],
    ["Development levy", "4% of assessable profit", "NTA 2025 (replaces TEL and similar levies)", "Confirm rate, base "
     "and exemption"],
    ["WHT - contracts / construction", "5%", "Deduction of Tax at Source (Withholding) Regulations", "Confirm"],
    ["WHT - professional & consultancy", "5%", "As above", "Confirm"],
    ["WHT - rent", "10%", "As above", "Confirm"],
    ["WHT - supply of goods", "2% (companies) / 1.5% (individuals)", "As above", "Confirm - exemptions apply"],
    ["WHT remittance deadline", "21st of the following month", "NTAA 2025", "Confirm"],
    ["PAYE remittance", "10th of the following month", "PITA / state revenue", "Confirm"],
    ["Pension", "Employee 8%, employer 10% of basic, housing and transport", "Pension Reform Act 2014", "Confirm"],
    ["ITF / NSITF / NHF", "1% / 1% / 2.5%", "ITF Act, ECA 2010, NHF Act", "Confirm applicability"],
    ["Capital allowances", "Class rates: 10% / 20% / 25% with 1% residual", "NTA 2025 (replacing initial + annual "
     "allowances)", "Confirm - applies to 2026 additions"],
    ["Deferred tax rate", "Rate expected to apply when the difference reverses", "IAS 12", "Confirm"],
], widths=[4.8, 4.2, 4.8, 2.8], fontsize=8.5)
callout("This is a configuration reference, not tax advice.",
        "Every rate above should be confirmed in writing by the company's tax adviser before it is used in a filing. "
        "The guide uses them so the system is built with the right structure; the numbers themselves must be signed "
        "off.")

pagebreak()

# =========================================================== APPENDIX C
h1("Appendix C - Opening trial balance at 1 January 2026 (corrected basis)")
table(["Code", "Account", "Debit (NGN)", "Credit (NGN)"], [
    ["1021", "Access Bank - A/C 1854793470", "24,194.10", ""],
    ["1091", "Furniture & Fittings", "165,000.00", ""],
    ["1093", "Computers & IT Equipment", "360,053.75", ""],
    ["1101", "Accumulated Depreciation - Furniture & Fittings", "", "56,184.75"],
    ["1103", "Accumulated Depreciation - Computers & IT", "", "122,637.25"],
    ["1045", "Director Current Account Receivable", "14,223,846.92", ""],
    ["1046", "Other Receivables", "332,672.02", ""],
    ["2011", "Trade Payables - Suppliers", "", "1,335,000.00"],
    ["2023", "Accrued Professional Fees", "", "300,000.00"],
    ["2031", "Contract Liabilities - Project Deposits", "", "3,460,000.00"],
    ["2045", "Income Tax Payable (CIT / minimum tax)", "", "188,852.14"],
    ["2046", "Development Levy Payable", "", "85,475.07"],
    ["2051", "Short-term Loan - Individual Lender", "", "230,000.00"],
    ["2052", "Short-term Working Capital Loan", "", "1,230,000.00"],
    ["3011", "Ordinary Share Capital", "", "1,000,000.00"],
    ["3071", "Capital Reserve", "", "2,000.00"],
    ["3090", "Retained Earnings", "", "7,095,617.58"],
    ["", "TOTAL", "15,105,766.79", "15,105,766.79"],
], widths=[1.8, 9.4, 3.0, 2.4], fontsize=8.5)
p("Net assets: N14,926,944.79 (total debits N15,105,766.79 less accumulated depreciation N178,822.00 presented as a "
  "contra-asset). Alternative route: load the AFS as issued (share capital 1,000,000; retained earnings 2,197,911; "
  "receivables 4,453,485; PPE 355,232 net; payables 1,635,000; cash 24,194) and post the IAS 8 adjustment journal of "
  "N4,897,706.58 shown in deliverable 01.")

h1("Appendix D - Journal import field map")
table(["CSV column", "Zoho Books field", "Mandatory", "Note"], [
    ["Journal Date", "Journal Date", "Yes", "dd-MMM-yyyy; must fall in an open period"],
    ["Journal#", "Journal Number", "No", "JE-YYYY-MM-###; must be unique or Zoho rejects the row"],
    ["Reference#", "Reference Number", "No", "The bank transaction ID - this is the audit trail back to the statement"],
    ["Notes", "Notes", "No", "Journal-level narration; the engine writes the bank narration here"],
    ["Currency", "Currency", "Yes", "NGN"],
    ["Exchange Rate", "Exchange Rate", "Yes", "1"],
    ["Account", "Account", "Yes", "Account name exactly as in the chart - the engine looks it up from the code"],
    ["Account Code", "Account Code (if mapped)", "No", "Leave unmapped if the importer does not offer it"],
    ["Debit", "Debit", "Per line", "One of debit/credit per line"],
    ["Credit", "Credit", "Per line", ""],
    ["Description", "Line item description", "No", "Explains the leg"],
    ["Contact", "Contact", "No", "Only if the contact exists; otherwise leave blank and tag later"],
    ["Project", "Project", "No", "Only if the project exists"],
], widths=[3.2, 4.4, 2.2, 6.8], fontsize=8.5)

h1("Appendix E - Documents to request from the client")
for t in [
    "The Access Bank statements for November 2026 onwards, to continue the posting journal already built for "
    "August-October 2026 (deliverable 02e-02i).",
    "VAT registration certificate and any FIRS correspondence, or confirmation that registration has not been done.",
    "Signed contracts and invoices for all live projects (CIIFA SafeHouse, The Victoria King, Hassan Umar Abba, "
    "Joshua Abolade, Kehinde Adeshakin, Ibrahim Junaidu, Ekpunobi Judith, Amaowoh, Famakinwa, Auxit Media, Venture "
    "Garden).",
    "Supplier and subcontractor invoices for the N1,335,000 of trade payables and the N34m of 2024-2025 direct cost.",
    "Loan agreements for the N2.5m working capital facility and the N230,000 individual loan.",
    "Director loan / current account agreement and a repayment plan for the N14.22m balance.",
    "Payroll register for 2024 and 2025, PAYE and pension remittance receipts.",
    "CAC annual return and the register of members (the 2024-2025 Ime Ubongabasi share event is not reflected in the "
    "AFS).",
    "The fixed-asset register supporting the N525,054 cost and the N9,000 evidence gap in 2025.",
    "Bank mandate in the correct legal name ('LIMITED', not 'LIMTED').",
]:
    bullet(t)

doc.add_paragraph()
p("End of document.", italic=True, size=9, colour=GREY)

doc.save(OUT)
print("wrote", OUT)
