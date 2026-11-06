# Jemai International Limited — Zoho Books review, chart of accounts and posting journal

Deliverables produced from the historical files in this repository:

| Source file | What it is |
|---|---|
| `Chart_of_Accounts Oct 2026.xlsx` | Zoho Books chart-of-accounts export, 275 accounts, 12-column upload layout |
| `Jemai_International_Master_Workbook_2024-2025.xlsx` | 35-tab audit working paper: 1,226 valued bank transactions (Access Bank A/C 1854793470, 31-May-2024 → 31-Dec-2025), trial balances, tax computations, corrections log, 24-point exception dashboard |
| `Jemai Int'nal 2024 AFS-Final-signed31102026.pdf` | Signed FY2024 audited financial statements |
| `Jemai Int'nal 2025 AFS-Final-signed31102026.pdf` | Signed FY2025 audited financial statements |
| `August to October 2026 Statement.pdf` | Access Bank A/C 1854793470, 26 pages, 01-Aug-2026 → 31-Oct-2026: opening N24,954,074.00, withdrawals N74,373,993.97, deposits N51,225,487.00, closing N1,805,567.88 |

## Deliverables (`/deliverables`)

### 1. Reviewed and improved chart of accounts
**`01_Chart_of_Accounts_Improved_Jemai_Oct2026.xlsx`** + **`01b_CoA_Zoho_Import.csv`**

- 275 accounts reviewed → **302 accounts**: 7 overlapping accounts deactivated, 27 added, 19 orphan accounts re-parented, 25 revenue accounts nested under the revenue header.
- **2 blocking defects fixed**: duplicate account code 2030 (Retained Earnings vs Customer/Project Deposits → retained earnings renumbered 3090) and one account with no code (License Fees → 6419).
- New accounts close the reporting gaps that produced 13 audit failures: deferred tax asset/liability, provisions, IFRS 16 lease pair, ITF/NSITF/NHF, share application money, EMTL and transfer commission, ECL on related-party receivables, related-party receivable, air travel, revaluation reserve, bank overdraft, wallet/interbank clearing.
- Tabs: `Zoho_Import_CoA` (upload-ready, parents-first), `Changes_Log` (every change + action in Zoho), `Validation_Tests` (20 structural tests, before/after), `Legacy_to_New_Map` (33-account controlled chart → new codes, with line counts), `Reporting_Map` (all 302 accounts → IFRS line, note, tax tag, cash-flow category), `Gap_Analysis`, `Open_Balances_1Jan2026` (opening TB + the IAS 8 prior-period adjustment alternative).
- Upload format is unchanged: the same 12 columns as the client export.

### 2. Monthly transaction posting journal — 1 Aug to 31 Oct 2026 (real statement)

**Import file: `02e_Journal_Aug-Oct_2026_ACTUAL_Zoho_Upload.csv`** — the complete, balanced journal in the Zoho Books journal-import layout, ready to upload.

| File | What it is |
|---|---|
| `02e_Journal_Aug-Oct_2026_ACTUAL_Zoho_Upload.csv` | All three months, **692 bank lines → 1,384 journal legs**, Dr = Cr = **N125,599,480.97** |
| `02f_Journal_Aug2026_ACTUAL.csv` / `02g_..._Sep2026` / `02h_..._Oct2026` | The same journal split by month (one journal per bank line, `JE-2026-MM-NNN`) |
| `02i_Posting_Journal_Aug-Oct_2026_ACTUAL_Workbook.xlsx` | 9 tabs: register, journal, trial balance, control checks, monthly summary, exceptions, VAT review, rules applied |
| `02_Jemai_Posting_Journal_Aug-Oct_2026.xlsx` | The reusable **engine** (110 rules, pre-wired register, formula-driven upload sheet) plus the Aug–Oct 2025 worked example that proves it |
| `02b` / `02c` / `02d` | Aug–Oct 2025 worked-example journal, blank 2026 upload template, blank input register |

**How the 2026 statement was turned into a journal.** The PDF was parsed from word coordinates (not text order), because Access prints each transaction's description wrapped over the lines above and below the money line. Every printed line was then reconciled: deposits agree to the statement's N51,225,487.00 to the kobo, and payments agree to N74,373,993.97 once the four failed IG-Ads/Marketing-Ads transfers (N1,000,053.75 each, reversed by the bank in the same block) are netted off. The journal's own bank ledger therefore rolls forward to N1,805,567.03; the statement prints N1,805,567.88, a N0.85 rounding difference inside the bank document itself.

**What the money did, Aug–Oct 2026.** Receipts N51.2m — corporate client collections N17.8m (AFRC N14.76m, OPTIMUS N3.0m), art / decor sales N16.7m (including N12.0m "Art and Candles" and N2.7m for purchase and hanging of artworks), gallery ticketing N3.0m, director repayment N5.0m, investor funds N5.0m, loans N3.0m. Payments N74.4m — payroll N17.6m, director current account N8.4m, rent prepayment N5.0m, artwork purchases N4.7m, PAYE arrears for April–July N3.5m, allowances N3.2m, a N3.0m tax clearance payment, N2.8m travel and visa, N2.7m furniture, N2.7m e-commerce website, N2.5m UK art logistics.

**Judgement calls, all on the `Exceptions` tab**: 2 payments left on 7999 (two Paystack checkouts with no purpose in the narration), 127 unlabelled individual receipts treated as gallery/art sales (6 of them ≥ N50,000 are itemised), and every override explained line by line. VAT is posted **gross** — the `VAT_Review` tab quantifies the N5,375 = N5,000 + 7.5% pattern and gives the reclass entry to post if the company is registered.

**Not audit-ready, by design**: this is bank evidence only. Accruals, depreciation, revenue cut-off, payroll liabilities and tax provisions still have to be journalled to close the year.

### 3. Zoho Books Premium setup guide
**`03_Zoho_Books_Premium_Setup_Guide_Jemai_International.docx`** — 26 sections, ~9,800 words, 44 tables. Covers organisation settings, chart import, opening balances, bank feeds and the priority order of bank rules, contacts and related-party discipline, projects and budgets, the sales and purchase cycles, WHT, payroll, tax configuration, approvals and roles, the IFRS reporting pack, a six-week go-live plan, close checklists, controls, risks and five appendices (sign-off checklist, tax parameters to confirm, opening trial balance, journal import field map, document request list).

## How to regenerate

```bash
pip install openpyxl xlsxwriter python-docx
cd build
python3 build_coa_xlsx.py     # deliverable 1
python3 build_journal.py      # deliverable 2 engine + 2025 worked example
python3 parse_statement.py    # parse the 2026 statement PDF -> /home/user/work/stmt2026.json
python3 build_journal_2026.py # deliverable 2 actual 2026 journal (02e-02i)
python3 build_guide.py        # deliverable 3
```

`build/journal_rules.py` holds the 110 rules of the reusable engine (2025 worked example); `build/journal_rules_2026.py` holds the 2026 rule set, including the explicit payee overrides that settle the judgement calls in the real statement. Edit and re-run to change either.

## Key findings carried into the design

1. **607 of 1,226 bank lines (49.5%) were posted to one account** (legacy 6050), which also absorbed site expenses, fuel, transport, training and refunds.
2. **N5.05m (2024) and N9.60m (2025) paid to the MD were expensed**; the corrected position is a **N14.22m receivable from the director**, undisclosed in the AFS.
3. **N3.46m of client advances were recognised as turnover** at both year ends.
4. **Tax was not provided** (minimum tax and TEL of N235,622 unpaid) and **VAT registration was missed** in 2025 (turnover N30.13m vs the N25m threshold; exposure ≈ N2.26m plus penalties).
5. The chart had **one duplicate code, one uncoded account, an inverted equity hierarchy and no deferred tax, provision or non-current liability accounts**.
