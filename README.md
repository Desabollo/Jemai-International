# Jemai International Limited — Zoho Books review, chart of accounts and posting journal

Deliverables produced from the historical files in this repository:

| Source file | What it is |
|---|---|
| `Chart_of_Accounts Oct 2026.xlsx` | Zoho Books chart-of-accounts export, 275 accounts, 12-column upload layout |
| `Jemai_International_Master_Workbook_2024-2025.xlsx` | 35-tab audit working paper: 1,226 valued bank transactions (Access Bank A/C 1854793470, 31-May-2024 → 31-Dec-2025), trial balances, tax computations, corrections log, 24-point exception dashboard |
| `Jemai Int'nal 2024 AFS-Final-signed31102026.pdf` | Signed FY2024 audited financial statements |
| `Jemai Int'nal 2025 AFS-Final-signed31102026.pdf` | Signed FY2025 audited financial statements |

## Deliverables (`/deliverables`)

### 1. Reviewed and improved chart of accounts
**`01_Chart_of_Accounts_Improved_Jemai_Oct2026.xlsx`** + **`01b_CoA_Zoho_Import.csv`**

- 275 accounts reviewed → **302 accounts**: 7 overlapping accounts deactivated, 27 added, 19 orphan accounts re-parented, 25 revenue accounts nested under the revenue header.
- **2 blocking defects fixed**: duplicate account code 2030 (Retained Earnings vs Customer/Project Deposits → retained earnings renumbered 3090) and one account with no code (License Fees → 6419).
- New accounts close the reporting gaps that produced 13 audit failures: deferred tax asset/liability, provisions, IFRS 16 lease pair, ITF/NSITF/NHF, share application money, EMTL and transfer commission, ECL on related-party receivables, related-party receivable, air travel, revaluation reserve, bank overdraft, wallet/interbank clearing.
- Tabs: `Zoho_Import_CoA` (upload-ready, parents-first), `Changes_Log` (every change + action in Zoho), `Validation_Tests` (20 structural tests, before/after), `Legacy_to_New_Map` (33-account controlled chart → new codes, with line counts), `Reporting_Map` (all 302 accounts → IFRS line, note, tax tag, cash-flow category), `Gap_Analysis`, `Open_Balances_1Jan2026` (opening TB + the IAS 8 prior-period adjustment alternative).
- Upload format is unchanged: the same 12 columns as the client export.

### 2. Monthly transaction posting journal
**`02_Jemai_Posting_Journal_Aug-Oct_2026.xlsx`** + CSVs

A tested posting engine that converts Access Bank statement lines into balanced Zoho Books journal imports on the improved chart: 12 tabs, **110 posting rules**, input register with drop-downs and direction controls, formula-driven journal output in Zoho's import layout, account validation, control checks and a Zoho upload guide.

**Important — the data position.** The repository contains bank statements to **31 December 2025 only**. There is no August–October 2026 statement, so 2026 entries cannot be extracted from evidence; inventing them would put fabricated transactions into a set of books. The workbook therefore ships with:

- an empty, pre-wired 2026 input register (paste the statement → the journal builds itself), and
- a **worked example** that proves the engine: the most recent three-month statement on file (**Aug–Oct 2025, 112 lines**, N4,431,000 in / N4,430,917.59 out) re-coded from the legacy 33-account chart into the 302-account chart, producing **224 balanced journal lines** that reconcile to the statement to the kobo (Aug +617.48, Sep +308.29, Oct −843.36). `02b_...csv` is that journal, ready to import.

### 3. Zoho Books Premium setup guide
**`03_Zoho_Books_Premium_Setup_Guide_Jemai_International.docx`** — 26 sections, ~9,800 words, 44 tables. Covers organisation settings, chart import, opening balances, bank feeds and the priority order of bank rules, contacts and related-party discipline, projects and budgets, the sales and purchase cycles, WHT, payroll, tax configuration, approvals and roles, the IFRS reporting pack, a six-week go-live plan, close checklists, controls, risks and five appendices (sign-off checklist, tax parameters to confirm, opening trial balance, journal import field map, document request list).

## How to regenerate

```bash
pip install openpyxl xlsxwriter python-docx
cd build
python3 build_coa_xlsx.py     # deliverable 1
python3 build_journal.py      # deliverable 2
python3 build_guide.py        # deliverable 3
```

`build/journal_rules.py` holds the 110 posting rules and the transaction-level judgements; edit it and re-run to change the engine.

## Key findings carried into the design

1. **607 of 1,226 bank lines (49.5%) were posted to one account** (legacy 6050), which also absorbed site expenses, fuel, transport, training and refunds.
2. **N5.05m (2024) and N9.60m (2025) paid to the MD were expensed**; the corrected position is a **N14.22m receivable from the director**, undisclosed in the AFS.
3. **N3.46m of client advances were recognised as turnover** at both year ends.
4. **Tax was not provided** (minimum tax and TEL of N235,622 unpaid) and **VAT registration was missed** in 2025 (turnover N30.13m vs the N25m threshold; exposure ≈ N2.26m plus penalties).
5. The chart had **one duplicate code, one uncoded account, an inverted equity hierarchy and no deferred tax, provision or non-current liability accounts**.
