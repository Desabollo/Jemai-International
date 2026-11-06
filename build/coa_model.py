"""
Jemai International Limited - Chart of Accounts model.
Loads the client's Zoho Books export ("Chart_of_Accounts Oct 2026.xlsx"),
applies the review fixes, and adds the missing accounts required for
IFRS-compliant reporting and for the 2026 NTA/NTAA tax regime.
"""
import json
import openpyxl

import os
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "Chart_of_Accounts Oct 2026.xlsx")
HDR = ['Account ID', 'Account Name', 'Account Code', 'Description', 'Account Type',
       'Mileage Rate', 'Mileage Unit', 'IsMileage', 'Account # ', 'Account Status',
       'Currency', 'Parent Account']


def load_source():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Accounts"]
    rows = []
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        rows.append(["" if c is None else str(c).strip() for c in r])
    hdr = [h.strip() for h in rows[0]]
    data = []
    for r in rows[1:]:
        d = dict(zip(hdr, r))
        if not d.get("Account Name"):
            continue
        data.append({
            "id": d["Account ID"],
            "name": d["Account Name"],
            "code": d["Account Code"],
            "desc": d["Description"],
            "type": d["Account Type"],
            "mileage_rate": "0.000",
            "mileage_unit": "",
            "is_mileage": "false",
            "acct_no": d["Account #"],
            "status": "Active",
            "currency": "NGN",
            "parent": d["Parent Account"],
        })
    return data


def k(name):
    """stable lookup key"""
    return name.strip().lower()


# ---------------------------------------------------------------------------
# NEW ACCOUNTS (missing from the client chart but required for reporting)
# ---------------------------------------------------------------------------
NEW_ACCOUNTS = [
    # code, name, type, parent, description, reason
    ("1029", "Bank Overdraft - Access Bank", "Bank Overdraft", "",
     "Overdraft balance on Access Bank A/C 1854793470. Actual credit balances are presented separately from operating cash so the bank position is never netted.",
     "AFS flag: recurring overdrafts / negative cash days are driving the going-concern qualification. A dedicated overdraft account is needed to disclose it under IAS 1.69-76."),
    ("1034", "Mobile Money / Wallet Clearing", "Other Current Asset", "Sales Collection Clearing",
     "Settlement clearing for OPay / Moniepoint / PalmPay / mobile wallet collections before credit to the bank account.",
     "2025 bank evidence shows heavy 'MOBILE TRF TO GTB/PVB/STL/FMB/MMF/FCL' channel usage. Collections land in the bank 1-2 days after the customer pays; without a clearing account cash is overstated at cut-off."),
    ("1035", "Interbank Transfer Clearing", "Other Current Asset", "Sales Collection Clearing",
     "Contra account for funds in transit between the company's own bank / wallet accounts and for transfers reversed or returned by the beneficiary bank.",
     "Needed once a second bank or wallet account is opened (recommended) and for the failed/returned transfers already visible in the 2024-2025 register."),
    ("1047", "Related Party Receivable - Family & Associates", "Other Current Asset", "",
     "Amounts paid to the MD's family and associates for company purposes, pending settlement or supporting documents.",
     "Payments to family/associates (e.g. 'Emmanuel Akinloluwa Akin-Ajayi - August Site Payment' N175,000 in Sep-2025) had no receivable account; the chart offered only a director receivable (1045) and a related-party PAYABLE (2062)."),
    ("1086", "Input VAT - Bank Charges & Fees", "Other Current Asset", "Tax Recoverable",
     "7.5% VAT charged by Access Bank on commissions, SMS alerts and maintenance fees - recoverable once the company is VAT-registered.",
     "258 of 1,226 historical bank lines carry deductible bank-charge VAT. Lumping it into the expense understates recoverable VAT by an estimated N150k-260k over two years."),
    ("1180", "Right-of-Use Assets (IFRS 16)", "Other Asset", "",
     "Capitalised right-of-use asset for the Ogba office lease and any equipment leases, net of amortisation.",
     "IFRS 16 requires lessees to recognise a lease liability and right-of-use asset. The AFS charges office rent of N200,000 p.a. straight to profit or loss with no IFRS 16 assessment."),
    ("1190", "Deferred Tax Asset", "Other Asset", "",
     "Deferred tax asset arising on deductible temporary differences and unused capital allowances, recognised only where future taxable profit is probable.",
     "The AFS has no IAS 12 disclosure at all. The chart had no deferred tax accounts, so the disclosure could not have been produced from the ledger."),
    ("2057", "Share Application Money Pending Allotment", "Equity", "",
     "Money received from a subscriber before the shares are allotted and the CAC return is filed.",
     "The 2024/2025 Ime Ubongabasi share event (N1,000 in, N99,000 refunded) had nowhere to sit and was mis-posted to revenue and capital."),
    ("2065", "Warranty / Defect Provision", "Other Current Liability", "Other Current Liabilities",
     "Provision for defect rectification on completed fit-out and furniture projects within the defects-liability period.",
     "IAS 37 provision. The chart only had a cost-of-sales rectification expense (5061) with no liability side, so any provision would break the balance sheet."),
    ("2066", "Provision for Claims, Fines & Penalties", "Other Current Liability", "Other Current Liabilities",
     "Provision for regulatory penalties and contested claims (e.g. VAT non-registration penalties, late-filing penalties).",
     "A N2.26m VAT exposure and unpaid minimum tax/TEL are identified in the working papers. IAS 37/IAS 10 require a recognised provision, not just a note."),
    ("2074", "ITF Contribution Payable", "Other Current Liability", "Payroll Liabilities",
     "Industrial Training Fund contribution (1% of annual payroll) due to ITF.",
     "Statutory payroll liability omitted from the chart; ITF applies to companies with 5+ employees or N50m+ turnover."),
    ("2075", "NSITF Contribution Payable", "Other Current Liability", "Payroll Liabilities",
     "Employees' Compensation Scheme contribution (1% of payroll) due to NSITF.",
     "Statutory payroll liability omitted from the chart."),
    ("2076", "NHF Payable", "Other Current Liability", "Payroll Liabilities",
     "National Housing Fund deductions (2.5% of basic salary) withheld from staff and payable to FMBN.",
     "Statutory payroll liability omitted from the chart."),
    ("2085", "Dividends Payable", "Other Current Liability", "Other Current Liabilities",
     "Dividends declared but not yet paid to shareholders.",
     "CAMA 2020 distribution mechanics; no equity-to-liability bridge existed."),
    ("2096", "Deferred Tax Liability", "Non Current Liability", "",
     "Deferred tax liability on taxable temporary differences (accelerated capital allowances over accounting depreciation).",
     "IAS 12. The chart had no non-current liability section at all before this change."),
    ("2097", "Long-term Borrowings", "Non Current Liability", "",
     "Loan principal repayable more than 12 months after the reporting date.",
     "Every borrowing account in the chart was classified 'Other Current Liability', which forces an incorrect current/non-current split under IAS 1.60-76."),
    ("2098", "Lease Liabilities - IFRS 16", "Non Current Liability", "",
     "Present value of remaining lease payments (current portion to be reclassified to 2052 at each year end).",
     "Counterpart to 1180; required if the IFRS 16 policy is adopted."),
    ("3072", "Revaluation / Fair Value Reserve", "Equity", "Reserves",
     "Surplus on revaluation of property, plant and equipment (IAS 16.39) or fair value of investments through OCI.",
     "The AFS presents a 'Revaluation gain' line in OCI but the chart had no reserve to hold it."),
    ("5038", "Project Insurance & Bonds", "Cost Of Goods Sold", "Site Operations Costs",
     "Contractor's all-risk insurance, performance bonds and project-specific premiums chargeable to a client contract.",
     "Project-direct cost omitted from the chart; such premiums would otherwise be mis-posted to administrative insurance (6218)."),
    ("6317", "Sales Commission & Referral Fees", "Expense", "Selling & Marketing Expenses",
     "Commission paid to introducers, agents and referrers on converted projects.",
     "No commission account existed; 2025 bank evidence shows referral-style payments that had to be forced into other expense buckets."),
    ("6515", "Air Travel & Flight Bookings", "Expense", "Transport & Travel",
     "Domestic and international flight bookings for staff and client-related travel.",
     "No air-fare account existed; flight bookings (N77,000 in Sep-2025, paid through the Paystack gateway) were forced into 'Hotel & Lodging' or general admin transport."),
    ("6822", "Depreciation - Right-of-Use Assets", "Expense", "Depreciation & Amortisation",
     "Amortisation of the IFRS 16 right-of-use asset.",
     "Counterpart to 1180."),
    ("6915", "Transfer Commission & Channel Charges", "Expense", "Finance Costs",
     "Per-transfer commission charged by Access Bank / switching partners on outbound NIP and mobile transfers.",
     "This is the single biggest mis-posting in the history: 607 of 1,226 transactions were forced into one account (legacy 6050). Commission alone accounts for the majority."),
    ("6916", "Electronic Money Transfer Levy (EMTL)", "Expense", "Finance Costs",
     "FGN Electronic Money Transfer Levy (N50 per qualifying electronic transfer) deducted by the bank.",
     "Separately disclosable statutory levy; 36 historical lines were buried in general bank charges."),
    ("6917", "Lease Interest - IFRS 16", "Expense", "Finance Costs",
     "Unwinding of the discount on IFRS 16 lease liabilities.",
     "Counterpart to 2098."),
    ("7013", "Rounding / Posting Differences", "Expense", "Other Costs / Suspense",
     "Immaterial differences arising on bank statement rounding, kobo write-offs and journal balancing.",
     "Prevents the 'contra/reversal' account (7012) being used as a plug, which destroys its audit value."),
    ("7024", "ECL - Related Party Receivables", "Expense", "Bad Debt / Impairment",
     "Expected credit loss allowance on amounts due from directors and related parties (IAS 24 / IFRS 9).",
     "A N14.22m director receivable is the largest item on the corrected balance sheet and its recoverability is in doubt; the chart could only have impaired trade debtors (7022)."),
]

# ---------------------------------------------------------------------------
# CHANGES TO EXISTING ACCOUNTS
# ---------------------------------------------------------------------------
# name -> dict of field overrides, plus reason/meta
CHANGES = {
    "Retained Earnings": dict(
        code="3090", parent="Equity",
        reason="Duplicate account code 2030 shared with 'Customer / Project Deposits'. Zoho Books rejects duplicate codes at import and mis-sorts the trial balance. Retained earnings is equity and belongs in the 3xxx block.",
        severity="Blocking", action="Edit account in Chart of Accounts, change code from 2030 to 3090."),
    "License Fees": dict(
        code="6419", parent="Professional & Regulatory Expenses",
        reason="Account uploaded with a blank Account Code - Zoho Books will not create a reliable posting target and any CSV import referencing it fails the lookup.",
        severity="Blocking", action="Edit account, set code 6419."),
    "Owner's Equity": dict(
        parent="Equity",
        reason="Hierarchy inverted: 3060 was made a child of 3061 'Director Equity Injections', so the equity header sits under its own detail account.",
        severity="High", action="Set 3060 parent to 3000 Equity."),
    "Director Equity Injections": dict(
        parent="Owner's Equity",
        reason="See above - becomes a child of 3060.",
        severity="High", action="Set 3061 parent to 3060."),
    "Director Personal Investment / Drawings": dict(
        parent="Director Drawings", status="Inactive",
        reason="Duplicates 3051 'Director Drawings - Esther Akin-Ajayi' and blurs the IAS 24 distinction between capital introduced and drawings taken.",
        severity="Medium", action="Set parent 3050 and mark Inactive."),
    "Trade Receivables - Project Clients": dict(
        parent="Accounts Receivable-Control",
        reason="Trade debtors were not subordinated to the AR control account, so the balance sheet would show the control and the detail side by side.",
        severity="High", action="Set parent to 1040."),
    "Retention Receivables": dict(
        parent="Accounts Receivable-Control",
        reason="Same as above; retention is a receivable, not a separate asset class.",
        severity="High", action="Set parent to 1040."),
    "WHT Receivable / Tax Credit Notes": dict(
        status="Inactive",
        reason="Duplicate of 1082 'WHT Credit Receivable'. Two accounts for the same receivable guarantee a split, reconciling balance.",
        severity="High", action="Mark Inactive; post all WHT credits to 1082."),
    "VAT Output Payable": dict(
        parent="VAT Control / Net VAT Payable",
        reason="Two parallel VAT-output accounts. Nesting 2041 under 2042 makes 7.5% output VAT from sales roll into one net control account at month end.",
        severity="High", action="Set parent to 2042; use 2041 as the Zoho VAT rate account."),
    "Retentions Payable to Subcontractors": dict(
        code="2015", parent="Accounts Payable-Control",
        reason="Sits under 'Customer / Project Deposits' although it is a supplier payable, not a contract liability - it distorts both the trade-payables and contract-liability notes.",
        severity="High", action="Edit code to 2015 and parent to 2010."),
    "Tax Recoverable": dict(
        parent="", parent_children=["VAT Input Receivable", "WHT Credit Receivable",
                                    "CIT Overpayment Receivable",
                                    "PAYE Recoverable / Staff Tax Receivable"],
        reason="1080 is used as a header but its four detail accounts (1081-1084) hang loose at top level.",
        severity="Medium", action="Set 1081-1084 parent to 1080."),
    "Salaries and Employee Cost": dict(
        parent="Staff Costs",
        reason="Orphan: the main salary account was not under the Staff Costs header, so payroll never totals in the P&L.",
        severity="Medium", action="Set parent to 6100."),
    "IT_Internet & Data": dict(
        parent="Admin & Office Expenses",
        reason="Orphan in the admin block.",
        severity="Low", action="Set parent to 6200."),
    "Cleaning & Sanitation": dict(
        parent="Admin & Office Expenses",
        reason="Orphan in the admin block.",
        severity="Low", action="Set parent to 6200."),
    "Advertising And Marketing": dict(
        parent="Selling & Marketing Expenses",
        reason="Orphan that duplicates the meaning of header 6300.",
        severity="Low", action="Set parent to 6300; use 6310 for paid media."),
    "Other Consultancy Fees": dict(
        parent="Professional & Regulatory Expenses",
        reason="Orphan in the professional-fees block.",
        severity="Low", action="Set parent to 6400."),
    "Transport & Logistics": dict(parent="Transport & Travel",
        reason="Orphan in the travel block.", severity="Low", action="Set parent to 6500."),
    "Local Travel": dict(parent="Transport & Travel",
        reason="Orphan in the travel block.", severity="Low", action="Set parent to 6500."),
    "Fuel & Vehicle Running": dict(parent="Transport & Travel",
        reason="Orphan in the travel block.", severity="Low", action="Set parent to 6500."),
    "Hotel & Lodging": dict(parent="Transport & Travel",
        reason="Orphan in the travel block.", severity="Low", action="Set parent to 6500."),
    "Depreciation - Motor Vehicles": dict(parent="Depreciation & Amortisation",
        reason="Orphan while its six siblings sit under 6800.",
        severity="Low", action="Set parent to 6800."),
    "Bank Charges & Maintenance Fees": dict(parent="Finance Costs",
        reason="Orphan - the highest-volume expense account in the ledger has no parent, so finance costs never aggregate.",
        severity="High", action="Set parent to 6900."),
    "Payment Gateway Charges": dict(parent="Finance Costs",
        reason="Orphan in the finance-cost block (Paystack/NIP gateway fees).",
        severity="Medium", action="Set parent to 6900."),
    "Bad Debt Expense": dict(parent="Bad Debt / Impairment",
        reason="Orphan that duplicates the 7020 header.",
        severity="Low", action="Set parent to 7020."),
    "Uncategorised Expense": dict(parent="Other Costs / Suspense",
        reason="Orphan suspense account.",
        severity="Low", action="Set parent to 7000."),
    "Cost of Sales - Telephone & Internet": dict(parent="Cost of Sales",
        reason="Orphan COGS account; project-dedicated data lines were not grouped under cost of sales.",
        severity="Low", action="Set parent to 5000."),
    "Capital Reserve": dict(parent="Reserves",
        reason="Orphan within equity.",
        severity="Low", action="Set parent to 3070."),
    "Project Tranche Revenue": dict(status="Inactive",
        reason="Not a revenue stream - stage/tranche is a billing pattern. Keeping it invites double counting against the service-line accounts and breaks IFRS 15 disaggregation.",
        severity="High", action="Mark Inactive after the 2025 conversion; bill tranches against 4011-4019."),
    "Corporate Client Revenue": dict(status="Inactive",
        reason="Revenue split by customer type instead of by performance obligation. It overlaps 4011-4019 and cannot be aggregated without double counting.",
        severity="High", action="Mark Inactive; use Customer Type / Reporting Tags for corporate vs residential."),
    "Residential Client Revenue": dict(status="Inactive",
        reason="See above.",
        severity="High", action="Mark Inactive; use a Reporting Tag."),
    "Retention Release Revenue": dict(status="Inactive",
        reason="Retention release is the collection of a receivable (1042), not new revenue. Posting it to income double counts the contract.",
        severity="High", action="Mark Inactive; release retentions against 1042."),
    "Business Entertainment": dict(status="Inactive",
        reason="Duplicate of 6314 'Client Entertainment' - two entertainment accounts split the same spend and confuse the non-deductibility review.",
        severity="Medium", action="Mark Inactive; use 6314."),
}

# accounts that must move under 4000 Income / Sales for a clean P&L
REVENUE_CHILDREN = ["Architectural Design Revenue", "Interior Design / Fit-out Revenue",
                    "Furniture Production & Supply Revenue", "Visual Art / Decor Sales",
                    "Space Planning & Design Consultation Revenue", "Project Management Fees",
                    "3D Rendering / Visualisation Revenue", "Shortlet / Property Interior Project Revenue",
                    "Installation & Site Supervision Fees", "Shipping Charges Billed"]

# accounts that cannot take the parent we would like, because Zoho requires
# parent and child to share an account type
TYPE_BLOCKED = {
    "Exchange Gain or Loss": ("Other Expense", "Type 'Other Expense' has no matching header in this chart; it must stay top-level and be presented under 'Other income / (expense)' in reports."),
    "Purchase Discounts": ("Expense", "'Purchase Discounts' is typed Expense but is economically a cost-of-sales contra. Zoho will not allow it under 5000 (Cost Of Goods Sold). Either leave top-level and reclassify in the P&L report, or re-create the account as type Cost Of Goods Sold."),
    "Suspense - Liability Review": ("Other Liability", "Type 'Other Liability' cannot be nested under 2000 'Liabilities' (typed Other Current Liability)."),
    "Stock accounts": ("Stock", "The Stock type does not accept a parent in Zoho Books; inventory stays flat."),
    "Cash accounts": ("Cash", "There is no Cash-type header, so 1011/1012/1013 cannot be grouped. Create a 'Cash & Cash Equivalents' header only if Zoho enables a Cash-type parent in your DC."),
}


def build():
    data = load_source()
    index = {k(a["name"]): a for a in data}

    # --- apply code / parent / status overrides -------------------------------
    for name, ov in CHANGES.items():
        a = index[k(name)]
        if "code" in ov:
            a["code"] = ov["code"]
        if "parent" in ov:
            a["parent"] = ov["parent"]
        if "status" in ov:
            a["status"] = ov["status"]
        if "parent_children" in ov:
            for c in ov["parent_children"]:
                index[k(c)]["parent"] = name

    for name in REVENUE_CHILDREN:
        index[k(name)]["parent"] = "Income / Sales"

    # --- append new accounts -------------------------------------------------
    added = []
    for code, name, typ, parent, desc, reason in NEW_ACCOUNTS:
        a = {"id": "", "name": name, "code": code, "desc": desc, "type": typ,
             "mileage_rate": "0.000", "mileage_unit": "", "is_mileage": "false",
             "acct_no": "", "status": "Active", "currency": "NGN", "parent": parent}
        data.append(a)
        index[k(name)] = a
        added.append(a)

    # --- depth sort so headers precede children in the import file ----------
    def depth(a, seen=None):
        d, p = 0, a["parent"]
        seen = seen or set()
        while p and p.lower() not in seen:
            seen.add(p.lower())
            pn = index.get(k(p))
            if pn is None:
                break
            d += 1
            p = pn["parent"]
        return d

    for a in data:
        a["depth"] = depth(a)

    data.sort(key=lambda a: (a["code"] == "", a["code"] or "9999", a["name"]))
    data.sort(key=lambda a: (a["depth"], a["code"] if a["code"] else "9999", a["name"]))

    return data, index, added


if __name__ == "__main__":
    data, index, added = build()
    print("total accounts:", len(data), "new:", len(added))
    # integrity re-check
    import collections
    codes = [a["code"] for a in data if a["code"]]
    dup = [c for c, n in collections.Counter(codes).items() if n > 1]
    print("duplicate codes:", dup)
    names = set(k(a["name"]) for a in data)
    print("missing parents:", [(a["code"], a["name"], a["parent"]) for a in data
                               if a["parent"] and k(a["parent"]) not in names])
    print("missing codes:", [(a["name"], a["type"]) for a in data if not a["code"]])
    print("inactive:", [(a["code"], a["name"]) for a in data if a["status"] == "Inactive"])
    print("max depth:", max(a["depth"] for a in data))
