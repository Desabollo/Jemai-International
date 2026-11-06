"""Posting rules for the REAL August - October 2026 Access Bank statement.

The statement is machine-parsed by ``parse_statement.py``; every economic line
(692 of them) is then mapped to the improved Jemai chart of accounts by the
rules below.  A rule matches on the narration keywords, in priority order, so
the highest-priority rules are the ones that must never be shadowed:

    1. bank charges        - every transfer carries a commission / VAT / EMTL leg
    2. statutory taxes     - PAYE, VAT, CIT, tax consultants
    3. payroll             - the statement labels salary runs explicitly
    4. direct project cost - materials, subcontractors, artwork purchases
    5. operating expense   - rent, utilities, marketing, travel, professional
    6. advances            - director / staff / site floats
    7. revenue             - client, gallery and art receipts

Rule ids are stable and appear in the journal register so every posted line is
traceable back to the rule that produced it.

VAT: the 2026 journal is posted GROSS (the company's VAT registration status is
not evidenced in the files).  Receipts whose amounts are demonstrably 7.5 %
VAT-inclusive are quantified separately on the ``VAT_Review`` tab, with the
reclass entry to post if registration is confirmed.  See ``vat_inclusive()``.
"""
import re
from decimal import Decimal

BANK = "1021"
VAT_RATE = 0.075


def _r(rid, direction, kws, code, note="", vat=0.0):
    return {"id": rid, "dir": direction, "kw": [k.upper() for k in kws],
            "code": code, "note": note, "vat": vat}


# --------------------------------------------------------------------------- #
# 1. bank charges - check these first, they sit inside almost every narration
# --------------------------------------------------------------------------- #
BANK_RULES = [
    _r("BNK-01", "Out", ["COMMISSION MOBILE TRF", "COMMISSION TRF", "HANDLING CHARGE",
                         "COMMISSION ON", "COMMISSION PAY"], "6915",
       "Per-transfer commission. Highest-volume expense line in the ledger."),
    _r("BNK-02", "Out", ["VAT MOBILE TRF", "VAT TRF", "REVERSAL OF-VAT", "VAT ON COMMISSION",
                         "VAT THIS IS AN AUTOMATED"],
       "6914", "7.5% VAT charged by the bank on its own fees - not recoverable unless "
               "a valid bank tax invoice is held (then 1086)."),
    _r("BNK-03", "Out", ["FGN STAMP DUTY", "ELECTRONIC MONEY TRANSFER LEVY", "EMTL"],
       "6916", "N50 per qualifying transfer - FGN Electronic Money Transfer Levy."),
    _r("BNK-04", "Out", ["ACCOUNT MAINTENANCE FEE", "MAINTENANCE FEE ON CURRENT"],
       "6910", "Account maintenance fee (VAT on it is a separate line - see BNK-02)."),
    _r("BNK-05", "Out", ["AUTOMATED TRANSACTION ALERT", "SMS ALERT"], "6910",
       "Alert / notification charge."),
    _r("BNK-07", "Out", ["088ZCOT"], "6914",
       "VAT on the monthly account-maintenance charge (the charge itself is on a separate line)."),
    _r("BNK-06", "Out", ["OFF-CYCLE", "EMBASSY STATEMENT", "STATEMENT REQUEST",
                         "CARD FEE", "PENAL CHARGE", "SERVICE AND VAT CHARGE"], "6910",
       "Other bank charge."),
]

# --------------------------------------------------------------------------- #
# 2. statutory taxes & payroll bullets
# --------------------------------------------------------------------------- #
TAX_RULES = [
    _r("TAX-01", "Out", ["PAYE REMITTANCE", "PAYE PAYMENT", "PAYE"], "2044",
       "Settles the PAYE liability declared on the monthly payroll. The narration "
       "refers to April/May/July runs remitted late - the payroll accrual must be "
       "raised in the month of the payroll, not the month of payment."),
    _r("TAX-02", "Out", ["TAX CLEARANCE"], "6413",
       "Tax clearance / consultant handling fees - professional cost of compliance, "
       "not the tax itself."),
    _r("TAX-03", "Out", ["TAX FILING", "2025 TAX"], "2045",
       "CIT settlement for the 2025 year of assessment (opening liability 2045)."),
    _r("TAX-04", "Out", ["VAT PAYMENT", "VAT REMITTANCE", "REMITA VAT"], "2041",
       "VAT remitted to FIRS / state."),
    _r("TAX-05", "Out", ["REMITA HOLDING FUND", "HOLDING FUND COLLECTION"], "2044",
       "Remita fund collection - PAYE / WHT remitted through Remita. Split the tax types "
       "from the Remita receipt; PAYE is the assumed dominant component."),
    _r("PAY-01", "Out", ["PENSION", "NPF", "PENCOM", "OLAYINKA"], "6117",
       "Employer pension contribution."),
    _r("PAY-02", "Out", ["SALARY", "SALARIES", "WAGES", "PAYROLL", "DRIVER"], "6110",
       "Salary / wage run. The statement labels the August, September and October "
       "runs explicitly."),
    _r("PAY-03", "Out", ["HOUSING ALLOWANCE", "ALLOWANCE", "BONUS", "STIPEND"], "6111",
       "Allowances and bonuses (taxable benefits - include in the PAYE computation)."),
    _r("PAY-04", "Out", ["CONTRACT STAFF", "INTERN STIPEND"], "6112", "Contract staff / intern stipend."),
]

# --------------------------------------------------------------------------- #
# 3. direct project cost (cost of sales)
# --------------------------------------------------------------------------- #
COS_RULES = [
    _r("COS-01", "Out", ["ARTWORK", "ART WORK", "PAINTING", "ART CERTIFICATE", "CANVAS",
                         "ART SUPPL", "ART FRAMING", "PURCHASE AND HANGING"], "5019",
       "Visual art purchased for resale / installation (artists, framing, certificates)."),
    _r("COS-02", "Out", ["CURTAIN", "VELVET", "THROW PILLOW", "BEDDING", "FABRIC",
                         "TABLE CLOTH", "TABLECLOTH", "UPHOLSTERY"], "5012",
       "Fabrics, upholstery and soft furnishings for client projects."),
    _r("COS-03", "Out", ["CARPENTRY", "FURNITURE", "WARDROBE", "CABINET", "BUNK BED",
                         "FOLDABLE TABLE", "SHELVING"], "5021",
       "Carpentry / furniture subcontractor and finished furniture purchases."),
    _r("COS-04", "Out", ["TILE", "CEMENT", "SCREEDING", "MASONRY"], "5013",
       "Tiles, cement and building materials."),
    _r("COS-05", "Out", ["ELECTRICAL", "WIRING", "SOCKET"], "5015",
       "Electrical materials and fittings."),
    _r("COS-06", "Out", ["PLUMBING", "PUMPING"], "5016", "Plumbing materials and works."),
    _r("COS-07", "Out", ["POP", "CEILING", "PLASTER"], "5025", "POP / plastering and ceiling works."),
    _r("COS-08", "Out", ["AC REPAIR", "AC INSTALL", "AC BALANCE", "AIR CONDITION", "FRIDGE",
                         "MICROWAVE", "APPLIANCE", "SPEAKER", "CCTV", "SCALE", "TOOLS"], "5018",
       "Interior appliances, equipment and AV kit installed in client / gallery space."),
    _r("COS-09", "Out", ["GENERATOR", "FUEL - SITE"], "5035",
       "Site / gallery generator: purchase, servicing and fuel."),
    _r("COS-10", "Out", ["SITE EXPENSES", "SITE MATERIAL", "SITE SUPPLIES"], "5031",
       "Site petty-cash spend (see also ADV-03 where the transfer is a float top-up)."),
    _r("COS-11", "Out", ["UK ART LOGISTICS", "ART LOGISTICS", "HANGING", "INSTALLATION LABOUR"],
       "5032", "Transport and delivery of artworks / project items to site or overseas."),
    _r("COS-12", "Out", ["PRINTING", "BROCHURE", "PRESENTATION BOARD"], "5042",
       "Printing and presentation material for projects."),
    _r("COS-14", "Out", ["DOOR"], "5021", "Gallery / project door joinery."),
    _r("COS-15", "Out", ["PAINT"], "5014", "Paint and finishing materials (artists' paints are COS-01)."),
    _r("COS-16", "Out", ["AUPPLIES", "GALLERY SUPPLIES", "SUPPLIES"], "5019",
       "Gallery / studio supplies."),
    _r("COS-13", "Out", ["PRODUCTION BALANCE", "PRODUCTION DEPOSIT"], "5041",
       "Content / 3D production delivered for the project."),
]

# --------------------------------------------------------------------------- #
# 4. operating expenses
# --------------------------------------------------------------------------- #
OPEX_RULES = [
    _r("OPS-01", "Out", ["RENT 1ST PAYMENT", "RENT PAYMENT", "OFFICE RENT", "SERVICE CHARGE"],
       "1071", "Rent paid in advance - held as a prepayment (1071) and released to 6210 "
               "over the lease period. Do not expense the full amount on day one."),
    _r("OPS-02", "Out", ["RENT REFUND"], "6613",
       "Refund of rent previously collected. Review against the original receipt before "
       "year-end; may belong against 1071 instead."),
    _r("OPS-03", "Out", ["AIRBNB", "SHORTLET"], "6514",
       "Short-let / serviced accommodation for staff or clients."),
    _r("OPS-04", "Out", ["ACCOMMODATION", "ACOMMODSTION", "ACOMMODATION"], "1091",
       "Furnishing of company accommodation - capitalised as furniture & fittings "
       "(or to 5021 if it belongs to a client project)."),
    _r("OPS-05", "Out", ["IG ADS", "MARKETING ADS", "FACEBOOK", "ADS"], "6313",
       "Paid social media advertising (Instagram / Facebook)."),
    _r("OPS-06", "Out", ["FLIGHT", "AIRPEACE", "AIR TICKET", "AIRFARE"], "6515",
       "Flight bookings - travel cost."),
    _r("OPS-07", "Out", ["VISA"], "6515", "Visa processing fees for staff / client travel."),
    _r("OPS-08", "Out", ["FUEL", "DIESEL", "PETROL"], "6512",
       "Fuel and vehicle running (site generator fuel is COS-09 instead)."),
    _r("OPS-09", "Out", ["BOLT", "UBER", "TAXI", "LOCAL TRAVEL"], "6511", "Local travel."),
    _r("OPS-10", "Out", ["LIGHT", "ELECTRICITY", "IKEDC", "EKO DISCO", "UTILITY", "WATER"],
       "6211", "Utilities."),
    _r("OPS-11", "Out", ["LOGISTIC", "LOGOSTIC", "DISPATCH", "DELIVERY", "COURIER"], "6510",
       "Transport & logistics (courier, dispatch, haulage)."),
    _r("OPS-12", "Out", ["MAINTENANCE", "REPAIR", "CAR WASH", "FUMIGATION"], "6216",
       "Repairs and maintenance."),
    _r("OPS-13", "Out", ["FLOWERS", "VASES", "DECOR", "CHRISTMAS"], "5019",
       "Decor purchases for the gallery / client projects."),
    _r("OPS-14", "Out", ["RENTAL", "HIRE"], "5034", "Equipment, chair and speaker hire."),
    _r("OPS-15", "Out", ["PAINT AND SIP SUPPLIES", "PAINT AND SIP CANDLE", "CUPS AND SWEET",
                         "DRINKS AND SNACK", "CANDLE", "SIP AND PAINT"], "6316",
       "Paint-and-sip event consumables - event cost."),
    _r("OPS-16", "Out", ["FOOD", "LUNCH", "REFRESHMENT", "WELFARE", "SNACK"], "6113",
       "Staff welfare and meals."),
    _r("OPS-17", "Out", ["RETAINER", "4TH QUARTER SERVICE FEE", "SERVICE FEE",
                         "NHF", "NSITF", "ITF", "BOOKKEEPING", "ACCOUNTING"], "6411",
       "Accounting / bookkeeping retainer including statutory registrations."),
    _r("OPS-18", "Out", ["AUDIT"], "6412", "External audit fee."),
    _r("OPS-19", "Out", ["LEGAL", "SOLICITOR", "LAWYER"], "6410", "Legal fees."),
    _r("OPS-20", "Out", ["PERMIT", "LICENCE", "LICENSE", "CAC"], "6415",
       "Permits, licences and CAC filings."),
    _r("OPS-21", "Out", ["DOMAIN", "HOSTING", "WEBSITE", "ECOMMERCE", "IT SUPPORT"], "6223",
       "Website / domain / hosting build and support."),
    _r("OPS-22", "Out", ["SOFTWARE", "SUBSCRIPTION", "ZOHO", "AUTOCAD", "SKETCHUP"], "6220",
       "Software subscriptions."),
    _r("OPS-23", "Out", ["BANK CHARGE"], "6910", "Bank charge."),
    _r("OPS-24", "Out", ["LAUNDRY", "CLEANING"], "6217", "Cleaning and sanitation."),
    _r("OPS-25", "Out", ["TRAVEL", "HOTEL", "LODGING"], "6514", "Hotel and lodging."),
    _r("OPS-26", "Out", ["INVOICE 358", "WLF COMMISSION"], "6317",
       "Sales commission / referral fee paid."),
    _r("OPS-27", "Out", ["CAR ACCESS", "CARACCESSORIES"], "6512", "Vehicle running costs."),
    _r("OPS-28", "Out", ["GAS AND REFILL", "GAS REFILL", "COOKING GAS"], "6211",
       "Gas refill for the gallery / office."),
    _r("OPS-29", "Out", ["TIP"], "6610", "Tip / gratuity - sundry."),
    _r("OPS-30", "Out", ["SALON"], "6610", "Personal-care item charged to the company - review."),
]

# --------------------------------------------------------------------------- #
# 5. advances and balance-sheet payments (checked before generic expense rules)
# --------------------------------------------------------------------------- #
ADV_RULES = [
    _r("ADV-01", "Out", ["ESTHER", "JESUTOFUNMI", "AKIN-AJAYI", "AKIN AJAYI"], "1045",
       "Payment to the director - posted to her current account, not to expense. "
       "N14.2m was already receivable at 1 Jan 2026; every naira paid out increases it."),
    _r("ADV-02", "Out", ["AKINAJAYI", "AKINLOLUWA", "EMMANUEL AKINOLUWA"], "1047",
       "Payment to a related party (family / associate)."),
    _r("ADV-05", "Out", ["JEMAI EXPENSES REFUND", "SITE EXPENSES REFUND"], "5031",
       "Reimbursement of site / company expenses to the person who incurred them."),
    _r("ADV-03", "Out", ["SITE EXPENSES", "SITE FLOAT"], "1013",
       "Transfer of a site float to the site team - an advance, expensed when the site "
       "returns are journalised."),
    _r("ADV-04", "Out", ["LOAN REPAYMENT", "LOAN"], "2052", "Loan principal repayment."),
]

# --------------------------------------------------------------------------- #
# 6. receipts
# --------------------------------------------------------------------------- #
REC_RULES = [
    _r("REC-03", "In", ["GALLERY SPACE", "PAYMENT FOR THE HALL", "SPACE BOOKING", "APPBOOKING",
                        "BOOKING", "FOR THE HALL"], "4030", "Gallery space / hall hire income."),
    _r("REC-01", "In", ["ART GALLERY ENTRY", "ENTRY FEE", "ENTRY TICKET", "GALLERY ENTRY",
                        "JEMAI ENTRY", "ENTRY"], "4010",
       "Gallery ticketing / entry income. Part of the gallery line launched in 2026."),
    _r("REC-02", "In", ["PAINT AND SIP", "SIP AND PAINT", "PAIHT AND SIP", "PAINTINF"], "4010",
       "Paint-and-sip event income."),
    _r("REC-04", "In", ["ARTWORK", "PURCHASE AND HANGING", "ART CERTIFICATE", "ART MAINTAINANCE",
                        "ART SUPPL", "WHERE LIGHT FALLS", "ART", "PICTURES", "WALL ART"], "4014",
       "Sale of art, pictures and decor; artwork hanging / restoration services."),
    _r("REC-05", "In", ["OFFICE DESIGN", "INTERIOR", "FIT-OUT", "FITOUT", "DESIGN"], "4012",
       "Interior design / fit-out project receipt."),
    _r("REC-12", "In", ["AFRC", "OPTIMUS PRO"], "4021",
       "Corporate client collection - project not identified in the narration."),
    _r("REC-13", "In", ["ESTHER", "JESUTOFUNMI", "AKIN-AJAYI"], "1045",
       "Repayment received from the director - reduces 1045."),
    _r("REC-06", "In", ["WLF COMMISSION", "COMMISSION"], "4010",
       "Commission earned on an art sale (where light falls)."),
    _r("REC-07", "In", ["LOAN"], "2052", "Loan / facility received."),
    _r("REC-08", "In", ["INVESTMENT", "SHARE CAPITAL"], "3061",
       "Funds injected by an investor / shareholder."),
    _r("REC-09", "In", ["EXPENSES REFUND", "EXPENSES BALANCE REFUND", "REFUND"], "1013",
       "Site float / expense refund returned to the bank."),
    _r("REC-10", "In", ["SUPPLIER REFUND"], "4901", "Refund received from a supplier."),

]

RULES = BANK_RULES + TAX_RULES + COS_RULES + OPEX_RULES + ADV_RULES + REC_RULES
RULE_INDEX = {r["id"]: r for r in RULES}

# --------------------------------------------------------------------------- #
# explicit overrides for lines where the narration alone is not enough
# key: (posted date, signed amount, direction)
# --------------------------------------------------------------------------- #
OVERRIDES = {
    # AFRC / OLAYEMI OLADEJO receipts - exterior/art fabrication contract for AFRC
    ("-", 0, ""): None,
}
OVERRIDES.pop(("-", 0, ""))

# payee-keyed overrides (substring of the narration, uppercase) - applied after rules
PAYEE_OVERRIDES = [
    # (substring, direction, code, note)
    ("ABOLLO HENRY EDAFE", "Out", "6413",
     "N3.0m paid to a tax consultant for 2025 tax clearance; treated as professional "
     "fees pending the consultant's invoice - reclassify to 2045 if it is the tax itself."),
    ("CLOUD ERP ACCOUNTING", "Out", "6411", "Accounting retainer + statutory registrations."),
    ("TAEILLO LIMITED", "Out", "1091",
     "Furniture for the company accommodation - capitalised. Move to 5021 if it is a "
     "client project."),
    ("KINGDAVID INTERIORS", "Out", "5012", "Curtains, delivery and installation - client project."),
    ("MARSA SIGNATURE", "Out", "5019", "Christmas decor items for the gallery."),
    ("KANMI PRODUCTION", "Out", "6311",
     "Content / film production for the brand - move to 5041 if it is client-project work."),
    ("QUONOS TECHNOLOGIES", "Out", "1121",
     "E-commerce website build and modifications - capitalised as an intangible "
     "(or 6223 if it is only support)."),
    ("VUGOHOMES", "Out", "6514", "Airbnb / short-let accommodation."),
    ("MINTANA GLOBAL LOGISTICS", "Out", "5032", "UK art logistics - shipping artworks."),
    ("TRAVELPLOTT", "Out", "6515", "UK visa processing / travel."),
    ("SQUAD KANA", "Out", "6512", "Fuel."),
    ("SUNBETH ENERGIES", "Out", "6512", "Fuel."),
    ("ENEMAKWU ENERGY", "Out", "6512", "Fuel."),
    ("BAIYEWU", "Out", "6110", "Wages."),
    ("SABIS DANIEL BIO", "Out", "6110", "Wages."),
    ("GODIYA", "Out", "6110", "Wages."),
    ("ADEYEYE OLUWADAMILARE", "Out", "5026", "Painting balance - finishing labour."),
    ("ADENIYI EMMANUEL", "Out", "5021", "Furniture balance - carpentry."),
    ("LOWE RENTAL", "Out", "6216", "Annual generator maintenance / re-location."),
    ("AJAO HAMMED", "Out", "2045", "2025 income tax filing payment."),
    ("AIICO", "Out", "6218", "Insurance premium."),
    ("ACCESS BANK LOAN", "Out", "2052", "Loan repayment to Access Bank."),
    ("LOAN/ESTHER", "In", "2053",
     "Loan received from the director - a payable to her, not revenue."),
    ("PRIMERO", "Out", "6418", "Consultancy."),
    ("OKPARAOCHA", "Out", "6110", "August wages."),
    ("HAMZA ALIYU", "Out", "6110", "Wages."),
    ("ODEYEMI MAYOWA", "Out", "6110", "Wages."),
    ("LAMISHI", "Out", "5019", "Art framing / artwork purchase."),
    ("SUSAN EFFANGA", "Out", "5019", "Artwork purchase."),
    ("ABDUL GENTECH", "Out", "5035", "Gallery generator supply, parts and servicing."),
    ("ACHIKE TOCHUKWU", "Out", "5018", "Mirror and fittings for the project."),
    ("ADUKUYUK ATINJAH", "Out", "5026", "Painting balance / finishing labour."),
]
PAYEE_OVERRIDES = [p for p in PAYEE_OVERRIDES if p[2]]

# payees that are staff - they appear inside salary runs and as ad-hoc wage lines
STAFF_HINTS = ("SALARY", "WAGES", "BONUS", "ALLOWANCE", "DRIVER", "CARETAKER")

# receipts that are demonstrably 7.5% VAT-inclusive
def vat_inclusive(amount):
    """True when the amount is exactly a whole-naira figure grossed up at 7.5%."""
    a = Decimal(str(amount))
    net = a / (Decimal("1") + Decimal("0.075"))
    return net == net.to_integral_value() and net > 0


def clean(desc):
    d = desc.upper()
    for bad in ("\u2019", "\u2018", "  "):
        d = d.replace(bad, " ")
    return " ".join(d.split())


PAY_RUN = [r for r in TAX_RULES if r["id"] in ("PAY-02", "PAY-03", "PAY-04")]


def _first(bucket, d, direction):
    for r in bucket:
        if r["dir"] != direction:
            continue
        for k in r["kw"]:
            if k in d:
                return (r["id"], r["code"], r["note"])
    return None


def classify(desc, direction, amount=0.0):
    """Returns (rule_id, code, note).

    Priority (each stage can only be reached if the previous did not match):
        1  bank charges      - a N50 commission on a salary transfer is a bank charge
        2  explicit payee    - judgement calls that the narration alone cannot settle
        3  payroll run       - 'September Salary / ESTHER' is payroll, not director advance
        4  advances          - director / related party / site float
        5  statutory taxes
        6  direct project cost
        7  operating expense
        8  revenue
    """
    d = clean(desc)
    hit = _first(BANK_RULES, d, direction)
    if hit:
        return hit
    for needle, dirn, code, note in PAYEE_OVERRIDES:
        if dirn == direction and needle.upper() in d:
            return ("OVR", code, note)
    hit = _first(PAY_RUN, d, direction)
    if hit:
        return hit
    for bucket in (ADV_RULES, TAX_RULES, COS_RULES, OPEX_RULES,
                   REC_RULES if direction == "In" else []):
        hit = _first(bucket, d, direction)
        if hit:
            return hit
    if direction == "In":
        return ("REC-99", "4010",
                "Unlabelled individual receipt - assumed gallery ticket / art sale "
                "(typical ticket values N5,375 = N5,000 + 7.5% VAT, N10,750, N21,500). "
                "Agree to the ticket register.")
    return ("EXP-99", "7999", "Payment not matched by any rule - reclassify before import.")


def payee_of(desc):
    """Best-effort counterparty: the last narration segment that looks like a name."""
    parts = [p.strip() for p in re.split(r"[/|]", desc) if p.strip()]
    for part in reversed(parts):
        letters = sum(1 for ch in part if ch.isalpha())
        if letters >= 4 and not part.replace(" ", "").isdigit():
            return re.sub(r"\s+", " ", part)[:60]
    return ""


def narrative(desc):
    """Tidy narration for the journal: collapse the double spaces the PDF leaves."""
    return re.sub(r"\s+", " ", desc).strip()
