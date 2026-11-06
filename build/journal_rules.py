"""Posting rules that convert Access Bank statement lines into double-entry journals
on the improved Jemai chart of accounts.

Each rule fixes the debit and credit account for one *shape* of transaction:
  - receipt : Dr 1021 Access Bank            / Cr <revenue, receivable or liability>
  - payment : Dr <expense, asset, liability> / Cr 1021 Access Bank

VAT modes
  NONE          no VAT in the bank line (the company is not VAT-registered, or the item is out of scope)
  OUT-INCL      gross receipt contains 7.5% output VAT -> split to 2041 (only when registered)
  IN-INCL       gross payment contains 7.5% input VAT -> split to 1081 (only when registered and invoice held)
"""

BANK = "1021"

# (rule_id, description, keywords, direction, dr, cr, vat, wht, note)
RULES = [
    # ---------------- receipts -------------------------------------------------
    ("REC-01", "Client receipt - interior design / fit-out project", "client deposit, first deposit, tranche, transfer from, interior deco",
     "In", BANK, "4012", "NONE", 0.0, "Recognise when the performance obligation is satisfied (IFRS 15). A deposit received before the work starts is a contract liability (2031), not revenue."),
    ("REC-02", "Client receipt - furniture production & supply", "furniture, wardrobe, cabinetry",
     "In", BANK, "4013", "NONE", 0.0, "Separate performance obligation from the fit-out where sold separately."),
    ("REC-03", "Client receipt - architectural / design fee", "architectural, design fee, drawings",
     "In", BANK, "4011", "NONE", 0.0, ""),
    ("REC-04", "Client receipt - consultancy & project management", "consultancy, project management, supervision",
     "In", BANK, "4016", "NONE", 0.0, ""),
    ("REC-05", "Client receipt - visual art / decor", "art, painting sale, decor",
     "In", BANK, "4014", "NONE", 0.0, ""),
    ("REC-06", "Retention released by client", "retention release, retention payment",
     "In", BANK, "1042", "NONE", 0.0, "Collection of a receivable - never revenue (that would double count the contract)."),
    ("REC-07", "Supplier refund received", "refund from supplier, supplier refund",
     "In", BANK, "4901", "NONE", 0.0, "Other income, or net against the original cost if the invoice is in the same period."),
    ("REC-08", "Director repays / funds the current account", "from esther, from the md, salary to jemai",
     "In", BANK, "1045", "NONE", 0.0, "Reduces the director current account receivable. If the director is genuinely funding the company, the balance may turn into a payable (2061) - classify on substance."),
    ("REC-09", "Related party repayment", "from akinyemi, from emmanuel",
     "In", BANK, "1047", "NONE", 0.0, ""),
    ("REC-10", "Loan received", "loan, facility, credit received",
     "In", BANK, "2052", "NONE", 0.0, "Interest accrues to 6912."),
    ("REC-11", "Share application money received", "share capital, share application",
     "In", BANK, "2057", "NONE", 0.0, "Move to 3011 only when the shares are allotted and the CAC return is filed."),
    ("REC-12", "Interest income", "interest credit, interest earned",
     "In", BANK, "4043", "NONE", 0.0, ""),
    ("REC-13", "Unidentified credit - suspense", "unidentified, unknown payer",
     "In", BANK, "4999", "NONE", 0.0, "Must be cleared before the year-end close."),

    # ---------------- bank charges --------------------------------------------
    ("BNK-01", "Bank commission on transfer", "commission trf, commission mobile trf, nip transfer charges",
     "Out", "6915", BANK, "NONE", 0.0, "Highest-volume line in the ledger - was buried in legacy 6050."),
    ("BNK-02", "VAT on bank charges - NOT recoverable (unregistered)", "vat trf, vat mobile trf, vat deduction",
     "Out", "6914", BANK, "NONE", 0.0, "Use this rule while Jemai is not VAT-registered. Switch to BNK-02R once registered."),
    ("BNK-02R", "VAT on bank charges - recoverable (registered)", "vat trf, vat mobile trf, vat deduction",
     "Out", "1086", BANK, "NONE", 0.0, "Use once VAT-registered and the bank tax invoice is held."),
    ("BNK-03", "FGN Electronic Money Transfer Levy", "electronic money transfer levy",
     "Out", "6916", BANK, "NONE", 0.0, "N50 per qualifying transfer. Separately disclosable levy."),
    ("BNK-04", "Account maintenance fee", "account maintenance",
     "Out", "6910", BANK, "NONE", 0.0, "The VAT element appears as a separate bank line - use BNK-02."),
    ("BNK-05", "SMS / alert charges", "sms alert",
     "Out", "6910", BANK, "NONE", 0.0, ""),
    ("BNK-06", "Other or penal bank charge", "penal charge, stamp duty, visa_business_card, card fee",
     "Out", "6910", BANK, "NONE", 0.0, ""),
    ("BNK-07", "Debit interest", "debit interest",
     "Out", "6911", BANK, "NONE", 0.0, "Overdraft / debit interest - finance cost."),
    ("BNK-08", "Loan interest", "loan interest",
     "Out", "6912", BANK, "NONE", 0.0, ""),
    ("BNK-09", "Payment gateway fee", "paystack fee, gateway charge, checkout fee",
     "Out", "6913", BANK, "NONE", 0.0, ""),

    # ---------------- direct cost: subcontractors ------------------------------
    ("COS-01", "Carpentry & furniture subcontractor", "furniture deposit, furniture balance, carpentry, agege part payment, wardrobe",
     "Out", "5021", BANK, "NONE", 5.0, "WHT 5% on contract values for unincorporated suppliers; 2%/1.5% on goods. Withhold at source and remit by the 21st."),
    ("COS-02", "Tiling & masonry subcontractor", "tiling, tiler, masonry, screeding labour",
     "Out", "5022", BANK, "NONE", 5.0, ""),
    ("COS-03", "Electrical works subcontractor", "electrical works, electrical installation, wiring",
     "Out", "5023", BANK, "NONE", 5.0, ""),
    ("COS-04", "Plumbing works subcontractor", "plumbing works, plumber",
     "Out", "5024", BANK, "NONE", 5.0, ""),
    ("COS-05", "POP / plastering & ceiling works", "pop, plastering, ceiling works, screeding",
     "Out", "5025", BANK, "NONE", 5.0, ""),
    ("COS-06", "Painting & finishing labour", "painting touch up, painting deposit, painting balance, painting",
     "Out", "5026", BANK, "NONE", 5.0, ""),
    ("COS-07", "Cladding / fluted panel works", "fluted panel, cladding, panel works",
     "Out", "5027", BANK, "NONE", 5.0, ""),
    ("COS-08", "Installation labour", "installation, installer, fitting labour",
     "Out", "5028", BANK, "NONE", 5.0, ""),
    ("COS-09", "Site security / temporary labour", "site security, temporary labour, guard",
     "Out", "5029", BANK, "NONE", 5.0, ""),

    # ---------------- direct cost: materials -----------------------------------
    ("COS-10", "Wood, boards & carpentry materials", "plywood, wood, board, mdf, material",
     "Out", "5011", BANK, "NONE", 2.0, "WHT on supply of goods: 2% (companies) / 1.5% (individuals) under the 2024 Regulations."),
    ("COS-11", "Fabrics, upholstery & soft furnishings", "curtain, fabric, upholstery, cushion, blind",
     "Out", "5012", BANK, "NONE", 2.0, ""),
    ("COS-12", "Tiles, cement & building materials", "tile, cement, building material, sand",
     "Out", "5013", BANK, "NONE", 2.0, ""),
    ("COS-13", "Paint, screeding & finishing materials", "paint supply, emulsion, screeding material",
     "Out", "5014", BANK, "NONE", 2.0, ""),
    ("COS-14", "Electrical materials & fittings", "electrical fittings, cables, socket, light fitting",
     "Out", "5015", BANK, "NONE", 2.0, ""),
    ("COS-15", "Plumbing materials & fixtures", "plumbing material, sanitary, water heater",
     "Out", "5016", BANK, "NONE", 2.0, ""),
    ("COS-16", "Glass, windows & aluminium", "mirror, glass, aluminium, window",
     "Out", "5017", BANK, "NONE", 2.0, ""),
    ("COS-17", "Interior appliances & equipment", "appliance, air conditioner, tv, fridge",
     "Out", "5018", BANK, "NONE", 2.0, ""),
    ("COS-18", "Visual art materials", "art material, canvas, frame",
     "Out", "5019", BANK, "NONE", 2.0, ""),

    # ---------------- direct cost: site operations -----------------------------
    ("COS-19", "Site petty cash", "site petty cash, petty cash - site",
     "Out", "5031", BANK, "NONE", 0.0, ""),
    ("COS-20", "Site transport & delivery", "transportation, transport, logistics, delivery, dispatch to site",
     "Out", "5032", BANK, "NONE", 0.0, "Only where the movement is chargeable to a client contract; staff movement goes to 6511."),
    ("COS-21", "Equipment / scaffold rental", "scaffold, rental, equipment hire",
     "Out", "5033", BANK, "NONE", 0.0, ""),
    ("COS-22", "Generator / fuel - site", "diesel, generator, fuel - site",
     "Out", "5035", BANK, "NONE", 0.0, ""),
    ("COS-23", "Project cleaning / waste disposal", "deep cleaning, cleaning, waste, refuse",
     "Out", "5037", BANK, "NONE", 0.0, ""),
    ("COS-24", "3D rendering & visualisation direct cost", "3d render, 3d visualization, rendering",
     "Out", "5041", BANK, "NONE", 5.0, "Direct cost of the design deliverable - not marketing."),
    ("COS-25", "Printing / presentation boards", "presentation board, printing - project, sample board",
     "Out", "5042", BANK, "NONE", 0.0, ""),
    ("COS-26", "Project insurance & bonds", "all risk insurance, performance bond, project insurance",
     "Out", "5038", BANK, "NONE", 0.0, ""),

    # ---------------- staff ----------------------------------------------------
    ("ADM-01", "Salaries and employee cost", "salary, salaries, payroll",
     "Out", "6110", BANK, "NONE", 0.0, "PAYE, pension (8%/10%) and NHF must be withheld and remitted separately."),
    ("ADM-02", "Staff allowances & bonuses", "allowance, bonus",
     "Out", "6111", BANK, "NONE", 0.0, ""),
    ("ADM-03", "Contract staff / intern stipend", "intern, stipend, contract staff",
     "Out", "6112", BANK, "NONE", 0.0, ""),
    ("ADM-04", "Staff welfare & meals", "welfare, staff lunch, refreshment",
     "Out", "6113", BANK, "NONE", 0.0, ""),
    ("ADM-05", "Medical & welfare", "medical, hospital, hmo",
     "Out", "6114", BANK, "NONE", 0.0, ""),
    ("ADM-06", "Staff training & development", "training, course, seminar",
     "Out", "6115", BANK, "NONE", 0.0, "N650,000 of 2025 training had no matching salary base - retain evidence."),
    ("ADM-07", "PAYE remittance to LIRS", "paye remittance, paye payment",
     "Out", "2044", BANK, "NONE", 0.0, ""),
    ("ADM-08", "Pension remittance to PFA", "pension remittance, pencom",
     "Out", "2072", BANK, "NONE", 0.0, ""),

    # ---------------- office & IT ---------------------------------------------
    ("ADM-09", "Rent & service charge", "rent, service charge",
     "Out", "6210", BANK, "NONE", 10.0, "WHT 10% on rent paid to a landlord. Test IFRS 16 (right-of-use)."),
    ("ADM-10", "Office utilities", "electricity, ekedc, water bill, utility",
     "Out", "6211", BANK, "NONE", 0.0, ""),
    ("ADM-11", "Internet & data", "internet, data subscription, broadband",
     "Out", "6212", BANK, "NONE", 0.0, ""),
    ("ADM-12", "Telephone & airtime", "mtn, airtel, airtime, glo, telephone",
     "Out", "6213", BANK, "NONE", 0.0, ""),
    ("ADM-13", "Office supplies & stationery", "stationery, office supplies",
     "Out", "6214", BANK, "NONE", 0.0, ""),
    ("ADM-14", "Printing & photocopying", "printing, photocopy",
     "Out", "6215", BANK, "NONE", 0.0, ""),
    ("ADM-15", "Repairs & maintenance", "repair, maintenance of office",
     "Out", "6216", BANK, "NONE", 0.0, ""),
    ("ADM-16", "Cleaning & sanitation - office", "office cleaning",
     "Out", "6217", BANK, "NONE", 0.0, ""),
    ("ADM-17", "Insurance - office", "insurance",
     "Out", "6218", BANK, "NONE", 0.0, ""),
    ("ADM-18", "Security - office", "security",
     "Out", "6219", BANK, "NONE", 0.0, ""),
    ("ADM-19", "Design software subscription", "autocad, sketchup, lumion, software subscription",
     "Out", "6221", BANK, "NONE", 0.0, ""),
    ("ADM-20", "Zoho Books / cloud software", "zoho",
     "Out", "6222", BANK, "NONE", 0.0, ""),
    ("ADM-21", "Website hosting & domain", "hosting, domain, website",
     "Out", "6223", BANK, "NONE", 0.0, ""),
    ("ADM-22", "IT support & maintenance", "it support, computer repair",
     "Out", "6224", BANK, "NONE", 0.0, ""),

    # ---------------- selling & marketing --------------------------------------
    ("ADM-23", "Advertising & paid media", "advert, billboard, radio, tv advert",
     "Out", "6310", BANK, "NONE", 0.0, ""),
    ("ADM-24", "Photography & videography", "photography, shoot, videography",
     "Out", "6311", BANK, "NONE", 0.0, ""),
    ("ADM-25", "Branding & graphics design", "branding, logo, graphics, design - brand",
     "Out", "6312", BANK, "NONE", 0.0, "Own-account marketing goes here; branding delivered to a client is 5041-5044."),
    ("ADM-26", "Social media promotions", "instagram, facebook, social media, boost",
     "Out", "6313", BANK, "NONE", 0.0, ""),
    ("ADM-27", "Client entertainment & gifts", "clients gift, gift, entertainment, lunch with client",
     "Out", "6314", BANK, "NONE", 0.0, "Restricted for CIT - keep a gifts register."),
    ("ADM-28", "Proposal / pitch expenses", "proposal, pitch, tender",
     "Out", "6315", BANK, "NONE", 0.0, ""),
    ("ADM-29", "Events & exhibition", "exhibition, event, launch",
     "Out", "6316", BANK, "NONE", 0.0, ""),
    ("ADM-30", "Sales commission & referral", "commission to agent, referral, introducer",
     "Out", "6317", BANK, "NONE", 5.0, ""),

    # ---------------- professional & regulatory --------------------------------
    ("ADM-31", "Legal fees", "legal, solicitor, barrister",
     "Out", "6410", BANK, "NONE", 5.0, ""),
    ("ADM-32", "Accounting & bookkeeping fees", "accounting fee, bookkeeping",
     "Out", "6411", BANK, "NONE", 5.0, ""),
    ("ADM-33", "Audit fees", "audit fee, auditor",
     "Out", "6412", BANK, "NONE", 5.0, ""),
    ("ADM-34", "Tax consultancy fees", "tax consultancy, tax advisory",
     "Out", "6413", BANK, "NONE", 5.0, ""),
    ("ADM-35", "CAC filing fees", "cac, annual return",
     "Out", "6414", BANK, "NONE", 0.0, ""),
    ("ADM-36", "Business permits & licences", "permit, licence, lasaa, sign permit",
     "Out", "6415", BANK, "NONE", 0.0, ""),
    ("ADM-37", "Professional memberships", "membership, subscription fee - body",
     "Out", "6416", BANK, "NONE", 0.0, ""),
    ("ADM-38", "Penalties & fines", "penalty, fine",
     "Out", "6417", BANK, "NONE", 0.0, "Non-allowable for CIT - isolate it."),
    ("ADM-39", "Other consultancy fees", "consultancy, consultant, covenant university",
     "Out", "6418", BANK, "NONE", 5.0, ""),

    # ---------------- travel ---------------------------------------------------
    ("ADM-40", "Transport & logistics - admin", "logistics", "Out", "6510", BANK, "NONE", 0.0, ""),
    ("ADM-41", "Local travel", "local travel, bolt, uber, taxi, site visit", "Out", "6511", BANK, "NONE", 0.0, ""),
    ("ADM-42", "Fuel & vehicle running", "fuel, petrol, diesel - vehicle", "Out", "6512", BANK, "NONE", 0.0, ""),
    ("ADM-43", "Dispatch & postage", "postage, courier, dhl", "Out", "6513", BANK, "NONE", 0.0, ""),
    ("ADM-44", "Hotel & lodging", "hotel, lodging, accommodation", "Out", "6514", BANK, "NONE", 0.0, ""),
    ("ADM-45", "Air travel & flight bookings", "flight, airline, air ticket", "Out", "6515", BANK, "NONE", 0.0, ""),

    # ---------------- general --------------------------------------------------
    ("ADM-46", "Petty cash & sundry expenses", "sundry, petty cash", "Out", "6610", BANK, "NONE", 0.0, ""),
    ("ADM-47", "Charitable donations", "church, donation, tithe, offering", "Out", "6612", BANK, "NONE", 0.0, "Non-allowable for CIT."),

    # ---------------- balance sheet movements ----------------------------------
    ("BS-01", "Payment to MD / director current account", "to esther, esther jesutofunmi, akinaJayi esther, e. akin-ajayi",
     "Out", "1045", BANK, "NONE", 0.0, "Director receivable. Non-allowable for CIT until business purpose is substantiated and repaid; assess ECL (7024)."),
    ("BS-02", "Payment to related party - family / associates", "akinloluwa, akinyemi, felicia, ime ubongabasi",
     "Out", "1047", BANK, "NONE", 0.0, ""),
    ("BS-03", "Loan principal repayment", "repayment, loan refund", "Out", "2052", BANK, "NONE", 0.0, ""),
    ("BS-04", "Supplier / materials advance", "deposit for materials, advance to supplier", "Out", "1072", BANK, "NONE", 0.0, ""),
    ("BS-05", "VAT remittance to FIRS", "vat payment, vat remittance", "Out", "2042", BANK, "NONE", 0.0, ""),
    ("BS-06", "WHT remittance to FIRS / state", "wht payment, withholding tax remittance", "Out", "2043", BANK, "NONE", 0.0, ""),
    ("BS-07", "CIT / development levy payment", "cit payment, income tax payment, tet payment", "Out", "2045", BANK, "NONE", 0.0, ""),
    ("BS-08", "Customer refund", "refund to client, client refund, refund", "Out", "4042", BANK, "NONE", 0.0, "Contra-revenue. If the refund relates to an earlier year, clear the refunds payable account 2064 first."),
    ("BS-09", "Fixed asset purchase", "laptop, computer, phone, furniture - office, equipment purchase", "Out", "1093", BANK, "NONE", 0.0, "Capitalise; do not expense. Depreciation per NTA 2025 class rates."),
    ("BS-10", "Security / rent deposit paid", "deposit - office, security deposit", "Out", "1131", BANK, "NONE", 0.0, ""),
    ("BS-11", "Interbank / wallet transfer out", "transfer to opay, transfer to moniepoint, transfer to palmpay", "Out", "1035", BANK, "NONE", 0.0, ""),
    ("BS-12", "Staff advance", "staff advance, imprest", "Out", "1044", BANK, "NONE", 0.0, ""),
    ("BS-13", "Unidentified payment - suspense", "unknown payee, unidentified", "Out", "7010", BANK, "NONE", 0.0, "Clear before period close."),
    ("BS-14", "Refund of share application money", "share refund", "Out", "2057", BANK, "NONE", 0.0, ""),
]

RULE_INDEX = {r[0]: r for r in RULES}

# explicit overrides for the archived Aug-Oct 2025 statement, based on the
# narration-by-narration review recorded in the working papers
OVERRIDES = {
    "TXN-2025-0676": "REC-01",   # Joshua Oluwadawolemi Abolade - project receipt
    "TXN-2025-0697": "REC-08",   # MD funds the July salary
    "TXN-2025-0703": "REC-08",   # MD funds Nissi's salary
    "TXN-2025-0711": "REC-01",   # The Victoria King - first deposit, interior deco
    "TXN-2025-0712": "COS-24",   # 3D render - Victor Israel Baiyewu
    "TXN-2025-0718": "REC-01",
    "TXN-2025-0719": "BS-01",    # 'Loan refund' to the MD - see flag
    "TXN-2025-0722": "ADM-41",   # Paystack checkout 'Site visit' - flagged
    "TXN-2025-0728": "COS-14",   # electrical fittings (materials) from De Great Ify
    "TXN-2025-0731": "COS-11",   # curtain deposit
    "TXN-2025-0732": "COS-01",   # furniture deposit - Adeniyi Emmanuel
    "TXN-2025-0735": "COS-06",   # painting deposit
    "TXN-2025-0742": "REC-01",
    "TXN-2025-0743": "ADM-45",   # flight booking
    "TXN-2025-0746": "BS-02",    # Emmanuel Akinloluwa Akin-Ajayi - related party
    "TXN-2025-0749": "COS-01",   # Agege part payment - Adeniyi Emmanuel
    "TXN-2025-0752": "BS-08",    # fluted panel refund to the client
    "TXN-2025-0753": "COS-06",   # painting balance
    "TXN-2025-0756": "ADM-27",   # client gift - De Adams Accessories
    "TXN-2025-0759": "COS-16",   # curved mirrors - glass
    "TXN-2025-0762": "COS-07",   # fluted panels
    "TXN-2025-0769": "REC-01",
    "TXN-2025-0774": "REC-01",   # Venture Garden Nigeria (VigiPay)
    "TXN-2025-0775": "COS-24",   # 3D visualisation - Victor Israel Baiyewu
}

FLAGS = {
    "TXN-2025-0719": "N121,900 described as 'Loan refund' to the MD. Posted to the director current account (1045) consistently with the working papers. If a loan agreement evidences a director loan, reclassify to 2053.",
    "TXN-2025-0722": "N209,505 paid through the Paystack gateway described 'Site visit' - unusually large for travel. Posted to 6511 per the working papers; obtain the Paystack settlement report and reclassify to the actual supplier.",
    "TXN-2025-0746": "N175,000 to Emmanuel Akinloluwa Akin-Ajayi, a related party (not the MD). Posted to 1047; IAS 24 disclosure required.",
    "TXN-2025-0701": "SMS alert fee line includes VAT in one amount; the bank did not split it. Posted gross to 6910 with a note - immaterial.",
    "TXN-2025-0716": "SMS alert fee 29/06-28/07 charged in September (prior period). Immaterial; expense in September.",
}
