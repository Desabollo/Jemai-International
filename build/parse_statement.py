"""Parses the Access Bank 'August to October 2026 Statement.pdf' into a normalised
transaction register.

The statement lays each transaction out as a table row (posted date, value date,
description, debit, credit, balance) but the description cell wraps over up to
three visual lines that are vertically centred on the money row. This parser
uses word coordinates, not text order, so wrapped descriptions are re-attached
to the correct transaction.
"""
import json
import os
import re
from collections import defaultdict

import pdfplumber

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PDF = os.path.join(ROOT, "August to October 2026 Statement.pdf")

DATE_RE = re.compile(r"^\d{2}-[A-Z]{3}-\d{2}$")
MONEY_RE = re.compile(r"^-?[\d,]+\.\d{2}$")

# x-coordinate bands (from the header row: Description 181, Debit 363, Credit 424, Balance 486)
B_POSTED = (40, 115)
B_VALUE = (115, 175)
B_DESC = (175, 358)
B_DEBIT = (358, 420)
B_CREDIT = (420, 480)
B_BAL = (480, 560)


def band(x, b):
    return b[0] <= x < b[1]


def money(t):
    """Money with its sign kept. A *negative* amount in the DEBIT column is how
    Access Bank shows a reversal of a failed transfer: it adds the money back."""
    t = t.replace(",", "").replace("\u2212", "-").strip()
    if not t:
        return None
    try:
        return float(t)
    except ValueError:
        return None


def rows_of(page):
    words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
    buckets = defaultdict(list)
    for w in words:
        buckets[round(w["top"] / 3)].append(w)   # 3pt vertical tolerance
    out = []
    for key in sorted(buckets):
        ws = sorted(buckets[key], key=lambda w: w["x0"])
        out.append({"y": key * 3, "words": ws})
    return out


def parse():
    txn_rows = []
    with pdfplumber.open(PDF) as pdf:
        for pno, page in enumerate(pdf.pages):
            rows = rows_of(page)
            # locate the column header; anything above it is page furniture
            hdr_idx = None
            for i, r in enumerate(rows):
                txt = " ".join(w["text"] for w in r["words"])
                if "Posted Date" in txt and "Description" in txt:
                    hdr_idx = i
                    break
            if hdr_idx is not None:
                rows = rows[hdr_idx + 1:]

            # split each row into columns
            for r in rows:
                rec = {"page": pno + 1, "y": r["y"], "posted": "", "value": "",
                       "desc": [], "debit": None, "credit": None, "balance": None,
                       "raw": " ".join(w["text"] for w in r["words"])}
                for w in r["words"]:
                    x, t = w["x0"], w["text"]
                    if band(x, B_POSTED):
                        if DATE_RE.match(t):
                            rec["posted"] = t
                        elif t not in ("-",):
                            rec["desc"].append(t)
                    elif band(x, B_VALUE):
                        if DATE_RE.match(t):
                            rec["value"] = t
                        elif t not in ("-",):
                            rec["desc"].append(t)
                    elif band(x, B_DESC):
                        rec["desc"].append(t)
                    elif band(x, B_DEBIT):
                        if MONEY_RE.match(t):
                            rec["debit"] = money(t)
                            rec["reversal"] = t.lstrip().startswith("-")
                        elif t == "-":
                            pass
                        else:
                            rec["desc"].append(t)
                    elif band(x, B_CREDIT):
                        if MONEY_RE.match(t):
                            rec["credit"] = money(t)
                            if t.lstrip().startswith("-"):
                                rec["reversal"] = True
                        elif t == "-":
                            pass
                        else:
                            rec["desc"].append(t)
                    elif band(x, B_BAL):
                        if MONEY_RE.match(t):
                            rec["balance"] = money(t)
                        else:
                            rec["desc"].append(t)
                rec["is_txn"] = bool(rec["posted"])
                txn_rows.append(rec)

    # ---- attach description fragments to the nearest transaction row -------
    # Two passes: assign first, then join, so that fragments printed *below* the
    # money row are not lost (a single pass computes the text too early).
    for i, r in enumerate(txn_rows):
        r["idx"] = i
    anchors = [r for r in txn_rows if r["is_txn"]]
    for a in anchors:
        a["gy"] = a["page"] * 10000 + a["y"]
        a["frags"] = [(a["gy"], " ".join(a["desc"]))] if a["desc"] else []
    for r in txn_rows:
        if r["is_txn"] or not r["desc"]:
            continue
        gy = r["page"] * 10000 + r["y"]
        # nearest money line in *points*; on a tie prefer the line below the
        # fragment (Access prints the first description line just above the
        # amount and the remainder just below it).
        best = min(anchors, key=lambda a: (abs(a["gy"] - gy), 0 if a["gy"] >= gy else 1)) if anchors else None
        if best is not None:
            best["frags"].append((gy, " ".join(r["desc"])))

    out = []
    for a in anchors:
        desc = " ".join(f[1] for f in sorted(a.get("frags", [])))
        desc = re.sub(r"\s+", " ", desc).strip()
        if not desc and not (a["debit"] or a["credit"]):
            continue
        out.append({
            "page": a["page"],
            "posted": a["posted"],
            "value": a["value"],
            "description": desc,
            "debit": a["debit"] or 0.0,
            "credit": a["credit"] or 0.0,
            "balance": a["balance"],
            "reversal": a.get("reversal", False),
        })
    return out


if __name__ == "__main__":
    txns = parse()
    # drop the opening balance row
    opening = None
    clean = []
    for t in txns:
        if "opening balance" in t["description"].lower():
            opening = t["balance"]
            continue
        if abs(t["debit"]) < 0.005 and abs(t["credit"]) < 0.005:
            continue
        clean.append(t)
    gross_dr = round(sum(t["debit"] for t in clean if not t["reversal"]), 2)
    gross_cr = round(sum(t["credit"] for t in clean if not t["reversal"]), 2)
    rev = round(-sum(t["debit"] + t["credit"] for t in clean if t["reversal"]), 2)
    net_dr = round(gross_dr - rev, 2)
    print("rows:", len(clean), " reversals:", sum(1 for t in clean if t["reversal"]))
    print("opening:", opening)
    print("gross debits   :", f"{gross_dr:,.2f}", " statement: 74,373,993.97")
    print("gross credits  :", f"{gross_cr:,.2f}", " statement: 51,225,487.00")
    print("reversal pairs :", f"{rev:,.2f}")
    print("net debits     :", f"{net_dr:,.2f}")
    print("closing computed:", f"{(opening or 0) + gross_cr - net_dr:,.2f}", " statement: 1,805,567.88")

    # running-balance reconciliation: where does the balance column stop footing?
    prev, drift, worst = opening, 0.0, []
    for i, t in enumerate(clean):
        if t["balance"] is None:
            continue
        exp = prev - t["debit"] + t["credit"]
        d = round(t["balance"] - exp, 2)
        if abs(d) > 0.005:
            worst.append((i, t["posted"], round(t["debit"], 2), round(t["credit"], 2),
                          t["balance"], round(exp, 2), d, t["description"][:70]))
        prev = t["balance"]
    print("balance mismatches:", len(worst), "/", len(clean))
    for w in worst[:12]:
        print("  ", w)
    json.dump({"opening": opening, "txns": clean}, open("/home/user/work/stmt2026.json", "w"), indent=1)
    print("saved")
