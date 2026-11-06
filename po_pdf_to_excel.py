#!/usr/bin/env python3
"""
po_pdf_to_excel.py  -  Convert a Shopify-style Purchase Order PDF into Excel.

Usage:
    python po_pdf_to_excel.py  input.pdf  [output.xlsx]
    python po_pdf_to_excel.py  folder_with_pdfs/           (batch: every *.pdf -> .xlsx)

Requirements:
    pip install pdfplumber openpyxl

What it does:
  * Reads the PO header (PO #, date, supplier, ship-to, bill-to, terms, currency, ETA)
  * Reads every product line (title, size/variant, SKU, qty, cost, tax, total),
    including lines whose text wraps across a page break
  * Reads the summary block (reference no., notes, taxes, subtotal, total)
  * Writes an Excel file laid out like the PDF, with live formulas for line totals,
    subtotal and total, and a check against the totals printed on the PDF
"""
import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SIZES = {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "2XL", "3XL", "OS", "FREE SIZE"}
MONEY = re.compile(r"^\$?-?[\d,]+\.\d{2}$")


# --------------------------------------------------------------------------- #
# PDF reading helpers
# --------------------------------------------------------------------------- #
def to_float(s):
    return float(s.replace("$", "").replace(",", ""))


def group_lines(words, tol=2.5):
    """Group pdfplumber words into visual lines (sorted top->bottom, left->right)."""
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(lines[-1]["top"] - w["top"]) <= tol:
            lines[-1]["words"].append(w)
        else:
            lines.append({"top": w["top"], "words": [w]})
    for ln in lines:
        ln["words"].sort(key=lambda w: w["x0"])
        ln["text"] = " ".join(w["text"] for w in ln["words"])
    return lines


def join_wrapped(parts):
    """Join wrapped title lines. A line ending in '-' directly after a character
    (e.g. 'ST2STE-') continues the same word; ' -' (space-dash) is a separator."""
    out = ""
    for p in parts:
        if not out:
            out = p
        elif out.endswith("-") and not out.endswith(" -"):
            out += p
        else:
            out += " " + p
    return out.strip()


def find_columns(header_line):
    """Return x-start of each table column from the header line."""
    pos = {}
    for w in header_line["words"]:
        t = w["text"].upper()
        if t == "PRODUCTS":
            pos["products"] = w["x0"]
        elif t == "SUPPLIER":
            pos["sku"] = w["x0"]
        elif t == "QTY":
            pos["qty"] = w["x0"]
        elif t == "COST":
            pos["cost"] = w["x0"]
        elif t == "TAX":
            pos["tax"] = w["x0"]
        elif t == "TOTAL":
            pos["total"] = w["x0"]
    return pos


# --------------------------------------------------------------------------- #
# Parser
# --------------------------------------------------------------------------- #
def parse_po(pdf_path):
    po = {"items": [], "header": {}, "summary": {}}
    with pdfplumber.open(pdf_path) as pdf:
        page_lines = [group_lines(p.extract_words(keep_blank_chars=False)) for p in pdf.pages]
        page_heights = [p.height for p in pdf.pages]

    # ---------------- header (page 1) ----------------
    p1 = page_lines[0]
    texts = [l["text"] for l in p1]
    for l in p1:
        m = re.search(r"#\s*PO\S+", l["text"])
        if m and "po_number" not in po["header"] and l["top"] > 40:
            po["header"]["po_number"] = m.group(0).replace(" ", "")
        if re.search(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", l["words"][-3:] and " ".join(w["text"] for w in l["words"][-3:])):
            po["header"]["date"] = " ".join(w["text"] for w in l["words"][-3:])

    # three-column address block
    i_sup = next(i for i, t in enumerate(texts) if t.startswith("SUPPLIER") and "SHIP" in t)
    i_pay = next(i for i, t in enumerate(texts) if t.startswith("PAYMENT TERMS"))
    i_hdr = next(i for i, t in enumerate(texts) if t.startswith("PRODUCTS"))
    hdr_words = p1[i_sup]["words"]
    x_ship = next(w["x0"] for w in hdr_words if w["text"] == "SHIP") - 3
    x_bill = next(w["x0"] for w in hdr_words if w["text"] == "BILL") - 3

    def split3(line):
        cols = ["", "", ""]
        for w in line["words"]:
            k = 0 if w["x0"] < x_ship else (1 if w["x0"] < x_bill else 2)
            cols[k] = (cols[k] + " " + w["text"]).strip()
        return cols

    sup, ship, bill = [], [], []
    for l in p1[i_sup + 1:i_pay]:
        c = split3(l)
        sup.append(c[0]), ship.append(c[1]), bill.append(c[2])
    po["header"]["supplier"] = [x for x in sup if x]
    po["header"]["ship_to"] = [x for x in ship if x]
    po["header"]["bill_to"] = [x for x in bill if x]
    # labels row + values row
    labels = split3(p1[i_pay])
    values = split3(p1[i_pay + 1])
    po["header"]["payment_terms"], po["header"]["currency"], po["header"]["eta"] = values

    # ---------------- table rows ----------------
    cols = find_columns(p1[i_hdr])
    x_sku, x_qty, x_cost, x_tax, x_tot = (cols[k] for k in ("sku", "qty", "cost", "tax", "total"))
    x_prod = cols["products"]

    current = None
    stop = False
    for pi, lines in enumerate(page_lines):
        h = page_heights[pi]
        for l in lines:
            if stop:
                break
            if pi == 0 and l["top"] <= p1[i_hdr]["top"] + 1:
                continue                      # header area of page 1
            if l["top"] < 45 or l["top"] > h - 70:
                continue                      # running page header / footer
            if l["text"].startswith("REFERENCE NUMBER"):
                stop = True
                break

            left = " ".join(w["text"] for w in l["words"] if w["x0"] < x_sku - 3)
            sku_w = [w for w in l["words"] if x_sku - 3 <= w["x0"] < x_qty - 3]
            qty_w = [w for w in l["words"] if x_qty - 3 <= w["x0"] < x_cost - 3]
            cost_w = [w for w in l["words"] if x_cost - 3 <= w["x0"] < x_tax - 3]
            tax_w = [w for w in l["words"] if x_tax - 3 <= w["x0"] < x_tot - 3]
            tot_w = [w for w in l["words"] if w["x0"] >= x_tot - 3]

            is_anchor = (sku_w and qty_w and cost_w and tot_w
                         and sku_w[0]["text"].isdigit() and qty_w[0]["text"].isdigit()
                         and MONEY.match(cost_w[0]["text"]) and MONEY.match(tot_w[0]["text"]))
            if is_anchor:
                current = {
                    "left": [left] if left else [],
                    "supplier_sku": sku_w[0]["text"],
                    "qty": int(qty_w[0]["text"]),
                    "cost": to_float(cost_w[0]["text"]),
                    "tax": tax_w[0]["text"] if tax_w else "0%",
                    "total": to_float(tot_w[0]["text"]),
                }
                po["items"].append(current)
            elif current is not None and left:
                current["left"].append(left)

    # post-process each item's left-column text -> title / size / product sku
    for it in po["items"]:
        parts = list(it.pop("left"))
        sku = it["supplier_sku"]
        if parts and re.fullmatch(r"\d{5,}", parts[-1]):
            sku = parts.pop()
        size = ""
        if parts and parts[-1].strip().upper() in SIZES:
            size = parts.pop().strip()
        it["title"] = join_wrapped(parts)
        it["size"] = size
        it["sku"] = sku

    # ---------------- summary (last page) ----------------
    last = page_lines[-1]
    s = po["summary"]
    for idx, l in enumerate(last):
        t = l["text"]
        if t.startswith("REFERENCE NUMBER"):
            nxt = last[idx + 1]["text"] if idx + 1 < len(last) else ""
            s["reference"] = " ".join(w["text"] for w in last[idx + 1]["words"] if w["x0"] < 290)
        if t.startswith("NOTES TO SUPPLIER"):
            # anything on following lines left of the cost summary (until blank gap) = notes
            notes = []
            for l2 in last[idx + 1:]:
                lt = " ".join(w["text"] for w in l2["words"] if w["x0"] < 290)
                if lt and not lt.startswith("NISHAT") and l2["top"] < l["top"] + 60:
                    notes.append(lt)
            s["notes"] = " ".join(notes)
        m = re.search(r"Taxes.*?(\$[\d,]+\.\d{2})", t)
        if m:
            s["taxes"] = to_float(m.group(1))
        m = re.search(r"Subtotal \((\d+) items?\).*?(\$[\d,]+\.\d{2})", t)
        if m:
            s["item_count"] = int(m.group(1))
            s["subtotal"] = to_float(m.group(2))
        m = re.match(r"^.*\bTotal\b.*?(\$[\d,]+\.\d{2})$", t)
        if m and "Subtotal" not in t and "TOTAL" not in t:
            s["total"] = to_float(m.group(1))
    # footer (company contact)
    foot = []
    grab = False
    for l in last:
        if l["text"].startswith("NISHAT USA") and l["top"] > 200:
            grab = True
        if grab and not l["text"].startswith("Powered by"):
            foot.append(l["text"])
    s["footer"] = foot
    return po


# --------------------------------------------------------------------------- #
# Excel writer
# --------------------------------------------------------------------------- #
def write_excel(po, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = (po["header"].get("po_number", "PO").lstrip("#") or "PO")[:31]

    F = "Arial"
    f_norm = Font(name=F, size=10)
    f_bold = Font(name=F, size=10, bold=True)
    f_title = Font(name=F, size=18, bold=True)
    f_small = Font(name=F, size=9, color="666666")
    f_note = Font(name=F, size=9, italic=True, color="666666")
    line = Side(style="thin", color="000000")
    light = Side(style="thin", color="D9D9D9")
    fill_head = PatternFill("solid", fgColor="F2F2F2")

    widths = {"A": 52, "B": 9, "C": 14, "D": 14, "E": 8, "F": 12, "G": 8, "H": 15}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

    def put(cell, value, font=f_norm, align=None, fmt=None):
        c = ws[cell]
        c.value = value
        c.font = font
        if align:
            c.alignment = align
        if fmt:
            c.number_format = fmt
        return c

    right = Alignment(horizontal="right", vertical="top")
    left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
    center_top = Alignment(horizontal="center", vertical="top")

    h = po["header"]
    # running header
    put("A1", "NISHAT USA", f_small)
    put("H1", h.get("po_number", ""), f_small, right)
    # title
    put("A3", "NISHAT USA", f_title)
    ws.row_dimensions[3].height = 26
    put("H3", h.get("po_number", ""), f_bold, right)
    put("H4", h.get("date", ""), f_norm, right)

    # address blocks
    put("A6", "SUPPLIER", f_bold)
    put("C6", "SHIP TO", f_bold)
    put("F6", "BILL TO", f_bold)
    n = max(len(h["supplier"]), len(h["ship_to"]), len(h["bill_to"]))
    for i in range(n):
        r = 7 + i
        if i < len(h["supplier"]):
            put(f"A{r}", h["supplier"][i])
        if i < len(h["ship_to"]):
            put(f"C{r}", h["ship_to"][i])
        if i < len(h["bill_to"]):
            put(f"F{r}", h["bill_to"][i])
    r = 7 + n + 1
    put(f"A{r}", "PAYMENT TERMS", f_bold)
    put(f"C{r}", "SUPPLIER CURRENCY", f_bold)
    put(f"F{r}", "ESTIMATED ARRIVAL", f_bold)
    put(f"A{r+1}", h.get("payment_terms", ""))
    put(f"C{r+1}", h.get("currency", ""))
    put(f"F{r+1}", h.get("eta", ""))

    # table header
    hr = r + 3
    heads = ["PRODUCTS", "SIZE", "SKU", "SUPPLIER SKU", "QTY", "COST", "TAX", "TOTAL (USD)"]
    for i, t in enumerate(heads):
        c = ws.cell(row=hr, column=i + 1, value=t)
        c.font = f_bold
        c.fill = fill_head
        c.border = Border(top=line, bottom=line)
        c.alignment = Alignment(horizontal="right" if i >= 4 else "left", vertical="center")
    ws.row_dimensions[hr].height = 20

    # product rows
    first = hr + 1
    row = first
    for it in po["items"]:
        ws.cell(row=row, column=1, value=it["title"]).font = f_bold
        ws.cell(row=row, column=2, value=it["size"]).font = f_norm
        c = ws.cell(row=row, column=3, value=int(it["sku"]) if it["sku"].isdigit() else it["sku"])
        c.font = f_norm
        c.number_format = "0"
        c.alignment = Alignment(horizontal="left")
        c = ws.cell(row=row, column=4, value=int(it["supplier_sku"]))
        c.font = f_norm
        c.number_format = "0"
        c.alignment = Alignment(horizontal="left")
        ws.cell(row=row, column=5, value=it["qty"]).font = f_norm
        c = ws.cell(row=row, column=6, value=it["cost"])
        c.font = f_norm
        c.number_format = "$#,##0.00"
        c = ws.cell(row=row, column=7, value=float(it["tax"].strip("%")) / 100)
        c.font = f_norm
        c.number_format = "0%"
        c = ws.cell(row=row, column=8, value=f"=ROUND(E{row}*F{row},2)")
        c.font = f_norm
        c.number_format = "$#,##0.00"
        for col in range(1, 9):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(bottom=light)
            if col == 1:
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            elif col >= 5:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        row += 1
    last = row - 1

    # summary
    s = po["summary"]
    r0 = last + 3
    put(f"A{r0}", "REFERENCE NUMBER", f_bold)
    put(f"A{r0+1}", s.get("reference", ""))
    put(f"A{r0+3}", "NOTES TO SUPPLIER", f_bold)
    put(f"A{r0+4}", s.get("notes", ""))
    put(f"F{r0}", "COST SUMMARY", f_bold)
    put(f"F{r0+1}", "Taxes (included)")
    put(f"H{r0+1}", s.get("taxes", 0.0), f_norm, right, "$#,##0.00")
    put(f"F{r0+2}", f'="Subtotal ("&COUNTA(D{first}:D{last})&" items)"')
    put(f"H{r0+2}", f"=SUM(H{first}:H{last})", f_norm, right, "$#,##0.00")
    put(f"F{r0+4}", "Total", f_bold)
    c = put(f"H{r0+4}", f"=H{r0+2}", f_bold, right, "$#,##0.00")
    c.border = Border(top=line)

    # footer block
    fr = r0 + 7
    for i, t in enumerate(s.get("footer", [])):
        put(f"A{fr+i}", t, f_bold if i == 0 else f_norm)
    put(f"A{fr+len(s.get('footer', []))+1}", "Powered by Shopify", f_small)

    # check vs PDF (outside the print layout, clearly labelled)
    cr = fr + len(s.get("footer", [])) + 4
    put(f"A{cr}", "CHECK AGAINST ORIGINAL PDF (values copied from the PDF's printed totals)", f_note)
    put(f"A{cr+1}", "Subtotal printed on PDF", f_note)
    put(f"H{cr+1}", s.get("subtotal"), f_note, right, "$#,##0.00")
    put(f"A{cr+2}", "Items printed on PDF", f_note)
    put(f"H{cr+2}", s.get("item_count"), f_note, right, "0")
    put(f"A{cr+3}", "Difference vs Excel subtotal (should be 0.00)", f_note)
    put(f"H{cr+3}", f"=ROUND(H{r0+2}-H{cr+1},2)", f_note, right, "$#,##0.00;-$#,##0.00;0.00")

    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:H{last}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hr}:{hr}"
    wb.save(out_path)
    return first, last


# --------------------------------------------------------------------------- #
# Validation
# --------------------------------------------------------------------------- #
def validate(po):
    problems = []
    items = po["items"]
    s = po["summary"]
    for it in items:
        calc = round(it["qty"] * it["cost"], 2)
        if abs(calc - it["total"]) > 0.011:
            problems.append(f"Line {it['supplier_sku']}: qty*cost={calc} but PDF total={it['total']}")
        if not it["title"]:
            problems.append(f"Line {it['supplier_sku']}: empty title")
    tot = round(sum(i["total"] for i in items), 2)
    if "subtotal" in s and abs(tot - s["subtotal"]) > 0.011:
        problems.append(f"Sum of line totals {tot:,.2f} != PDF subtotal {s['subtotal']:,.2f}")
    if "item_count" in s and len(items) != s["item_count"]:
        problems.append(f"Parsed {len(items)} lines but PDF says {s['item_count']} items")
    return problems, tot


def convert(pdf_path, out_path=None):
    pdf_path = Path(pdf_path)
    out_path = Path(out_path) if out_path else pdf_path.with_suffix(".xlsx")
    po = parse_po(pdf_path)
    problems, tot = validate(po)
    write_excel(po, out_path)
    print(f"{pdf_path.name}: {len(po['items'])} lines, sum of totals ${tot:,.2f} -> {out_path}")
    if problems:
        print("  WARNINGS:")
        for p in problems:
            print("   -", p)
    else:
        print("  All checks passed (line math, item count and subtotal match the PDF).")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = Path(sys.argv[1])
    if src.is_dir():
        for f in sorted(src.glob("*.pdf")):
            convert(f)
    else:
        convert(src, sys.argv[2] if len(sys.argv) > 2 else None)
