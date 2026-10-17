"""สกัดใบรับสินค้า (PDF) -> ไฟล์ .xlsx ใบวางบิล 2 ชุด + รายการสินค้า

วิธีใช้:  python invoice_to_xlsx.py "ไฟล์.pdf" [-o ผลลัพธ์.xlsx]
- PDF ที่มี text layer: อ่านตำแหน่งคำโดยตรง (แม่นยำกว่า OCR)
- PDF ที่เป็นภาพสแกน (ไม่มี text): ใช้ Tesseract OCR (ภาษา tha+eng) เป็นทางเลือกสำรอง
"""
import argparse
import re
import sys
from collections import OrderedDict
from pathlib import Path

import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break

TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ขอบเขตคอลัมน์ของตารางรายการ (พิกัด x ของหน้า PDF)
COLS = [
    ("seq", 0, 36), ("code", 36, 78), ("desc", 78, 205), ("size", 205, 255),
    ("unit", 255, 300), ("qty", 300, 355), ("free", 355, 382), ("unit2", 382, 410),
    ("price", 410, 460), ("pct", 460, 482), ("disc", 482, 522), ("amount", 522, 900),
]
TONE_FIX = {"É": "\u0e48", "Ê": "\u0e49", "Ë": "\u0e4a", "Ì": "\u0e4b"}  # ฟอนต์ใน PDF map วรรณยุกต์เพี้ยน


def fix_thai(s: str) -> str:
    for a, b in TONE_FIX.items():
        s = s.replace(a, b)
    return s.replace("\u0e4d\u0e32", "\u0e33")  # ํา -> ำ


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except ValueError:
        return None


def page_words(page):
    words = page.get_text("words")
    if words:
        return [(w[0], w[1], fix_thai(w[4])) for w in words]
    # สำรอง: OCR
    import pytesseract
    from PIL import Image
    pytesseract.pytesseract.tesseract_cmd = TESSERACT
    zoom = 3
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    d = pytesseract.image_to_data(img, lang="tha+eng", config="--psm 6",
                                  output_type=pytesseract.Output.DICT)
    return [(d["left"][i] / zoom, d["top"][i] / zoom, d["text"][i])
            for i in range(len(d["text"])) if d["text"][i].strip()]


def value_right_of(words, label_re, x_min, tol=4):
    """หาข้อความที่อยู่ในบรรทัดเดียวกับ label และอยู่ทางขวา x_min"""
    for x, y, t in words:
        if re.search(label_re, t):
            same = sorted((w for w in words if abs(w[1] - y) <= tol and w[0] > x_min), key=lambda w: w[0])
            if same:
                return " ".join(w[2] for w in same)
    return ""


def parse_page(words):
    inv = {}
    inv["no"] = next((t for x, y, t in words if x > 480 and 78 <= y <= 84), "")
    inv["date"] = next((t for x, y, t in words if x > 480 and 96 <= y <= 102), "")
    inv["vendor_code"] = next((t for x, y, t in words if 85 <= x <= 95 and 96 <= y <= 102), "")
    inv["vendor"] = " ".join(t for x, y, t in sorted(words, key=lambda w: w[0])
                             if 85 <= x <= 300 and 114 <= y <= 120)
    inv["wh_code"] = next((t for x, y, t in words if 85 <= x <= 95 and 78 <= y <= 84), "")
    inv["branch"] = " ".join(t for x, y, t in sorted(words, key=lambda w: w[0])
                             if 150 <= x <= 440 and 78 <= y <= 84)
    inv["po"] = next((t for x, y, t in words if 85 <= x <= 200 and 133 <= y <= 139), "")
    inv["inv_ref"] = next((t for x, y, t in words if x > 480 and 114 <= y <= 120), "")

    # --- รายการสินค้า: อยู่ระหว่างหัวตาราง (y~155) กับบรรทัด "ยอดรวมก่อนหัก" ---
    end_y = next((y for x, y, t in words if "ยอดรวมก่อนหัก" in t), 600)
    body = [w for w in words if 165 < w[1] < end_y - 5]
    seqs = sorted(y for x, y, t in body if x < 36 and t.isdigit())
    items = []
    for i, sy in enumerate(seqs):
        lo = sy - 6
        hi = (seqs[i + 1] - 6) if i + 1 < len(seqs) else end_y - 5
        row = {}
        for name, a, b in COLS:
            ws = sorted((w for w in body if lo <= w[1] < hi and a <= w[0] < b), key=lambda w: (round(w[1] / 4), w[0]))
            row[name] = " ".join(w[2] for w in ws)
        items.append(row)
    inv["items"] = items

    # --- ยอดสรุป ---
    def total_at(label):
        for x, y, t in words:
            if label in t:
                c = [w for w in words if abs(w[1] - y) <= 4 and w[0] > 520 and num(w[2]) is not None]
                if c:
                    return num(c[0][2])
        return None
    inv["net"] = total_at("จำนวนเงินสุทธิ") if any("จำนวนเงินสุทธิ" in t for _, _, t in words) else None
    inv["vat"] = total_at("ภาษีมูลค่าเพิ่ม")
    inv["text"] = " ".join(t for x, y, t in sorted(words, key=lambda w: w[0]) if 100 <= x <= 360 and 708 <= y <= 714)
    inv["page"] = "".join(t for x, y, t in words if 25 <= y <= 45 and x > 520 and t.isdigit())
    return inv


def read_invoices(pdf_path):
    doc = pymupdf.open(pdf_path)
    invoices = OrderedDict()
    for pno, page in enumerate(doc, 1):
        words = page_words(page)
        if not any("ใบรับสินค้า" in w[2] for w in words):
            print(f"  ข้ามหน้า {pno}: ไม่ใช่ใบรับสินค้า")
            continue
        p = parse_page(words)
        if not p["no"]:
            print(f"  คำเตือน หน้า {pno}: อ่านเลขที่เอกสารไม่ได้")
            continue
        cur = invoices.get(p["no"])
        if cur is None:
            p["pages"] = [pno]
            invoices[p["no"]] = p
        else:  # ใบเดียวกันหลายหน้า: รวมรายการ ใช้ยอดสุทธิจากหน้าท้าย
            cur["items"] += p["items"]
            cur["pages"].append(pno)
            if p["net"] is not None:
                cur["net"], cur["text"] = p["net"], p["text"]
    return list(invoices.values())


def baht_text(amount):
    """แปลงจำนวนเงินเป็นตัวอักษรไทย เช่น 131.20 -> หนึ่งร้อยสามสิบเอ็ดบาทยี่สิบสตางค์"""
    from decimal import Decimal, ROUND_HALF_UP
    digits = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"]
    places = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"]

    def read(n):  # n < 1,000,000 หรือมากกว่า (แบ่งทีละ 6 หลัก)
        if n >= 1_000_000:
            hi, lo = divmod(n, 1_000_000)
            return read(hi) + "ล้าน" + (read(lo) if lo else "")
        s, out = str(n), ""
        for i, ch in enumerate(s):
            d, p = int(ch), len(s) - i - 1
            if d == 0:
                continue
            if p == 1:
                out += ("" if d == 1 else "ยี่" if d == 2 else digits[d]) + "สิบ"
            elif p == 0:
                out += "เอ็ด" if d == 1 and len(s) > 1 and s[-2] != "0" else \
                       "เอ็ด" if d == 1 and len(s) > 1 else digits[d]
            else:
                out += digits[d] + places[p]
        return out

    amt = Decimal(str(amount)).quantize(Decimal("0.01"), ROUND_HALF_UP)
    baht, satang = divmod(int(amt * 100), 100)
    if baht == 0 and satang == 0:
        return "ศูนย์บาทถ้วน"
    text = (read(baht) + "บาท") if baht else ""
    return text + (read(satang) + "สตางค์" if satang else "ถ้วน")


def d_key(s):  # dd/mm/yyyy -> sortable
    m = re.match(r"(\d+)/(\d+)/(\d+)", s or "")
    return (int(m[3]), int(m[2]), int(m[1])) if m else (0, 0, 0)


# ---------------------------------------------------------------- Excel
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
WARN_FILL = PatternFill("solid", fgColor="FFC7CE")
FONT = "Tahoma"


def bill_sheet(ws, invs, copy_label, branches=None):
    ws.sheet_view.showGridLines = False
    vendor = invs[0]["vendor"]
    dates = sorted((i["date"] for i in invs), key=d_key)
    ws["A1"] = "ใบวางบิล" + (f"  ({copy_label})" if copy_label else "")
    ws["A1"].font = Font(name=FONT, size=20, bold=True)
    ws.merge_cells("A1:G1")
    ws["A1"].alignment = Alignment(horizontal="center")
    info = [("วางบิลถึง", "บริษัท สุวรรณชาด จำกัด"),
            ("ช่วงวันที่รับสินค้า", f"{dates[0]} - {dates[-1]}")]
    for r, (k, v) in enumerate(info, 3):
        ws.cell(r, 1, k).font = Font(name=FONT, size=11, bold=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(r, 3, v).font = Font(name=FONT, size=11)
    hdr = ["ลำดับ", "วันที่รับ", "สาขา", "เลขที่ใบรับสินค้า", "อ้างถึงใบสั่งซื้อ", "ใบแจ้งหนี้อ้างอิง", "จำนวนเงิน (บาท)"]
    r0 = 7
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(r0, c, h)
        cell.font = Font(name=FONT, size=11, bold=True)
        cell.fill, cell.border = HEAD_FILL, BOX
        cell.alignment = Alignment(horizontal="center")
    for n, inv in enumerate(sorted(invs, key=lambda i: (d_key(i["date"]), i["no"])), 1):
        r = r0 + n
        vals = [n, inv["date"], (branches or {}).get(inv["wh_code"], (inv["branch"], ""))[0].replace("สาขา", "", 1), inv["no"], inv["po"], inv["inv_ref"], inv["net"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font, cell.border = Font(name=FONT, size=11), BOX
            cell.alignment = Alignment(horizontal="center" if c != 3 and c < 7 else ("left" if c == 3 else "right"), wrap_text=(c == 3), vertical="center")
        ws.cell(r, 7).number_format = "#,##0.00"
    last = r0 + len(invs)
    tr = last + 1
    ws.cell(tr, 6, "รวมทั้งสิ้น").font = Font(name=FONT, size=11, bold=True)
    ws.cell(tr, 6).alignment = Alignment(horizontal="right")
    t = ws.cell(tr, 7, f"=SUM(G{r0 + 1}:G{last})")
    t.font, t.border, t.number_format = Font(name=FONT, size=11, bold=True), BOX, "#,##0.00"
    ws.cell(tr + 3, 2, "ผู้วางบิล ____________________").font = Font(name=FONT, size=11)
    ws.cell(tr + 3, 5, "ผู้รับวางบิล ____________________").font = Font(name=FONT, size=11)
    for col, w in zip("ABCDEFG", [8, 14, 30, 20, 20, 22, 16]):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation, ws.page_setup.paperSize = "portrait", ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.6)


def detail_sheet(ws, invs):
    hdr = ["วันที่รับ", "เลขที่", "ลำดับ", "รหัสสินค้า", "รายการ", "ขนาดบรรจุ", "หน่วย", "ปริมาณ",
           "ราคา/หน่วย", "ส่วนลด", "จำนวนเงิน", "ตรวจยอด"]
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(1, c, h)
        cell.font, cell.fill, cell.border = Font(name=FONT, size=11, bold=True), HEAD_FILL, BOX
    r = 2
    for inv in sorted(invs, key=lambda i: (d_key(i["date"]), i["no"])):
        line_sum = round(sum(num(i["amount"]) or 0 for i in inv["items"]), 2)
        ok = inv["net"] is not None and abs(line_sum - inv["net"]) < 0.005
        for it in inv["items"]:
            vals = [inv["date"], inv["no"], it["seq"], it["code"], it["desc"], it["size"], it["unit"],
                    num(it["qty"]), num(it["price"]), num(it["disc"]), num(it["amount"]),
                    "ตรง" if ok else f"ยอดสุทธิ {inv['net']} ≠ ผลรวม {line_sum}"]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(r, c, v)
                cell.font, cell.border = Font(name=FONT, size=11), BOX
                if not ok:
                    cell.fill = WARN_FILL
            for c in (9, 10, 11):
                ws.cell(r, c).number_format = "#,##0.00"
            r += 1
    for col, w in zip("ABCDEFGHIJKL", [12, 18, 7, 12, 40, 12, 8, 10, 12, 10, 14, 26]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


BLUE = "2E96D0"
CONFIG_KEYS = ("sender_name", "sender_address", "sender_tax_id", "deliverer_name", "customer_company")


def app_dir():
    """โฟลเดอร์ที่เก็บ config.json / branches.csv (ข้าง .exe หรือข้างสคริปต์)"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BRANCH_CSV = app_dir() / "branches.csv"
CONFIG_JSON = app_dir() / "config.json"


class ConfigError(Exception):
    pass


def load_config():
    """อ่านข้อมูลผู้ส่งของจาก config.json (ไม่เก็บใน git)"""
    import json
    if not CONFIG_JSON.exists():
        raise ConfigError("ไม่พบ config.json  ให้ตั้งค่าในแท็บ 'ตั้งค่า' หรือก๊อป config.example.json เป็น config.json")
    with open(CONFIG_JSON, encoding="utf-8-sig") as f:
        cfg = json.load(f)
    missing = [k for k in CONFIG_KEYS if not str(cfg.get(k, "")).strip()]
    if missing:
        raise ConfigError("config.json ยังไม่มีค่า: " + ", ".join(missing))
    return cfg


def save_config(cfg):
    import json
    with open(CONFIG_JSON, "w", encoding="utf-8-sig") as f:
        json.dump({k: cfg.get(k, "") for k in CONFIG_KEYS}, f, ensure_ascii=False, indent=2)


SENDER_NAME = SENDER_ADDR = SENDER_TAX_ID = DELIVERER_NAME = CUSTOMER_COMPANY = ""


def apply_config(cfg):
    global SENDER_NAME, SENDER_ADDR, SENDER_TAX_ID, DELIVERER_NAME, CUSTOMER_COMPANY
    SENDER_NAME, SENDER_ADDR = cfg["sender_name"], cfg["sender_address"]
    SENDER_TAX_ID, DELIVERER_NAME = cfg["sender_tax_id"], cfg["deliverer_name"]
    CUSTOMER_COMPANY = cfg["customer_company"]


ROWS_PER_PAGE = 16   # จำนวนรายการสินค้าต่อหน้า A4
BLOCK = 34           # จำนวนแถวต่อ 1 หน้า


def read_branches():
    """อ่าน branches.csv -> {รหัสคลัง: (ชื่อสาขา, ที่อยู่)}"""
    import csv
    data = {}
    if BRANCH_CSV.exists():
        with open(BRANCH_CSV, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                data[row["รหัสคลัง"].strip()] = (row["ชื่อสาขา"].strip(), row["ที่อยู่"].strip())
    return data


def save_branches(data):
    import csv
    with open(BRANCH_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["รหัสคลัง", "ชื่อสาขา", "ที่อยู่"])
        for k, (n, a) in data.items():
            w.writerow([k, n, a])


def load_branches(invs):
    """อ่าน branches.csv แล้วเติมรหัสคลังใหม่ที่พบใน PDF (ใช้ชื่อจาก PDF, ที่อยู่ว่าง) พร้อมบันทึก"""
    data = read_branches()
    changed = False
    for i in invs:
        if i["wh_code"] not in data:
            data[i["wh_code"]] = (i["branch"], "")
            changed = True
    if changed or not BRANCH_CSV.exists():
        save_branches(data)
    return data


def missing_addresses(invs, branches):
    """รายการ (รหัส, ชื่อ) ของสาขาที่พบใน PDF แต่ยังไม่มีที่อยู่"""
    seen, out = set(), []
    for i in invs:
        code = i["wh_code"]
        name, addr = branches.get(code, (i["branch"], ""))
        if code not in seen and not addr.strip():
            out.append((code, name))
        seen.add(code)
    return out


def mismatched(invs):
    """เลขที่ใบที่ยอดสุทธิไม่ตรงกับผลรวมรายการ"""
    return [i["no"] for i in invs if i["net"] is None or
            abs(sum(num(x["amount"]) or 0 for x in i["items"]) - i["net"]) > 0.005]


def delivery_sheet(ws, invs, branches, copy_label):
    """ใบส่งของ A4 ใบละ 1 หน้า (ตามแบบ Flowaccount) ไม่มีที่อยู่/ชื่อร้านผู้ขาย"""
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFG", [6, 13, 40, 10, 10, 14, 16]):
        ws.column_dimensions[col].width = w
    blue = Font(name=FONT, size=10, color=BLUE)
    black = Font(name=FONT, size=11)
    bold = Font(name=FONT, size=11, bold=True)
    r = 1
    for inv in sorted(invs, key=lambda i: (d_key(i["date"]), i["no"])):
        name, addr = branches.get(inv["wh_code"], (inv["branch"], ""))
        items = inv["items"] or [{}]
        chunks = [items[k:k + ROWS_PER_PAGE] for k in range(0, len(items), ROWS_PER_PAGE)]
        for pg, chunk in enumerate(chunks, 1):
            for rr in range(r, r + BLOCK):
                ws.row_dimensions[rr].height = 20
            ws.row_dimensions[r].height = 40      # แถวชื่อเรื่อง (ฟอนต์ 20pt) กันสระ ใ ถูกตัด
            ws.row_dimensions[r + 1].height = 22
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
            c = ws.cell(r, 1, "ใบส่งของ")
            c.font, c.alignment = Font(name=FONT, size=20, bold=True, color=BLUE), Alignment(horizontal="center")
            ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=7)
            c = ws.cell(r + 1, 1, copy_label + (f"  (หน้า {pg}/{len(chunks)})" if len(chunks) > 1 else ""))
            c.font, c.alignment = blue, Alignment(horizontal="center")
            # ซ้าย: ชื่อสาขา + ที่อยู่ (ไม่มีคำว่า "ลูกค้า")
            for k, text in enumerate([f"ชื่อผู้ส่งของ: {SENDER_NAME}", SENDER_ADDR, f"เลขประจำตัวผู้เสียภาษีอากร: {SENDER_TAX_ID}"]):
                ws.merge_cells(start_row=r + 3 + k, start_column=1, end_row=r + 3 + k, end_column=5)
                ws.cell(r + 3 + k, 1, text).font = black
            ws.merge_cells(start_row=r + 6, start_column=1, end_row=r + 6, end_column=5)
            ws.cell(r + 6, 1, CUSTOMER_COMPANY).font = bold
            ws.merge_cells(start_row=r + 7, start_column=1, end_row=r + 7, end_column=5)
            ws.cell(r + 7, 1, name).font = bold
            ws.merge_cells(start_row=r + 8, start_column=1, end_row=r + 9, end_column=5)
            c = ws.cell(r + 8, 1, addr)
            c.font, c.alignment = black, Alignment(wrap_text=True, vertical="top")
            # ขวา: เลขที่/วันที่/อ้างอิง
            for k, (lab, val) in enumerate([("เลขที่", inv["no"]), ("วันที่", inv["date"]),
                                            ("อ้างอิง", inv["po"]), ("ใบแจ้งหนี้", inv["inv_ref"])]):
                ws.cell(r + 3 + k, 6, lab).font = blue
                v = ws.cell(r + 3 + k, 7, val)
                v.font, v.alignment = black, Alignment(horizontal="left")
                for cc in (6, 7):
                    ws.cell(r + 3 + k, cc).border = Border(bottom=Side(style="thin", color="BBBBBB"))
            # หัวตาราง
            h = r + 10
            for col, t in enumerate(["#", "รหัสสินค้า", "รายละเอียด", "จำนวน", "หน่วย", "ราคา/หน่วย", "จำนวนเงิน"], 1):
                c = ws.cell(h, col, t)
                c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor=BLUE)
                c.alignment = Alignment(horizontal="center" if col != 3 else "left")
            for k, it in enumerate(chunk):
                row = h + 1 + k
                n = (pg - 1) * ROWS_PER_PAGE + k + 1
                vals = [n if it else None, it.get("code"), it.get("desc"), num(it.get("qty")), it.get("unit"), num(it.get("price")), num(it.get("amount"))]
                for col, v in enumerate(vals, 1):
                    c = ws.cell(row, col, v)
                    c.font = black
                    c.alignment = Alignment(horizontal="center" if col in (1, 2, 4, 5) else "right" if col in (6, 7) else "left")
                    if col in (4, 6, 7):
                        c.number_format = "#,##0.00" if col in (6, 7) else "General"
                    c.border = Border(bottom=Side(style="hair", color="BBBBBB"))
            if pg == len(chunks):
                tr = h + 1 + ROWS_PER_PAGE
                ws.cell(tr, 3, "จำนวน").font = blue
                ws.cell(tr, 3).alignment = Alignment(horizontal="right")
                ws.cell(tr, 4, len(items)).font = black
                ws.cell(tr, 4).alignment = Alignment(horizontal="center")
                ws.cell(tr, 5, "รายการ").font = black
                total = inv["net"] if inv["net"] is not None else sum(num(i.get("amount")) or 0 for i in items)
                c = ws.cell(tr, 7, total)
                c.font, c.number_format = bold, "#,##0.00"
                c.border = Border(top=Side(style="thin"), bottom=Side(style="double"))
                ws.merge_cells(start_row=tr + 1, start_column=1, end_row=tr + 1, end_column=7)
                c = ws.cell(tr + 1, 1, f"({baht_text(total)})")
                c.font, c.alignment = bold, Alignment(horizontal="center")
            # ลายเซ็น
            s = r + BLOCK - 5
            ws.merge_cells(start_row=s + 3, start_column=1, end_row=s + 3, end_column=3)
            ws.cell(s + 3, 1, "ผู้รับสินค้า ______________________   วันที่ ____________").font = black
            ws.merge_cells(start_row=s + 3, start_column=4, end_row=s + 3, end_column=7)
            ws.cell(s + 3, 4, f"ผู้ส่งสินค้า  {DELIVERER_NAME}      วันที่ ________").font = black
            ws.row_breaks.append(Break(id=r + BLOCK - 1))
            r += BLOCK
    ws.print_area = f"A1:G{r - 1}"
    ws.page_setup.orientation, ws.page_setup.paperSize = "portrait", ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)


def build(invs, out):
    wb = Workbook()
    branches = load_branches(invs)
    # แยกตามผู้ขาย: ปกติไฟล์หนึ่งเป็นร้านเดียว ถ้ามีหลายร้านให้ทำชีตแยกร้าน
    by_vendor = OrderedDict()
    for i in invs:
        by_vendor.setdefault(i["vendor_code"], []).append(i)
    first = True
    for code, group in by_vendor.items():
        for label in ("ต้นฉบับ", "สำเนา"):
            ws = wb.active if first else wb.create_sheet()
            first = False
            ws.title = f"ใบวางบิล-{code}-{label}"[:31]
            bill_sheet(ws, group, label, branches)
    for label in ("ต้นฉบับ", "สำเนา"):
        delivery_sheet(wb.create_sheet(f"ใบส่งของ-{label}"), invs, branches, label)
    detail_sheet(wb.create_sheet("รายการสินค้า"), invs)
    wb.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    try:
        apply_config(load_config())
    except ConfigError as e:
        sys.exit(str(e))
    out = a.out or str(Path(a.pdf).with_suffix(".xlsx"))
    invs = read_invoices(a.pdf)
    if not invs:
        sys.exit("ไม่พบใบรับสินค้าในไฟล์")
    build(invs, out)
    bad = mismatched(invs)
    print(f"อ่านได้ {len(invs)} ใบ  ยอดรวม {sum(i['net'] or 0 for i in invs):,.2f} บาท -> {out}")
    print("ยอดไม่ตรง (ตรวจด้วยตา): " + (", ".join(bad) if bad else "ไม่มี"))
    for code, name in missing_addresses(invs, read_branches()):
        print(f"คำเตือน: สาขา {code} {name} ยังไม่มีที่อยู่ใน branches.csv")


if __name__ == "__main__":
    main()
