"""ฐานข้อมูล SQLite เก็บใบรับสินค้าที่สกัดแล้ว: กันซ้ำด้วยเลขที่ใบ, ดึงกลับไปสร้างเอกสาร, สรุป dashboard

- ไฟล์ .db อยู่ใน %LOCALAPPDATA%\\InvoiceToXlsx (ไม่อยู่ใน OneDrive เพื่อกันไฟล์เสียจากการซิงก์ขณะเปิด)
- สำรองไฟล์ไปที่โฟลเดอร์ backup/ ข้างโปรแกรมหลังบันทึกแต่ละครั้ง (ใช้ sqlite backup API ได้ไฟล์ที่สมบูรณ์เสมอ)
- ยอดเงินเก็บเป็นสตางค์ (จำนวนเต็ม) วันที่เก็บเป็น ISO ค.ศ. (YYYY-MM-DD)
"""
import os
import re
import sqlite3
from datetime import datetime
from pathlib import Path

import invoice_to_xlsx as core

KEEP_BACKUPS = 15

SCHEMA = """
CREATE TABLE IF NOT EXISTS imports(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_file TEXT, imported_at TEXT, n_new INTEGER, n_dup INTEGER);
CREATE TABLE IF NOT EXISTS invoices(
    no TEXT PRIMARY KEY,
    date_iso TEXT NOT NULL,
    vendor_code TEXT, wh_code TEXT, branch_pdf TEXT, po TEXT, inv_ref TEXT,
    net_satang INTEGER,
    source_file TEXT, import_id INTEGER REFERENCES imports(id));
CREATE TABLE IF NOT EXISTS invoice_items(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_no TEXT NOT NULL REFERENCES invoices(no) ON DELETE CASCADE,
    seq INTEGER, code TEXT, descr TEXT, qty REAL, unit TEXT,
    price_satang INTEGER, amount_satang INTEGER);
CREATE INDEX IF NOT EXISTS ix_inv_date ON invoices(date_iso);
CREATE INDEX IF NOT EXISTS ix_inv_branch ON invoices(wh_code);
CREATE INDEX IF NOT EXISTS ix_item_inv ON invoice_items(invoice_no);
"""


def db_path():
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "InvoiceToXlsx"
    base.mkdir(parents=True, exist_ok=True)
    return base / "invoices.db"


def connect(path=None):
    conn = sqlite3.connect(str(path or db_path()))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA)
    return conn


# ---------------------------------------------------------------- helpers
def be_to_iso(d):
    """'30/03/2569' (พ.ศ.) -> '2026-03-30'"""
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", d or "")
    if not m:
        raise ValueError(f"วันที่อ่านไม่ได้: {d!r}")
    day, mon, yr = int(m[1]), int(m[2]), int(m[3])
    if yr > 2400:
        yr -= 543
    return f"{yr:04d}-{mon:02d}-{day:02d}"


def iso_to_be(iso):
    y, m, d = iso.split("-")
    return f"{int(d):02d}/{int(m):02d}/{int(y) + 543}"


def satang(v):
    n = core.num(v)
    return None if n is None else int(round(n * 100))


def money(sat):
    return None if sat is None else sat / 100


# ---------------------------------------------------------------- import
def import_invoices(conn, invs, source_file=""):
    """บันทึกใบที่ยังไม่มีในฐานข้อมูล คืน (จำนวนใหม่, จำนวนซ้ำที่ข้าม)"""
    new = dup = 0
    now = datetime.now().isoformat(timespec="seconds")
    with conn:
        cur = conn.execute("INSERT INTO imports(source_file, imported_at, n_new, n_dup) VALUES (?,?,0,0)",
                           (Path(source_file).name, now))
        imp_id = cur.lastrowid
        for inv in invs:
            if conn.execute("SELECT 1 FROM invoices WHERE no=?", (inv["no"],)).fetchone():
                dup += 1
                continue
            conn.execute(
                "INSERT INTO invoices(no,date_iso,vendor_code,wh_code,branch_pdf,po,inv_ref,net_satang,"
                "source_file,import_id) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (inv["no"], be_to_iso(inv["date"]), inv.get("vendor_code"), inv["wh_code"], inv.get("branch"),
                 inv.get("po"), inv.get("inv_ref"),
                 None if inv.get("net") is None else int(round(inv["net"] * 100)),
                 Path(source_file).name, imp_id))
            for k, it in enumerate(inv["items"], 1):
                conn.execute(
                    "INSERT INTO invoice_items(invoice_no,seq,code,descr,qty,unit,price_satang,amount_satang) "
                    "VALUES (?,?,?,?,?,?,?,?)",
                    (inv["no"], k, it.get("code"), it.get("desc"), core.num(it.get("qty")), it.get("unit"),
                     satang(it.get("price")), satang(it.get("amount"))))
            new += 1
        conn.execute("UPDATE imports SET n_new=?, n_dup=? WHERE id=?", (new, dup, imp_id))
    return new, dup


def backup(conn, backup_dir):
    """คัดลอกฐานข้อมูลที่สมบูรณ์ไปโฟลเดอร์สำรอง เก็บล่าสุด KEEP_BACKUPS ไฟล์"""
    backup_dir = Path(backup_dir)
    backup_dir.mkdir(parents=True, exist_ok=True)
    out = backup_dir / f"invoices-{datetime.now():%Y%m%d-%H%M%S}.db"
    dst = sqlite3.connect(str(out))
    try:
        conn.backup(dst)
    finally:
        dst.close()
    for old in sorted(backup_dir.glob("invoices-*.db"))[:-KEEP_BACKUPS]:
        old.unlink(missing_ok=True)
    return out


# ---------------------------------------------------------------- load back
def months(conn):
    """รายการเดือนที่มีข้อมูล ['2026-08', ...] เรียงใหม่ -> เก่า"""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT substr(date_iso,1,7) FROM invoices ORDER BY 1 DESC")]


def month_label(ym):
    y, m = ym.split("-")
    names = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    return f"{names[int(m)]} {int(y) + 543}"


def load_invoices(conn, month=None):
    """ดึงใบกลับมาในรูปเดียวกับที่ตัวสกัดคืน (ใช้สร้าง xlsx/PDF ได้ทันที) month='YYYY-MM' หรือ None=ทั้งหมด"""
    q = "SELECT * FROM invoices"
    args = ()
    if month:
        q += " WHERE substr(date_iso,1,7)=?"
        args = (month,)
    q += " ORDER BY date_iso, no"
    out = []
    for r in conn.execute(q, args).fetchall():
        items = []
        for it in conn.execute("SELECT * FROM invoice_items WHERE invoice_no=? ORDER BY seq", (r["no"],)):
            items.append({
                "seq": str(it["seq"]), "code": it["code"] or "", "desc": it["descr"] or "",
                "size": "", "unit": it["unit"] or "", "unit2": it["unit"] or "", "free": "", "pct": "", "disc": "0.00",
                "qty": f"{it['qty']:g}" if it["qty"] is not None else "",
                "price": f"{it['price_satang'] / 100:.2f}" if it["price_satang"] is not None else "",
                "amount": f"{it['amount_satang'] / 100:.2f}" if it["amount_satang"] is not None else "",
            })
        out.append({
            "no": r["no"], "date": iso_to_be(r["date_iso"]), "vendor_code": r["vendor_code"] or "",
            "vendor": "", "wh_code": r["wh_code"], "branch": r["branch_pdf"] or "", "po": r["po"] or "",
            "inv_ref": r["inv_ref"] or "", "net": money(r["net_satang"]), "items": items,
            "text": "", "pages": [],
        })
    return out


# ---------------------------------------------------------------- dashboard
def summary(conn):
    r = conn.execute("SELECT COUNT(*), COALESCE(SUM(net_satang),0), MIN(date_iso), MAX(date_iso) FROM invoices").fetchone()
    return {"count": r[0], "total": r[1] / 100, "first": r[2], "last": r[3]}


def monthly(conn, last_n=12, fill=True):
    """[(YYYY-MM, ยอดรวม, จำนวนใบ)] เรียงเก่า -> ใหม่ เฉพาะ last_n เดือนล่าสุดที่มีข้อมูล
    fill=True เติมเดือนที่ไม่มีข้อมูลระหว่างช่วงเป็น 0 เพื่อให้กราฟต่อเนื่อง"""
    rows = conn.execute(
        "SELECT substr(date_iso,1,7) ym, SUM(net_satang), COUNT(*) FROM invoices GROUP BY ym ORDER BY ym DESC LIMIT ?",
        (last_n,)).fetchall()
    data = {r[0]: (r[1] / 100, r[2]) for r in rows}
    if not data:
        return []
    if not fill:
        return [(m, data[m][0], data[m][1]) for m in sorted(data)]
    first, last = min(data), max(data)
    y, m = int(first[:4]), int(first[5:])
    out = []
    while f"{y:04d}-{m:02d}" <= last:
        ym = f"{y:04d}-{m:02d}"
        v, c = data.get(ym, (0.0, 0))
        out.append((ym, v, c))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def by_branch(conn, month=None):
    q = "SELECT wh_code, SUM(net_satang), COUNT(*) FROM invoices"
    args = ()
    if month:
        q += " WHERE substr(date_iso,1,7)=?"
        args = (month,)
    q += " GROUP BY wh_code ORDER BY 2 DESC"
    return [(r[0], r[1] / 100, r[2]) for r in conn.execute(q, args)]


def by_product(conn, month=None):
    q = ("SELECT i.code, MAX(i.descr), SUM(i.qty), SUM(i.amount_satang) FROM invoice_items i "
         "JOIN invoices v ON v.no=i.invoice_no")
    args = ()
    if month:
        q += " WHERE substr(v.date_iso,1,7)=?"
        args = (month,)
    q += " GROUP BY i.code ORDER BY 4 DESC"
    return [(r[0], r[1], r[2] or 0, (r[3] or 0) / 100) for r in conn.execute(q, args)]


def recent_imports(conn, n=8):
    return [dict(r) for r in conn.execute(
        "SELECT source_file, imported_at, n_new, n_dup FROM imports ORDER BY id DESC LIMIT ?", (n,))]
