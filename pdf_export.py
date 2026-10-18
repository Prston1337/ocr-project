"""ส่งออกใบวางบิล / ใบส่งของเป็น PDF (A4 แบ่งหน้าตายตัว) โดยพิมพ์ HTML ผ่าน Microsoft Edge แบบ headless

ใช้เมื่อไม่มี Excel: PDF แบ่งหน้าเหมือนกันทุกเครื่อง ไม่ขึ้นกับ Google Sheets/Excel
"""
import html
import subprocess
import tempfile
from pathlib import Path

import invoice_to_xlsx as core

ROWS_PER_PAGE = 18
EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

CSS = """
@page { size: A4; margin: 12mm 12mm 12mm 12mm; }
* { box-sizing: border-box; }
body { font-family: 'Tahoma', 'Leelawadee UI', sans-serif; font-size: 11pt; color: #111; margin: 0; }
.page { height: 268mm; display: flex; flex-direction: column; page-break-after: always; break-after: page; }
.page:last-child { page-break-after: auto; break-after: auto; }
h1 { text-align: center; color: #2E96D0; font-size: 22pt; margin: 0; line-height: 1.5; }
.copy { text-align: center; color: #2E96D0; font-size: 10pt; margin-bottom: 6mm; }
.head { display: flex; justify-content: space-between; gap: 8mm; }
.head .left { flex: 1; line-height: 1.7; }
.head .right { width: 62mm; }
.head .right div { display: flex; border-bottom: 0.3mm solid #bbb; padding: 0.8mm 0; line-height: 1.5; }
.head .right span { color: #2E96D0; width: 20mm; font-size: 10pt; }
.b { font-weight: bold; }
table { border-collapse: collapse; width: 100%; }
.items { margin-top: 7mm; }
.items th { background: #2E96D0; color: #fff; padding: 1.6mm 2mm; text-align: center; font-weight: bold; }
.items td { padding: 1.4mm 2mm; border-bottom: 0.2mm solid #ccc; }
.c { text-align: center; } .r { text-align: right; } .l { text-align: left; }
.fill { flex: 1; }
.total { display: flex; justify-content: flex-end; align-items: baseline; gap: 6mm; margin-top: 4mm; }
.total .cnt { color: #2E96D0; }
.total .sum { border-top: 0.3mm solid #000; border-bottom: 1mm double #000; padding: 0.5mm 2mm; font-weight: bold; min-width: 34mm; text-align: right; }
.baht { text-align: center; font-weight: bold; margin: 3mm 0 16mm; }
.sign { display: flex; justify-content: space-between; padding-bottom: 6mm; }
.bill { font-size: 9.5pt; }
.bill td, .bill th { padding: 1mm 1.5mm; border: 0.2mm solid #999; white-space: nowrap; }
.bill td.wrap { white-space: normal; line-height: 1.25; }
.bill th { background: #DDEBF7; }
.bill tr { page-break-inside: avoid; }
.info { line-height: 1.7; margin-bottom: 4mm; }
.grand td { border: none; font-weight: bold; }
"""


def e(x):
    return html.escape(str(x if x is not None else ""))


def money(v):
    return "" if v is None else f"{v:,.2f}"


def _qty(v):
    n = core.num(v)
    if n is None:
        return e(v)
    return f"{n:g}"


def delivery_pages(invs, branches, label):
    out = []
    for inv in sorted(invs, key=lambda i: (core.d_key(i["date"]), i["no"])):
        name, addr = branches.get(inv["wh_code"], (inv["branch"], ""))
        items = inv["items"] or [{}]
        chunks = [items[k:k + ROWS_PER_PAGE] for k in range(0, len(items), ROWS_PER_PAGE)]
        total = inv["net"] if inv["net"] is not None else sum(core.num(i.get("amount")) or 0 for i in items)
        for pg, chunk in enumerate(chunks, 1):
            rows = ""
            for k, it in enumerate(chunk):
                n = (pg - 1) * ROWS_PER_PAGE + k + 1
                rows += (f"<tr><td class='c'>{n}</td><td class='c'>{e(it.get('code'))}</td>"
                         f"<td class='l'>{e(it.get('desc'))}</td><td class='c'>{_qty(it.get('qty'))}</td>"
                         f"<td class='c'>{e(it.get('unit'))}</td><td class='r'>{money(core.num(it.get('price')))}</td>"
                         f"<td class='r'>{money(core.num(it.get('amount')))}</td></tr>")
            foot = ""
            if pg == len(chunks):
                foot = (f"<div class='total'><span class='cnt'>จำนวน</span><span>{len(items)} รายการ</span>"
                        f"<span class='sum'>{money(total)}</span></div>"
                        f"<div class='baht'>({e(core.baht_text(total))})</div>")
            else:
                foot = "<div class='baht'>&nbsp;</div>"
            pgtxt = f"  (หน้า {pg}/{len(chunks)})" if len(chunks) > 1 else ""
            out.append(f"""
<div class='page'>
  <h1>ใบส่งของ</h1><div class='copy'>{e(label)}{pgtxt}</div>
  <div class='head'>
    <div class='left'>
      <div>ชื่อผู้ส่งของ: {e(core.SENDER_NAME)}</div>
      <div>{e(core.SENDER_ADDR)}</div>
      <div>เลขประจำตัวผู้เสียภาษีอากร: {e(core.SENDER_TAX_ID)}</div>
      <div class='b'>{e(core.CUSTOMER_COMPANY)}</div>
      <div class='b'>{e(name)}</div>
      <div>{e(addr)}</div>
    </div>
    <div class='right'>
      <div><span>เลขที่</span>{e(inv['no'])}</div>
      <div><span>วันที่</span>{e(inv['date'])}</div>
      <div><span>อ้างอิง</span>{e(inv['po'])}</div>
      <div><span>ใบแจ้งหนี้</span>{e(inv['inv_ref'])}</div>
    </div>
  </div>
  <table class='items'>
    <tr><th style='width:8mm'>#</th><th style='width:22mm'>รหัสสินค้า</th><th class='l'>รายละเอียด</th>
        <th style='width:16mm'>จำนวน</th><th style='width:14mm'>หน่วย</th>
        <th style='width:27mm;white-space:nowrap'>ราคา/หน่วย</th><th style='width:26mm'>จำนวนเงิน</th></tr>
    {rows}
  </table>
  <div class='fill'></div>
  {foot}
  <div class='sign'><div>ผู้รับสินค้า ______________________   วันที่ ____________</div>
  <div>ผู้ส่งสินค้า  {e(core.DELIVERER_NAME)}      วันที่ ________</div></div>
</div>""")
    return "".join(out)


def bill_pages(invs, branches, label):
    rows = ""
    total = 0
    for n, inv in enumerate(sorted(invs, key=lambda i: (core.d_key(i["date"]), i["no"])), 1):
        br = branches.get(inv["wh_code"], (inv["branch"], ""))[0].replace("สาขา", "", 1)
        total += inv["net"] or 0
        rows += (f"<tr><td class='c'>{n}</td><td class='c'>{e(inv['date'])}</td><td class='l wrap'>{e(br)}</td>"
                 f"<td class='c'>{e(inv['no'])}</td><td class='c'>{e(inv['po'])}</td>"
                 f"<td class='c'>{e(inv['inv_ref'])}</td><td class='r'>{money(inv['net'])}</td></tr>")
    dates = sorted((i["date"] for i in invs), key=core.d_key)
    return f"""
<div class='page' style='height:auto'>
  <h1 style='color:#111'>ใบวางบิล <span style='font-size:14pt'>({e(label)})</span></h1>
  <div class='info'>
    <div><b>วางบิลถึง</b>&nbsp;&nbsp;{e(core.CUSTOMER_COMPANY)}</div>
    <div><b>ช่วงวันที่รับสินค้า</b>&nbsp;&nbsp;{e(dates[0])} - {e(dates[-1])}</div>
  </div>
  <table class='bill'>
    <colgroup><col style='width:12mm'><col style='width:23mm'><col><col style='width:30mm'><col style='width:30mm'>
      <col style='width:34mm'><col style='width:22mm'></colgroup>
    <tr><th>ลำดับ</th><th>วันที่รับ</th><th>สาขา</th><th>เลขที่ใบรับ</th><th>อ้างอิง PO</th>
        <th>ใบแจ้งหนี้</th><th>จำนวนเงิน</th></tr>
    {rows}
    <tr class='grand'><td colspan='6' class='r'>รวมทั้งสิ้น</td><td class='r' style='border:0.3mm solid #000'>{money(total)}</td></tr>
  </table>
  <div class='sign' style='margin-top:18mm'><div>ผู้วางบิล ____________________</div><div>ผู้รับวางบิล ____________________</div></div>
</div>"""


def find_browser():
    for p in EDGE_PATHS:
        if Path(p).exists():
            return p
    return None


def html_to_pdf(body, out_pdf):
    browser = find_browser()
    if not browser:
        raise RuntimeError("ไม่พบ Microsoft Edge/Chrome ในเครื่อง จึงสร้าง PDF ไม่ได้")
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "doc.html"
        src.write_text(f"<!doctype html><html lang='th'><head><meta charset='utf-8'><style>{CSS}</style></head>"
                       f"<body>{body}</body></html>", encoding="utf-8")
        cmd = [browser, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
               f"--user-data-dir={Path(tmp) / 'profile'}", f"--print-to-pdf={out_pdf}", src.as_uri()]
        subprocess.run(cmd, check=True, timeout=120, creationflags=0x08000000,  # CREATE_NO_WINDOW
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if not Path(out_pdf).exists():
        raise RuntimeError("สร้างไฟล์ PDF ไม่สำเร็จ")


def export_pdfs(invs, base_path):
    """สร้าง <base>-ใบวางบิล.pdf และ <base>-ใบส่งของ.pdf คืน list ของไฟล์ที่สร้าง"""
    base = Path(base_path)
    branches = core.load_branches(invs)
    bill = base.with_name(base.stem + "-ใบวางบิล.pdf")
    dele = base.with_name(base.stem + "-ใบส่งของ.pdf")
    html_to_pdf(bill_pages(invs, branches, "ต้นฉบับ") + bill_pages(invs, branches, "สำเนา"), str(bill))
    html_to_pdf(delivery_pages(invs, branches, "ต้นฉบับ") + delivery_pages(invs, branches, "สำเนา"), str(dele))
    return [str(bill), str(dele)]

