"""หน้าต่างโปรแกรม: เลือก PDF -> ตรวจรายการ -> สร้าง .xlsx (ใบวางบิล + ใบส่งของ)"""
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import invoice_to_xlsx as core

FONT = ("Tahoma", 10)
FONT_BOLD = ("Tahoma", 10, "bold")
FONT_BIG = ("Tahoma", 12, "bold")
CFG_LABELS = [
    ("sender_name", "ชื่อผู้ส่งของ"),
    ("sender_address", "ที่อยู่ผู้ส่งของ"),
    ("sender_tax_id", "เลขประจำตัวผู้เสียภาษีอากร"),
    ("deliverer_name", "ชื่อที่พิมพ์ในช่อง 'ผู้ส่งสินค้า'"),
    ("customer_company", "ชื่อบริษัทที่ส่งสินค้าให้"),
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("แปลงใบรับสินค้า → ใบวางบิล / ใบส่งของ")
        self.geometry("1020x680")
        self.minsize(880, 560)
        style = ttk.Style(self)
        style.configure(".", font=FONT)
        style.configure("Treeview", font=FONT, rowheight=24)
        style.configure("Treeview.Heading", font=FONT_BOLD)
        style.configure("Accent.TButton", font=FONT_BOLD)

        self.pdf_path = None
        self.invs = []
        self.branches = core.read_branches()
        self.cfg_ok = False

        nb = self.nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_run, self.tab_br, self.tab_cfg = ttk.Frame(nb), ttk.Frame(nb), ttk.Frame(nb)
        nb.add(self.tab_run, text="  สกัดเอกสาร  ")
        nb.add(self.tab_br, text="  สาขา  ")
        nb.add(self.tab_cfg, text="  ตั้งค่า  ")
        self._build_run()
        self._build_branches()
        self._build_config()
        self._load_cfg(initial=True)

    # ------------------------------------------------------------ tab 1
    def _build_run(self):
        f = self.tab_run
        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 6))
        ttk.Button(top, text="1. เลือกไฟล์ PDF…", style="Accent.TButton", command=self.choose_pdf).pack(side="left")
        self.lbl_file = ttk.Label(top, text="ยังไม่ได้เลือกไฟล์", foreground="#666")
        self.lbl_file.pack(side="left", padx=12)

        self.lbl_sum = ttk.Label(f, text="", font=FONT_BIG)
        self.lbl_sum.pack(anchor="w", padx=12, pady=(6, 2))

        cols = ("date", "branch", "no", "po", "items", "amount", "status")
        heads = ("วันที่รับ", "สาขา", "เลขที่ใบรับสินค้า", "อ้างอิงใบสั่งซื้อ", "รายการ", "จำนวนเงิน", "สถานะ")
        widths = (90, 260, 130, 130, 60, 100, 160)
        box = ttk.Frame(f)
        box.pack(fill="both", expand=True, padx=12, pady=6)
        self.tree = ttk.Treeview(box, columns=cols, show="headings", selectmode="browse")
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="e" if c in ("amount", "items") else "w")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("bad", background="#FFC7CE")
        self.tree.tag_configure("warn", background="#FFF2CC")

        self.warn_frame = ttk.Frame(f)
        self.warn_frame.pack(fill="x", padx=12)
        self.lbl_warn = ttk.Label(self.warn_frame, text="", foreground="#B00020", wraplength=720, justify="left")
        self.lbl_warn.pack(side="left")
        self.btn_goto = ttk.Button(self.warn_frame, text="ไปกรอกที่อยู่ที่แท็บสาขา",
                                   command=lambda: self.nb.select(self.tab_br))

        bottom = ttk.Frame(f)
        bottom.pack(fill="x", padx=12, pady=12)
        self.btn_make = ttk.Button(bottom, text="2. สร้างไฟล์ .xlsx…", style="Accent.TButton",
                                   command=self.make_xlsx, state="disabled")
        self.btn_make.pack(side="right")
        self.lbl_status = ttk.Label(bottom, text="", foreground="#666")
        self.lbl_status.pack(side="left")

    def choose_pdf(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ใบรับสินค้า (PDF)",
            initialdir=str(Path(self.pdf_path).parent) if self.pdf_path else None,
            filetypes=[("PDF", "*.pdf"), ("ทุกไฟล์", "*.*")])
        if not path:
            return
        self.lbl_status.config(text="กำลังอ่านไฟล์…")
        self.btn_make.config(state="disabled")
        self.update_idletasks()
        threading.Thread(target=self._read_thread, args=(path,), daemon=True).start()

    def _read_thread(self, path):
        try:
            invs = core.read_invoices(path)
            err = None
        except Exception as e:  # noqa: BLE001
            invs, err = [], e
        self.after(0, self.load_result, path, invs, err)

    def load_result(self, path, invs, err):
        self.lbl_status.config(text="")
        if err is not None:
            messagebox.showerror("อ่านไฟล์ไม่ได้", f"{type(err).__name__}: {err}")
            return
        if not invs:
            messagebox.showwarning("ไม่พบใบรับสินค้า", "ไฟล์นี้ไม่มีหน้าที่เป็นใบรับสินค้า")
            return
        self.pdf_path, self.invs = path, invs
        self.lbl_file.config(text=Path(path).name, foreground="#000")
        # สาขาใหม่ที่ยังไม่มีในตาราง: เติมในหน้าสาขา (ยังไม่บันทึกจนกว่าจะสร้างไฟล์หรือกดบันทึก)
        for i in invs:
            if i["wh_code"] not in self.branches:
                self.branches[i["wh_code"]] = (i["branch"], "")
        self._refresh_branch_tree()
        self._refresh_preview()

    def _refresh_preview(self):
        self.tree.delete(*self.tree.get_children())
        bad = set(core.mismatched(self.invs))
        miss = {c for c, _ in core.missing_addresses(self.invs, self.branches)}
        for inv in sorted(self.invs, key=lambda i: (core.d_key(i["date"]), i["no"])):
            name = self.branches.get(inv["wh_code"], (inv["branch"], ""))[0]
            if inv["no"] in bad:
                status, tag = "ยอดไม่ตรง", "bad"
            elif inv["wh_code"] in miss:
                status, tag = "ขาดที่อยู่สาขา", "warn"
            else:
                status, tag = "ตรงกัน", ""
            net = inv["net"]
            self.tree.insert("", "end", tags=(tag,), values=(
                inv["date"], name, inv["no"], inv["po"], len(inv["items"]),
                f"{net:,.2f}" if net is not None else "-", status))
        total = sum(i["net"] or 0 for i in self.invs)
        self.lbl_sum.config(text=f"อ่านได้ {len(self.invs)} ใบ   ยอดรวม {total:,.2f} บาท   "
                                 f"ยอดไม่ตรง {len(bad)} ใบ")
        missing = core.missing_addresses(self.invs, self.branches)
        if missing:
            self.lbl_warn.config(text="สาขาที่ยังไม่มีที่อยู่: " + ", ".join(f"{c} {n}" for c, n in missing))
            self.lbl_warn.pack(side="left")
            self.btn_goto.pack(side="right")
        else:
            self.lbl_warn.pack_forget()
            self.btn_goto.pack_forget()
        self.btn_make.config(state="normal" if self.cfg_ok else "disabled")
        if not self.cfg_ok:
            self.lbl_status.config(text="ต้องกรอกแท็บ 'ตั้งค่า' ให้ครบก่อนสร้างไฟล์", foreground="#B00020")

    def make_xlsx(self):
        if not self.cfg_ok:
            self.nb.select(self.tab_cfg)
            return
        missing = core.missing_addresses(self.invs, self.branches)
        if missing:
            names = "\n".join(f"  • {c} {n}" for c, n in missing)
            ans = messagebox.askyesnocancel(
                "มีสาขาที่ยังไม่มีที่อยู่",
                f"สาขาต่อไปนี้ยังไม่มีที่อยู่ ใบส่งของของสาขานี้จะไม่มีที่อยู่:\n{names}\n\n"
                "ใช่ = สร้างไฟล์ต่อ (เว้นที่อยู่ว่าง)\nไม่ = ไปกรอกที่อยู่ก่อน\nยกเลิก = กลับ")
            if ans is None:
                return
            if ans is False:
                self.nb.select(self.tab_br)
                return
        src = Path(self.pdf_path)
        out = filedialog.asksaveasfilename(
            title="บันทึกไฟล์ .xlsx", initialdir=str(src.parent), initialfile=src.stem + ".xlsx",
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not out:
            return
        try:
            core.save_branches(self.branches)
            core.build(self.invs, out)
        except PermissionError:
            messagebox.showerror("บันทึกไม่ได้", "ไฟล์นี้เปิดค้างอยู่ใน Excel  ปิดไฟล์ก่อน หรือเลือกชื่อใหม่")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("สร้างไฟล์ไม่สำเร็จ", f"{type(e).__name__}: {e}")
            return
        self.lbl_status.config(text=f"บันทึกแล้ว: {Path(out).name}", foreground="#1B7F3B")
        if messagebox.askyesno("เสร็จแล้ว", f"สร้างไฟล์เรียบร้อย\n{out}\n\nเปิดโฟลเดอร์ที่เก็บไฟล์ไหม?"):
            os.startfile(Path(out).parent)

    # ------------------------------------------------------------ tab 2
    def _build_branches(self):
        f = self.tab_br
        ttk.Label(f, text="ชื่อและที่อยู่สาขา (แสดงในใบส่งของ)", font=FONT_BIG).pack(anchor="w", padx=12, pady=(12, 4))
        box = ttk.Frame(f)
        box.pack(fill="both", expand=True, padx=12, pady=4)
        self.btree = ttk.Treeview(box, columns=("code", "name", "addr"), show="headings", selectmode="browse")
        for c, h, w in (("code", "รหัสคลัง", 80), ("name", "ชื่อสาขา", 300), ("addr", "ที่อยู่", 560)):
            self.btree.heading(c, text=h)
            self.btree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.btree.yview)
        self.btree.configure(yscrollcommand=sb.set)
        self.btree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.btree.tag_configure("warn", background="#FFF2CC")
        self.btree.bind("<<TreeviewSelect>>", self._pick_branch)

        form = ttk.Frame(f)
        form.pack(fill="x", padx=12, pady=6)
        self.v_code, self.v_name, self.v_addr = tk.StringVar(), tk.StringVar(), tk.StringVar()
        for r, (lab, var, w) in enumerate((("รหัสคลัง", self.v_code, 14), ("ชื่อสาขา", self.v_name, 60),
                                           ("ที่อยู่", self.v_addr, 90))):
            ttk.Label(form, text=lab).grid(row=r, column=0, sticky="w", pady=2)
            ttk.Entry(form, textvariable=var, width=w).grid(row=r, column=1, sticky="we", pady=2, padx=6)
        form.columnconfigure(1, weight=1)

        btns = ttk.Frame(f)
        btns.pack(fill="x", padx=12, pady=(2, 12))
        ttk.Button(btns, text="เพิ่ม / อัปเดตแถว", command=self._upsert_branch).pack(side="left")
        ttk.Button(btns, text="ลบแถว", command=self._delete_branch).pack(side="left", padx=6)
        ttk.Button(btns, text="บันทึก", style="Accent.TButton", command=self._save_branches).pack(side="right")
        self.lbl_br = ttk.Label(btns, text="", foreground="#666")
        self.lbl_br.pack(side="right", padx=10)
        self._refresh_branch_tree()

    def _refresh_branch_tree(self):
        self.btree.delete(*self.btree.get_children())
        for code, (name, addr) in self.branches.items():
            self.btree.insert("", "end", iid=code, tags=("" if addr.strip() else "warn",), values=(code, name, addr))

    def _pick_branch(self, _evt=None):
        sel = self.btree.selection()
        if sel:
            code = sel[0]
            name, addr = self.branches[code]
            self.v_code.set(code)
            self.v_name.set(name)
            self.v_addr.set(addr)

    def _upsert_branch(self):
        code = self.v_code.get().strip()
        if not code:
            messagebox.showwarning("ยังไม่ได้กรอก", "กรอกรหัสคลังก่อน")
            return
        self.branches[code] = (self.v_name.get().strip(), self.v_addr.get().strip())
        self._refresh_branch_tree()
        self.btree.selection_set(code)
        self.lbl_br.config(text="มีการแก้ไขที่ยังไม่ได้บันทึก", foreground="#B00020")
        if self.invs:
            self._refresh_preview()

    def _delete_branch(self):
        sel = self.btree.selection()
        if sel and messagebox.askyesno("ลบสาขา", f"ลบสาขารหัส {sel[0]} ?"):
            self.branches.pop(sel[0], None)
            self._refresh_branch_tree()
            self.lbl_br.config(text="มีการแก้ไขที่ยังไม่ได้บันทึก", foreground="#B00020")

    def _save_branches(self):
        core.save_branches(self.branches)
        self.lbl_br.config(text="บันทึกแล้ว", foreground="#1B7F3B")

    # ------------------------------------------------------------ tab 3
    def _build_config(self):
        f = self.tab_cfg
        ttk.Label(f, text="ข้อมูลผู้ส่งของ (พิมพ์ที่หัวใบส่งของ)", font=FONT_BIG).pack(anchor="w", padx=12, pady=(12, 8))
        grid = ttk.Frame(f)
        grid.pack(fill="x", padx=12)
        self.cfg_vars = {}
        for r, (key, lab) in enumerate(CFG_LABELS):
            ttk.Label(grid, text=lab).grid(row=r, column=0, sticky="w", pady=5)
            v = tk.StringVar()
            ttk.Entry(grid, textvariable=v, width=80).grid(row=r, column=1, sticky="we", padx=8, pady=5)
            self.cfg_vars[key] = v
        grid.columnconfigure(1, weight=1)
        row = ttk.Frame(f)
        row.pack(fill="x", padx=12, pady=12)
        ttk.Button(row, text="บันทึกการตั้งค่า", style="Accent.TButton", command=self._save_cfg).pack(side="right")
        self.lbl_cfg = ttk.Label(row, text="", foreground="#666")
        self.lbl_cfg.pack(side="left")

    def _load_cfg(self, initial=False):
        try:
            cfg = core.load_config()
        except core.ConfigError as e:
            self.cfg_ok = False
            self.lbl_cfg.config(text=str(e), foreground="#B00020")
            if initial:
                self.nb.select(self.tab_cfg)
            # ถ้ามีไฟล์แต่ค่าไม่ครบ ให้เติมช่องที่มีไว้ก่อน
            try:
                import json
                with open(core.CONFIG_JSON, encoding="utf-8-sig") as fh:
                    cfg = json.load(fh)
                for k, v in self.cfg_vars.items():
                    v.set(cfg.get(k, ""))
            except Exception:  # noqa: BLE001
                pass
            return
        for k, v in self.cfg_vars.items():
            v.set(cfg[k])
        core.apply_config(cfg)
        self.cfg_ok = True
        self.lbl_cfg.config(text="โหลดการตั้งค่าแล้ว", foreground="#666")

    def _save_cfg(self):
        cfg = {k: v.get().strip() for k, v in self.cfg_vars.items()}
        empty = [lab for k, lab in CFG_LABELS if not cfg[k]]
        if empty:
            messagebox.showwarning("ข้อมูลไม่ครบ", "ยังไม่ได้กรอก: " + ", ".join(empty))
            return
        core.save_config(cfg)
        core.apply_config(cfg)
        self.cfg_ok = True
        self.lbl_cfg.config(text="บันทึกแล้ว", foreground="#1B7F3B")
        self.lbl_status.config(text="", foreground="#666")
        if self.invs:
            self.btn_make.config(state="normal")


if __name__ == "__main__":
    App().mainloop()
