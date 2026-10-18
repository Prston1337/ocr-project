"""หน้าต่างโปรแกรม (ธีม Dark สไตล์ Quantix): เลือก PDF -> ตรวจรายการ -> สร้าง .xlsx"""
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

import invoice_to_xlsx as core
import db
import pdf_export

# ---- design tokens (ตามธีม Dark ที่ตกลงไว้) ----
BG = "#08080B"
SIDEBAR = "#0C0C11"
CARD = "#101016"
CARD_2 = "#15151C"
BORDER = "#1F1F29"
TEXT = "#ECECF1"
MUTED = "#8B8B9B"
ACCENT = "#5B6CF0"
ACCENT_HOVER = "#6C7BF5"
ACCENT_SOFT = "#1B1D3A"
OK_BG, OK_FG = "#10301F", "#5FD38D"
WARN_BG, WARN_FG = "#352A0E", "#F2C14E"
BAD_BG, BAD_FG = "#381518", "#F0707A"
FAMILY = "Leelawadee UI"

CFG_LABELS = [
    ("sender_name", "ชื่อผู้ส่งของ"),
    ("sender_address", "ที่อยู่ผู้ส่งของ"),
    ("sender_tax_id", "เลขประจำตัวผู้เสียภาษีอากร"),
    ("deliverer_name", "ชื่อที่พิมพ์ในช่อง 'ผู้ส่งสินค้า'"),
    ("customer_company", "ชื่อบริษัทที่ส่งสินค้าให้"),
]

ctk.set_appearance_mode("dark")


def font(size=13, bold=False):
    return ctk.CTkFont(family=FAMILY, size=size, weight="bold" if bold else "normal")


def card(parent, **kw):
    return ctk.CTkFrame(parent, fg_color=CARD, border_width=1, border_color=BORDER, corner_radius=14, **kw)


def primary_btn(parent, text, command, **kw):
    return ctk.CTkButton(parent, text=text, command=command, font=font(13, True), fg_color=ACCENT,
                         hover_color=ACCENT_HOVER, text_color="#FFFFFF", corner_radius=10, height=38, **kw)


def ghost_btn(parent, text, command, **kw):
    return ctk.CTkButton(parent, text=text, command=command, font=font(13), fg_color=CARD_2, hover_color="#1D1D27",
                         border_width=1, border_color=BORDER, text_color=TEXT, corner_radius=10, height=38, **kw)


def entry(parent, var, **kw):
    return ctk.CTkEntry(parent, textvariable=var, font=font(13), fg_color=CARD_2, border_color=BORDER,
                        text_color=TEXT, corner_radius=10, height=38, **kw)


class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=BG)
        self.title("แปลงใบรับสินค้า → ใบวางบิล / ใบส่งของ")
        self.geometry("1180x740")
        self.minsize(1000, 640)
        self._style_tree()

        self.pdf_path = None
        self.invs = []
        self.branches = core.read_branches()
        self.cfg_ok = False
        self.conn = db.connect()
        self.saved = False

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_sidebar()
        self.main = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(0, weight=1)
        self.pages = {}
        self._build_run()
        self._build_dash()
        self._build_branches()
        self._build_config()
        self.show("run")
        self._load_cfg(initial=True)

    # ------------------------------------------------------------ shell
    def _style_tree(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("Dark.Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT, borderwidth=0,
                     rowheight=34, font=(FAMILY, 10), relief="flat")
        st.configure("Dark.Treeview.Heading", background=CARD, foreground=MUTED, borderwidth=0,
                     font=(FAMILY, 10), relief="flat", padding=(8, 10))
        st.map("Dark.Treeview", background=[("selected", ACCENT_SOFT)], foreground=[("selected", TEXT)])
        st.map("Dark.Treeview.Heading", background=[("active", CARD)])
        st.layout("Dark.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        st.configure("Dark.Vertical.TScrollbar", background=BORDER, troughcolor=CARD, bordercolor=CARD,
                     arrowcolor=MUTED, lightcolor=CARD, darkcolor=CARD, gripcount=0)

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, fg_color=SIDEBAR, corner_radius=0, width=250, border_width=0)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        ctk.CTkFrame(self, fg_color=BORDER, width=1, corner_radius=0).grid(row=0, column=0, sticky="nse")

        ctk.CTkLabel(sb, text="Invoice → Xlsx", font=font(18, True), text_color=TEXT, anchor="w").pack(
            fill="x", padx=24, pady=(28, 0))
        ctk.CTkLabel(sb, text="ใบวางบิล / ใบส่งของ", font=font(11), text_color=MUTED, anchor="w").pack(
            fill="x", padx=24, pady=(0, 18))
        ctk.CTkFrame(sb, fg_color=BORDER, height=1).pack(fill="x", padx=24)
        ctk.CTkLabel(sb, text="เมนู", font=font(11), text_color=MUTED, anchor="w").pack(
            fill="x", padx=24, pady=(22, 6))

        self.nav = {}
        for key, label in (("run", "สกัดเอกสาร"), ("dash", "Dashboard"), ("branches", "สาขา"), ("config", "ตั้งค่า")):
            b = ctk.CTkButton(sb, text="   " + label, anchor="w", height=44, corner_radius=10, font=font(14),
                              fg_color="transparent", hover_color=CARD_2, text_color=MUTED, border_width=1,
                              border_color=SIDEBAR, command=lambda k=key: self.show(k))
            b.pack(fill="x", padx=16, pady=3)
            self.nav[key] = b
        ctk.CTkLabel(sb, text="v0.2.0 · Dark", font=font(10), text_color=MUTED, anchor="w").pack(
            side="bottom", fill="x", padx=24, pady=20)

    def show(self, key):
        if key == "dash" and "dash" in self.pages:
            self.refresh_dash()
        for k, p in self.pages.items():
            if k == key:
                p.grid(row=0, column=0, sticky="nsew")
            else:
                p.grid_remove()
        for k, b in self.nav.items():
            on = k == key
            b.configure(fg_color=CARD_2 if on else "transparent", text_color=TEXT if on else MUTED,
                        border_color=BORDER if on else SIDEBAR)

    def _page(self, key, crumb, title):
        p = ctk.CTkFrame(self.main, fg_color=BG, corner_radius=0)
        p.columnconfigure(0, weight=1)
        self.pages[key] = p
        ctk.CTkLabel(p, text=f"ภาพรวม  /  {crumb}", font=font(12), text_color=MUTED, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=32, pady=(26, 0))
        ctk.CTkLabel(p, text=title, font=font(26, True), text_color=TEXT, anchor="w").grid(
            row=1, column=0, sticky="ew", padx=32, pady=(4, 14))
        return p

    def _tree(self, parent, cols, heads, widths, anchors=None):
        box = tk.Frame(parent, bg=CARD)
        tree = ttk.Treeview(box, columns=cols, show="headings", selectmode="browse", style="Dark.Treeview")
        for i, (c, h, w) in enumerate(zip(cols, heads, widths)):
            tree.heading(c, text=h, anchor="w")
            tree.column(c, width=w, anchor=(anchors or {}).get(c, "w"))
        sb = ttk.Scrollbar(box, orient="vertical", command=tree.yview, style="Dark.Vertical.TScrollbar")
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=6)
        sb.pack(side="right", fill="y", pady=6)
        return box, tree

    # ------------------------------------------------------------ run page
    def _build_run(self):
        p = self._page("run", "สกัดเอกสาร", "นำเข้าใบรับสินค้า")
        p.rowconfigure(5, weight=1)

        top = ctk.CTkFrame(p, fg_color="transparent")
        top.grid(row=2, column=0, sticky="ew", padx=32)
        primary_btn(top, "เลือกไฟล์ PDF…", self.choose_pdf, width=170).pack(side="left")
        self.lbl_file = ctk.CTkLabel(top, text="ยังไม่ได้เลือกไฟล์", font=font(13), text_color=MUTED)
        self.lbl_file.pack(side="left", padx=16)

        stats = ctk.CTkFrame(p, fg_color="transparent")
        stats.grid(row=3, column=0, sticky="ew", padx=32, pady=(16, 8))
        for i in range(3):
            stats.columnconfigure(i, weight=1, uniform="s")
        self.stat = {}
        for i, (key, label) in enumerate((("count", "ใบรับสินค้า"), ("total", "ยอดรวม (บาท)"),
                                          ("check", "ต้องตรวจ"))):
            c = card(stats)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0 if i == 2 else 8))
            ctk.CTkLabel(c, text=label, font=font(12), text_color=MUTED, anchor="w").pack(
                fill="x", padx=20, pady=(16, 0))
            v = ctk.CTkLabel(c, text="–", font=font(26, True), text_color=TEXT, anchor="w")
            v.pack(fill="x", padx=20, pady=(2, 16))
            self.stat[key] = v

        self.warn = ctk.CTkFrame(p, fg_color=WARN_BG, corner_radius=12, border_width=1, border_color="#4A3B14")
        self.lbl_warn = ctk.CTkLabel(self.warn, text="", font=font(12), text_color=WARN_FG, anchor="w",
                                     justify="left", wraplength=640)
        self.lbl_warn.pack(side="left", padx=16, pady=10)
        ctk.CTkButton(self.warn, text="ไปกรอกที่อยู่", width=120, height=30, font=font(12), corner_radius=8,
                      fg_color="#4A3B14", hover_color="#5C4A1A", text_color=WARN_FG,
                      command=lambda: self.show("branches")).pack(side="right", padx=12)
        self.warn.grid(row=4, column=0, sticky="ew", padx=32, pady=(0, 8))
        self.warn.grid_remove()

        tc = card(p)
        tc.grid(row=5, column=0, sticky="nsew", padx=32, pady=(0, 8))
        tc.rowconfigure(0, weight=1)
        tc.columnconfigure(0, weight=1)
        box, self.tree = self._tree(
            tc, ("date", "branch", "no", "po", "items", "amount", "status"),
            ("วันที่รับ", "สาขา", "เลขที่ใบรับสินค้า", "อ้างอิงใบสั่งซื้อ", "รายการ", "จำนวนเงิน", "สถานะ"),
            (100, 280, 140, 140, 70, 110, 140), {"amount": "e", "items": "e"})
        box.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self.tree.tag_configure("bad", background=BAD_BG, foreground=BAD_FG)
        self.tree.tag_configure("warn", background=WARN_BG, foreground=WARN_FG)

        bottom = ctk.CTkFrame(p, fg_color="transparent")
        bottom.grid(row=6, column=0, sticky="ew", padx=32, pady=(4, 22))
        self.btn_make = primary_btn(bottom, "สร้างไฟล์ .xlsx…", self.make_xlsx, width=190)
        self.btn_make.pack(side="right")
        self.btn_make.configure(state="disabled")
        self.btn_pdf = ghost_btn(bottom, "สร้าง PDF พร้อมพิมพ์…", self.make_pdf, width=210)
        self.btn_pdf.pack(side="right", padx=(0, 10))
        self.btn_pdf.configure(state="disabled")
        self.btn_db = ghost_btn(bottom, "บันทึกเข้าฐานข้อมูล", self.save_db, width=190)
        self.btn_db.pack(side="right", padx=(0, 10))
        self.btn_db.configure(state="disabled")
        self.lbl_status = ctk.CTkLabel(bottom, text="", font=font(12), text_color=MUTED)
        self.lbl_status.pack(side="left")

    def choose_pdf(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ใบรับสินค้า (PDF)",
            initialdir=str(Path(self.pdf_path).parent) if self.pdf_path else None,
            filetypes=[("PDF", "*.pdf"), ("ทุกไฟล์", "*.*")])
        if not path:
            return
        self.lbl_status.configure(text="กำลังอ่านไฟล์…", text_color=MUTED)
        self.btn_make.configure(state="disabled")
        self.btn_pdf.configure(state="disabled")
        self.btn_db.configure(state="disabled")
        self.update_idletasks()
        threading.Thread(target=self._read_thread, args=(path,), daemon=True).start()

    def _read_thread(self, path):
        try:
            invs, err = core.read_invoices(path), None
        except Exception as e:  # noqa: BLE001
            invs, err = [], e
        self.after(0, self.load_result, path, invs, err)

    def load_result(self, path, invs, err):
        self.lbl_status.configure(text="")
        if err is not None:
            messagebox.showerror("อ่านไฟล์ไม่ได้", f"{type(err).__name__}: {err}")
            return
        if not invs:
            messagebox.showwarning("ไม่พบใบรับสินค้า", "ไฟล์นี้ไม่มีหน้าที่เป็นใบรับสินค้า")
            return
        self.pdf_path, self.invs = path, invs
        self.saved = False
        self.btn_db.configure(state="normal", text="บันทึกเข้าฐานข้อมูล")
        self.lbl_file.configure(text=Path(path).name, text_color=TEXT)
        for i in invs:
            if i["wh_code"] not in self.branches:
                self.branches[i["wh_code"]] = (i["branch"], "")
        self._refresh_branch_tree()
        self._refresh_preview()

    def _refresh_preview(self):
        self.tree.delete(*self.tree.get_children())
        bad = set(core.mismatched(self.invs))
        missing = core.missing_addresses(self.invs, self.branches)
        miss = {c for c, _ in missing}
        for inv in sorted(self.invs, key=lambda i: (core.d_key(i["date"]), i["no"])):
            name = self.branches.get(inv["wh_code"], (inv["branch"], ""))[0]
            if inv["no"] in bad:
                status, tag = "● ยอดไม่ตรง", "bad"
            elif inv["wh_code"] in miss:
                status, tag = "● ขาดที่อยู่สาขา", "warn"
            else:
                status, tag = "● ตรงกัน", ""
            net = inv["net"]
            self.tree.insert("", "end", tags=(tag,), values=(
                inv["date"], name, inv["no"], inv["po"], len(inv["items"]),
                f"{net:,.2f}" if net is not None else "-", status))
        total = sum(i["net"] or 0 for i in self.invs)
        self.stat["count"].configure(text=f"{len(self.invs):,}")
        self.stat["total"].configure(text=f"{total:,.2f}")
        n_check = len(bad) + sum(1 for i in self.invs if i["wh_code"] in miss and i["no"] not in bad)
        self.stat["check"].configure(text=f"{n_check:,}", text_color=OK_FG if n_check == 0 else
                                     (BAD_FG if bad else WARN_FG))
        if missing:
            self.lbl_warn.configure(text="สาขาที่ยังไม่มีที่อยู่:  " + ",  ".join(f"{c} {n}" for c, n in missing))
            self.warn.grid()
        else:
            self.warn.grid_remove()
        self.btn_make.configure(state="normal" if self.cfg_ok else "disabled")
        self.btn_pdf.configure(state="normal" if self.cfg_ok else "disabled")
        if not self.cfg_ok:
            self.lbl_status.configure(text="ต้องกรอกหน้า 'ตั้งค่า' ให้ครบก่อนสร้างไฟล์", text_color=BAD_FG)

    def _confirm_missing(self, invs=None):
        """เตือนถ้ามีสาขาไม่มีที่อยู่ คืน True ถ้าให้สร้างต่อ"""
        missing = core.missing_addresses(self.invs if invs is None else invs, self.branches)
        if not missing:
            return True
        names = "\n".join(f"  • {c} {n}" for c, n in missing)
        ans = messagebox.askyesnocancel(
            "มีสาขาที่ยังไม่มีที่อยู่",
            f"สาขาต่อไปนี้ยังไม่มีที่อยู่ ใบส่งของของสาขานี้จะไม่มีที่อยู่:\n{names}\n\n"
            "ใช่ = สร้างไฟล์ต่อ (เว้นที่อยู่ว่าง)\nไม่ = ไปกรอกที่อยู่ก่อน\nยกเลิก = กลับ")
        if ans is False:
            self.show("branches")
        return ans is True

    def make_pdf(self):
        """ส่งออก PDF (ใบวางบิล + ใบส่งของ) แบ่งหน้า A4 ตายตัว ไม่ต้องใช้ Excel"""
        if not self.cfg_ok:
            self.show("config")
            return
        if not self._confirm_missing():
            return
        src = Path(self.pdf_path)
        out = filedialog.asksaveasfilename(
            title="บันทึก PDF (จะได้ 2 ไฟล์: ใบวางบิล และ ใบส่งของ)", initialdir=str(src.parent),
            initialfile=src.stem + ".pdf", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        core.save_branches(self.branches)
        self.lbl_status.configure(text="กำลังสร้าง PDF… (ประมาณ 10 วินาที)", text_color=MUTED)
        self.btn_pdf.configure(state="disabled")
        threading.Thread(target=self._pdf_thread, args=(out,), daemon=True).start()

    def _pdf_thread(self, out):
        try:
            files, err = pdf_export.export_pdfs(self.invs, out), None
        except Exception as e:  # noqa: BLE001
            files, err = [], e
        self.after(0, self._pdf_done, files, err)

    def _pdf_done(self, files, err):
        self.btn_pdf.configure(state="normal")
        if err is not None:
            self.lbl_status.configure(text="สร้าง PDF ไม่สำเร็จ", text_color=BAD_FG)
            messagebox.showerror("สร้าง PDF ไม่สำเร็จ", f"{type(err).__name__}: {err}")
            return
        self.lbl_status.configure(text="สร้าง PDF แล้ว: " + ", ".join(Path(f).name for f in files),
                                  text_color=OK_FG)
        if messagebox.askyesno("เสร็จแล้ว", "สร้าง PDF เรียบร้อย\n" + "\n".join(files) +
                               "\n\nเปิดโฟลเดอร์ที่เก็บไฟล์ไหม?"):
            os.startfile(Path(files[0]).parent)

    def make_xlsx(self):
        if not self.cfg_ok:
            self.show("config")
            return
        if not self._confirm_missing():
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
        self.lbl_status.configure(text=f"บันทึกแล้ว: {Path(out).name}", text_color=OK_FG)
        if messagebox.askyesno("เสร็จแล้ว", f"สร้างไฟล์เรียบร้อย\n{out}\n\nเปิดโฟลเดอร์ที่เก็บไฟล์ไหม?"):
            os.startfile(Path(out).parent)


    # ------------------------------------------------------------ database
    def save_db(self):
        if not self.invs:
            return
        try:
            new, dup = db.import_invoices(self.conn, self.invs, self.pdf_path or "")
            db.backup(self.conn, core.app_dir() / "backup")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("บันทึกฐานข้อมูลไม่ได้", f"{type(e).__name__}: {e}")
            return
        self.saved = True
        self.btn_db.configure(state="disabled", text="บันทึกแล้ว ✓")
        msg = f"บันทึกเข้าฐานข้อมูลแล้ว: ใหม่ {new} ใบ"
        if dup:
            msg += f" · ข้ามใบที่มีอยู่แล้ว {dup} ใบ"
        self.lbl_status.configure(text=msg, text_color=OK_FG)

    def _build_dash(self):
        p = self._page("dash", "Dashboard", "สรุปรายได้")
        body = ctk.CTkScrollableFrame(p, fg_color=BG, corner_radius=0)
        body.grid(row=2, column=0, sticky="nsew", padx=(18, 8))
        p.rowconfigure(2, weight=1)
        body.columnconfigure((0, 1), weight=1, uniform="d")

        top = ctk.CTkFrame(body, fg_color="transparent")
        top.grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 8))
        ctk.CTkLabel(top, text="ช่วงเวลา", font=font(12), text_color=MUTED).pack(side="left")
        self.dash_month = ctk.CTkOptionMenu(top, values=["ทั้งหมด"], width=170, height=34, font=font(12),
                                            dropdown_font=font(12), fg_color=CARD_2, button_color=CARD_2,
                                            button_hover_color="#1D1D27", command=lambda _v: self._dash_fill())
        self.dash_month.pack(side="left", padx=10)
        self.lbl_dash_hint = ctk.CTkLabel(top, text="", font=font(12), text_color=MUTED)
        self.lbl_dash_hint.pack(side="left", padx=8)

        kp = ctk.CTkFrame(body, fg_color="transparent")
        kp.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 8))
        for i in range(4):
            kp.columnconfigure(i, weight=1, uniform="k")
        self.kpi = {}
        for i, (key, label) in enumerate((("total", "รายได้ (บาท)"), ("count", "ใบรับสินค้า"),
                                          ("avg", "เฉลี่ยต่อใบ (บาท)"), ("delta", "เทียบเดือนก่อน*"))):
            c = card(kp)
            c.grid(row=0, column=i, sticky="ew", padx=8)
            ctk.CTkLabel(c, text=label, font=font(12), text_color=MUTED, anchor="w").pack(fill="x", padx=18, pady=(14, 0))
            v = ctk.CTkLabel(c, text="–", font=font(22, True), text_color=TEXT, anchor="w")
            v.pack(fill="x", padx=18, pady=(2, 14))
            self.kpi[key] = v

        mc = card(body)
        mc.grid(row=2, column=0, sticky="nsew", padx=14, pady=8)
        ctk.CTkLabel(mc, text="รายได้รายเดือน", font=font(14, True), anchor="w").pack(fill="x", padx=18, pady=(14, 0))
        self.cv_month = tk.Canvas(mc, height=250, bg=CARD, highlightthickness=0)
        self.cv_month.pack(fill="both", expand=True, padx=10, pady=10)
        self.cv_month.bind("<Configure>", lambda _e: self._dash_draw())

        bc = card(body)
        bc.grid(row=2, column=1, sticky="nsew", padx=14, pady=8)
        ctk.CTkLabel(bc, text="รายได้ตามสาขา", font=font(14, True), anchor="w").pack(fill="x", padx=18, pady=(14, 0))
        self.cv_branch = tk.Canvas(bc, height=250, bg=CARD, highlightthickness=0)
        self.cv_branch.pack(fill="both", expand=True, padx=10, pady=10)
        self.cv_branch.bind("<Configure>", lambda _e: self._dash_draw())

        pc = card(body)
        pc.grid(row=3, column=0, columnspan=2, sticky="ew", padx=14, pady=8)
        ctk.CTkLabel(pc, text="สินค้า", font=font(14, True), anchor="w").pack(fill="x", padx=18, pady=(14, 4))
        box, self.ptree = self._tree(pc, ("code", "name", "qty", "amount"),
                                     ("รหัส", "สินค้า", "จำนวน", "ยอดเงิน"), (110, 420, 90, 130),
                                     {"qty": "e", "amount": "e"})
        self.ptree.configure(height=4)
        box.pack(fill="x", padx=8, pady=(0, 8))

        ec = card(body)
        ec.grid(row=4, column=0, columnspan=2, sticky="ew", padx=14, pady=(8, 20))
        ctk.CTkLabel(ec, text="ส่งออกเอกสารจากฐานข้อมูล (ตามช่วงเวลาที่เลือกด้านบน)", font=font(14, True),
                     anchor="w").pack(fill="x", padx=18, pady=(14, 2))
        ctk.CTkLabel(ec, text="ไม่ต้องเปิด PDF ซ้ำ: สร้างใบวางบิล/ใบส่งของจากข้อมูลที่บันทึกไว้", font=font(12),
                     text_color=MUTED, anchor="w").pack(fill="x", padx=18)
        eb = ctk.CTkFrame(ec, fg_color="transparent")
        eb.pack(fill="x", padx=18, pady=14)
        primary_btn(eb, "ส่งออก .xlsx…", lambda: self.export_db("xlsx"), width=170).pack(side="left")
        ghost_btn(eb, "ส่งออก PDF พร้อมพิมพ์…", lambda: self.export_db("pdf"), width=220).pack(side="left", padx=10)
        self.lbl_exp = ctk.CTkLabel(eb, text="", font=font(12), text_color=MUTED)
        self.lbl_exp.pack(side="left", padx=10)
        self._dash_data = {}

    def refresh_dash(self):
        ms = db.months(self.conn)
        labels = ["ทั้งหมด"] + [db.month_label(m) for m in ms]
        self._month_map = {db.month_label(m): m for m in ms}
        cur = self.dash_month.get()
        self.dash_month.configure(values=labels)
        self.dash_month.set(cur if cur in labels else "ทั้งหมด")
        self._dash_fill()

    def _dash_ym(self):
        return self._month_map.get(self.dash_month.get())

    def _dash_fill(self):
        ym = self._dash_ym()
        s = db.summary(self.conn)
        if not s["count"]:
            self.lbl_dash_hint.configure(text="ยังไม่มีข้อมูล: นำเข้า PDF แล้วกด 'บันทึกเข้าฐานข้อมูล' ที่หน้าสกัดเอกสาร")
        else:
            self.lbl_dash_hint.configure(text=f"ข้อมูล {db.iso_to_be(s['first'])} – {db.iso_to_be(s['last'])}   ·   "
                                                  "* เดือนที่นำเข้าข้อมูลไม่ครบ เทียบแล้วจะไม่แม่น")
        if ym:
            r = self.conn.execute("SELECT COUNT(*), COALESCE(SUM(net_satang),0) FROM invoices "
                                  "WHERE substr(date_iso,1,7)=?", (ym,)).fetchone()
            count, total = r[0], r[1] / 100
        else:
            count, total = s["count"], s["total"]
        self.kpi["total"].configure(text=f"{total:,.2f}")
        self.kpi["count"].configure(text=f"{count:,}")
        self.kpi["avg"].configure(text=f"{total / count:,.2f}" if count else "–")
        mo = db.monthly(self.conn, 24, fill=False)
        delta, color = "–", TEXT
        if ym:
            idx = [m for m, _v, _c in mo].index(ym) if ym in [m for m, _v, _c in mo] else -1
            if idx > 0 and mo[idx - 1][1]:
                pct = (mo[idx][1] - mo[idx - 1][1]) / mo[idx - 1][1] * 100
                delta, color = f"{pct:+.1f}%", OK_FG if pct >= 0 else BAD_FG
        elif len(mo) >= 2 and mo[-2][1]:
            pct = (mo[-1][1] - mo[-2][1]) / mo[-2][1] * 100
            delta, color = f"{pct:+.1f}%", OK_FG if pct >= 0 else BAD_FG
        self.kpi["delta"].configure(text=delta, text_color=color)
        names = {c: n for c, (n, _a) in self.branches.items()}
        self._dash_data = {
            "monthly": [(db.month_label(m).replace(" ", "\n"), v) for m, v, _c in db.monthly(self.conn, 12)],
            "branch": [(names.get(c, c).replace("สาขา", "", 1), v) for c, v, _n in db.by_branch(self.conn, ym)[:8]],
        }
        self.ptree.delete(*self.ptree.get_children())
        for code, name, qty, amt in db.by_product(self.conn, ym):
            self.ptree.insert("", "end", values=(code, name, f"{qty:g}", f"{amt:,.2f}"))
        self._dash_draw()

    def _dash_draw(self):
        d = self._dash_data
        if not d:
            return
        self._draw_vbars(self.cv_month, d["monthly"])
        self._draw_hbars(self.cv_branch, d["branch"])

    @staticmethod
    def _fmt(v):
        return f"{v / 1000:.1f}k" if v >= 1000 else f"{v:,.0f}"

    def _draw_vbars(self, cv, data):
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if not data or w < 60:
            cv.create_text(w / 2, h / 2, text="ยังไม่มีข้อมูล", fill=MUTED, font=(FAMILY, 11))
            return
        top, bot = 26, 44
        mx = max(v for _l, v in data) or 1
        slot = (w - 20) / len(data)
        bw = min(46, slot * 0.6)
        for i, (lab, v) in enumerate(data):
            x0 = 10 + slot * i + (slot - bw) / 2
            bh = (h - top - bot) * v / mx
            y1 = h - bot
            cv.create_rectangle(x0, y1 - bh, x0 + bw, y1, fill=ACCENT if i == len(data) - 1 else "#3C4690", width=0)
            cv.create_text(x0 + bw / 2, y1 - bh - 10, text=self._fmt(v), fill=TEXT, font=(FAMILY, 9))
            cv.create_text(x0 + bw / 2, y1 + 18, text=lab, fill=MUTED, font=(FAMILY, 9), justify="center")

    def _draw_hbars(self, cv, data):
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if not data or w < 60:
            cv.create_text(w / 2, h / 2, text="ยังไม่มีข้อมูล", fill=MUTED, font=(FAMILY, 11))
            return
        mx = max(v for _l, v in data) or 1
        row = (h - 10) / len(data)
        lab_w = 150
        for i, (lab, v) in enumerate(data):
            y = 5 + row * i + row / 2
            cv.create_text(lab_w - 8, y, text=lab if len(lab) < 22 else lab[:21] + "…", anchor="e", fill=MUTED,
                           font=(FAMILY, 9))
            bwid = (w - lab_w - 70) * v / mx
            cv.create_rectangle(lab_w, y - 8, lab_w + max(bwid, 2), y + 8, fill=ACCENT if i == 0 else "#3C4690", width=0)
            cv.create_text(lab_w + bwid + 8, y, text=self._fmt(v), anchor="w", fill=TEXT, font=(FAMILY, 9))

    def export_db(self, kind):
        if not self.cfg_ok:
            self.show("config")
            return
        ym = self._dash_ym()
        invs = db.load_invoices(self.conn, ym)
        if not invs:
            messagebox.showinfo("ไม่มีข้อมูล", "ยังไม่มีใบรับสินค้าในช่วงเวลานี้ นำเข้า PDF ก่อน")
            return
        for i in invs:  # ใช้ชื่อสาขาปัจจุบันจากตารางสาขา
            if i["wh_code"] not in self.branches:
                self.branches[i["wh_code"]] = (i["branch"], "")
        if not self._confirm_missing(invs):
            return
        tag = ym or "ทั้งหมด"
        ext = ".xlsx" if kind == "xlsx" else ".pdf"
        out = filedialog.asksaveasfilename(
            title="บันทึกไฟล์", initialfile=f"ใบรับสินค้า-{tag}{ext}", defaultextension=ext,
            filetypes=[("Excel", "*.xlsx")] if kind == "xlsx" else [("PDF", "*.pdf")])
        if not out:
            return
        core.save_branches(self.branches)
        self.lbl_exp.configure(text="กำลังสร้างไฟล์…", text_color=MUTED)
        self.update_idletasks()
        try:
            if kind == "xlsx":
                core.build(invs, out)
                files = [out]
            else:
                files = pdf_export.export_pdfs(invs, out)
        except PermissionError:
            messagebox.showerror("บันทึกไม่ได้", "ไฟล์นี้เปิดค้างอยู่ ปิดก่อนหรือเลือกชื่อใหม่")
            self.lbl_exp.configure(text="")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("สร้างไฟล์ไม่สำเร็จ", f"{type(e).__name__}: {e}")
            self.lbl_exp.configure(text="")
            return
        self.lbl_exp.configure(text=f"สร้างแล้ว {len(invs)} ใบ", text_color=OK_FG)
        if messagebox.askyesno("เสร็จแล้ว", "สร้างไฟล์เรียบร้อย\n" + "\n".join(files) + "\n\nเปิดโฟลเดอร์ไหม?"):
            os.startfile(Path(files[0]).parent)

    # ------------------------------------------------------------ branches page
    def _build_branches(self):
        p = self._page("branches", "สาขา", "ชื่อและที่อยู่สาขา")
        p.rowconfigure(2, weight=1)
        tc = card(p)
        tc.grid(row=2, column=0, sticky="nsew", padx=32, pady=(0, 10))
        tc.rowconfigure(0, weight=1)
        tc.columnconfigure(0, weight=1)
        box, self.btree = self._tree(tc, ("code", "name", "addr"), ("รหัสคลัง", "ชื่อสาขา", "ที่อยู่"),
                                     (90, 320, 600))
        box.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self.btree.tag_configure("warn", background=WARN_BG, foreground=WARN_FG)
        self.btree.bind("<<TreeviewSelect>>", self._pick_branch)

        fc = card(p)
        fc.grid(row=3, column=0, sticky="ew", padx=32, pady=(0, 8))
        fc.columnconfigure(1, weight=1)
        self.v_code, self.v_name, self.v_addr = tk.StringVar(), tk.StringVar(), tk.StringVar()
        for r, (lab, var, w) in enumerate((("รหัสคลัง", self.v_code, 140), ("ชื่อสาขา", self.v_name, 0),
                                           ("ที่อยู่", self.v_addr, 0))):
            ctk.CTkLabel(fc, text=lab, font=font(12), text_color=MUTED, anchor="w", width=80).grid(
                row=r, column=0, sticky="w", padx=(20, 8), pady=(14 if r == 0 else 4, 4 if r < 2 else 14))
            e = entry(fc, var, **({"width": w} if w else {}))
            e.grid(row=r, column=1, sticky="w" if w else "ew", padx=(0, 20),
                   pady=(14 if r == 0 else 4, 4 if r < 2 else 14))

        bt = ctk.CTkFrame(p, fg_color="transparent")
        bt.grid(row=4, column=0, sticky="ew", padx=32, pady=(0, 22))
        ghost_btn(bt, "เพิ่ม / อัปเดตแถว", self._upsert_branch, width=160).pack(side="left")
        ghost_btn(bt, "ลบแถว", self._delete_branch, width=100).pack(side="left", padx=8)
        primary_btn(bt, "บันทึก", self._save_branches, width=130).pack(side="right")
        self.lbl_br = ctk.CTkLabel(bt, text="", font=font(12), text_color=MUTED)
        self.lbl_br.pack(side="right", padx=14)
        self._refresh_branch_tree()

    def _refresh_branch_tree(self):
        self.btree.delete(*self.btree.get_children())
        for code, (name, addr) in self.branches.items():
            self.btree.insert("", "end", iid=code, tags=("" if addr.strip() else "warn",), values=(code, name, addr))

    def _pick_branch(self, _evt=None):
        sel = self.btree.selection()
        if sel:
            name, addr = self.branches[sel[0]]
            self.v_code.set(sel[0])
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
        self.lbl_br.configure(text="มีการแก้ไขที่ยังไม่ได้บันทึก", text_color=WARN_FG)
        if self.invs:
            self._refresh_preview()

    def _delete_branch(self):
        sel = self.btree.selection()
        if sel and messagebox.askyesno("ลบสาขา", f"ลบสาขารหัส {sel[0]} ?"):
            self.branches.pop(sel[0], None)
            self._refresh_branch_tree()
            self.lbl_br.configure(text="มีการแก้ไขที่ยังไม่ได้บันทึก", text_color=WARN_FG)

    def _save_branches(self):
        core.save_branches(self.branches)
        self.lbl_br.configure(text="บันทึกแล้ว", text_color=OK_FG)

    # ------------------------------------------------------------ config page
    def _build_config(self):
        p = self._page("config", "ตั้งค่า", "ข้อมูลผู้ส่งของ")
        c = card(p)
        c.grid(row=2, column=0, sticky="ew", padx=32, pady=(0, 10))
        c.columnconfigure(1, weight=1)
        ctk.CTkLabel(c, text="ข้อมูลนี้จะพิมพ์ที่หัวใบส่งของ  ·  เก็บในไฟล์ config.json ข้างโปรแกรม",
                     font=font(12), text_color=MUTED, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(16, 6))
        self.cfg_vars = {}
        for r, (key, lab) in enumerate(CFG_LABELS, 1):
            ctk.CTkLabel(c, text=lab, font=font(12), text_color=MUTED, anchor="w", width=230).grid(
                row=r, column=0, sticky="w", padx=(20, 8), pady=6)
            v = tk.StringVar()
            entry(c, v).grid(row=r, column=1, sticky="ew", padx=(0, 20), pady=6)
            self.cfg_vars[key] = v
        ctk.CTkFrame(c, fg_color="transparent", height=10).grid(row=len(CFG_LABELS) + 1, column=0)
        row = ctk.CTkFrame(p, fg_color="transparent")
        row.grid(row=3, column=0, sticky="ew", padx=32)
        primary_btn(row, "บันทึกการตั้งค่า", self._save_cfg, width=170).pack(side="right")
        self.lbl_cfg = ctk.CTkLabel(row, text="", font=font(12), text_color=MUTED)
        self.lbl_cfg.pack(side="left")

    def _load_cfg(self, initial=False):
        try:
            cfg = core.load_config()
        except core.ConfigError as e:
            self.cfg_ok = False
            self.lbl_cfg.configure(text=str(e), text_color=BAD_FG)
            if initial:
                self.show("config")
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
        self.lbl_cfg.configure(text="โหลดการตั้งค่าแล้ว", text_color=MUTED)

    def _save_cfg(self):
        cfg = {k: v.get().strip() for k, v in self.cfg_vars.items()}
        empty = [lab for k, lab in CFG_LABELS if not cfg[k]]
        if empty:
            messagebox.showwarning("ข้อมูลไม่ครบ", "ยังไม่ได้กรอก: " + ", ".join(empty))
            return
        core.save_config(cfg)
        core.apply_config(cfg)
        self.cfg_ok = True
        self.lbl_cfg.configure(text="บันทึกแล้ว", text_color=OK_FG)
        self.lbl_status.configure(text="")
        if self.invs:
            self.btn_make.configure(state="normal")
            self.btn_pdf.configure(state="normal")


if __name__ == "__main__":
    App().mainloop()

