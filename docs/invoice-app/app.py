# -*- coding: utf-8 -*-
"""Office Go Plan - 書類管理アプリ（tkinter GUI）"""
from __future__ import annotations

import datetime
import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk

import config
import db
import docgen

# ---- ユーティリティ ----

def _fmt_date(d) -> str:
    if d is None: return ""
    if isinstance(d, datetime.date): return d.strftime("%Y-%m-%d")
    s = str(d)
    return s[:10] if s else ""


def _fmt_money(n) -> str:
    try:   return f"¥{int(n):,}"
    except: return ""


def open_path(path: str) -> None:
    if not path or not os.path.exists(path):
        messagebox.showwarning("ファイルが見つかりません", path or "パスが未設定です")
        return
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


# ---- メインウィンドウ ----

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Go Plan  書類管理")
        self.geometry("1200x700")
        self.minsize(960, 520)
        self._all_txns: list[dict] = []
        self._current_txn: dict | None = None
        self._setup_style()
        self._build_ui()
        self._refresh()

    # ---- スタイル ----

    def _setup_style(self) -> None:
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",             font=("Meiryo UI", 10))
        s.configure("H.TLabel",      font=("Meiryo UI", 11, "bold"))
        s.configure("Muted.TLabel",  font=("Meiryo UI", 9),  foreground="#6B7280")
        s.configure("Bold.TLabel",   font=("Meiryo UI", 10, "bold"))
        s.configure("Issued.TLabel", font=("Meiryo UI", 9,  "bold"), foreground="#059669")
        s.configure("Treeview",      rowheight=22, font=("Meiryo UI", 9))
        s.configure("Treeview.Heading", font=("Meiryo UI", 9, "bold"))
        s.map("Treeview", background=[("selected", "#3B82F6")],
              foreground=[("selected", "white")])

    # ---- UI 構築 ----

    def _build_ui(self) -> None:
        # ツールバー
        tb = tk.Frame(self, bg="#1A2332", pady=7, padx=10)
        tb.pack(fill="x")
        tk.Label(tb, text="Office Go Plan  書類管理",
                 bg="#1A2332", fg="white",
                 font=("Meiryo UI", 12, "bold")).pack(side="left")
        for text, cmd in [
            ("📊 Excel を開く",   lambda: open_path(config.EXCEL_PATH)),
            ("📂 出力フォルダ",   lambda: open_path(config.OUTPUT_DIR)),
            ("↻ 更新",           self._refresh),
            ("＋ 新規取引",       self._new_txn),
        ]:
            tk.Button(tb, text=text, command=cmd,
                      bg="#2D3F55", fg="white", relief="flat",
                      font=("Meiryo UI", 9), padx=8, pady=2,
                      cursor="hand2").pack(side="right", padx=3)

        # メイン: PanedWindow
        pw = ttk.PanedWindow(self, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=4, pady=4)

        left = ttk.Frame(pw)
        pw.add(left, weight=1)
        self._build_list(left)

        right = ttk.Frame(pw)
        pw.add(right, weight=2)
        self._build_detail_container(right)

    def _build_list(self, parent) -> None:
        ttk.Label(parent, text="取引一覧", style="H.TLabel").pack(
            anchor="w", padx=6, pady=(4, 2))

        # 検索バー
        ff = ttk.Frame(parent)
        ff.pack(fill="x", padx=4, pady=2)
        ttk.Label(ff, text="検索:").pack(side="left")
        self._filter_var = tk.StringVar()
        self._filter_var.trace_add("write", lambda *_: self._apply_filter())
        ttk.Entry(ff, textvariable=self._filter_var, width=20).pack(
            side="left", padx=4)

        # ステータスフィルタ
        self._status_var = tk.StringVar(value="すべて")
        status_cb = ttk.Combobox(ff, textvariable=self._status_var, width=12,
                                 state="readonly",
                                 values=["すべて"] + list(db.STATUS.values()))
        status_cb.pack(side="left", padx=4)
        status_cb.bind("<<ComboboxSelected>>", lambda *_: self._apply_filter())

        # Treeview
        cols = ("取引ID", "顧客名", "金額", "ステータス", "作成日")
        self._tree = ttk.Treeview(parent, columns=cols, show="headings",
                                  selectmode="browse")
        for c, w in zip(cols, [145, 145, 80, 95, 90]):
            self._tree.heading(c, text=c,
                               command=lambda _c=c: self._sort_by(_c))
            self._tree.column(c, width=w, minwidth=60)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True, padx=4, pady=2)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        self._tree.tag_configure("alt", background="#F8FAFC")

    def _build_detail_container(self, parent) -> None:
        canvas = tk.Canvas(parent, highlightthickness=0, bg="#FFFFFF")
        vsb    = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)
        self._detail_inner = ttk.Frame(canvas, padding=12)
        self._detail_win   = canvas.create_window((0, 0), window=self._detail_inner,
                                                  anchor="nw")
        self._detail_inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfig(self._detail_win, width=e.width))
        self._render_detail(None)

    # ---- 詳細パネル描画 ----

    def _render_detail(self, txn: dict | None) -> None:
        for w in self._detail_inner.winfo_children():
            w.destroy()
        if txn is None:
            ttk.Label(self._detail_inner, text="← 取引を選択してください",
                      style="Muted.TLabel").pack(pady=30)
            return

        f = self._detail_inner
        f.columnconfigure(0, weight=1)

        # ── 顧客情報 ──
        ttk.Label(f, text="顧客情報", style="H.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 4))

        cf = ttk.LabelFrame(f, text="", padding=8)
        cf.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for r, (lbl, val) in enumerate([
            ("取引ID",  txn.get("取引ID", "")),
            ("作成日",  _fmt_date(txn.get("作成日"))),
            ("顧客名",  txn.get("顧客名", "")),
            ("住　所",  txn.get("顧客住所", "")),
        ]):
            ttk.Label(cf, text=f"{lbl}:", style="Muted.TLabel",
                      width=8, anchor="e").grid(row=r, column=0, sticky="e", padx=(0,6))
            ttk.Label(cf, text=val).grid(row=r, column=1, sticky="w")

        # ── 品目 ──
        ttk.Label(f, text="品目", style="H.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, 4))

        pf = ttk.LabelFrame(f, text="", padding=8)
        pf.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        cols = ("品目", "数量", "単価", "金額")
        it = ttk.Treeview(pf, columns=cols, show="headings", height=3,
                          selectmode="none")
        it.column("品目", width=230); it.column("数量", width=50, anchor="e")
        it.column("単価", width=80, anchor="e"); it.column("金額", width=80, anchor="e")
        for c in cols: it.heading(c, text=c)

        qty    = int(txn.get("数量", 1) or 1)
        unit   = int(txn.get("単価", 0) or 0)
        amount = int(txn.get("金額", 0) or 0)
        it.insert("", "end", values=(
            txn.get("品目", ""), f"{qty:,}", f"¥{unit:,}", f"¥{amount:,}"))
        it.pack(fill="x")
        ttk.Label(pf, text=f"合計:  ¥{amount:,}  （消費税不課税）",
                  style="Bold.TLabel").pack(anchor="e", pady=(4, 0))

        # ── 書類 ──
        ttk.Label(f, text="書類", style="H.TLabel").grid(
            row=4, column=0, sticky="w", pady=(0, 4))

        df = ttk.LabelFrame(f, text="", padding=10)
        df.grid(row=5, column=0, sticky="ew", pady=(0, 8))
        df.columnconfigure(0, weight=1)

        self._doc_row(df, 0, "見積書", "見積番号", "見積日",
                      docgen.DocType.QUOTE, txn)
        ttk.Separator(df, orient="horizontal").grid(
            row=1, column=0, columnspan=6, sticky="ew", pady=4)
        self._doc_row(df, 2, "請求書", "請求番号", "請求日",
                      docgen.DocType.INVOICE, txn)
        ttk.Separator(df, orient="horizontal").grid(
            row=3, column=0, columnspan=6, sticky="ew", pady=4)
        self._doc_row(df, 4, "領収書", "領収番号", "領収日",
                      docgen.DocType.RECEIPT, txn)

        # ── ステータス ──
        sf = ttk.LabelFrame(f, text="ステータス / 備考", padding=8)
        sf.grid(row=6, column=0, sticky="ew")
        status = txn.get("ステータス", "新規")
        ttk.Label(sf, text=f"状態: {status}", style="Bold.TLabel").grid(
            row=0, column=0, sticky="w")
        notes = txn.get("備考") or ""
        if notes:
            ttk.Label(sf, text=f"備考: {notes}", style="Muted.TLabel",
                      wraplength=460).grid(row=1, column=0, sticky="w", pady=(4, 0))

        # ── 削除 ──
        ttk.Separator(f, orient="horizontal").grid(
            row=7, column=0, sticky="ew", pady=8)
        ttk.Button(f, text="この取引を削除",
                   command=lambda t=txn: self._delete_txn(t)).grid(
            row=8, column=0, sticky="e")

    def _doc_row(self, parent, row, label, num_key, date_key, doc_type, txn) -> None:
        doc_num  = (txn.get(num_key) or "").strip()
        doc_date = txn.get(date_key)
        issued   = bool(doc_num)
        lbl_map  = {"見積書": "quotes", "請求書": "invoices", "領収書": "receipts"}

        ttk.Label(parent, text=label, style="Bold.TLabel",
                  width=6).grid(row=row, column=0, sticky="w", padx=(0, 10))

        if issued:
            ttk.Label(parent, text=doc_num,
                      style="Issued.TLabel").grid(row=row, column=1, sticky="w", padx=(0, 8))
            ttk.Label(parent, text=_fmt_date(doc_date),
                      style="Muted.TLabel").grid(row=row, column=2, sticky="w", padx=(0, 8))
            subdir  = lbl_map[label]
            pdf_p   = os.path.join(config.OUTPUT_DIR, subdir, f"{doc_num}.pdf")
            docx_p  = os.path.join(config.OUTPUT_DIR, subdir, f"{doc_num}.docx")
            open_p  = pdf_p if os.path.exists(pdf_p) else docx_p
            ext_lbl = "PDF" if os.path.exists(pdf_p) else "Word"
            if os.path.exists(open_p):
                ttk.Button(parent, text=f"▶ {ext_lbl}",
                           command=lambda p=open_p: open_path(p)).grid(
                    row=row, column=3, padx=2)
            ttk.Button(parent, text="再発行",
                       command=lambda dt=doc_type, t=txn:
                           self._issue_doc(dt, t, reissue=True)).grid(
                row=row, column=4, padx=2)
        else:
            ttk.Label(parent, text="未発行",
                      style="Muted.TLabel").grid(row=row, column=1, sticky="w", padx=(0, 8))
            ttk.Button(parent, text=f"  {label}を発行  ",
                       command=lambda dt=doc_type, t=txn:
                           self._issue_doc(dt, t)).grid(
                row=row, column=2, columnspan=2, padx=2)

    # ---- データ操作 ----

    def _refresh(self) -> None:
        self._all_txns = db.get_all()
        self._apply_filter()
        if self._current_txn:
            updated = db.get_by_id(self._current_txn["取引ID"])
            self._render_detail(updated)

    def _apply_filter(self) -> None:
        q      = self._filter_var.get().lower()
        status = self._status_var.get()
        rows   = [
            t for t in self._all_txns
            if (not q or q in str(t.get("顧客名","")).lower()
                      or q in str(t.get("取引ID","")).lower()
                      or q in str(t.get("品目","")).lower())
            and (status == "すべて" or t.get("ステータス") == status)
        ]
        self._tree.delete(*self._tree.get_children())
        for i, t in enumerate(rows):
            self._tree.insert("", "end", iid=t["取引ID"],
                              tags=("alt",) if i % 2 else (),
                              values=(t.get("取引ID",""), t.get("顧客名",""),
                                      _fmt_money(t.get("金額",0)),
                                      t.get("ステータス",""),
                                      _fmt_date(t.get("作成日"))))

    def _sort_by(self, col: str) -> None:
        items = [(self._tree.set(k, col), k) for k in self._tree.get_children()]
        items.sort()
        for idx, (_, k) in enumerate(items):
            self._tree.move(k, "", idx)

    def _on_select(self, _=None) -> None:
        sel = self._tree.selection()
        if not sel: return
        self._current_txn = db.get_by_id(sel[0])
        self._render_detail(self._current_txn)

    # ---- ダイアログ起動 ----

    def _new_txn(self) -> None:
        dlg = NewTransactionDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            txn_id = db.create_transaction(**dlg.result)
            self._refresh()
            try:
                self._tree.selection_set(txn_id)
                self._tree.see(txn_id)
                self._on_select()
            except Exception:
                pass

    def _issue_doc(self, doc_type: docgen.DocType, txn: dict,
                   reissue: bool = False) -> None:
        dlg = IssueDialog(self, doc_type, txn, reissue=reissue)
        self.wait_window(dlg)
        if dlg.result:
            self._refresh()

    def _delete_txn(self, txn: dict) -> None:
        if not messagebox.askyesno(
                "削除確認",
                f"取引 {txn['取引ID']} を削除しますか？\n"
                "（Excel の行を削除します。生成済みファイルは残ります）",
                parent=self):
            return
        wb = db._load()
        ws = wb["取引管理"]
        for row in ws.iter_rows(min_row=2):
            if row[0].value == txn["取引ID"]:
                ws.delete_rows(row[0].row)
                break
        wb.save(config.EXCEL_PATH)
        self._current_txn = None
        self._render_detail(None)
        self._refresh()


# ---- 新規取引ダイアログ ----

class NewTransactionDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("新規取引")
        self.resizable(False, False)
        self.result = None
        self._build()
        self.grab_set(); self.transient(parent)
        self.focus_force()

    def _build(self) -> None:
        pad = {"padx": 8, "pady": 4}
        ttk.Label(self, text="新規取引", font=("Meiryo UI", 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))

        # 顧客情報
        cf = ttk.LabelFrame(self, text="顧客情報", padding=8)
        cf.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=4)
        ttk.Label(cf, text="顧客名 *").grid(row=0, column=0, sticky="e", **pad)
        self._name = tk.StringVar()
        ttk.Entry(cf, textvariable=self._name, width=34).grid(
            row=0, column=1, sticky="w", **pad)
        ttk.Label(cf, text="住　所 *").grid(row=1, column=0, sticky="e", **pad)
        self._addr = tk.StringVar()
        ttk.Entry(cf, textvariable=self._addr, width=42).grid(
            row=1, column=1, sticky="w", **pad)

        # 品目
        pf = ttk.LabelFrame(self, text="品目", padding=8)
        pf.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=4)
        for c, (h, w) in enumerate([("品目名", 26), ("数量", 7), ("単価（円）", 12)]):
            ttk.Label(pf, text=h, font=("Meiryo UI", 9, "bold")).grid(
                row=0, column=c, padx=4)

        self._desc  = tk.StringVar(value=config.PRODUCT_DEFAULT)
        self._qty   = tk.StringVar(value="1")
        self._unit  = tk.StringVar(value=str(config.UNIT_PRICE_DEFAULT))
        ttk.Entry(pf, textvariable=self._desc,  width=28).grid(row=1, column=0, padx=4, pady=2)
        ttk.Entry(pf, textvariable=self._qty,   width=8 ).grid(row=1, column=1, padx=4, pady=2)
        ttk.Entry(pf, textvariable=self._unit,  width=14).grid(row=1, column=2, padx=4, pady=2)
        self._amt_lbl = ttk.Label(pf, text="金額:  ¥5,000", style="Bold.TLabel")
        self._amt_lbl.grid(row=2, column=0, columnspan=3, sticky="e", pady=(4, 0))
        for v in (self._qty, self._unit):
            v.trace_add("write", lambda *_: self._update_amount())

        # 備考
        nf = ttk.LabelFrame(self, text="備考", padding=8)
        nf.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=4)
        self._notes = tk.StringVar()
        ttk.Entry(nf, textvariable=self._notes, width=52).grid(
            row=0, column=0, sticky="ew", padx=4)

        # ボタン
        bf = ttk.Frame(self)
        bf.grid(row=4, column=0, columnspan=2, pady=12)
        ttk.Button(bf, text="キャンセル", command=self.destroy).pack(side="left", padx=8)
        ttk.Button(bf, text="  作成  ",   command=self._submit ).pack(side="left", padx=8)

    def _update_amount(self) -> None:
        try:
            qty  = int(self._qty.get())
            unit = int(self._unit.get().replace(",", ""))
            self._amt_lbl.config(text=f"金額:  ¥{qty * unit:,}")
        except ValueError:
            self._amt_lbl.config(text="金額:  ---")

    def _submit(self) -> None:
        name = self._name.get().strip()
        addr = self._addr.get().strip()
        if not name or not addr:
            messagebox.showwarning("入力エラー", "顧客名と住所は必須です。", parent=self)
            return
        try:
            qty  = int(self._qty.get())
            unit = int(self._unit.get().replace(",", ""))
        except ValueError:
            messagebox.showwarning("入力エラー", "数量・単価は整数で入力してください。", parent=self)
            return
        self.result = {
            "customer_name":    name,
            "customer_address": addr,
            "items": [{"description": self._desc.get(), "qty": qty, "unit_price": unit}],
            "notes": self._notes.get().strip(),
        }
        self.destroy()


# ---- 書類発行ダイアログ ----

class IssueDialog(tk.Toplevel):
    _LABELS = {
        docgen.DocType.QUOTE:   ("見積書", "Q"),
        docgen.DocType.INVOICE: ("請求書", "I"),
        docgen.DocType.RECEIPT: ("領収書", "R"),
    }
    _NUM_KEY  = {"見積書": "見積番号", "請求書": "請求番号", "領収書": "領収番号"}
    _DATE_KEY = {"見積書": "見積日",   "請求書": "請求日",   "領収書": "領収日"}

    def __init__(self, parent, doc_type: docgen.DocType, txn: dict,
                 reissue: bool = False):
        super().__init__(parent)
        self.doc_type = doc_type
        self.txn      = txn
        self.reissue  = reissue
        self.result   = None
        lbl, _        = self._LABELS[doc_type]
        self.title(f"{'再' if reissue else ''}{lbl}を発行")
        self.resizable(False, False)
        self._build()
        self.grab_set(); self.transient(parent)
        self.focus_force()

    def _build(self) -> None:
        pad   = {"padx": 8, "pady": 4}
        today = datetime.date.today()
        lbl, _ = self._LABELS[self.doc_type]
        ttk.Label(self, text=f"{lbl}を発行", font=("Meiryo UI", 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))

        # 発行情報
        inf = ttk.LabelFrame(self, text="発行情報", padding=8)
        inf.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=4)
        ttk.Label(inf, text="発行日:").grid(row=0, column=0, sticky="e", **pad)
        self._issue = tk.StringVar(value=today.strftime("%Y-%m-%d"))
        ttk.Entry(inf, textvariable=self._issue, width=14).grid(
            row=0, column=1, sticky="w", **pad)
        ttk.Label(inf, text="（YYYY-MM-DD）",
                  foreground="#6B7280").grid(row=0, column=2, sticky="w")

        self._due = self._valid = None
        if self.doc_type == docgen.DocType.INVOICE:
            due = today + datetime.timedelta(days=30)
            ttk.Label(inf, text="入金期限:").grid(row=1, column=0, sticky="e", **pad)
            self._due = tk.StringVar(value=due.strftime("%Y-%m-%d"))
            ttk.Entry(inf, textvariable=self._due, width=14).grid(
                row=1, column=1, sticky="w", **pad)
        if self.doc_type == docgen.DocType.QUOTE:
            valid = today + datetime.timedelta(days=30)
            ttk.Label(inf, text="有効期限:").grid(row=1, column=0, sticky="e", **pad)
            self._valid = tk.StringVar(value=valid.strftime("%Y-%m-%d"))
            ttk.Entry(inf, textvariable=self._valid, width=14).grid(
                row=1, column=1, sticky="w", **pad)

        # 発行内容サマリ
        sf = ttk.LabelFrame(self, text="発行内容", padding=8)
        sf.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=4)
        qty    = int(self.txn.get("数量", 1) or 1)
        unit   = int(self.txn.get("単価", 0) or 0)
        amount = qty * unit
        for r, (k, v) in enumerate([
            ("顧客",             self.txn.get("顧客名", "")),
            ("品目",             self.txn.get("品目", "")),
            ("数量 / 単価 / 金額",
             f"{qty:,}  /  ¥{unit:,}  /  ¥{amount:,}"),
        ]):
            ttk.Label(sf, text=f"{k}:", foreground="#6B7280",
                      width=18, anchor="e").grid(row=r, column=0, sticky="e", padx=4)
            ttk.Label(sf, text=v).grid(row=r, column=1, sticky="w", padx=4)

        # ボタン
        bf = ttk.Frame(self)
        bf.grid(row=3, column=0, columnspan=2, pady=12)
        ttk.Button(bf, text="キャンセル", command=self.destroy).pack(side="left", padx=8)
        ttk.Button(bf, text="  発行  ",   command=self._submit ).pack(side="left", padx=8)

    @staticmethod
    def _parse_date(s: str) -> datetime.date:
        try:
            return datetime.date.fromisoformat(s.strip())
        except ValueError:
            raise ValueError(f"日付形式が正しくありません: {s}（YYYY-MM-DD）")

    def _submit(self) -> None:
        try:
            issue_date = self._parse_date(self._issue.get())
            due_date   = self._parse_date(self._due.get())   if self._due   else None
            valid_until= self._parse_date(self._valid.get()) if self._valid else None
        except ValueError as e:
            messagebox.showwarning("入力エラー", str(e), parent=self)
            return

        lbl, prefix = self._LABELS[self.doc_type]
        num_key     = self._NUM_KEY[lbl]
        existing    = (self.txn.get(num_key) or "").strip()
        doc_number  = existing if (existing and not self.reissue) else db.next_number(prefix)

        qty  = int(self.txn.get("数量", 1) or 1)
        unit = int(self.txn.get("単価", 0) or 0)
        params = docgen.DocumentParams(
            doc_type         = self.doc_type,
            doc_number       = doc_number,
            issue_date       = issue_date,
            customer_name    = self.txn.get("顧客名", ""),
            customer_address = self.txn.get("顧客住所", ""),
            items            = [docgen.LineItem(self.txn.get("品目", ""), qty, unit)],
            due_date         = due_date,
            valid_until      = valid_until,
        )

        try:
            result = docgen.generate(params)
        except Exception as e:
            messagebox.showerror("生成エラー", str(e), parent=self)
            return

        # Excel 更新
        date_key    = self._DATE_KEY[lbl]
        status_map  = {
            docgen.DocType.QUOTE:   db.STATUS["quoted"],
            docgen.DocType.INVOICE: db.STATUS["invoiced"],
            docgen.DocType.RECEIPT: db.STATUS["completed"],
        }
        upd = {num_key: doc_number, date_key: issue_date,
               "ステータス": status_map[self.doc_type]}
        if self.doc_type == docgen.DocType.INVOICE and due_date:
            upd["入金期限"] = due_date
        db.update(self.txn["取引ID"], **upd)

        # ファイルを開く
        path = result.get("pdf") or result.get("docx")
        if path:
            open_path(path)

        ext = "PDF" if result.get("pdf") else "Word（PDF 変換不可）"
        messagebox.showinfo("発行完了",
                            f"{lbl} {doc_number} を発行しました。\n"
                            f"出力形式: {ext}", parent=self)
        self.result = result
        self.destroy()


# ---- エントリポイント ----

if __name__ == "__main__":
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    App().mainloop()
