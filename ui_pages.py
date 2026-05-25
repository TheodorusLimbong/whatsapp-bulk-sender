import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import requests
import pywhatkit as kit
import threading
import time
import random
import re
import os
import json
from utils import normalize_phone
from logic_worker import thread_safe_askstring, thread_safe_update_label
from i18n import lang

WINDOW_W = 1366
WINDOW_H = 768


class WhatsAppBulkSenderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WhatsApp Bulk Sender")
        x = (self.winfo_screenwidth() - WINDOW_W) // 2
        y = (self.winfo_screenheight() - WINDOW_H) // 2
        self.geometry(f"{WINDOW_W}x{WINDOW_H}+{x}+{y}")
        self.minsize(1000, 600)

        self.df = None
        self.excel_path = None
        self.selected_source = None
        self.sending_thread = None
        self.last_page = None
        self.stop_event = threading.Event()

        self.success = 0
        self.fail = 0
        self.skipped = 0

        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.page_welcome = None
        self.page_source = None
        self.page_main = None

        self._build_pages()

        self._load_gsheet_config()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self._save_gsheet_config()
        self.destroy()

    def _load_gsheet_config(self):
        self._save_gsheet_config()
        try:
            if os.path.exists("config.json"):
                with open("config.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.entry_api_key.delete(0, tk.END)
                    self.entry_api_key.insert(0, data.get("api_key", ""))
                    self.entry_spreadsheet_id.delete(0, tk.END)
                    self.entry_spreadsheet_id.insert(0, data.get("spreadsheet_id", ""))
        except Exception as e:
            self.log(f"Config load error: {e}")

    def _save_gsheet_config(self):
        try:
            data = {
                "api_key": self.entry_api_key.get().strip(),
                "spreadsheet_id": self.entry_spreadsheet_id.get().strip()
            }
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            self.log(f"Config save error: {e}")

    # ---------------------------
    # Pages: build and navigation
    # ---------------------------
    def _build_pages(self):
        if self.page_welcome:
            self.page_welcome.destroy()
            self.page_source.destroy()
            self.page_main.destroy()

        self.page_welcome = tk.Frame(self.container, bg="white")
        self.page_source = tk.Frame(self.container, bg="white")
        self.page_main = tk.Frame(self.container)

        for p in (self.page_welcome, self.page_source, self.page_main):
            p.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._build_welcome_page()
        self._build_source_page()
        self._build_main_page()
        self.show_welcome()

    def show_welcome(self):
        self.last_page = self.page_welcome
        self.page_welcome.lift()

    def show_source(self):
        self.last_page = self.page_source
        self.page_source.lift()

    def show_main(self):
        self.last_page = self.page_main
        self._refresh_main_source_ui()
        self.page_main.lift()

    def safe_show_page(self, page):
        if page:
            self.after(0, lambda: page.lift())

    def _switch_language(self, new_lang):
        lang.lang = new_lang
        self._build_pages()

    # ---------------------------
    # Welcome page
    # ---------------------------
    def _build_welcome_page(self):
        p = self.page_welcome
        header = tk.Frame(p, height=60, bg="white")
        header.pack(fill="x")

        body = tk.Frame(p, bg="white")
        body.pack(fill="both", expand=True)

        title_size = min(86, max(36, WINDOW_H // 12))
        lbl_title = tk.Label(body, text=lang.tr("welcome_title"), font=("Segoe UI", title_size, "bold"), bg="white")
        lbl_title.pack(pady=(40, 0))

        lbl_sub = tk.Label(body, text=lang.tr("welcome_subtitle"), font=("Segoe UI", 28), bg="white", justify="center")
        lbl_sub.pack(pady=(8, 20))

        lang_frame = tk.Frame(body, bg="white")
        lang_frame.pack(pady=10)
        tk.Label(lang_frame, text="Language / Bahasa:", font=("Segoe UI", 12), bg="white").pack(side="left", padx=(0, 10))
        lang_combo = ttk.Combobox(lang_frame, values=lang.get_lang_options(), state="readonly", width=10)
        lang_combo.set(lang.lang)
        lang_combo.pack(side="left")
        lang_combo.bind("<<ComboboxSelected>>", lambda e: self._switch_language(lang_combo.get()))

        btn_next = ttk.Button(
            body,
            text=lang.tr("btn_next"),
            command=self.show_source
        )
        btn_next.pack(pady=20, ipadx=25, ipady=22)
        btn_next.configure(width=30)
        btn_next.focus_set()
        self.bind('<Return>', lambda e: self.show_source())

        footer = tk.Frame(p, height=50, bg="white")
        footer.pack(side="bottom", fill="x")
        lbl_footer = tk.Label(footer, text=lang.tr("footer"), bg="white")
        lbl_footer.pack(side="bottom", pady=6)

    # ---------------------------
    # Source page (page 2)
    # ---------------------------
    def _build_source_page(self):
        p = self.page_source
        header = tk.Frame(p, height=80, bg="white")
        header.pack(fill="x")

        style = ttk.Style()
        style.configure(
            "Large.TButton",
            font=("Segoe UI", 18, "bold"),
            padding=14
        )

        btn_back = ttk.Button(header, text=lang.tr("btn_back"), command=self.show_welcome)
        btn_back.pack(side="left", padx=14, pady=6)
        btn_back.configure(width=10)

        lbl = tk.Label(
            p,
            text=lang.tr("select_source"),
            font=("Segoe UI", 24, "bold"),
            bg="white"
        )
        lbl.pack(pady=(24, 10))

        box = tk.Frame(p, bg="white")
        box.pack(expand=True)

        btn_gsheet = ttk.Button(
            box,
            text=lang.tr("btn_gsheet"),
            command=lambda: self._select_source_and_continue("gsheet"),
            style="Large.TButton"
        )

        btn_excel = ttk.Button(
            box,
            text=lang.tr("btn_excel"),
            command=lambda: self._select_source_and_continue("excel"),
            style="Large.TButton"
        )

        btn_gsheet.grid(row=0, column=0, padx=100, pady=40, ipadx=20, ipady=14)
        btn_excel.grid(row=0, column=1, padx=100, pady=40, ipadx=20, ipady=14)

        footer = tk.Frame(p, height=40, bg="white")
        footer.pack(side="bottom", fill="x")
        tk.Label(footer, text=lang.tr("footer"), bg="white").pack(side="bottom", pady=6)

    def _select_source_and_continue(self, mode):
        self.reset_log()
        self.selected_source = mode
        self.df = None
        self.excel_path = None
        self.show_main()

    # ---------------------------
    # Main page (page 3) build
    # ---------------------------
    def _build_main_page(self):
        p = self.page_main

        topbar = tk.Frame(p)
        topbar.pack(fill="x", padx=10, pady=6)

        btn_back_main = ttk.Button(topbar, text=lang.tr("btn_back"), command=self._back_to_source)
        btn_back_main.pack(side="left", padx=4)
        btn_back_main.configure(width=10)

        self.source_label = tk.Label(topbar, text=f"{lang.tr('source')}: -", font=("Segoe UI", 10, "bold"))
        self.source_label.pack(side="left", padx=8)

        controls = tk.Frame(p)
        controls.pack(fill="x", padx=10, pady=6)

        controls.grid_columnconfigure(0, minsize=167)
        controls.grid_columnconfigure(1, weight=4, minsize=440)

        tk.Label(controls, text=lang.tr("rows_start_end"), width=20, anchor="w")\
            .grid(row=3, column=0, sticky="w", pady=6)

        rows_frame = tk.Frame(controls)
        rows_frame.grid(row=3, column=1, sticky="w", padx=(6, 0))

        self.row_start = tk.Entry(rows_frame, width=6)
        self.row_start.pack(side="left", padx=(0, 4))

        self.row_end = tk.Entry(rows_frame, width=6)
        self.row_end.pack(side="left", padx=(0, 0))

        self.excel_frame = tk.Frame(controls)
        self.excel_frame.grid(row=0, column=0, columnspan=10, sticky="w")

        self.gsheet_frame = tk.Frame(controls)
        self.gsheet_frame.grid(row=1, column=0, columnspan=10, sticky="w")

        # --- Excel file UI ---
        self.excel_frame.grid_columnconfigure(0, minsize=167)
        self.excel_frame.grid_columnconfigure(1, weight=4, minsize=440)
        self.excel_frame.grid_columnconfigure(2, weight=0)

        tk.Label(self.excel_frame, text=lang.tr("file_csv_excel"), anchor="w")\
            .grid(row=0, column=0, sticky="w")

        self.entry_file = tk.Entry(self.excel_frame)
        self.entry_file.grid(row=0, column=1, padx=(6, 6), sticky="ew")

        self.btn_browse = ttk.Button(self.excel_frame, text=lang.tr("btn_browse"), command=self._browse_file)
        self.btn_browse.grid(row=0, column=2, padx=6, sticky="w")

        tk.Label(self.excel_frame, text=lang.tr("sheet"), anchor="w")\
            .grid(row=1, column=0, sticky="w")

        self.sheet_combo = ttk.Combobox(self.excel_frame)
        self.sheet_combo.grid(row=1, column=1, padx=(6, 6), sticky="ew")

        self.btn_load_sheet = ttk.Button(self.excel_frame, text=lang.tr("btn_load_sheet"), command=self._load_sheet_from_excel)
        self.btn_load_sheet.grid(row=1, column=2, padx=6, sticky="w")

        # --- Google Sheets UI ---
        tk.Label(self.gsheet_frame, text=lang.tr("api_key"), width=20, anchor="w").grid(row=0, column=0, sticky="w")
        self.entry_api_key = tk.Entry(self.gsheet_frame, width=53)
        self.entry_api_key.grid(row=0, column=1, padx=6, columnspan=3, sticky="ew")

        tk.Label(self.gsheet_frame, text=lang.tr("spreadsheet_id"), width=20, anchor="w").grid(row=1, column=0, sticky="w")
        self.entry_spreadsheet_id = tk.Entry(self.gsheet_frame, width=53)
        self.entry_spreadsheet_id.grid(row=1, column=1, padx=6, columnspan=3, sticky="ew")

        self.btn_getsheets = ttk.Button(self.gsheet_frame, text=lang.tr("btn_get_sheets"), command=self._get_gs_sheets)
        self.btn_getsheets.grid(row=1, column=4, sticky="w", padx=6)

        self.gs_sheet_combo = ttk.Combobox(self.gsheet_frame)
        self.gs_sheet_combo.grid(row=2, column=1, padx=6, columnspan=3, sticky="ew")

        self.btn_load_gs = ttk.Button(self.gsheet_frame, text=lang.tr("btn_load_gs"), command=self._load_from_gs)
        self.btn_load_gs.grid(row=2, column=4, sticky="w", padx=6)

        self.gsheet_frame.columnconfigure(1, weight=1)
        self.gsheet_frame.columnconfigure(2, weight=1)
        self.gsheet_frame.columnconfigure(3, weight=1)
        self.gsheet_frame.columnconfigure(4, weight=0)

        # Column mapping area
        map_frame = tk.Frame(p)
        map_frame.pack(fill="x", padx=12, pady=(6, 0))

        def mk_row(r, label):
            lbl = tk.Label(map_frame, text=label, width=20, anchor="w")
            lbl.grid(row=r, column=0, sticky="w", pady=6)
            cmb = ttk.Combobox(map_frame)
            cmb.grid(row=r, column=1, columnspan=3, sticky="ew", padx=6)
            return cmb

        self.cmb_name = mk_row(1, lang.tr("col_name"))
        self.cmb_no = mk_row(2, lang.tr("col_phone"))
        self.cmb_addr = mk_row(3, lang.tr("col_address"))

        tk.Label(map_frame, text=lang.tr("col_status"), width=20, anchor="w")\
            .grid(row=4, column=0, sticky="w", pady=6)

        self.cmb_status_col = ttk.Combobox(map_frame)
        self.cmb_status_col.grid(row=4, column=1, columnspan=3, sticky="ew", padx=6)
        self.cmb_status_col.bind("<<ComboboxSelected>>", self._on_status_col_selected)

        tk.Label(map_frame, text=lang.tr("filter_status"), width=20, anchor="w")\
            .grid(row=5, column=0, sticky="w", pady=6)

        self.filter_status_combo = ttk.Combobox(map_frame, width=50)
        self.filter_status_combo.grid(row=5, column=1, columnspan=3, sticky="w", padx=6)

        btn_auto_map = ttk.Button(map_frame, text=lang.tr("btn_auto_map"), command=self._auto_map_cols)
        btn_auto_map.grid(row=6, column=1, pady=10, sticky="w")

        # Delay controls
        delay_frame = tk.Frame(p)
        delay_frame.pack(fill="x", padx=12, pady=6)
        tk.Label(delay_frame, text=lang.tr("delay_min")).grid(row=0, column=0, sticky="w")
        self.delay_min = tk.Entry(delay_frame, width=6)
        self.delay_min.insert(0, "3")
        self.delay_min.grid(row=0, column=1, sticky="w")
        tk.Label(delay_frame, text=lang.tr("delay_max")).grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.delay_max = tk.Entry(delay_frame, width=6)
        self.delay_max.insert(0, "5")
        self.delay_max.grid(row=0, column=3, sticky="w")

        # Split left (template) and right (log)
        split_frame = tk.Frame(p)
        split_frame.pack(fill="both", expand=True, padx=12, pady=6)

        left_frame = tk.Frame(split_frame)
        left_frame.pack(side="left", fill="both", expand=True)

        tk.Label(left_frame, text=lang.tr("template_label")).pack(anchor="w")
        self.template_text = tk.Text(left_frame, height=3, width=50)
        self.template_text.insert("1.0", lang.tr("default_template"))
        self.template_text.pack(fill="both", expand=True, pady=(0, 6))

        right_frame = tk.Frame(split_frame)
        right_frame.pack(side="left", fill="both", expand=True, padx=(12, 0))

        tk.Label(right_frame, text=lang.tr("log_preview")).pack(anchor="w")
        self.log_text = tk.Text(right_frame, height=3, width=50)
        self.log_text.pack(fill="both", expand=True)

        btns_frame = tk.Frame(p)
        btns_frame.pack(fill="x", padx=12, pady=6)
        self.btn_preview = ttk.Button(btns_frame, text=lang.tr("btn_preview"), command=self._preview, width=20)
        self.btn_preview.pack(side="left", padx=6)
        self.btn_start = ttk.Button(btns_frame, text=lang.tr("btn_start"), command=self._start_sending, width=20)
        self.btn_start.pack(side="left", padx=6)
        self.btn_stop = ttk.Button(btns_frame, text=lang.tr("btn_stop"), command=self._stop_sending, width=12)
        self.btn_stop.pack(side="left", padx=6)

        prog_frame = tk.Frame(p)
        prog_frame.pack(fill="x", padx=12, pady=6)
        tk.Label(prog_frame, text=lang.tr("progress")).pack(side="left")
        self.progress = ttk.Progressbar(prog_frame, orient="horizontal", length=600, mode="determinate")
        self.progress.pack(side="left", padx=8)
        self.status_label = tk.Label(prog_frame, text=lang.tr("status_idle"))
        self.status_label.pack(side="left", padx=12)

        self.stat_label = tk.Label(p, text=lang.tr("stat_label", success=0, fail=0, skipped=0),
                                   font=("Arial", 10), fg="white", bg="#333")
        self.stat_label.pack(fill="x", padx=12, pady=(0, 6))

    # ---------------------------
    # Navigation helpers
    # ---------------------------
    def _back_to_source(self):
        self.selected_source = None
        self.show_source()

    # ---------------------------
    # Excel functions
    # ---------------------------
    def _browse_file(self):
        self.reset_log()
        f = filedialog.askopenfilename(filetypes=[(lang.tr("file_types"), lang.tr("file_types_pattern"))])
        if not f:
            return
        self.excel_path = f
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, f)

        try:
            xls = pd.ExcelFile(f)
            self.sheet_combo['values'] = xls.sheet_names
            messagebox.showinfo("File loaded", f"File loaded. {len(xls.sheet_names)} sheet(s) found.")
        except Exception as e:
            self.sheet_combo['values'] = []
            messagebox.showwarning("Warning",
                                f"Not an Excel file or could not read sheets: {e}. If CSV, use Load Sheet anyway.")

    def _load_sheet_from_excel(self):
        self.reset_log()
        if not self.excel_path:
            messagebox.showwarning(lang.tr("msg_no_data"), lang.tr("msg_no_file"))
            return
        sheet = self.sheet_combo.get()
        try:
            if sheet:
                self.df = pd.read_excel(self.excel_path, sheet_name=sheet)
            else:
                if self.excel_path.lower().endswith('.csv'):
                    self.df = pd.read_csv(self.excel_path)
                else:
                    messagebox.showwarning(lang.tr("msg_no_data"), lang.tr("msg_no_sheet"))
                    return
            self._after_load()
            messagebox.showinfo("Loaded", "Data loaded from Excel.")
        except Exception as e:
            messagebox.showerror("Error", lang.tr("msg_load_sheet_fail", e=e))

    # ---------------------------
    # Google Sheets functions
    # ---------------------------
    def _get_gs_sheets(self):
        self.reset_log()
        api = self.entry_api_key.get().strip()
        ss = self.entry_spreadsheet_id.get().strip()
        if not api or not ss:
            messagebox.showwarning(lang.tr("msg_no_data"), lang.tr("msg_missing_api"))
            return
        try:
            url = f"https://sheets.googleapis.com/v4/spreadsheets/{ss}?fields=sheets.properties.title&key={api}"
            r = requests.get(url, timeout=10)
            r.raise_for_status()
            js = r.json()
            sheets = [s['properties']['title'] for s in js.get("sheets", [])]
            self.gs_sheet_combo['values'] = sheets
            messagebox.showinfo("OK", f"Found {len(sheets)} sheets.")
        except Exception as e:
            messagebox.showerror("Error", lang.tr("msg_fetch_sheets_fail", e=e))

    def _load_from_gs(self):
        self.reset_log()
        api = self.entry_api_key.get().strip()
        ss = self.entry_spreadsheet_id.get().strip()
        sheet = self.gs_sheet_combo.get().strip()

        if not (api and ss and sheet):
            messagebox.showwarning(lang.tr("msg_no_data"), lang.tr("msg_missing_gs"))
            return

        rng = "A:ZZZ"

        try:
            values_range = f"{sheet}!{rng}"
            url = f"https://sheets.googleapis.com/v4/spreadsheets/{ss}/values/{values_range}?key={api}"

            r = requests.get(url, timeout=15)
            r.raise_for_status()
            js = r.json()

            vals = js.get("values", [])
            if not vals:
                messagebox.showwarning(lang.tr("msg_no_data"), lang.tr("msg_no_data_gs"))
                return

            max_len = max(len(row) for row in vals)
            vals = [row + [""] * (max_len - len(row)) for row in vals]

            self.df = pd.DataFrame(vals[1:], columns=vals[0])
            self._after_load()

            messagebox.showinfo("Loaded", "Data loaded from Google Sheets (all columns).")

        except Exception as e:
            messagebox.showerror("Error", lang.tr("msg_load_gs_fail", e=e))

    # ---------------------------
    # After loading any source
    # ---------------------------
    def _after_load(self):
        cols = list(self.df.columns)

        for cmb in (self.cmb_name, self.cmb_no, self.cmb_addr):
            cmb['values'] = cols

        self.cmb_status_col['values'] = cols
        self.cmb_status_col.set("")

        self.filter_status_combo.set("")
        self.filter_status_combo['values'] = []

        self.log(f"Data loaded. Columns: {cols}")
        self.status_label.config(text="Data loaded")
        self._auto_map_cols()

    def _on_status_col_selected(self, event=None):
        col = self.cmb_status_col.get()
        if not col or self.df is None or col not in self.df.columns:
            self.filter_status_combo['values'] = []
            self.filter_status_combo.set("")
            return

        values = (
            self.df[col]
            .dropna()
            .astype(str)
            .str.strip()
            .sort_values()
            .unique()
            .tolist()
            )

        self.filter_status_combo['values'] = [""] + values
        self.filter_status_combo.set("")

    def _auto_map_cols(self):
        if self.df is None:
            return

        cols_lower = [c.lower() for c in self.df.columns]

        def find_like(keys):
            for k in keys:
                for i, c in enumerate(cols_lower):
                    if k in c:
                        return self.df.columns[i]
            return ""

        self.cmb_name.set(find_like(['name', 'nama', 'nm']))
        self.cmb_no.set(find_like(['phone', 'hp', 'no', 'tel']))
        self.cmb_addr.set(find_like(['address', 'alamat', 'addr']))

        self.cmb_status_col['values'] = list(self.df.columns)
        self.cmb_status_col.set("")
        self.filter_status_combo.set("")
        self.filter_status_combo['values'] = []

    def _update_filter_status_options(self):
        try:
            col = self.cmb_status_col.get()
            if not col or self.df is None:
                return

            unique_vals = sorted(list(set(str(x).strip() for x in self.df[col].dropna())))
            self.filter_status_combo['values'] = [""] + unique_vals
        except:
            pass

    # ---------------------------
    # Preview
    # ---------------------------
    def _preview(self):
        self.reset_log()

        if self.df is None:
            messagebox.showwarning(lang.tr("msg_no_data"), lang.tr("msg_no_valid_send"))
            return

        template = self.template_text.get("1.0", tk.END)
        filter_status = self.filter_status_combo.get().strip()

        row_start = int(self.row_start.get()) if self.row_start.get().strip() else 1
        row_end = int(self.row_end.get()) if self.row_end.get().strip() else len(self.df)

        row_start = max(1, row_start)
        row_end = min(len(self.df), row_end)

        valid_rows = []

        for i in range(row_start - 1, row_end):
            row = self.df.iloc[i]
            actual_row = i + 1

            if filter_status:
                if str(row.get(self.cmb_status_col.get(), "")).strip() != filter_status:
                    continue

            phone = normalize_phone(row.get(self.cmb_no.get(), ""))
            if not phone:
                continue

            valid_rows.append((actual_row, row, phone))

            if len(valid_rows) >= 5:
                break

        if not valid_rows:
            self.log(lang.tr("msg_no_valid_preview"))
            return

        for actual_row, row, phone in valid_rows:
            name = row.get(self.cmb_name.get(), "")
            address = row.get(self.cmb_addr.get(), "")

            try:
                msg = template.format(name=name, phone=phone, address=address)
            except:
                msg = "[Template error]"

            self.log(f"Preview row {actual_row} -> {phone}\n{msg}\n")

    # ---------------------------
    # Sending logic
    # ---------------------------
    def _start_sending(self):
        self.reset_log()
        if self.df is None:
            messagebox.showwarning(lang.tr("msg_no_data"), lang.tr("msg_no_valid_send"))
            return
        if self.sending_thread and self.sending_thread.is_alive():
            messagebox.showinfo(lang.tr("msg_sending_running"), lang.tr("msg_sending_running"))
            return
        if not self.cmb_name.get() or not self.cmb_no.get():
            messagebox.showwarning(lang.tr("msg_mapping_required"), lang.tr("msg_mapping_required"))
            return
        self.success = 0
        self.fail = 0
        self.skipped = 0
        self.progress['value'] = 0
        self.stop_event.clear()
        self.sending_thread = threading.Thread(target=self._send_worker, daemon=True)
        self.sending_thread.start()

    def _stop_sending(self):
        if self.sending_thread and self.sending_thread.is_alive():
            self.stop_event.set()
            self.log("Stop requested...")
        else:
            self.log(lang.tr("msg_no_process"))

    def _send_worker(self):
        rows_total = len(self.df)
        try:
            start_idx = int(self.row_start.get()) if self.row_start.get().strip() else 1
        except:
            start_idx = 1

        try:
            end_idx = int(self.row_end.get()) if self.row_end.get().strip() else rows_total
        except:
            end_idx = rows_total

        start_idx = max(1, start_idx)
        end_idx = min(rows_total, end_idx)

        total_to_process = max(0, end_idx - start_idx + 1)

        self.progress['maximum'] = total_to_process
        self.progress['value'] = 0

        self.log(f"Start sending. Range rows: {start_idx}..{end_idx} | Total: {total_to_process}")
        self.status_label.config(text="Running")

        filter_status_val = ""
        if hasattr(self, "filter_status_combo"):
            try:
                filter_status_val = self.filter_status_combo.get().strip()
            except:
                filter_status_val = ""

        first = True

        try:
            for i in range(start_idx - 1, end_idx):

                if self.stop_event.is_set():
                    self.log(lang.tr("msg_stopped"))
                    break

                row = self.df.iloc[i]
                actual_row = i + 1

                if filter_status_val:
                    row_status = str(row.get(self.cmb_status_col.get(), "")).strip()
                    if row_status != filter_status_val:
                        self.skipped += 1
                        self.log(lang.tr("msg_skipped_status", row=actual_row))
                        self.progress['value'] += 1
                        continue

                raw_no = row.get(self.cmb_no.get(), "")
                phone = normalize_phone(raw_no)
                if not phone:
                    self.fail += 1
                    self.log(lang.tr("msg_fail_invalid_phone", row=actual_row, phone=raw_no))
                    self._update_stats()
                    self.progress['value'] += 1
                    continue

                name = row.get(self.cmb_name.get(), "")
                address = row.get(self.cmb_addr.get(), "")
                template = self.template_text.get("1.0", tk.END)
                try:
                    message = template.format(name=name, phone=phone, address=address)
                except Exception as e:
                    self.fail += 1
                    self.log(lang.tr("msg_template_error", row=actual_row, e=e))
                    self._update_stats()
                    self.progress['value'] += 1
                    continue

                if first:
                    self.log(lang.tr("msg_preparing_wa"))
                    time.sleep(5)
                    first = False

                self.log(lang.tr("msg_sending_to", phone=phone, row=actual_row))

                try:
                    kit.sendwhatmsg_instantly(phone, message, wait_time=10, tab_close=True)
                    self.success += 1
                    self.log(lang.tr("msg_success", row=actual_row, phone=phone))
                except Exception as e:
                    self.fail += 1
                    self.log(lang.tr("msg_send_error", row=actual_row, e=e))

                self.progress['value'] += 1
                self._update_stats()

                try:
                    dmin = int(self.delay_min.get())
                    dmax = int(self.delay_max.get())
                except:
                    dmin, dmax = 3, 5

                if dmax < dmin:
                    dmax = dmin

                delay = random.randint(dmin, dmax)
                self.log(lang.tr("msg_wait", delay=delay))

                for _ in range(delay):
                    if self.stop_event.is_set():
                        break
                    time.sleep(1)

                if self.stop_event.is_set():
                    self.log("Stop requested. Exiting.")
                    break

        finally:
            self.progress['value'] = self.progress['maximum']
            self.status_label.config(text=lang.tr("status_idle"))
            self.log(lang.tr("msg_sending_finished"))
            self._update_stats()

            if self.last_page:
                def bring_last_page():
                    self.last_page.lift()
                    self.last_page.focus_force()
                    self.attributes('-topmost', True)
                    self.attributes('-topmost', False)
                self.after(0, bring_last_page)

    # ---------------------------
    # Helpers
    # ---------------------------
    def _refresh_main_source_ui(self):
        if self.selected_source == "excel":
            self.excel_frame.grid()
            self.gsheet_frame.grid_remove()
        elif self.selected_source == "gsheet":
            self.gsheet_frame.grid()
            self.excel_frame.grid_remove()
        self.log_text.delete("1.0", tk.END)
        self.source_label.config(text=f"{lang.tr('source')}: {self.selected_source}")

    def _update_stats(self):
        thread_safe_update_label(self, self.stat_label,
                                 lang.tr("stat_label", success=self.success, fail=self.fail, skipped=self.skipped))

    def log(self, msg):
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n")
        self.log_text.see("end")

    def reset_log(self):
        self.log_text.delete("1.0", tk.END)
