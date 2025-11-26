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

# Window size chosen: 1366 x 768 (not too tall)
WINDOW_W = 1366
WINDOW_H = 768

class WATemplateSenderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Whatsapp Automation")
        # center geometry
        x = (self.winfo_screenwidth() - WINDOW_W) // 2
        y = (self.winfo_screenheight() - WINDOW_H) // 2
        self.geometry(f"{WINDOW_W}x{WINDOW_H}+{x}+{y}")
        self.minsize(1000, 600)

        # state
        self.df = None
        self.excel_path = None
        self.selected_source = None  # 'excel' or 'gsheet' - set from page2
        self.sending_thread = None
        self.stop_event = threading.Event()

        # stats
        self.success = 0
        self.fail = 0
        self.skipped = 0

        self._build_pages()
        self.show_welcome()
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
        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.page_welcome = tk.Frame(self.container, bg="white")
        self.page_source = tk.Frame(self.container, bg="white")
        self.page_main = tk.Frame(self.container)

        for p in (self.page_welcome, self.page_source, self.page_main):
            p.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._build_welcome_page()
        self._build_source_page()
        self._build_main_page()

    def show_welcome(self):
        self.page_welcome.lift()

    def show_source(self):
        self.page_source.lift()

    def show_main(self):
        # update main page widgets according to selected_source
        self._refresh_main_source_ui()
        self.page_main.lift()

    # ---------------------------
    # Welcome page
    # ---------------------------
    def _build_welcome_page(self):
        p = self.page_welcome
        header = tk.Frame(p, height=60, bg="white")
        header.pack(fill="x")

        body = tk.Frame(p, bg="white")
        body.pack(fill="both", expand=True)

        # responsive font size (example)
        title_size = min(86, max(36, WINDOW_H // 12))
        lbl_title = tk.Label(body, text="Welcome", font=("Segoe UI", title_size, "bold"), bg="white")
        lbl_title.pack(pady=(40, 0))

        lbl_sub = tk.Label(body, text="To\nWhatsApp Automation", font=("Segoe UI", 28), bg="white", justify="center")
        lbl_sub.pack(pady=(8, 20))

        btn_next = ttk.Button(
            body,
            text="Next",
            command=self.show_source
        )
        btn_next.pack(pady=20, ipadx=25, ipady=22)   # memperbesar fisik tombol
        btn_next.configure(width=30)
        btn_next.focus_set()
        self.bind('<Return>', lambda e: self.show_source())

        footer = tk.Frame(p, height=50, bg="white")
        footer.pack(side="bottom", fill="x")
        lbl_footer = tk.Label(footer, text="© Iconnet. All rights reserved", bg="white")
        lbl_footer.pack(side="bottom", pady=6)


    # ---------------------------
    # Source page (page 2)
    # ---------------------------
    def _build_source_page(self):
        p = self.page_source
        header = tk.Frame(p, height=60, bg="white")
        header.pack(fill="x")

        btn_back = ttk.Button(header, text="←", command=self.show_welcome)
        btn_back.pack(side="left", padx=8, pady=8)

        lbl = tk.Label(p, text="Pilih Inputan Data yang akan di proses", font=("Segoe UI", 22, "bold"), bg="white")
        lbl.pack(pady=(18, 6))

        box = tk.Frame(p, bg="white")
        box.pack(expand=True)

        # Big buttons (enlarged)
        btn_gsheet = ttk.Button(box, text="Google Sheets", command=lambda: self._select_source_and_continue("gsheet"))
        btn_excel = ttk.Button(box, text="Excel (.xlsx)", command=lambda: self._select_source_and_continue("excel"))

        # place side by side with gap
        btn_gsheet.grid(row=0, column=0, padx=80, pady=30)
        btn_excel.grid(row=0, column=1, padx=80, pady=30)

        # enlarge fonts on the buttons
        btn_gsheet.configure(style="Large.TButton")
        btn_excel.configure(style="Large.TButton")

        # style for big buttons
        style = ttk.Style()
        style.configure("Large.TButton", font=("Segoe UI", 16), padding=10)

        footer = tk.Frame(p, height=40, bg="white")
        footer.pack(side="bottom", fill="x")
        tk.Label(footer, text="© Iconnet. All rights reserved", bg="white").pack(side="bottom", pady=6)

    def _select_source_and_continue(self, mode):
        self.reset_log()
        # mode is "excel" or "gsheet"
        self.selected_source = mode
        # clear previous loaded data state
        self.df = None
        self.excel_path = None
        self.show_main()

    # ---------------------------
    # Main page (page 3) build
    # ---------------------------
    def _build_main_page(self):
        p = self.page_main

        # top frame holds back button + source label
        topbar = tk.Frame(p)
        topbar.pack(fill="x", padx=10, pady=6)

        btn_back_main = ttk.Button(topbar, text="← Back", command=self._back_to_source)
        btn_back_main.pack(side="left", padx=4)
        btn_back_main.configure(width=10)

        self.source_label = tk.Label(topbar, text="Source: -", font=("Segoe UI", 10, "bold"))
        self.source_label.pack(side="left", padx=8)

        # content frames
        controls = tk.Frame(p)
        controls.pack(fill="x", padx=10, pady=6)

        # ==== NEW: frame khusus excel & gsheet ====
        self.excel_frame = tk.Frame(controls)
        self.excel_frame.grid(row=0, column=0, columnspan=10, sticky="w")

        self.gsheet_frame = tk.Frame(controls)
        self.gsheet_frame.grid(row=1, column=0, columnspan=10, sticky="w")


        # --- Excel file UI (shown only if selected_source == 'excel') ---
        # ---- EXCEL UI ----
        tk.Label(self.excel_frame, text="File (CSV/Excel):").grid(row=0, column=0, sticky="w")
        self.entry_file = tk.Entry(self.excel_frame, width=70)
        self.entry_file.grid(row=0, column=1, padx=6)
        self.btn_browse = ttk.Button(self.excel_frame, text="Browse", command=self._browse_file)
        self.btn_browse.grid(row=0, column=2, padx=6)

        tk.Label(self.excel_frame, text="Sheet:").grid(row=1, column=0, sticky="w")
        self.sheet_combo = ttk.Combobox(self.excel_frame, width=30)
        self.sheet_combo.grid(row=1, column=1, padx=6)
        self.btn_load_sheet = ttk.Button(self.excel_frame, text="Load Sheet", command=self._load_sheet_from_excel)
        self.btn_load_sheet.grid(row=1, column=2, padx=6)


        # --- Google Sheets UI (shown only if selected_source == 'gsheet') ---
        # ---- GOOGLE SHEETS UI ----
        tk.Label(self.gsheet_frame, text="API_KEY:").grid(row=0, column=0, sticky="w")
        self.entry_api = tk.Entry(self.gsheet_frame, width=50)
        self.entry_api.grid(row=0, column=1, padx=6)

        tk.Label(self.gsheet_frame, text="SPREADSHEET_ID:").grid(row=1, column=0, sticky="w")
        self.entry_ss = tk.Entry(self.gsheet_frame, width=50)
        self.entry_ss.grid(row=1, column=1, padx=6)

        self.btn_getsheets = ttk.Button(self.gsheet_frame, text="Get Sheet List", command=self._get_gs_sheets)
        self.btn_getsheets.grid(row=0, column=2, padx=6)

        self.gs_sheet_combo = ttk.Combobox(self.gsheet_frame, width=30)
        self.gs_sheet_combo.grid(row=1, column=2, padx=6)

        self.btn_load_gs = ttk.Button(self.gsheet_frame, text="Load from Google Sheets", command=self._load_from_gs)
        self.btn_load_gs.grid(row=1, column=3, padx=6)


        # Rows selection
        tk.Label(controls, text="Rows (start-end):").grid(row=4, column=0, sticky="w", pady=6)
        self.row_start = tk.Entry(controls, width=6)
        self.row_start.grid(row=4, column=1, sticky="w")
        self.row_end = tk.Entry(controls, width=6)
        self.row_end.grid(row=4, column=1, sticky="e")

        # Column mapping area (bigger controls)
        map_frame = tk.Frame(p)
        map_frame.pack(fill="x", padx=12, pady=(6,0))

        def mk_row(r, label):
            lbl = tk.Label(map_frame, text=label, width=20, anchor="w")
            lbl.grid(row=r, column=0, sticky="w", pady=6)
            cmb = ttk.Combobox(map_frame, width=50)
            cmb.grid(row=r, column=1, columnspan=3, sticky="w", padx=6)
            return cmb

        self.cmb_name = mk_row(0, "Kolom Nama:")
        self.cmb_no = mk_row(1, "Kolom No HP:")
        self.cmb_addr = mk_row(2, "Kolom Alamat:")
        self.cmb_status = mk_row(3, "Kolom Status:")

        # Status Pengambilan filter dropdown (A option)
        filter_frame = tk.Frame(p)
        filter_frame.pack(fill="x", padx=12, pady=(6,0))
        # NEW: Filter Status otomatis (untuk kolom status biasa)
        tk.Label(filter_frame, text="Filter Status:", anchor="w").pack(side="left", padx=10)
        self.filter_status_combo = ttk.Combobox(filter_frame, width=30)
        self.filter_status_combo.pack(side="left", padx=6)


        btn_auto_map = ttk.Button(map_frame, text="Isi Kolom Otomatis", command=self._auto_map_cols)
        btn_auto_map.grid(row=5, column=1, pady=8, sticky="w")

        # Delay controls
        delay_frame = tk.Frame(p)
        delay_frame.pack(fill="x", padx=12, pady=6)
        tk.Label(delay_frame, text="Delay per message (seconds) min:").grid(row=0, column=0, sticky="w")
        self.delay_min = tk.Entry(delay_frame, width=6)
        self.delay_min.insert(0, "3")
        self.delay_min.grid(row=0, column=1, sticky="w")
        tk.Label(delay_frame, text="max:").grid(row=0, column=2, sticky="w", padx=(8,0))
        self.delay_max = tk.Entry(delay_frame, width=6)
        self.delay_max.insert(0, "5")
        self.delay_max.grid(row=0, column=3, sticky="w")

        # --- SPLIT LEFT (Template) & RIGHT (Log) ---
        split_frame = tk.Frame(p)
        split_frame.pack(fill="both", expand=True, padx=12, pady=6)

        # ========== LEFT: Template Pesan ==========
        left_frame = tk.Frame(split_frame)
        left_frame.pack(side="left", fill="both", expand=True)

        tk.Label(left_frame, text="Template pesan (pakai {nama}, {no_hp}, {alamat}):").pack(anchor="w")
        self.template_text = tk.Text(left_frame, height=3, width=50)
        default_template = (
            "Selamat Pagi, Saya dari pihak Iconnet ingin melakukan Dismantle/Penarikan Modem, mohon maaf jika pesan ini "
            "sudah pernah terkirim sebelumnya. Dikarenakan adanya kesalahan teknis yang menyebabkan pelanggan dihubungi lebih dari sekali.\n\n"
            "Nama  : {nama}\nNo HP : {no_hp}\nAlamat: {alamat}\n\n"
            "Apakah anda sedang ada ditempat, atau kami bisa mendapatkan waktu lain untuk pengambilan modem?\nTerima kasih atas perhatian dan kerja samanya🙏"
        )
        self.template_text.insert("1.0", default_template)
        self.template_text.pack(fill="both", expand=True, pady=(0, 6))

        # ========== RIGHT: Log Area ==========
        right_frame = tk.Frame(split_frame)
        right_frame.pack(side="left", fill="both", expand=True, padx=(12,0))

        tk.Label(right_frame, text="Log / Preview:").pack(anchor="w")
        self.log_text = tk.Text(right_frame, height=3, width=50)
        self.log_text.pack(fill="both", expand=True)

        # Buttons large
        btns_frame = tk.Frame(p)
        btns_frame.pack(fill="x", padx=12, pady=6)
        self.btn_preview = ttk.Button(btns_frame, text="Preview (5 rows)", command=self._preview, width=20)
        self.btn_preview.pack(side="left", padx=6)
        self.btn_start = ttk.Button(btns_frame, text="Start Sending", command=self._start_sending, width=20)
        self.btn_start.pack(side="left", padx=6)
        self.btn_stop = ttk.Button(btns_frame, text="Stop", command=self._stop_sending, width=12)
        self.btn_stop.pack(side="left", padx=6)

        # Progress & status
        prog_frame = tk.Frame(p)
        prog_frame.pack(fill="x", padx=12, pady=6)
        tk.Label(prog_frame, text="Progress:").pack(side="left")
        self.progress = ttk.Progressbar(prog_frame, orient="horizontal", length=600, mode="determinate")
        self.progress.pack(side="left", padx=8)
        self.status_label = tk.Label(prog_frame, text="Idle")
        self.status_label.pack(side="left", padx=12)

        # Statistik bawah (buat 1x saja)
        self.stat_label = tk.Label(p, text="Success: 0 | Fail: 0 | Skipped: 0",
                                   font=("Arial", 10), fg="white", bg="#333")
        self.stat_label.pack(fill="x", padx=12, pady=(0, 6))

        # # Log area
        # log_frame = tk.Frame(p)
        # log_frame.pack(fill="both", expand=True, padx=12, pady=6)
        # tk.Label(log_frame, text="Log / Preview:").pack(anchor="w")
        # self.log_text = tk.Text(log_frame)
        # self.log_text.pack(fill="both", expand=True)

    # ---------------------------
    # Navigation helpers
    # ---------------------------
    def _back_to_source(self):
        # user pressed back from main page -> reset selection and go back to page2
        self.selected_source = None
        self.show_source()

    # ---------------------------
    # Excel functions
    # ---------------------------
    def _browse_file(self):
        self.reset_log()
        f = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not f:
            return
        self.excel_path = f
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, f)

        try:
            xls = pd.ExcelFile(f)
            # Tidak dibatasi — semua sheet akan dimasukkan
            self.sheet_combo['values'] = xls.sheet_names
            messagebox.showinfo("File loaded", f"File loaded. {len(xls.sheet_names)} sheet(s) found.")
        except Exception as e:
            self.sheet_combo['values'] = []
            messagebox.showwarning("Warning",
                                f"Not an Excel file or could not read sheets: {e}. If CSV, use Load Sheet anyway.")

    def _load_sheet_from_excel(self):
        self.reset_log()
        if not self.excel_path:
            messagebox.showwarning("No file", "Pilih file Excel terlebih dulu.")
            return
        sheet = self.sheet_combo.get()
        try:
            if sheet:
                self.df = pd.read_excel(self.excel_path, sheet_name=sheet)
            else:
                # maybe CSV
                if self.excel_path.lower().endswith('.csv'):
                    self.df = pd.read_csv(self.excel_path)
                else:
                    messagebox.showwarning("No sheet", "Pilih sheet dari dropdown.")
                    return
            self._after_load()
            messagebox.showinfo("Loaded", "Data loaded from Excel.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal load sheet: {e}")

    # ---------------------------
    # Google Sheets functions
    # ---------------------------
    def _get_gs_sheets(self):
        self.reset_log()
        api = self.entry_api.get().strip()
        ss = self.entry_ss.get().strip()
        if not api or not ss:
            messagebox.showwarning("Missing", "Isi API_KEY dan SPREADSHEET_ID.")
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
            messagebox.showerror("Error", f"Gagal ambil sheet list: {e}")

    def _load_from_gs(self):
        self.reset_log()
        api = self.entry_api.get().strip()
        ss = self.entry_ss.get().strip()
        sheet = self.gs_sheet_combo.get().strip()

        if not (api and ss and sheet):
            messagebox.showwarning("Missing", "Isi API_KEY, SPREADSHEET_ID, dan pilih sheet.")
            return

        # RANGE AUTO: ambil semua kolom dari sheet
        rng = "A:ZZZ"

        try:
            values_range = f"{sheet}!{rng}"
            url = f"https://sheets.googleapis.com/v4/spreadsheets/{ss}/values/{values_range}?key={api}"

            r = requests.get(url, timeout=15)
            r.raise_for_status()
            js = r.json()

            vals = js.get("values", [])
            if not vals:
                messagebox.showwarning("Empty", "Tidak ada data pada sheet tersebut.")
                return

            # Normalisasi panjang tiap baris
            max_len = max(len(row) for row in vals)
            vals = [row + [""] * (max_len - len(row)) for row in vals]

            self.df = pd.DataFrame(vals[1:], columns=vals[0])
            self._after_load()

            messagebox.showinfo("Loaded", "Data loaded from Google Sheets (all columns).")

        except Exception as e:
            messagebox.showerror("Error", f"Gagal load data: {e}")



    # ---------------------------
    # After loading any source
    # ---------------------------
    def _after_load(self):
        cols = list(self.df.columns)
        for cmb in (self.cmb_name, self.cmb_no, self.cmb_addr, self.cmb_status):
            cmb['values'] = cols
                # NEW: auto fill unique status
        try:
            status_col = self.cmb_status.get()
            if status_col in self.df.columns:
                unique_status = sorted(self.df[status_col].dropna().unique().astype(str))
                self.filter_status_combo['values'] = [""] + unique_status
                self.filter_status_combo.set("")  # default no filter
        except:
            pass

        self.log(f"Data loaded. Columns: {cols}")
        self.status_label.config(text="Data loaded")
        # select sensible defaults
        self._auto_map_cols()
        self._update_filter_status_options()


    def _auto_map_cols(self):
        if self.df is None:
            messagebox.showwarning("No data", "Load data terlebih dulu.")
            return
        cols_lower = [c.lower() for c in self.df.columns]
        def find_like(keys):
            for k in keys:
                for i,c in enumerate(cols_lower):
                    if k in c:
                        return self.df.columns[i]
            return ""
        self.cmb_name.set(find_like(['name','nama','nm']))
        self.cmb_no.set(find_like(['phone','hp','no','tel']))
        self.cmb_addr.set(find_like(['alamat','address','addr']))
        self.cmb_status.set(find_like(['status']))
        self._update_filter_status_options()

    def _update_filter_status_options(self):
        """Mengisi filter_status_combo berdasarkan kolom status yang dipilih."""
        try:
            col = self.cmb_status.get()
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
            messagebox.showwarning("No data", "Load data dulu.")
            return

        template = self.template_text.get("1.0", tk.END)
        filter_status = self.filter_status_combo.get().strip()

        # Baca batas row dari input user
        row_start = int(self.row_start.get()) if self.row_start.get().strip() else 1
        row_end = int(self.row_end.get()) if self.row_end.get().strip() else len(self.df)

        # Pastikan tidak melebihi batas
        row_start = max(1, row_start)
        row_end = min(len(self.df), row_end)

        valid_rows = []

        # Loop hanya dari row_start–row_end
        for i in range(row_start - 1, row_end):
            row = self.df.iloc[i]
            actual_row = i + 1

            # Filter status
            if filter_status:
                if str(row.get(self.cmb_status.get(), "")).strip() != filter_status:
                    continue

            # Validasi nomor HP
            phone = normalize_phone(row.get(self.cmb_no.get(), ""))
            if not phone:
                continue

            valid_rows.append((actual_row, row, phone))

            # Ambil hanya 5 preview
            if len(valid_rows) >= 5:
                break

        # Tidak ada row valid
        if not valid_rows:
            self.log("Tidak ada baris valid untuk preview di range tersebut.")
            return

        # Tampilkan preview
        for actual_row, row, phone in valid_rows:
            nama = row.get(self.cmb_name.get(), "")
            alamat = row.get(self.cmb_addr.get(), "")

            try:
                msg = template.format(nama=nama, no_hp=phone, alamat=alamat)
            except:
                msg = "[Template error]"

            self.log(f"Preview row {actual_row} -> {phone}\n{msg}\n")



    # ---------------------------
    # Sending logic
    # ---------------------------
    def _start_sending(self):
        self.reset_log()
        if self.df is None:
            messagebox.showwarning("No data", "Load data dulu.")
            return
        if self.sending_thread and self.sending_thread.is_alive():
            messagebox.showinfo("Running", "Sending already running.")
            return
        if not self.cmb_name.get() or not self.cmb_no.get():
            messagebox.showwarning("Mapping", "Pastikan Kolom Nama dan Kolom No HP terpilih.")
            return
        # reset stats
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
            self.log("No running process to stop.")

    
    def _send_worker(self):
        rows_total = len(self.df)
        self.progress['maximum'] = rows_total
        self.log(f"Start sending. Total rows: {rows_total}")
        self.status_label.config(text="Running")

        start_idx = int(self.row_start.get()) if self.row_start.get().strip() else 1
        end_idx = int(self.row_end.get()) if self.row_end.get().strip() else rows_total
        start_idx = max(1, start_idx)
        end_idx = min(rows_total, end_idx)

        # Ambil filter status satu kali
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
                    self.log("Stopped by user.")
                    break

                row = self.df.iloc[i]
                actual_row = i + 1

                # -------------------------
                # FILTER STATUS
                # -------------------------
                if filter_status_val:
                    row_status = str(row.get(self.cmb_status.get(), "")).strip()
                    if row_status != filter_status_val:
                        self.skipped += 1
                        self.log(f"Row {actual_row} skipped (status mismatch)")
                        self.progress['value'] += 1
                        continue

                # -------------------------
                # VALIDASI PHONE
                # -------------------------
                raw_no = row.get(self.cmb_no.get(), "")
                phone = normalize_phone(raw_no)
                if not phone:
                    self.fail += 1
                    self.log(f"Row {actual_row} fail: invalid phone ({raw_no})")
                    self._update_stats()
                    self.progress['value'] += 1
                    continue

                nama = row.get(self.cmb_name.get(), "")
                alamat = row.get(self.cmb_addr.get(), "")
                template = self.template_text.get("1.0", tk.END)
                try:
                    pesan = template.format(nama=nama, no_hp=phone, alamat=alamat)
                except Exception as e:
                    self.fail += 1
                    self.log(f"Row {actual_row} template error: {e}")
                    self._update_stats()
                    self.progress['value'] += 1
                    continue

                # -------------------------
                # FIRST SEND WAIT FIX
                # -------------------------
                if first:
                    self.log("Preparing WhatsApp Web... waiting 5 seconds...")
                    time.sleep(5)
                    first = False

                # -------------------------
                # SEND MESSAGE (SETIAP ITERASI)
                # -------------------------
                self.log(f"Sending to {phone} (row {actual_row}) ...")

                try:
                    kit.sendwhatmsg_instantly(phone, pesan, wait_time=10, tab_close=True)
                    self.success += 1
                    self.log(f"Row {actual_row} success -> {phone}")
                except Exception as e:
                    self.fail += 1
                    self.log(f"Row {actual_row} SEND ERROR: {e}")

                self.progress['value'] += 1
                self._update_stats()

                # -------------------------
                # DELAY ANTAR PESAN
                # -------------------------
                try:
                    dmin = int(self.delay_min.get())
                    dmax = int(self.delay_max.get())
                except:
                    dmin, dmax = 3, 5

                if dmax < dmin:
                    dmax = dmin

                delay = random.randint(dmin, dmax)
                self.log(f"Wait {delay} sec ...")

                for _ in range(delay):
                    if self.stop_event.is_set():
                        break
                    time.sleep(1)

                if self.stop_event.is_set():
                    self.log("Stop requested. Exiting.")
                    break

        finally:
            self.status_label.config(text="Idle")
            self.log("Sending finished.")
            self._update_stats()


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



    def _update_stats(self):
        # Thread-safe update (do not recreate widgets here)
        thread_safe_update_label(self, self.stat_label,
                                 f"Success: {self.success} | Fail: {self.fail} | Skipped: {self.skipped}")

    def log(self, msg):
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n")
        self.log_text.see("end")
    
    def reset_log(self):
        """Clear log setiap kali user menjalankan aksi baru."""
        self.log_text.delete("1.0", tk.END)