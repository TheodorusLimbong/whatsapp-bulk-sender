# ui_pages.py
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

from utils import normalize_phone
from logic_worker import thread_safe_askstring, thread_safe_update_label

# Window size chosen: 1366 x 768 (not too tall)
WINDOW_W = 1366
WINDOW_H = 768

# Predefined status pengambilan dropdown options (you can edit)
STATUS_PENG_OPTIONS = ["", "BELUM DIAMBIL", "SUDAH DIAMBIL", "PERLU CATATAN", "BELUM", "SUDAH"]

class WATemplateSenderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WA Template Sender - Iconnet")
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

        lbl_title = tk.Label(body, text="Welcome", font=("Segoe UI", 48, "bold"), bg="white")
        lbl_title.pack(pady=(60, 0))

        lbl_sub = tk.Label(body, text="To\nWhatsApp Automation", font=("Segoe UI", 20), bg="white", justify="center")
        lbl_sub.pack(pady=(8, 20))

        btn_next = ttk.Button(body, text="Next", command=self.show_source)
        btn_next.pack(pady=12)
        btn_next.configure(width=20)

        footer = tk.Frame(p, height=40, bg="white")
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

        # --- Excel file UI (shown only if selected_source == 'excel') ---
        lbl_file = tk.Label(controls, text="File (CSV / Excel):")
        lbl_file.grid(row=0, column=0, sticky="w", pady=4)
        self.entry_file = tk.Entry(controls, width=70)
        self.entry_file.grid(row=0, column=1, columnspan=3, sticky="w")
        self.btn_browse = ttk.Button(controls, text="Browse", command=self._browse_file)
        self.btn_browse.grid(row=0, column=4, padx=6)
        self.sheet_combo = ttk.Combobox(controls, width=30)
        self.sheet_combo.grid(row=1, column=1, sticky="w")
        self.btn_load_sheet = ttk.Button(controls, text="Load Sheet", command=self._load_sheet_from_excel)
        self.btn_load_sheet.grid(row=1, column=2, padx=6)

        # --- Google Sheets UI (shown only if selected_source == 'gsheet') ---
        lbl_api = tk.Label(controls, text="API_KEY:")
        lbl_api.grid(row=2, column=0, sticky="w", pady=4)
        self.entry_api = tk.Entry(controls, width=50)
        self.entry_api.grid(row=2, column=1, columnspan=2, sticky="w")
        lbl_ss = tk.Label(controls, text="SPREADSHEET_ID:")
        lbl_ss.grid(row=3, column=0, sticky="w", pady=4)
        self.entry_ss = tk.Entry(controls, width=50)
        self.entry_ss.grid(row=3, column=1, columnspan=2, sticky="w")
        self.btn_getsheets = ttk.Button(controls, text="Get Sheet List", command=self._get_gs_sheets)
        self.btn_getsheets.grid(row=2, column=3, padx=6)
        self.gs_sheet_combo = ttk.Combobox(controls, width=30)
        self.gs_sheet_combo.grid(row=2, column=4, sticky="w")
        self.btn_load_gs = ttk.Button(controls, text="Load from Google Sheets", command=self._load_from_gs)
        self.btn_load_gs.grid(row=3, column=3, padx=6)

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
        self.cmb_status_peng = mk_row(4, "Kolom Status Pengambilan:")

        # Status Pengambilan filter dropdown (A option)
        filter_frame = tk.Frame(p)
        filter_frame.pack(fill="x", padx=12, pady=(6,0))
        tk.Label(filter_frame, text="Filter Status Pengambilan (choose):", anchor="w").pack(side="left", padx=4)
        self.filter_status_peng_combo = ttk.Combobox(filter_frame, values=STATUS_PENG_OPTIONS, width=30)
        self.filter_status_peng_combo.pack(side="left", padx=6)

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
        for cmb in (self.cmb_name, self.cmb_no, self.cmb_addr, self.cmb_status, self.cmb_status_peng):
            cmb['values'] = cols
        self.log(f"Data loaded. Columns: {cols}")
        self.status_label.config(text="Data loaded")
        # select sensible defaults
        self._auto_map_cols()

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
        self.cmb_status_peng.set(find_like(['pengambilan','ambil','status peng']))

    # ---------------------------
    # Preview
    # ---------------------------
    def _preview(self):
        if self.df is None:
            messagebox.showwarning("No data", "Load data dulu.")
            return
        try:
            n = 5
            start = int(self.row_start.get()) if self.row_start.get().strip() else 1
            end = int(self.row_end.get()) if self.row_end.get().strip() else start + n - 1
            subset = self.df.iloc[start-1:end]
            template = self.template_text.get("1.0", tk.END)
            self.log_text.delete("1.0", tk.END)
            for idx, r in subset.iterrows():
                nama = r.get(self.cmb_name.get(), "")
                no = normalize_phone(r.get(self.cmb_no.get(), "")) or ""
                alamat = r.get(self.cmb_addr.get(), "")
                try:
                    msg = template.format(nama=nama, no_hp=no, alamat=alamat)
                except Exception as e:
                    msg = f"[Template fmt error: {e}]"
                self.log_text.insert(tk.END, f"Row {idx+1} -> {no}\n{msg}\n\n")
        except Exception as e:
            messagebox.showerror("Error", f"Preview error: {e}")

    # ---------------------------
    # Sending logic
    # ---------------------------
    def _start_sending(self):
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

        # filters from UI
        filter_status_val = None
        if self.cmb_status.get().strip():
            filter_status_val = thread_safe_askstring(self, "Filter", f"Enter value to filter column '{self.cmb_status.get()}': (leave blank = no filter)")
        filter_status_peng_val = self.filter_status_peng_combo.get().strip()

        try:
            for i in range(start_idx-1, end_idx):
                if self.stop_event.is_set():
                    self.log("Stopped by user.")
                    break
                row = self.df.iloc[i]
                # apply status filter if specified
                if filter_status_val:
                    if str(row.get(self.cmb_status.get(), "")).strip() != filter_status_val:
                        self.skipped += 1
                        self.log(f"Row {i+1} skipped (status mismatch).")
                        self._update_stats()
                        self.progress['value'] += 1
                        continue
                # apply status pengambilan filter (from dropdown)
                if filter_status_peng_val:
                    if str(row.get(self.cmb_status_peng.get(), "")).strip() != filter_status_peng_val:
                        self.skipped += 1
                        self.log(f"Row {i+1} skipped (status pengambilan mismatch).")
                        self._update_stats()
                        self.progress['value'] += 1
                        continue

                raw_no = row.get(self.cmb_no.get(), "")
                phone = normalize_phone(raw_no)
                if not phone:
                    self.fail += 1
                    self.log(f"Row {i+1} fail: invalid phone ({raw_no})")
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
                    self.log(f"Row {i+1} fail: template error: {e}")
                    self._update_stats()
                    self.progress['value'] += 1
                    continue

                self.log(f"Sending to {phone} (row {i+1}) ...")
                try:
                    kit.sendwhatmsg_instantly(phone, pesan, wait_time=10, tab_close=True)
                    self.success += 1
                    self.log(f"Row {i+1} success -> {phone}")
                except Exception as e:
                    self.fail += 1
                    self.log(f"Row {i+1} error sending: {e}")

                self.progress['value'] += 1
                self._update_stats()

                # delay loop allows stop request to be honored promptly
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
        # Called when showing main page: show/hide UI depending on self.selected_source
        s = self.selected_source or "excel"
        self.source_label.config(text=f"Source: {s.upper()}")
        if s == "excel":
            # show file controls, hide google sheets controls
            self.entry_file.configure(state="normal")
            self.btn_browse.configure(state="normal")
            self.sheet_combo.configure(state="normal")
            self.btn_load_sheet.configure(state="normal")
            # hide google sheets widgets by disabling them
            self.entry_api.delete(0, tk.END)
            self.entry_ss.delete(0, tk.END)
            self.entry_api.configure(state="disabled")
            self.entry_ss.configure(state="disabled")
            self.btn_getsheets.configure(state="disabled")
            self.gs_sheet_combo.configure(state="disabled")
            self.btn_load_gs.configure(state="disabled")
        else:
            # gsheet: disable file browse, enable api fields
            self.entry_file.delete(0, tk.END)
            self.excel_path = None
            self.entry_file.configure(state="disabled")
            self.btn_browse.configure(state="disabled")
            self.sheet_combo.configure(state="disabled")
            self.btn_load_sheet.configure(state="disabled")
            # enable gs widgets
            self.entry_api.configure(state="normal")
            self.entry_ss.configure(state="normal")
            self.btn_getsheets.configure(state="normal")
            self.gs_sheet_combo.configure(state="normal")
            self.btn_load_gs.configure(state="normal")

    def _update_stats(self):
        # Thread-safe update (do not recreate widgets here)
        thread_safe_update_label(self, self.stat_label,
                                 f"Success: {self.success} | Fail: {self.fail} | Skipped: {self.skipped}")

    def log(self, msg):
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n")
        self.log_text.see("end")
