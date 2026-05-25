LANGUAGES = {
    "en": {
        # Welcome page
        "welcome_title": "Welcome",
        "welcome_subtitle": "To\nWhatsApp Bulk Sender",
        "btn_next": "Next",
        "footer": "All rights reserved",

        # Source page
        "select_source": "Select Input Data Source",
        "btn_back": "\u2190 Back",
        "btn_gsheet": "Google Sheets",
        "btn_excel": "Excel (.xlsx)",

        # Main page
        "source": "Source",
        "rows_start_end": "Rows (start-end):",
        "file_csv_excel": "File (CSV/Excel):",
        "btn_browse": "Browse",
        "sheet": "Sheet:",
        "btn_load_sheet": "Load Sheet",
        "api_key": "API_KEY:",
        "spreadsheet_id": "SPREADSHEET_ID:",
        "btn_get_sheets": "Get Sheet List",
        "btn_load_gs": "Load from Google Sheets",
        "col_name": "Name Column:",
        "col_phone": "Phone Column:",
        "col_address": "Address Column:",
        "col_status": "Status Column:",
        "filter_status": "Filter Status:",
        "btn_auto_map": "Auto Fill Columns",
        "delay_min": "Delay per message (seconds) min:",
        "delay_max": "max:",
        "template_label": "Message template (use {name}, {phone}, {address}):",
        "log_preview": "Log / Preview:",
        "btn_preview": "Preview (5 rows)",
        "btn_start": "Start Sending",
        "btn_stop": "Stop",
        "progress": "Progress:",
        "status_idle": "Idle",
        "stat_label": "Success: {success} | Fail: {fail} | Skipped: {skipped}",

        # Messages
        "msg_no_file": "Please select an Excel file first.",
        "msg_no_sheet": "Please select a sheet from the dropdown.",
        "msg_load_sheet_fail": "Failed to load sheet: {e}",
        "msg_missing_api": "Please fill API_KEY and SPREADSHEET_ID.",
        "msg_fetch_sheets_fail": "Failed to fetch sheet list: {e}",
        "msg_missing_gs": "Please fill API_KEY, SPREADSHEET_ID, and select a sheet.",
        "msg_no_data_gs": "No data found in the selected sheet.",
        "msg_load_gs_fail": "Failed to load data: {e}",
        "msg_no_data": "Load data first.",
        "msg_no_valid_preview": "No valid rows to preview in the selected range.",
        "msg_no_valid_send": "Load data first.",
        "msg_mapping_required": "Please select Name and Phone columns.",
        "msg_sending_running": "Sending already running.",
        "msg_stopped": "Stopped by user.",
        "msg_sending_finished": "Sending finished.",
        "msg_no_process": "No running process to stop.",
        "msg_preparing_wa": "Preparing WhatsApp Web... waiting 5 seconds...",
        "msg_sending_to": "Sending to {phone} (row {row}) ...",
        "msg_success": "Row {row} success -> {phone}",
        "msg_send_error": "Row {row} SEND ERROR: {e}",
        "msg_skipped_status": "Row {row} skipped (status mismatch)",
        "msg_fail_invalid_phone": "Row {row} fail: invalid phone ({phone})",
        "msg_template_error": "Row {row} template error: {e}",
        "msg_wait": "Wait {delay} sec ...",

        # File dialog
        "file_types": "Excel/CSV",
        "file_types_pattern": "*.xlsx *.xls *.csv",

        # Default template
        "default_template": "Hello, this is an automated message.\n\nName  : {name}\nPhone : {phone}\nAddress: {address}\n\nThank you.",

        # Stats label
        "success": "Success",
        "fail": "Fail",
        "skipped": "Skipped",
    },
    "id": {
        # Welcome page
        "welcome_title": "Selamat Datang",
        "welcome_subtitle": "Di\nWhatsApp Bulk Sender",
        "btn_next": "Lanjut",
        "footer": "Hak cipta dilindungi",

        # Source page
        "select_source": "Pilih Sumber Data Input",
        "btn_back": "\u2190 Kembali",
        "btn_gsheet": "Google Sheets",
        "btn_excel": "Excel (.xlsx)",

        # Main page
        "source": "Sumber",
        "rows_start_end": "Baris (awal-akhir):",
        "file_csv_excel": "File (CSV/Excel):",
        "btn_browse": "Cari",
        "sheet": "Sheet:",
        "btn_load_sheet": "Muat Sheet",
        "api_key": "API_KEY:",
        "spreadsheet_id": "SPREADSHEET_ID:",
        "btn_get_sheets": "Dapatkan Daftar Sheet",
        "btn_load_gs": "Muat dari Google Sheets",
        "col_name": "Kolom Nama:",
        "col_phone": "Kolom No HP:",
        "col_address": "Kolom Alamat:",
        "col_status": "Kolom Status:",
        "filter_status": "Filter Status:",
        "btn_auto_map": "Isi Kolom Otomatis",
        "delay_min": "Jeda per pesan (detik) min:",
        "delay_max": "maks:",
        "template_label": "Template pesan (pakai {name}, {phone}, {address}):",
        "log_preview": "Log / Pratinjau:",
        "btn_preview": "Pratinjau (5 baris)",
        "btn_start": "Mulai Kirim",
        "btn_stop": "Berhenti",
        "progress": "Progres:",
        "status_idle": "Siaga",
        "stat_label": "Berhasil: {success} | Gagal: {fail} | Dilewati: {skipped}",

        # Messages
        "msg_no_file": "Pilih file Excel terlebih dulu.",
        "msg_no_sheet": "Pilih sheet dari dropdown.",
        "msg_load_sheet_fail": "Gagal muat sheet: {e}",
        "msg_missing_api": "Isi API_KEY dan SPREADSHEET_ID.",
        "msg_fetch_sheets_fail": "Gagal ambil daftar sheet: {e}",
        "msg_missing_gs": "Isi API_KEY, SPREADSHEET_ID, dan pilih sheet.",
        "msg_no_data_gs": "Tidak ada data pada sheet tersebut.",
        "msg_load_gs_fail": "Gagal muat data: {e}",
        "msg_no_data": "Muat data terlebih dulu.",
        "msg_no_valid_preview": "Tidak ada baris valid untuk pratinjau di rentang tersebut.",
        "msg_no_valid_send": "Muat data terlebih dulu.",
        "msg_mapping_required": "Pastikan Kolom Nama dan Kolom No HP terpilih.",
        "msg_sending_running": "Pengiriman sedang berjalan.",
        "msg_stopped": "Dihentikan oleh pengguna.",
        "msg_sending_finished": "Pengiriman selesai.",
        "msg_no_process": "Tidak ada proses yang berjalan untuk dihentikan.",
        "msg_preparing_wa": "Menyiapkan WhatsApp Web... tunggu 5 detik...",
        "msg_sending_to": "Mengirim ke {phone} (baris {row}) ...",
        "msg_success": "Baris {row} berhasil -> {phone}",
        "msg_send_error": "Baris {row} GAGAL KIRIM: {e}",
        "msg_skipped_status": "Baris {row} dilewati (status tidak cocok)",
        "msg_fail_invalid_phone": "Baris {row} gagal: nomor HP tidak valid ({phone})",
        "msg_template_error": "Baris {row} error template: {e}",
        "msg_wait": "Tunggu {delay} detik ...",

        # File dialog
        "file_types": "Excel/CSV",
        "file_types_pattern": "*.xlsx *.xls *.csv",

        # Default template
        "default_template": "Halo, ini adalah pesan otomatis.\n\nNama  : {name}\nNo HP : {phone}\nAlamat: {address}\n\nTerima kasih.",

        # Stats label
        "success": "Berhasil",
        "fail": "Gagal",
        "skipped": "Dilewati",
    },
}


class LanguageManager:
    def __init__(self, default_lang="en"):
        self._lang = default_lang if default_lang in LANGUAGES else "en"

    @property
    def lang(self):
        return self._lang

    @lang.setter
    def lang(self, value):
        if value in LANGUAGES:
            self._lang = value

    def tr(self, key, **kwargs):
        text = LANGUAGES[self._lang].get(key, LANGUAGES["en"].get(key, key))
        if kwargs:
            try:
                text = text.format(**kwargs)
            except KeyError:
                pass
        return text

    def get_lang_options(self):
        return list(LANGUAGES.keys())


lang = LanguageManager()
