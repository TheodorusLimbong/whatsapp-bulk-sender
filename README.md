# WhatsApp Bulk Sender

A desktop GUI application to send bulk WhatsApp messages using data from Excel (.xlsx/.csv) or Google Sheets. Built with Python, Tkinter, and pywhatkit.

## Features

- **Dual Language** — English and Indonesian (switchable from welcome page)
- **Data Sources** — Load from Excel (.xlsx, .xls, .csv) or Google Sheets (via REST API)
- **Smart Column Mapping** — Auto-detects name, phone, and address columns by keyword
- **Status Filter** — Filter rows by a status column (e.g., only send to "Pending" rows)
- **Custom Template** — Compose messages with `{name}`, `{phone}`, `{address}` placeholders
- **Row Range** — Select specific row ranges to send
- **Adjustable Delay** — Set min/max delay between messages (randomized)
- **Live Preview** — Preview the first 5 messages before sending
- **Progress Tracker** — Real-time success/fail/skipped stats

## Installation

### Prerequisites

- Python 3.8+
- Tkinter (included with Python)
- WhatsApp Web logged in on your default browser

### Setup

```bash
# Clone the repository
git clone https://github.com/TheodorusLimbong/whatsapp-bulk-sender.git
cd whatsapp-bulk-sender

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
python main.py
```

### Steps

1. **Select Language** — Choose English or Indonesian on the welcome page
2. **Choose Data Source** — Google Sheets or Excel
3. **Load Data**
   - *Excel*: Browse for file → select sheet → click "Load Sheet"
   - *Google Sheets*: Enter API key + Spreadsheet ID → get sheet list → load
4. **Map Columns** — Verify auto-detected columns for Name, Phone, Address (optional: Status column + filter value)
5. **Set Row Range** — Optional start/end row numbers (defaults: all rows)
6. **Compose Template** — Use `{name}`, `{phone}`, `{address}` placeholders
7. **Preview** — Click "Preview (5 rows)" to test
8. **Send** — Click "Start Sending"

> **Important**: You must be logged into **WhatsApp Web** (`https://web.whatsapp.com`) in your default browser before sending.

## Download Executable (No Python Required)

Non-technical users can download the pre-built `.exe` from the **Releases** page:
https://github.com/TheodorusLimbong/whatsapp-bulk-sender/releases

Just download `WhatsAppBulkSender.exe` and double-click to run.

## Building Executable (PyInstaller)

```bash
pip install pyinstaller
pyinstaller WhatsAppBulkSender.spec
```

The executable will be created in `dist/WhatsAppBulkSender.exe`.

> **Note**: The `.exe` is not tracked in git (too large for version control). It's distributed via GitHub Releases instead.

## Project Structure

```
whatsapp-bulk-sender/
  main.py              # Entry point
  ui_pages.py          # GUI logic and pages
  i18n.py              # Language support (EN/ID)
  utils.py             # Phone number normalization
  logic_worker.py      # Thread-safe Tkinter helpers
  requirements.txt     # Python dependencies
  WhatsAppBulkSender.spec  # PyInstaller config
```

## Dependencies

- `pandas` — Data processing (Excel/CSV)
- `requests` — Google Sheets API calls
- `pywhatkit` — WhatsApp Web automation
- `openpyxl` — Excel file support
- `tkinter` — GUI framework (standard library)

## License

[MIT](LICENSE)
