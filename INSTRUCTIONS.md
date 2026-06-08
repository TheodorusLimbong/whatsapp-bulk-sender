# Instructions — Push ke GitHub

## 1. Revoke API Key Lama
Buka https://console.cloud.google.com → APIs & Services → Credentials
- Cari API key `AIzaSyDKQxyjXCH2QXPAcn_feuKsbk-hbEy-8Wc`
- Klik **Delete** atau **Disable**

## 2. Buat API Key Baru (Optional)
- Di halaman Credentials yang sama, klik **Create Credentials** → **API Key**
- Restrict keynya hanya untuk Google Sheets API
- Simpan key baru di `config.json` lokal (tidak akan ter-track)

## 3. Rename Repo di GitHub
- Buka https://github.com/TheodorusLimbong/whatsapp-automation
- Settings → Repository name → ganti ke `whatsapp-bulk-sender`
- Klik **Rename**

## 4. Push ke GitHub

```bash
cd D:\backup\data d\dismantle\whatsapp-bulk-sender

# Update remote URL ke repo yang sudah di-rename
git remote set-url origin https://github.com/TheodorusLimbong/whatsapp-bulk-sender.git

# Add semua file
git add .

# Commit
git commit -m "Refactor: public-ready with dual language, security clean, English placeholders"

# Force push (karena history sudah di-rewrite)
git push origin main --force
```

## 5. Build .exe untuk User Non-Teknis

Di komputer Anda (yang ada Python):

```bash
pip install pyinstaller
pyinstaller WhatsAppBulkSender.spec
```

Hasil: folder `dist/WhatsAppBulkSender.exe`

## 6. Upload .exe ke GitHub Release (Recommended)

Kenapa .exe tidak di-track di git (tapi di-upload manual ke Release):
- .exe ukuran 6-7 MB — tidak cocok di git (membengkakkan repo, setiap clone/download jadi besar)
- Source code tetap kecil (0.1 MB) — mudah di-download developer
- User non-teknis download .exe dari halaman Release — tanpa perlu install Python/git

Langkah upload:
1. Buka repo di GitHub → **Releases** → **Create a new release**
2. Tag: `v1.0.0`
3. Title: `v1.0.0 - Public Release`
4. Description: "WhatsApp Bulk Sender — Kirim WA massal dari Excel / Google Sheets. Dual language EN/ID."
5. Attach file: `dist/WhatsAppBulkSender.exe`
6. Klik **Publish release**

User cukup buka halaman Release → download `.exe` → double-click.

## 7. Update Local Original Repo
Folder asli (`whatsapp-automation`) tidak perlu diubah. Biarkan sebagai backup.

## Rekomendasi Alur Publikasi
1. Revoke API key lama (langkah 1)
2. Rename repo GitHub (langkah 3)
3. Push repo bersih ke GitHub (langkah 4)
4. Build .exe di komputer lokal (langkah 5)
5. Upload .exe ke GitHub Release (langkah 6)
6. Selesai — user download dari Release, developer clone dari git

## Struktur File Final
```
whatsapp-bulk-sender/
  .gitignore
  AGENTS.md
  INSTRUCTIONS.md       <--- file ini
  LICENSE
  README.md
  WhatsAppBulkSender.spec
  i18n.py
  logic_worker.py
  main.py
  requirements.txt
  ui_pages.py
  utils.py
```
