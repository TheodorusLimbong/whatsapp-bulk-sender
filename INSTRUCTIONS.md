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

## 5. Buat GitHub Release (Optional)
- Buka repo di GitHub → **Releases** → **Create a new release**
- Tag: `v1.0.0`
- Title: `v1.0.0 - Public Release`
- Description: fitur-fitur utama

## 6. Update Local Original Repo
Folder asli (`whatsapp-automation`) tidak perlu diubah. Biarkan sebagai backup.

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
