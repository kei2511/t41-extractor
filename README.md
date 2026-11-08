# Ekstraktor Ruang Lingkup T.4.1 ke T.4.2

Tool otomasi untuk mengekstrak poin ruang lingkup pengujian (fungsi-fungsi yang diuji) dari dokumen **Rencana SIT (T.4.1 - PDF)** dan menginjeksinya langsung ke dokumen template **Laporan SIT (T.4.2 - DOCX)** atau menyalinnya dalam format teks rapi ke clipboard.

---

## 🚀 Fitur Utama

- **Ekstraksi Otomatis dari PDF**: Membaca dokumen PDF T.4.1 dan mendeteksi poin ruang lingkup pengujian / fungsi yang diuji secara akurat.
- **Copy ke Clipboard**: Teks hasil ekstraksi diformat rapi dan siap ditempelkan di dokumen atau catatan pengujian.
- **Injeksi Langsung ke DOCX**: Mendukung pengisian otomatis tabel fungsi pengujian ke file template Word T.4.2 (.docx).
- **Multi-Platform**:
  1. **Web App Standalone** (`app.html`)
  2. **Ekstensi Browser / Chrome Extension** (folder `extension/`)
  3. **Script Python CLI** (`ekstrak_t41_ke_t42.py`)

---

## 📂 Struktur Proyek

```text
├── app.html                  # Aplikasi web standalone
├── popup.css                 # Styling untuk Web App
├── popup.js                  # Logika ekstraksi PDF & DOCX injector (Web)
├── ekstrak_t41_ke_t42.py      # Script ekstraksi berbasis Python CLI
├── libs/                     # Pustaka JavaScript pihak ketiga (PDF.js, JSZip)
├── extension/                # File siap pakai untuk Chrome Extension
│   ├── manifest.json
│   ├── popup.html
│   ├── popup.css
│   ├── popup.js
│   ├── icons/
│   └── libs/
└── README.md
```

---

## 🛠️ Cara Penggunaan

### 1. Web App Standalone
1. Buka file `app.html` langsung di browser favorit Anda (Chrome, Edge, Firefox, dll).
2. Upload dokumen PDF T.4.1 pada kotak yang tersedia.
3. Review hasil ekstraksi di kolom teks atau klik **Salin ke Clipboard**.
4. (Opsional) Upload template DOCX T.4.2 lalu klik **Injeksi & Download File T.4.2**.

### 2. Browser Extension (Chrome / Edge)
1. Buka `chrome://extensions/` atau `edge://extensions/`.
2. Aktifkan **Developer mode** (Mode pengembang) di pojok kanan atas.
3. Klik tombol **Load unpacked** (Muat ekstensi yang belum dikemas).
4. Pilih folder `extension` dari repository ini.
5. Ekstensi siap digunakan melalui ikon di toolbar browser.

### 3. Script Python
Jalankan script melalui terminal:
```bash
python ekstrak_t41_ke_t42.py --pdf "path/to/dokumen_t41.pdf" --template "path/to/template_t42.docx" --output "Laporan_T42.docx"
```
*(Jalankan `python ekstrak_t41_ke_t42.py --help` untuk opsi lengkap)*
