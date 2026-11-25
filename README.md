# Ekstraktor Ruang Lingkup T.4.1 ke T.4.2

Tool otomasi untuk mengekstrak poin ruang lingkup pengujian (fungsi-fungsi yang diuji) dari dokumen **Rencana SIT (T.4.1)** dan menginjeksinya langsung ke dokumen template **Laporan SIT (T.4.2 - DOCX)** atau menyalinnya dalam format teks/rich-text ke clipboard.

> [!IMPORTANT]
> **REKOMENDASI FORMAT INPUT:**
> **Sangat disarankan mengutamakan upload file T.4.1 dalam format DOCX (`.docx`)**.  
> Dengan format Word (`.docx`), sistem dapat membaca langsung pohon OpenXML sehingga **bold, italic, spasi, bullet point, dan tata letak paragraf dijamin 100% presisi sama persis seperti aslinya**.  
> Jika Anda hanya memiliki file PDF bertandatangan (`.pdf`), sistem tetap menyediakannya melalui *pipeline ekstraksi cerdas* (rekonstruksi font & koordinat posisi).

---

## 🚀 Fitur Utama

- **Dual-Input (DOCX & PDF)**:
  - **DOCX Input (100% Presisi)**: Mengekstrak teks beserta formatting bawaan dokumen Word asli.
  - **PDF Input (Ekstraksi Cerdas)**: Mendeteksi font (bold/italic) dan koordinat posisi teks.
- **Copy with Formatting**: Teks hasil ekstraksi disalin ke clipboard dalam format Rich Text (HTML) dan Plain Text, sehingga saat di-paste (`Ctrl+V`) ke Microsoft Word atau Google Docs, tebal (*bold*) dan miring (*italic*) langsung menempel otomatis.
- **Rich Text Editor**: Area preview hasil ekstraksi dapat langsung diedit dengan tampilan tebal/miring sebelum disalin atau diinjeksi.
- **Injeksi Langsung ke DOCX T.4.2**: Pengisian otomatis tabel ruang lingkup pengujian ke file template Word T.4.2 (.docx) dengan tetap mempertahankan format tebal/miring.
- **Multi-Platform**:
  1. **Web App Standalone** (`app.html`)
  2. **Ekstensi Browser / Chrome Extension** (folder `extension/`)
  3. **Script Python CLI** (`ekstrak_t41_ke_t42.py`)

---

## 📂 Struktur Proyek

```text
├── app.html                  # Aplikasi web standalone (mendukung DOCX & PDF)
├── popup.css                 # Styling modern untuk Web App & Extension
├── popup.js                  # Engine ekstraksi DOCX (JSZip) & PDF (PDF.js)
├── ekstrak_t41_ke_t42.py      # Script Python CLI dengan dukungan DOCX & PDF
├── libs/                     # Pustaka JavaScript (PDF.js, JSZip)
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
2. Upload dokumen T.4.1 (pilih **.docx** jika ada, atau **.pdf**).
3. Review hasil ekstraksi di kolom editor rich text (Anda bisa mengedit teksnya secara langsung).
4. Klik **Copy with Formatting** untuk menyalin ke clipboard (format bold/italic akan ikut saat di-paste di Word).
5. (Opsional) Upload template DOCX T.4.2 lalu klik **Injeksi & Download File T.4.2**.

### 2. Browser Extension (Chrome / Edge)
1. Buka `chrome://extensions/` atau `edge://extensions/`.
2. Aktifkan **Developer mode** (Mode pengembang) di pojok kanan atas.
3. Klik tombol **Load unpacked** (Muat ekstensi yang belum dikemas).
4. Pilih folder `extension` dari repository ini.
5. Ekstensi siap digunakan melalui ikon di toolbar browser.

### 3. Script Python
Jalankan script melalui terminal:
```bash
# Otomatis deteksi file DOCX (diutamakan) atau PDF di folder saat ini:
python ekstrak_t41_ke_t42.py

# Atau tentukan path file secara manual:
python ekstrak_t41_ke_t42.py --t41 "path/to/dokumen_t41.docx" --t42 "path/to/template_t42.docx"
```
*(Jalankan `python ekstrak_t41_ke_t42.py --help` untuk opsi lengkap)*
