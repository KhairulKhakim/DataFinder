# KampusReach Dashboard

Aplikasi pencarian dan follow-up database kampus dengan alur yang mengikuti situs referensi: identitas PIC, Search Projects, hasil per proyek, exception-based review, Master Database, Download Center, dan Settings & Safety. Fitur Buyer Intelligence tidak disertakan.

## Membuka

Buka `index.html` di browser. Untuk preview melalui server lokal, jalankan server statis dari folder ini lalu buka alamat yang diberikan server.

Data awal di `app.js` berisi contoh kontak. Gunakan tombol **Import CSV/TSV** pada halaman Master Database untuk menambahkan data sendiri. Proyek, hasil review, pengaturan, dan riwayat unduhan disimpan di browser menggunakan local storage.

Pencarian baru memakai Google Apps Script sebagai crawler serverless dan Google Sheets sebagai penyimpanan. Pengguna memasukkan URL publik; crawler membaca halaman kontak/profil yang relevan, mengekstrak email/telepon/alamat, lalu menyimpan hasilnya di spreadsheet. Website tidak melakukan crawling ke LinkedIn, halaman login, CAPTCHA, jaringan lokal, atau halaman yang dilarang `robots.txt`.

## Menghubungkan ke Google Sheets

1. Buka spreadsheet tujuan, pilih **Ekstensi → Apps Script**.
2. Salin isi `google-apps-script.gs` ke editor Apps Script, lalu simpan.
3. Jalankan fungsi `setupCrawler` satu kali dan izinkan akses. Fungsi ini membuat tab `Search_Projects`, `Crawl_Queue`, dan `Crawl_Results`, serta pemicu lima-menit.
4. Pilih **Deploy → Manage deployments**, buat versi baru untuk deployment Web App yang sudah ada, dan gunakan akses publik yang sama.
5. Buka **Settings & Safety** di dashboard untuk mengetes URL Web App berakhiran `/exec`.

Setiap pencarian baru disimpan di `Search_Projects`; URL masuk ke `Crawl_Queue`; kontak hasil ekstraksi masuk ke `Crawl_Results`; dan ringkasannya tetap dicatat di `Sheet1`.
