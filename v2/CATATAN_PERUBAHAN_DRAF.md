# Draft, belum dipublikasikan — Catatan Perubahan: PE Checker v2 (dibanding PE Checker v1)

**Status:** Draft, belum dipublikasikan. Draf ini tidak dikirim, tidak diunggah, dan tidak ditandai sebagai rilis.

## Ringkasan

PE Checker v2 mengganti tampilan editor Program dengan versi baru yang lebih cepat untuk Program dengan banyak step (40–70+): daftar step di kiri dengan pencarian dan filter, editor di kanan hanya untuk step yang dipilih, error tampil langsung di bawah isian, dan penyimpanan CSV dapat dipratinjau sebelum disimpan. Beberapa bug dari versi lama diperbaiki, dan cara program menyimpan file dibuat lebih aman. Untuk pengguna, cara kerja inti tidak berubah: format file config.json, session.json, operators.json, hasil scan dan validasi, isi CSV dan .dat, password, dan Link IP semuanya tetap sama, jadi file lama bisa langsung dipakai di folder yang sama dengan exe. Semua item di bawah berasal dari laporan Dev/Reviewer (73 tes lulus di wine); pengujian pada Windows asli belum dilakukan.

## Baru

Bukti semua item: dari laporan Dev/Reviewer (dites di wine, 73 tes lulus).

- **Navigasi samping baru di editor Program.** Menu samping berisi: Step, Penyimpanan CSV, File .dat, Umum.
- **Daftar step di kiri dengan editor di kanan.** Daftar step bisa dicari, difilter per jenis step, dan difilter "yang bermasalah" (step yang punya kesalahan). Editor kanan hanya menampilkan step yang dipilih — lebih cepat untuk Program 40–70+ step.
- **Aksi pada daftar step:** Tambah, Duplikat, Geser atas, Geser bawah, menu ⋯ (Salin, Tempel, Hapus, Alat cepat), dan Urungkan perubahan. Perintah lama "copy step 2 for step 3 to 73" tetap ada, sekarang di menu Alat cepat.
- **Tiap Data tampil sebagai kartu yang bisa dilipat.** Tiga validasi (Excel, Karakter, Unik) masing-masing bisa dilipat, punya saklar Aktif, dan ringkasan satu baris. Chip E/K/U di daftar step menandai validasi yang aktif.
- **Error tampil langsung di bawah isian**, bukan sebagai pop-up. Tombol Simpan nonaktif selama masih ada isian yang salah. Ada tautan "Ke masalah pertama ›" di bar bawah, dan lencana jumlah masalah di menu samping.
- **Penyimpanan CSV diperbarui:** pratinjau nama file langsung, daftar kolom dengan kotak cari serta tombol Pilih semua/Kosongkan, dan "Isi tetap" untuk Lini dan Hasil Tes.
- **Peringatan jika data terpotong** — batas 12 Data per step dan 1000 step tidak berubah, tetapi kini ada peringatan yang menyebut jumlah data yang terpotong saat dialog dibuka (sebelumnya terpotong diam-diam).
- **Pop-up konfirmasi diganti pesan status kecil** di bar bawah, sehingga alur kerja tidak terputus.
- **Tampilan jendela utama dipoles ringan:** tombol, kotak isian, dan header tabel. Alur scan operator tidak berubah.

## Diperbaiki

Bukti semua item: dari laporan Dev/Reviewer (bug dari versi lama, dites di wine).

- Pindah tab yang salah saat validasi gagal tidak terjadi lagi.
- Pesan error Excel yang kosong tidak terjadi lagi.
- File Excel yang rusak tidak lagi membuat program berhenti (crash) saat menekan Simpan.
- Karakter `"` sekarang ditolak pada teks nama file CSV (sebelumnya bisa menyebabkan masalah).
- File config.json atau session.json yang rusak, atau tanpa bagian "steps", tidak lagi membuat program crash.
- Tanggal pada nama file mengikuti waktu siklus selesai (sebelumnya tidak).
- Pencarian database Excel lebih cepat, dengan hasil yang sama seperti sebelumnya.

## Tidak berubah (penting untuk pengguna)

Bukti: dari laporan Dev/Reviewer.

- Format file **config.json, session.json, operators.json** — tidak diubah.
- **Hasil scan, validasi, isi CSV, dan isi .dat** — tidak berubah.
- **Password** — tidak berubah. Password disimpan sebagai hash secara internal, tetapi password yang dipakai TETAP SAMA seperti sebelumnya.
- **Link IP** — tidak berubah.
- **File lama bisa langsung dipakai:** letakkan config.json, session.json, dan operators.json lama di folder yang sama dengan PE_Checker_v2.exe.

### Perubahan perilaku kecil yang disengaja

Bukti: dari laporan Dev/Reviewer.

- Menyalin step atau mengganti jenis step sekarang menyegarkan daftar kolom CSV. Versi lama bisa menyimpan label usang seperti "Kode produk 2 (Step 3)" atau kehilangan kolom baru.
- Hapus, geser, atau duplikat step memindahkan pilihan kolom dan nama file CSV ke step yang benar. Pilihan milik step yang dihapus dibuang, dengan pesan.

## Keamanan dan ketahanan (tanpa perubahan perilaku yang terlihat)

Bukti: dari laporan Dev/Reviewer.

- Penyimpanan JSON memakai file sementara unik — aman bila beberapa PC menyimpan bersamaan.
- JSON yang rusak atau berisi byte tidak valid tidak membuat program crash.
- Nama file CSV/.dat disanitasi, termasuk nama khusus Windows (NUL, CON, dan lainnya), dan dicek agar tidak keluar dari folder tujuan.
- Link IP memanggil ping melalui path penuh di Windows.

## Perlu dikonfirmasi sebelum rilis

Keputusan berikut belum diambil user. Jangan dianggap sudah selesai. Bukti: dari laporan Dev/Reviewer.

| # | Hal yang perlu diputuskan | Status |
|---|---|---|
| 1 | Netralisasi formula CSV untuk barcode berawalan `=`, `+`, `@` | Belum diputuskan |
| 2 | Mengganti password "1234", menambah lockout, dan password untuk Reset Total | Belum diputuskan |
| 3 | Tanda tangan (signature) file config.json | Belum diputuskan |
| 4 | Kunci file CSV antar-PC | Belum diputuskan |

## Belum diuji

Bukti: dari laporan Dev/Reviewer. Semua pengujian selama ini di wine (73 tes lulus).

- Windows asli: DPI, font Segoe UI, tema gelap.
- Scanner fisik.
- Validasi Excel dengan file milik user.
- Link IP sungguhan.
- Folder share jaringan yang lambat.
- Peringatan SmartScreen/antivirus — exe tidak ditandatangani. Dibangun dengan PySide6 6.7.3, berjalan di Windows 10/11.

## Checklist uji penerimaan

Pemilik dan tanggal belum ditentukan; diisi oleh user sebelum rilis.

| # | Langkah uji | Hasil yang diharapkan (Ya/Tidak) | Pemilik | Tanggal |
|---|---|---|---|---|
| 1 | Buka PE_Checker_v2.exe dengan config.json, session.json, operators.json lama di folder yang sama | Program terbuka dan semua data lama terbaca: Ya | User (belum ditentukan) | belum ditentukan |
| 2 | Buka Program dengan 40–70+ step, cari dan filter daftar step di kiri | Daftar cepat dan filter bekerja; editor hanya menampilkan step terpilih: Ya | User (belum ditentukan) | belum ditentukan |
| 3 | Isi satu isian dengan nilai salah, lalu lihat bar bawah | Error tampil di bawah isian, Simpan nonaktif, tautan "Ke masalah pertama ›" menuju kesalahan: Ya | User (belum ditentukan) | belum ditentukan |
| 4 | Buka Penyimpanan CSV, lihat pratinjau nama file, cari kolom, Pilih semua/Kosongkan, isi "Isi tetap" | Pratinjau nama sesuai, kolom dan Isi tetap benar: Ya | User (belum ditentukan) | belum ditentukan |
| 5 | Simpan CSV, lalu bandingkan isinya dengan hasil PE Checker v1 untuk Program yang sama | Isi CSV identik dengan v1 (kecuali perbaikan daftar kolom yang disengaja): Ya | User (belum ditentukan) | belum ditentukan |
| 6 | Jalankan scan penuh dengan scanner fisik | Hasil scan sama seperti v1, nama file CSV memakai tanggal waktu siklus selesai: Ya | User (belum ditentukan) | belum ditentukan |
| 7 | Masuk dengan password lama | Password lama masih diterima: Ya | User (belum ditentukan) | belum ditentukan |
| 8 | Gunakan Link IP ke perangkat sungguhan | Ping/Link IP bekerja seperti di v1: Ya | User (belum ditentukan) | belum ditentukan |
| 9 | Salin step dan ganti jenis step, lalu cek daftar kolom CSV | Daftar kolom ikut tersegarkan, tidak ada label usang: Ya | User (belum ditentukan) | belum ditentukan |
| 10 | Jalankan exe pertama kali di Windows asli (perhatikan SmartScreen/antivirus) | Program berjalan; catat apakah muncul peringatan SmartScreen/antivirus: Ya | User (belum ditentukan) | belum ditentukan |

## Fakta berkas

- PE_Checker_v2.exe (~35 MB, satu file) dan PE_Checker_v2.py.
- Letakkan config.json, session.json, dan operators.json lama di folder yang sama dengan exe.
