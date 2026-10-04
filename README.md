# Format-Guard

Format-Guard adalah aplikasi web untuk memeriksa dan merapikan format dokumen Word (`.docx`). Dokumen diproses di browser pengguna; file tidak dikirim ke server Format-Guard. Tidak ada akun atau basis data yang diperlukan.

**Demo:** https://duhaalul.github.io/Format-Guard/

## Fitur

- Unggah atau tarik file `.docx` hingga 25 MB.
- Periksa margin, ukuran kertas, jenis dan ukuran font, serta spasi baris.
- Rapikan paragraf isi: rata kiri-kanan, inden baris pertama 1,25 cm, dan jarak sebelum/sesudah paragraf 0 pt.
- Periksa dan perbaiki urutan judul yang **sudah memiliki penanda teks** seperti `BAB 1`, `BAB 2`, `1.1`, `1.2`, dan `2.1`.
- Pilih posisi nomor halaman: pertahankan seperti dokumen asal, bawah tengah, atau bawah kanan.
- Lihat temuan per paragraf dan unduh file baru dengan akhiran `_rapi.docx`. File hasil diperiksa ulang sebelum ditawarkan untuk diunduh.

Profil bawaan saat ini bernama **Standar Universitas Andalas**. Nilainya: margin kiri dan atas 4 cm, kanan dan bawah 3 cm; kertas A4; Times New Roman 12 pt; spasi 1,5. Nilai ini adalah preset aplikasi, bukan pengganti pedoman resmi institusi.

## Menjalankan secara lokal

Tidak ada proses build atau instalasi paket. Jalankan server HTTP statis dari direktori proyek:

```powershell
python -m http.server 8765
```

Buka `http://localhost:8765/` di browser. Gunakan server HTTP karena aplikasi memakai modul JavaScript; membuka `index.html` melalui `file://` dapat membuat modul gagal dimuat.

## Cara mencoba

1. Pilih file `.docx` atau gunakan tombol **Coba dengan dokumen contoh**.
2. Pilih profil dan posisi nomor halaman.
3. Tinjau daftar temuan di panel kanan. Nomor halaman dalam laporan adalah perkiraan.
4. Klik **Perbaiki & unduh**.
5. Buka file `_rapi.docx` di Word atau aplikasi lain untuk membandingkan tata letaknya.

## Struktur proyek

| Berkas | Fungsi |
| --- | --- |
| `index.html` | Halaman dan alur antarmuka |
| `styles.css` | Tampilan responsif |
| `app.js` | Interaksi, laporan, dan unduhan |
| `docx.js` | Audit dan perbaikan XML DOCX |
| `jszip.min.js` | Pembacaan dan penulisan paket DOCX |
| `LICENSE-JSZip.md` | Lisensi JSZip |

Aplikasi saat ini menggunakan JavaScript di browser dan JSZip. Rancangan awal pernah menyebut backend Python/FastAPI, tetapi implementasi MVP ini sengaja berjalan tanpa backend agar dokumen tetap di perangkat pengguna.

## Batasan

- Pemeriksaan urutan BAB/subbab berlaku untuk nomor yang tertulis sebagai teks judul. Penomoran otomatis bawaan Word (`numbering.xml`) belum dianalisis atau diubah.
- Judul tanpa penanda `BAB` atau `1.1` tidak diberi nomor otomatis agar isi dokumen tidak salah ditafsirkan.
- Aturan paragraf isi tidak diterapkan pada judul, daftar, tabel, caption, dan elemen khusus lain.
- Pratinjau situs hanya menampilkan teks, bukan hasil tata letak Word yang persis. Nomor halaman dapat berubah setelah dokumen dibuka atau dicetak.
- Dokumen Word dengan elemen kompleks seperti kotak teks, objek tertanam, atau aturan format khusus perlu diperiksa manual setelah diunduh.
- Waktu pemrosesan bergantung pada ukuran dokumen dan kemampuan perangkat; target di bawah 5 detik untuk dokumen kurang dari 50 halaman belum menjadi jaminan.

## Lisensi dependensi

JSZip disertakan bersama berkas lisensinya di `LICENSE-JSZip.md`.
