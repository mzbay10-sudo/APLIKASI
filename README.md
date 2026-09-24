# Flow Bot (Google Chrome dengan profil terpisah)

## Aplikasi FlowBot.exe

Buka `FlowBot.exe` (atau `run-app.ps1` selama pengembangan). Pilih profil di bagian atas, lalu pilih beberapa Excel/CSV sekaligus. Daftar dapat diurutkan dengan tombol Naik/Turun. Aplikasi langsung menampilkan jumlah prompt per file dan total seluruh prompt, kemudian panel Progres menunjukkan nama file, urutan prompt, serta nomor baris Excel yang sedang dikerjakan.

Profil awal bernama **Renegade Immortal** dan memakai login lama di `runtime/browser-profile`. Tombol **Tambah profil** membuat penyimpanan login baru yang terisolasi; sesudah menambah profil, pilih **Login profil** satu kali.

Excel tidak perlu diubah manual menjadi CSV. Bot membaca `.xlsx`, `.xlsm`, dan `.csv` secara langsung. Beberapa file diproses bertahap sesuai urutan daftar. Untuk `RG_848-850.xlsx`, hasil masuk ke `downloads/RG_848-850/` dengan nama `RG_848-850 1.jpeg`, dan untuk `RG_851-854.xlsx` masuk ke `downloads/RG_851-854/` dengan penomoran sendiri.

Untuk membuat ulang EXE setelah kode diperbarui, jalankan `build-exe.ps1` (menghasilkan `FlowBot-v7.exe`).

Bot ini membaca setiap baris CSV/XLSX dan menjalankan field yang memiliki nilai. Nilai kosong, `0`, `None`, `null`, `nil`, `NaN`, `N/A`, atau `-` dilewati. Jadi `Character = 0` tidak membuka atau menggeser pemilih karakter, sedangkan Artifact/Skill tetap dapat diproses.

Setelah Generate, bot menunggu dua gambar baru selesai lalu menyimpan JPEG 1K ke `downloads/RG_848-850`. Penomoran mengikuti urutan output: baris data pertama menjadi `RG_848-850 1.jpeg` dan `RG_848-850 2.jpeg`, baris berikutnya menjadi nomor 3 dan 4, dan seterusnya.

Pemrosesan selalu berurutan per baris: Create, pantau persentase kedua hasil, tunggu proses selesai, unduh hasil yang berhasil, lalu pindahkan hasil yang sudah tersimpan ke sampah. Bot tidak mengirim prompt berikutnya saat hasil sebelumnya masih dibuat atau belum diunduh.

## Jadwal model otomatis

Bot memilih model berdasarkan jumlah klik Generate yang benar-benar berhasil dikirim, bukan nomor baris Excel:

- Generate ke-1 sampai ke-100 memakai **Nano Banana 2**.
- Setelah generate ke-100 selesai diproses, bot me-reload Google Flow satu kali dan menunggu halaman pulih. Kondisi siap diverifikasi dari `document.readyState`, hilangnya indikator loading, editor prompt yang dapat diedit, tombol tambah referensi, dan pemilih model yang stabil selama beberapa pemeriksaan berturut-turut.
- Setelah halaman siap, bot memilih dan memverifikasi **Nano Banana 2 Lite**. Semua generate berikutnya memakai model Lite.
- Jika reload lambat, bot dapat menunggu hingga 10 menit per percobaan dan mencoba pemulihan tiga kali. Bot berhenti jika model Lite tidak dapat diverifikasi agar prompt berikutnya tidak terkirim dengan model yang salah.

Pengaturan ini ada di bagian `model_schedule` pada `config.json`. Pemilihan dilakukan dari teks/label kontrol (`Nano Banana Pro`, `Nano Banana 2`, `Nano Banana 2 Lite`) dan tidak bergantung pada koordinat layar. Pemilihan model berlangsung sebelum referensi karakter dan teks prompt dipasang, sehingga alur reference image tetap utuh.

Google Chrome yang sudah terpasang di Windows dijalankan oleh Playwright dengan profil persisten di `runtime/browser-profile`. Folder ini hanya milik bot dan tidak memakai profil, history, extension, atau sesi Chrome pribadi.

## Kemudahan saat generate (v8)

- **Perkiraan sisa waktu:** baris status di bawah menampilkan `sisa ±X menit` dari rata-rata waktu per prompt.
- **Ulang otomatis sekali:** setelah semua file selesai, scene yang `GAGAL` (error sementara) langsung dicoba sekali lagi dengan profil yang sama. Kalau tidak ada satu pun scene yang berhasil, pengulangan tidak dilakukan karena masalahnya bukan sementara. Bisa dimatikan dengan `"auto_retry_failed": false` di `config.json`.
- **Ringkasan akhir:** jumlah SELESAI / SEBAGIAN / GAGAL / DITOLAK / BELUM per Excel ditampilkan di panel Progres dan jendela Selesai.

## Uji coba tanpa Google Flow

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Menguji pembacaan Excel, dry-run, penanda Excel (termasuk saat Excel terbuka), sinkron karakter antar profil dengan profil tiruan, serta alur aplikasi (pindah profil, ulang otomatis, ringkasan) dengan bot tiruan. Untuk ikut menguji Excel sungguhan (memakai salinan, file asli tidak diubah): `$env:FLOWBOT_SAMPLE_XLSX = "data\Episode_041.xlsx"` sebelum perintah di atas.

## Tanda di Excel + narasi (v7)

Setiap baris yang selesai langsung ditandai di Excel sumber, pada 4 kolom baru di sebelah kanan:

| Kolom | Isi |
|---|---|
| `STATUS` | `SELESAI (2/2)` hijau, `SEBAGIAN (1/2)` kuning, `GAGAL (0/2) - alasan` merah |
| `FILE GAMBAR` | Nama file gambar untuk scene itu, misalnya `Episode_041 S001-1.jpeg` dan `Episode_041 S001-2.jpeg` |
| `NARASI` | Teks narasi dari .docx sesuai `Source Ref` (`¶0`, `¶16b`, `¶102-103`, ...) |
| `WAKTU` | Waktu baris ditandai |

- **Nama file per scene:** `downloads/Episode_041/Episode_041 S001-1.jpeg`. Nomor `S001` diambil dari kolom `Scene #`, jadi nomornya tidak bergeser walaupun ada baris yang gagal. Kalau tidak ada kolom `Scene #`, nomor dihitung dari baris Excel (baris 2 = S001).
- **Narasi:** letakkan file .docx narasi di folder yang sama dengan Excel. Bot memilih .docx yang namanya memuat nama Excel atau nomor episodenya, misalnya `Episode_041.xlsx` dengan `BTS_MODE_FINAL_Episode_041_TTS_SAFE.docx`. `¶N` menunjuk ke paragraf ke-N (mulai 0). Source Ref tanpa nomor, seperti `¶soulsearcha`, dibiarkan kosong.
- **Lanjut otomatis:** saat dijalankan ulang, baris `SELESAI`/`SEBAGIAN` dilewati dan hanya baris `GAGAL` atau yang belum bertanda yang diproses. Untuk mengulang satu scene, kosongkan sel `STATUS`-nya.
- **Excel sedang dibuka:** saat klik Mulai, aplikasi mengingatkan kalau Excel masih terbuka. Kalau tetap lanjut, tanda disimpan sementara ke `downloads/<nama>/<nama>_TANDA.xlsx`. Pada run berikutnya tanda itu otomatis digabung ke Excel asli dan scene yang sudah selesai **tidak di-generate ulang**. Setelah tergabung, file `_TANDA` dihapus.
- **Cadangan:** sebelum tanda pertama ditulis, salinan Excel asli disimpan sebagai `downloads/<nama>/<nama>_ASLI.xlsx`.
- **Format kolom:** Excel dengan header `Prompt` yang tidak berada di kolom pertama (format STORY MODE: `Scene #`, `Source Ref`, `Prompt`, `Character 1..3`) dibaca berdasarkan nama header. Format lama (prompt di kolom 1) tetap dibaca per posisi seperti sebelumnya. `Key Artifact` tidak dipasang sebagai referensi.

## Urutan CapCut dari narasi Word + SRT (v7)

Tombol **Buat Urutan CapCut** menambahkan sheet `Urutan CapCut` (Urutan | Nama file | Mulai | Selesai | Durasi | Catatan) ke Excel yang dipilih. Sheet ini langsung dibaca CapCut Studio Sync.

1. Taruh Excel prompt, Word narasi (`..._TTS_SAFE.docx`), dan SRT dubbing (`...Episode_041.srt`) di folder yang sama, misalnya `data/`.
2. Setelah generate selesai, klik Excel-nya di daftar, lalu tekan **Buat Urutan CapCut**.
3. Cara kerjanya:
   - Setiap kata SRT dicocokkan ke paragraf Word, sehingga setiap `¶` punya waktu mulai yang tepat.
   - Scene dengan `¶` yang sama (¶0, ¶0 atau ¶3a, ¶3b) dibagi di awal kalimat SRT, sebanding dengan panjang story cue masing-masing scene.
   - Scene ≥ 6 detik memakai 2 gambar (`-1` lalu `-2`), berganti di awal kalimat atau setelah koma. Scene pendek memakai 1 gambar.
   - Scene tanpa gambar (GAGAL) durasinya disambung ke gambar sebelumnya, jadi timeline tetap tanpa celah dan berakhir di subtitle terakhir.
4. Di CapCut Studio Sync, masukkan gambar dari `downloads/<nama Excel>/` ke timeline, lalu pilih Excel dan SRT yang sama. Gambar yang tidak tercantum akan dikeluarkan dari timeline oleh aplikasi itu.
5. Kalau ada scene yang di-generate ulang, tekan tombol ini lagi supaya sheet diperbarui.

Dari command line: `.\.venv\Scripts\python.exe capcut_plan.py data\Episode_041.xlsx` (tambahkan `--check` untuk hanya memeriksa).

## URL project per profil (wajib untuk profil baru)

Project Google Flow hanya bisa dibuka oleh akun Google pemiliknya. Kalau profil lain (misalnya **BTS**) membuka URL project milik akun Renegade Immortal, Flow menampilkan **Project not found** (`flow.google.com/404?reason=project`).

- Klik **URL project** di bagian Profil, lalu tempel URL project Flow milik akun profil itu. Pakai project yang berisi karakter referensinya (Xu Qing, dst.). URL disimpan di `profile_urls.json` per profil.
- Saat pertama dibuka, URL lama dari `config.json` otomatis dipakai untuk profil yang memakai `runtime/browser-profile` (Renegade Immortal).
- Kalau URL belum diatur, aplikasi memintanya sebelum mulai. Kalau project tidak ditemukan, bot langsung berhenti dengan pesan yang jelas, tidak menunggu sampai timeout.

## Instalasi Windows

1. Pastikan Python 3.10+ tersedia melalui perintah `py`.
2. Pastikan Google Chrome biasa sudah terpasang.
3. Klik kanan `setup.ps1`, lalu pilih **Run with PowerShell** (atau jalankan dari PowerShell). Instalasi ini tidak mengunduh Chromium.
4. Edit `config.json`, terutama URL dan selector agar sesuai halaman yang tampil.
5. Jalankan `run-login.ps1` satu kali. Skrip ini membuka Chrome secara normal tanpa Playwright agar login Google tidak diblokir. Login, buka Flow sampai berhasil, lalu tutup seluruh jendela Chrome profil bot.
6. Data contoh tersedia di `data/input.csv`. Untuk memakai Excel, letakkan file di `data/input.xlsx` lalu ubah `input_file` di konfigurasi.
7. Jalankan `run-dry.ps1` untuk mengecek pembacaan data tanpa klik.
8. Jalankan `run-test.ps1` untuk mencoba satu baris pertama saja. Jika berhasil, gunakan `run-bot.ps1` untuk seluruh data.

## Kolom data

Untuk format `RG_848-850.xlsx`, kolom `1` adalah prompt, sedangkan kolom `2`, `3`, dan `4` adalah Character 1, Character 2, dan Character 3. Setiap slot karakter yang kosong/0/None dilewati. Field tambahan bisa ditambahkan ke `fields` tanpa mengubah kode.

## Menyesuaikan selector

Selector contoh bersifat titik awal karena skrip AHK lama hanya merekam koordinat mouse, bukan identitas elemen situs. Gunakan Playwright Inspector atau DevTools untuk menemukan selector stabil (`data-testid`, role, atau label lebih baik daripada koordinat). Template selector opsi wajib memakai `{value}`.

Untuk merekam interaksi sebagai referensi selector:

```powershell
.\.venv\Scripts\playwright.exe codegen --channel=chrome --user-data-dir=runtime/codegen-profile https://labs.google/fx/tools/flow
```

Jangan arahkan `profile_dir` atau `--user-data-dir` ke folder Chrome pribadi. Cookie login bot tersimpan lokal dalam `runtime/browser-profile`; lindungi folder itu seperti kata sandi. Nilai `browser_channel` harus tetap `chrome` untuk memakai Google Chrome biasa. Chrome manual dan bot tidak boleh memakai profil ini bersamaan; tutup jendela login sebelum menjalankan `run-bot.ps1`.

## Log dan kegagalan

Log disimpan di `runtime/logs/bot.log`. Jika satu baris gagal, screenshot otomatis masuk ke `runtime/screenshots`. Secara default bot melanjutkan baris berikutnya; ubah `stop_on_error` menjadi `true` untuk berhenti pada kegagalan pertama.

Untuk setiap baris, bot hanya menekan Create satu kali dan mengunduh semua hasil yang berhasil. Jika dua hasil berhasil, keduanya diunduh dan dihapus dari Flow, lalu bot langsung melanjutkan tanpa refresh. Jika hanya satu berhasil, hasil itu diunduh dan dihapus, kemudian halaman direfresh sebelum prompt berikutnya. Jika semuanya gagal, tidak ada file kosong yang dibuat dan halaman direfresh untuk pemulihan. Setelah itu bot terus berjalan sampai data habis.

Kartu aset dengan ikon orang di pojok kiri atas dikenali sebagai **Karakter** dan selalu dilindungi: tidak dianggap sebagai hasil generate, tidak diunduh, dan tidak dihapus. Panel progres mencatat `KARAKTER DILINDUNGI` atau `HASIL GENERATE` untuk memudahkan audit.

## Karakter antar profil (otomatis, tanpa tab terpisah)

Aplikasi hanya punya satu layar: **Generate Gambar**. Tab **Pindah Karakter antar Profil** sudah dihapus karena tidak diperlukan lagi. Cukup pilih Excel lalu klik **Mulai generate**. Karakter yang dipakai Excel tetapi belum ada di profil yang sedang dipakai otomatis dicari di profil lain lalu dibuat di profil ini sebelum generate dimulai (lihat **Sinkron karakter otomatis sebelum generate** di bawah).

Flow berbahasa Inggris maupun Indonesia sama-sama didukung.

## Tahan perubahan tampilan Flow (cara 1 → cara 2 → dst.)

Setiap langkah penting punya beberapa cara. Bot mencoba cara 1 dulu. Kalau gagal, bot lanjut ke cara 2, dan seterusnya. Kalau bot memakai cara cadangan, di Progres muncul `CARA CADANGAN | ...`.

| Langkah | Cara yang dicoba berurutan |
|---|---|
| Buka menu Karakter | menu samping → tautan/tab lain → panel aset (+ > Karakter, tampilan lama) |
| Daftar karakter | grid menu Karakter (gulir sampai habis) → panel aset |
| Unduh gambar karakter | fetch di halaman → request browser → ukuran asli thumbnail → screenshot kartu |
| Buka halaman karakter baru | alamat `/character` → tombol "Karakter baru" |
| Upload gambar | tombol Upload → input file langsung |
| Isi nama | kolom Nama karakter → judul di atas → ikon pensil + ketik → isi langsung kolom nama → isi langsung judul → klik 3x + tempel → set lewat script → muat ulang halaman lalu cara 1 / 2 / script (10 cara, tiap cara dicek dari judul halaman) |
| Pastikan nama tersimpan | cari di daftar karakter → buka lagi halaman karakter, isi ulang nama, buka ulang untuk cek (2 putaran) |
| Hapus draf gagal | tombol Hapus → ikon delete → tombol bertuliskan Hapus → menu opsi; lalu dicek benar-benar terhapus |
| Simpan karakter | Done editing → Selesai/Done → tombol kembali |
| Cek karakter sudah dibuat | kotak cari → daftar lengkap |
| Tampilan generate | Semua media → tautan Semua media → Gambar |
| Mode Agen di prompt | dimatikan otomatis |
| Pemilih model | chip model → klik/luaskan prompt → lanjut dengan model aktif di Flow |

Jendela **"Hak untuk menggunakan gambar ini"** (akun baru) diklik **Saya setuju** otomatis. Kalau tombolnya tidak ketemu, bot menunggu kamu mengklik sendiri.

## Kredit habis / pembatasan & pindah profil otomatis

1. Di **Nano Banana 2**: kalau kredit habis atau kena pembatasan ("aktivitas tidak biasa"), bot langsung pindah ke **Nano Banana 2 Lite** di profil yang sama dan mengulang scene itu.
2. Di **Nano Banana 2 Lite**: kalau kredit habis atau pembatasan terjadi **2 kali berturut-turut**, bot langsung pindah ke profil berikutnya. Kalau percobaan kedua berhasil, bot lanjut biasa dan hitungannya diulang dari nol.
3. Kegagalan lain, misalnya prompt ditolak kebijakan Google Flow, tidak dihitung. Scene itu dilewati dan ditandai, lalu bot lanjut.
4. Kalau **semua profil** kena kredit habis atau pembatasan, bot **berhenti**. Tidak ada jeda 30 menit.
5. **Berkelanjutan:** klik Mulai berikutnya dimulai dari **profil terakhir yang dipakai**, lalu lanjut ke bawah daftar dan memutar ke atas. Setiap profil dipakai sekali per putaran. Profil terakhir disimpan di `profile_rotation.json` (`last_profile`).

Pilih profil yang ikut bergantian di kotak **Pindah profil otomatis**. Karakter otomatis disinkronkan ke tiap profil (lihat bagian di bawah).

## Sinkron karakter otomatis sebelum generate

Setiap klik **Mulai generate**, dan setiap kali bot pindah ke profil lain, bot menjalankan langkah ini dulu:

1. Membaca nama karakter di kolom Character 1..10, hanya dari baris Excel yang belum selesai.
2. Mengecek karakter mana yang sudah ada di profil yang dipakai.
3. Karakter yang belum ada dicari:
   - **Cara 1:** di folder karakter yang sudah pernah diunduh (`downloads\_KARAKTER\<profil>`).
   - **Cara 2:** membuka profil lain satu per satu. Hasil cek disimpan di `downloads\_KARAKTER\_katalog.json`, jadi profil yang baru saja dicek (kurang dari 6 jam) dan tidak punya karakter itu tidak dibuka lagi.
4. Karakter yang ketemu dipindahkan dulu ke profil ini.
5. Karakter yang tidak ada di profil mana pun tidak masalah: generate tetap jalan, karakter diambil dari teks prompt, dan referensinya dilewati tanpa membuka panel.

Kalau sinkron gagal karena alasan apa pun, generate tetap jalan.

Tutup jendela Chrome profil lain saat bot berjalan. Profil yang Chrome-nya sedang terbuka dilewati saat pencarian.

## Karakter tanpa nama ("Karakter tanpa judul")

Kalau upload gambar ditolak Flow atau nama gagal tersimpan, bot sekarang:

1. Mengisi ulang nama (10 cara cadangan) lalu membuka ulang halaman karakter untuk memastikan nama benar-benar tersimpan.
2. Kalau tetap gagal, bot menghapus **hanya draf yang barusan dibuatnya sendiri**, dan hanya kalau judulnya masih kosong. Karakter yang punya nama tidak pernah dihapus.

Untuk membersihkan sisa karakter tanpa nama dari kegagalan sebelumnya, pilih profil di **Profil** (atas), lalu klik **Hapus karakter tanpa nama** (kanan bawah bagian Generate Gambar). Pilih **YES** untuk menghapus, atau **NO** untuk cek saja (hanya dihitung, tidak ada yang dihapus). Bot mengecek label dan gambar kartu (supaya kartu tidak tertukar), membuka setiap kartu tanpa nama, menunggu halamannya termuat, dan membaca namanya 2 kali. Karakter hanya dihapus kalau judulnya memang kosong. Kalau namanya tidak terbaca dengan pasti, karakter itu dilewati.
