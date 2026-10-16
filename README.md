# AutoSensor

AutoSensor adalah macro VBA untuk CorelDRAW yang membantu menyiapkan **sensor hitam** dan **WhiteSpace putih** pada halaman aktif.

Mode **Sensor** membuat empat ellipse dengan diameter yang dapat diatur. Mode **WhiteSpace** membuat objek putih yang dipusatkan pada anchor regmark dan `qr.svg`. Hasil dikumpulkan dalam satu group di `Layer 1`, dengan pengaturan Z-order objek serta pemulihan state dokumen.

Macro menggunakan area virtual **325 × 485 mm** yang dipusatkan pada halaman aktif. Ukuran halaman sebenarnya tetap dipertahankan.

## Main Features

### Sensor

Pilih `optSensor`, lalu tentukan diameter melalui `txbSensorSize`. Process membuat **empat ellipse hitam**, tanpa outline, dengan fill **RGB `#000000`**.

Diameter default adalah **7 mm**. Saat kolom kosong, form menampilkan hint abu-abu `7 mm`; hint hilang saat kolom mendapat fokus dan kembali saat kolom kosong ditinggalkan. Process memakai diameter default jika kolom kosong atau hint masih aktif.

Input diameter harus berupa angka **lebih besar dari 0**, tanpa satuan atau pemisah ribuan. Titik dan koma dapat digunakan sebagai pemisah desimal, misalnya `5.5` atau `5,5`. Jika mengetik nilai sendiri, masukkan `7`, bukan `7 mm`.

Posisi sensor ditentukan oleh diameter:

| Diameter | Pusat empat ellipse `(X; Y)`, dalam mm |
| --- | --- |
| **≤ 5 mm** | `(10; 10)`, `(10; 475)`, `(315; 10)`, `(315; 475)` |
| **> 5 mm** | `(11; 11)`, `(11; 474)`, `(314; 11)`, `(314; 474)` |

Koordinat mengacu pada area virtual, bukan pada selection atau posisi anchor. Mode Sensor tidak memerlukan layer `scpro2_regmarks` maupun `scpro2_printonly`.

### WhiteSpace

Pilih `optWhiteSpace` untuk membuat **empat ellipse putih berdiameter 5,5 mm** dan **satu rectangle putih berukuran 10,5 × 10,5 mm**. Semua objek memakai fill **RGB `#FFFFFF`**, tanpa outline.

Pada mode ini, `txbSensorSize` dinonaktifkan dan nilainya tidak digunakan. Saat kembali ke Sensor, kolom diaktifkan tanpa menghapus nilai yang tersimpan.

Macro mencari anchor pada halaman aktif sesuai tabel berikut:

| Layer anchor | Nama anchor | Pusat area pencarian `(X; Y)`, dalam mm | Ukuran area pencarian | Objek hasil per anchor |
| --- | --- | --- | --- | --- |
| `scpro2_regmarks` | Bebas | `(6; 10)`, `(6; 475)`, `(319; 10)`, `(319; 475)` | Persegi **10 × 10 mm** pada setiap posisi | Ellipse diameter **5,5 mm** |
| `scpro2_printonly` | `qr.svg` | `(162,5; 11,5)` | Persegi **12 × 12 mm** | Rectangle **10,5 × 10,5 mm** |

Ukuran pada kolom area pencarian adalah ukuran **zona pencarian pusat anchor**, bukan ukuran yang wajib dimiliki anchor.

Setiap zona harus memiliki **tepat satu kandidat**. Objek putih ditempatkan pada **pusat bounding box anchor yang ditemukan**, sehingga posisinya mengikuti anchor aktual di dalam zona tersebut.

### Validasi Anchor

Pencarian WhiteSpace mengikuti aturan berikut:

- Pusat bounding box anchor harus berada di dalam atau tepat pada batas zona pencarian. Toleransi numeriknya **0,000001 mm**.
- Empat anchor pada `scpro2_regmarks` tidak dibatasi nama, jenis, atau ukuran.
- Anchor pada `scpro2_printonly` harus bernama `qr.svg`.
- Pencocokan nama layer dan `qr.svg` tidak membedakan huruf besar/kecil.
- Hanya objek yang berada langsung pada layer yang diperiksa. Group dihitung sebagai satu kandidat; anggota group tidak ditelusuri secara terpisah.
- Seluruh lima anchor harus valid sebelum objek hasil pertama dibuat.

Jika layer anchor tidak ditemukan, satu zona tidak mempunyai kandidat, atau satu zona mempunyai lebih dari satu kandidat, Process dibatalkan dengan `MsgBox`. Nama layer yang dicari juga harus unik; layer dengan nama sama lebih dari satu menyebabkan validasi gagal.

Anchor tetap berada pada posisi serta layer asalnya, dan tidak digroup bersama hasil WhiteSpace.

## Area Virtual dan Koordinat

Area virtual **325 × 485 mm** dipusatkan pada halaman aktif. Sudut kiri bawahnya dihitung sebagai:

```text
Origin X = Page.CenterX − 162,5 mm
Origin Y = Page.CenterY − 242,5 mm
```

Semua koordinat dalam tabel mengacu pada sudut kiri bawah area virtual tersebut. **X bertambah ke kanan** dan **Y bertambah ke atas**. Koordinat Sensor adalah pusat objek; koordinat WhiteSpace adalah pusat zona pencarian anchor.

Posisi origin ruler tidak mengubah acuan penempatan terhadap pusat halaman. Macro tidak memeriksa apakah ukuran halaman sama dengan area virtual dan tidak mengubah ukuran halaman. Pada halaman yang lebih kecil, sebagian hasil dapat berada di luar halaman.

## Hasil, Layer, dan State Dokumen

### Penempatan Hasil

- Setiap Process yang berhasil membuat **satu group baru** di `Layer 1` pada halaman aktif. Layer dibuat jika belum ada.
- Process berulang menambahkan group baru; hasil yang sudah ada tidak diganti.
- Group hasil ditempatkan paling atas di dalam `Layer 1`.
- Urutan `Layer 1` terhadap layer lain tetap dipertahankan. Pengaturan Z-order hanya menaikkan group hasil di dalam `Layer 1`.

### Pemulihan State

Pemrosesan menggunakan satuan milimeter sementara. Setelah selesai, macro memulihkan:

- Layer aktif.
- Satuan dokumen serta satuan ruler horizontal dan vertikal.
- Status `Editable` layer anchor dan hasil yang dibuka sementara.

Setelah generate berhasil, hanya **group hasil** yang terseleksi; objek anggota group dan selection awal tidak ikut terseleksi. Ukuran halaman, posisi dan isi anchor, serta flag `Visible`/`Printable` layer tetap dipertahankan.

Seluruh perubahan dokumen dikelompokkan dalam satu `BeginCommandGroup`/`EndCommandGroup`, termasuk perubahan sementara satuan dan status kunci layer. Hasil Process yang selesai normal dapat dibatalkan dengan **satu Undo**.

### Penanganan Kegagalan

Jika terjadi kegagalan, macro mencoba menghapus objek atau group hasil parsial dan menghapus `Layer 1` jika layer itu baru dibuat oleh Process. Selection, layer aktif, satuan, serta status kunci layer juga dicoba dipulihkan.

`MsgBox` menampilkan tahap operasi, nomor error, sumber error, dan penjelasannya. Kegagalan pemulihan dilaporkan bersama pesan pembatalan agar hasil di dokumen dapat diperiksa sebelum melanjutkan.

## Kontrol Form

| Kontrol | Fungsi |
| --- | --- |
| `optSensor` | Memilih pembuatan empat sensor hitam. Dipilih secara bawaan jika belum ada pilihan mode saat form dibuka. |
| `optWhiteSpace` | Memilih pembuatan objek putih berdasarkan lima anchor. |
| `txbSensorSize` | Diameter sensor dalam mm; hanya digunakan pada mode Sensor. |
| `cmdProcess` | Menjalankan mode terpilih pada halaman aktif. Tombol dinonaktifkan selama pemrosesan untuk mencegah Process berulang. |
| `cmdClose` | Menutup form ketika tidak sedang memproses. |

Process memerlukan tepat satu mode terpilih. Setelah Process selesai, form tetap terbuka untuk digunakan kembali.

## Cara Menggunakan

### Membuat Sensor

1. Buka dokumen CorelDRAW dan aktifkan halaman tujuan. Tidak perlu memilih objek.
2. Jalankan macro `AutoSensor` untuk membuka `AutoSensorMenu`.
3. Pilih `optSensor`.
4. Biarkan diameter kosong untuk memakai **7 mm**, atau masukkan diameter positif pada `txbSensorSize`.
5. Tekan `cmdProcess`, lalu periksa empat sensor hitam di `Layer 1`.
6. Gunakan `cmdClose` untuk menutup form.

### Membuat WhiteSpace

1. Aktifkan halaman yang mempunyai layer `scpro2_regmarks` dan `scpro2_printonly`.
2. Pastikan empat anchor regmark berada pada zona pencarian masing-masing dan anchor bernama `qr.svg` berada pada zona pencarian `scpro2_printonly`.
3. Jalankan `AutoSensor` dan pilih `optWhiteSpace`. Tidak perlu memilih anchor secara manual.
4. Tekan `cmdProcess`.
5. Periksa empat ellipse dan satu rectangle putih yang mengikuti pusat anchor. Jika muncul pesan pembatalan, sesuaikan layer atau kandidat sesuai pesan tersebut sebelum mencoba kembali.

## Pemasangan dan Source

Repository menyediakan standard module, macro pembuka, dan code-behind UserForm. File `.vba` berisi kode form; layout visual UserForm lengkap tidak disertakan.

1. Import [AutoSensorGenerator.bas](src/modules/AutoSensorGenerator.bas) sebagai standard module pada project GMS CorelDRAW.
2. Siapkan UserForm bernama `AutoSensorMenu`, buat kontrol sesuai tabel **Kontrol Form**, lalu pasang kode dari [AutoSensorMenu.vba](src/forms/AutoSensorMenu.vba).
3. Import [AutoSensorWizard.bas](src/wizard/AutoSensorWizard.bas) ke project yang sama.
4. Jalankan macro publik `AutoSensor`. Form dibuka dengan `vbModal`.

| Lokasi | Peran |
| --- | --- |
| `src/forms/AutoSensorMenu.vba` | Pilihan mode, input diameter dan hint, Process, serta Close. |
| `src/modules/AutoSensorGenerator.bas` | `ASGenerate`, validasi ukuran/layer/anchor, pembuatan dan pengelompokan hasil, Z-order objek, serta pemulihan state dokumen. |
| `src/wizard/AutoSensorWizard.bas` | Entry point `AutoSensor` untuk membuka form. |

## Lisensi dan Masukan

Project ini menggunakan **MIT License**; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
