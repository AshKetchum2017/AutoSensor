# AutoSensor

Macro CorelDRAW VBA untuk membuat sensor hitam atau objek WhiteSpace putih pada halaman aktif. Area virtual **325 Ã— 485 mm** dipusatkan pada halaman aktif tanpa memeriksa atau mengubah ukuran halaman sebenarnya.

## Koordinat

Semua koordinat di bawah adalah pusat objek atau pusat area pencarian, dalam milimeter, relatif terhadap sudut kiri bawah area virtual. X bertambah ke kanan dan Y bertambah ke atas. Posisi origin ruler tidak menjadi titik acuan macro.

Sudut kiri bawah area virtual dihitung dari pusat halaman dikurangi `(162.5, 242.5)` mm. Jika halaman lebih kecil dari area virtual, sebagian hasil dapat berada di luar halaman.

## Mode Sensor

Pilih `optSensor`, lalu isi diameter di `txbSensorSize`. Default adalah **7 mm**: textbox kosong menampilkan hint abu-abu `7 mm`, yang hilang saat mendapat fokus dan kembali saat textbox kosong ditinggalkan. Process memakai 7 mm bila input kosong atau hint aktif. Angka yang diisi pengguna menggantikan default. Ukuran harus lebih dari 0 mm. Desimal titik maupun koma diterima; jangan gunakan satuan atau pemisah ribuan.

| Diameter | Empat pusat ellipse (X, Y) |
| --- | --- |
| â‰¤5 mm | (10, 10), (10, 475), (315, 10), (315, 475) |
| >5 mm | (11, 11), (11, 474), (314, 11), (314, 474) |

Keempat ellipse berdiameter sesuai input, memakai fill **RGB #000000**, tanpa outline.

## Mode WhiteSpace

Pilih `optWhiteSpace`. `txbSensorSize` dinonaktifkan dan nilainya tidak dipakai; kembali ke Sensor mengaktifkan textbox tanpa menghapus nilainya.

| Layer anchor | Pusat area pencarian (X, Y) | Area pencarian | Hasil |
| --- | --- | --- | --- |
| `scpro2_regmarks` | (6, 10), (6, 475), (319, 10), (319, 475) | Persegi 10 Ã— 10 mm pada setiap posisi | Empat ellipse diameter 5.5 mm |
| `scpro2_printonly` | (162.5, 11.5) | Persegi 12 Ã— 12 mm | Satu rectangle 10.5 Ã— 10.5 mm |

Anchor pada `scpro2_printonly` harus bernama `qr.svg`. Keempat anchor regmark tidak dibatasi nama, jenis, atau ukuran. Pusat bounding box anchor harus berada di dalam atau tepat di batas area pencarian, dengan toleransi numerik 0.000001 mm. Group dihitung sebagai satu kandidat; anggota group tidak diperiksa terpisah.

Seluruh area wajib memiliki tepat satu kandidat. Anchor hilang, kandidat ganda, atau layer anchor yang hilang menyebabkan pembatalan dengan `MsgBox`, sebelum objek hasil dibuat. Hasil dipusatkan pada posisi aktual anchor, memakai fill **RGB #FFFFFF**, tanpa outline. Anchor tidak dipindahkan atau digroup bersama hasil.

## Hasil dan state dokumen

- Setiap Process membuat satu group baru di `Layer 1`; layer dibuat jika belum ada. Process berulang menambahkan group baru.
- Group ditempatkan paling atas di `Layer 1`. `Layer 1` juga dinaikkan ke atas layer biasa lainnya pada halaman aktif; isi lama layer tersebut ikut naik. Layer sistem tidak dijadikan acuan pemindahan.
- Seluruh perubahan dokumen berada dalam satu `BeginCommandGroup`/`EndCommandGroup`, termasuk perubahan sementara unit dan status kunci layer.
- Layer anchor/hasil yang terkunci dibuka sementara, lalu status `Editable` asal dipulihkan. Selection, layer aktif, unit dokumen, dan kedua unit ruler dipulihkan.
- Jika gagal, macro mencoba menghapus hasil parsial dan layer hasil yang baru dibuat, memulihkan urutan layer asal, lalu menutup CommandGroup. Kegagalan pemulihan ikut dilaporkan di `MsgBox`.
- Ukuran halaman, posisi/isi anchor, serta flag visible/printable layer tidak diubah.

## Pemasangan dan penggunaan

1. Pasang [AutoSensorGenerator.bas](src/modules/AutoSensorGenerator.bas) sebagai standard module pada project GMS yang sama.
2. Perbarui kode UserForm `AutoSensorMenu` dari [AutoSensorMenu.vba](src/forms/AutoSensorMenu.vba). Form menggunakan kontrol `optSensor`, `optWhiteSpace`, `txbSensorSize`, `cmdProcess`, dan `cmdClose`; file `.vba` ini hanya kode form, bukan resource layout.
3. Pasang atau pertahankan [AutoSensorWizard.bas](src/wizard/AutoSensorWizard.bas), lalu jalankan macro publik `AutoSensor`.
4. Pilih mode dan klik Process. Close menutup form.
