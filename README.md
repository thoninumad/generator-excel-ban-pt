# generator-excel-ban-pt
generator an excel file from database based on SAPTO template (IAPS 4.0) using phpspreadsheet

Hak Cipta oleh pengembang (thoninumad) bersama tim Direktorat Pengembangan Teknologi dan Sistem Informasi (DPTSI) ITS

Keterangan:

1. Folder blank berisi file excel kosong yang akan menampung data dari database. Struktur kolomnya sudah disesuaikan dengan kolom-kolom apa saja yang akan ditarik dari db dan kolom-kolom apa saja yang dibutuhkan oleh LKPS.
2. Folder raw berisi file excel yang sudah terisi data mentah dari database. Data raw dari database sudah terfilter berdasarkan id prodi, dan/atau kolom kategorisasi lain yang dianggap perlu untuk menyesuaikan dgn kebutuhan tabel LKPS.
3. Folder formatted berisi file excel yang sudah dilakukan perubahan, seperti perubahan struktur kolom atau isian data menyesuaikan kebutuhan LKPS.
4. Folder result berisi file excel LKPS yang sudah terisi dari excel-excel di folder raw atau formatted. (Bila lgsg dari folder raw, kebutuhan datanya berarti sudah cocok dari db nya)
5. Terdapat file sapto_aps9.xlsx di folder blank. File itu merupakan file excel kosong LKPS (tanpa ada contoh isian data dari BAN-PT) yang siap menampung data-data dari excel.
6. Menjalankan file sapto_aspek1.php hingga sapto_aspek10.php (terdapat 10 aspek dari 8 kriteria borang LKPS)
7. Running file sapto_aspek1.php akan menggunakan sapto_aps9.xlsx di folder blank (memulai dari awal/kosongan) untuk diisi. 
8. Running file sapto_aspek2.php sampai sapto_aspek10.php akan menggunakan sapto_aps9 (F).xlsx di folder result (hasil isian dari aspek1 atau aspek-aspek sebelumnya) untuk diisi.
9. Bila ingin digunakan untuk laporan id prodi lain, yg harus dijalankan dulu yaitu sapto_aspek1.php (mengisi dari excel kosongan). Baru kemudian aspek-aspek lain (boleh acak, aspek9 dulu baru aspek8, dll) karena phpspreadsheet sifatnya menimpa dan menambahkan, tidak bakal menghapus isian data yang lama jika tidak dalam satu cell yang sama dengan data baru di excel.
10. Project bisa didownload atau clone dari github.com/thoninumad/generator-excel-ban-pt
11. Setelah diclone, bisa dirun perintah "composer install" (tanpa tanda petik) di cmd direktori project
12. Sementara ini id prodi diatur dalam global variable di tiap aspek. Namun, sudah disediakan pula perintah cmd dgn menggunakan parameter id prodi, sehingga mempermudah eksekusi php tiap aspek sesuai prodi yg diinginkan. Opsi tersebut di-comment di code bagian atas.
