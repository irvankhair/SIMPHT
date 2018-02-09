# SIMPHT Project
### Tugas Pertemuan Ke-2
- [x] Memanipulasi ulang transfer data nominatif (Excel -> mdb)
- [x] Form 1 DataGrid data nominatif
- [x] Form 2 DataGrid Resume
- [ ] Rapih - Rapih Software
- [x] Kolom Di dalam dan luar trase
- [x] Kolom lain lain



SIMPHT

Kamis, 08 Februari 2018
04.29

SIMPHT

Jumat, 26 Januari 2018
15.15

	1. [x]Input daftar nominatif
		[x]a. Format baku awal
		[x]b. Koneksi Excel ke recorder
		[x]c. Kodifikasi grup nomor induk
	2. Kalkulasi NPW (Nilai Penggantian Wajar)
		[x]a. Fisik
			i. Tanah
				1) Pembagian nilai zona
					a) Pinggir jalan besar
					b) Pinggir jalan kecil
					c) Dalam
				2) Metode input
					a) Menentukan zona dengan harganya (maks 20 zona)
					b) Menetapkan tipe zona untuk tiap NIB
					c) Kalkulasi NPW tanah tiap bidang
			ii. Bangunan
				1) Menentukan klasifikasi baku bangunan I-V dengan hargai masing-masing bisa M2 persegi atau ml meter lari
				2) Menampilkan semua bidang yang ada bangunan nya
				3) Mengisi klasifikasi tiap bidang berserta penyusutan nya
				4) Kalkulasi otomatis NPW bangunan
			iii. Tanaman
				1) Memunculkan daftar tanaman yang ada pada bidang (group)
				2) Menampilkan harga standar tiap tanaman (input dari perda)
				3) Update NPW tanaman
			iv. Benda lain
				1) Idem bangunan
		b. Non fisik
			i. Kerugian usaha
				1) Memilih NIB yang ada usahanya secara manual
				2) Menentukan kerugian per bulan dengan jumlah bulan
			ii. Solatium
				1) Data diambil dari nominal bangunan
				2) Perhitungan rumus x=nilai hitung tanah x nilai bangunan
				3) Nilai hitung tanah dihitung dengan menetapkan bangunan induk luas bidang dan luas bangunan
			iii. Pindah
				1) Maksimal 1%
				2) Rumus pindah=total fisik x 1%
			iv. Pajak
				1) Semua bidang kena pajak
				2) Rumus=(total nilai fisik-x)x5%
				3) X=NOPTKP ditetapkan oleh daerah masing-masing
			v. Masa tunggu
				1) Semua kena
				2) Rumus=nilai fisik x faktor masa tunggu
			vi. Kerugian sisa tanah
				1) Bidang di pilih manual
				2) Di isi nominal kerugian untuk masing-masing bidang terpilih
			vii. Potensi kerugian
				1) Idem kerugian sisa tanah
			viii. Potensi kenaikan harga
				1) Rumus=harga (tanah x luas)
				2) Dipilih manual
