document.getElementById('process').addEventListener('click', function() {
    const input = document.getElementById('upload');
    if (!input.files[0]) {
        alert('Harap upload file terlebih dahulu');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Cari indeks kolom "F8" dan "f1101"
        const headerRow = jsonData[0];
        const statusIndex = headerRow.indexOf("F8"); // Status
        const startJobIndex = headerRow.indexOf("F502"); // Start Job
        const incomeIndex = headerRow.indexOf("F505"); // Pendapatan perbulan
        const provinceIndex = headerRow.indexOf("F5a1"); // wilayah tempat bekerja
        const kabupatenIndex = headerRow.indexOf("F5a2"); // kabupaten tempat bekerja
        const jenisPerusahaanIndex = headerRow.indexOf("F1101"); // Jenis Perusahaan

        const companyNameIndex = headerRow.indexOf("F5b"); // kabupaten tempat bekerja
        const positionIfSelfEmployedIndex = headerRow.indexOf("F5c"); 
        const levelCompanyIndex = headerRow.indexOf("F5d"); 
        const sourceOfFundIndex = headerRow.indexOf("F18a"); 
        const collegeNameIndex = headerRow.indexOf("F18b"); 
        const programStudyIndex = headerRow.indexOf("F18c"); 
        const startDateIndex = headerRow.indexOf("F18d"); 
        const biayaKuliahIndex = headerRow.indexOf("F1201"); 
        const hubunganBidangStudiIndex = headerRow.indexOf("F14"); 
        const tingkatPendidikanYangTepatIndex = headerRow.indexOf("F15"); 
        const etikaAIndex = headerRow.indexOf("F1761"); 
        const etikaBIndex = headerRow.indexOf("F1762"); 
        const keahlianAIndex = headerRow.indexOf("F1763"); 
        const keahlianBIndex = headerRow.indexOf("F1764"); 
        const englishAIndex = headerRow.indexOf("F1765"); 
        const englishBIndex = headerRow.indexOf("F1766"); 
        const useItAIndex = headerRow.indexOf("F1767"); 
        const useItBIndex = headerRow.indexOf("F1768"); 
        const komunikasiAIndex = headerRow.indexOf("F1769"); 
        const komunikasiBIndex = headerRow.indexOf("F1770"); 
        const kerjaSamaAIndex = headerRow.indexOf("F1771"); 
        const kerjaSamaBIndex = headerRow.indexOf("F1772"); 
        const pengembanganAIndex = headerRow.indexOf("F1773"); 
        const pengembanganBIndex = headerRow.indexOf("F1774"); 
        const perkuliahanIndex = headerRow.indexOf("F21"); 
        const demonstrasiIndex = headerRow.indexOf("F22"); 
        const projekAkhirIndex = headerRow.indexOf("F23"); 
        const magangIndex = headerRow.indexOf("F24"); 
        const praktikumIndex = headerRow.indexOf("F25"); 
        const kerjaLapanganIndex = headerRow.indexOf("F26"); 
        const diskusiIndex = headerRow.indexOf("F27"); 
        const f301Index = headerRow.indexOf("F301");
        const f302Index = headerRow.indexOf("F302");
        const f303Index = headerRow.indexOf("F303"); 
        const f401Index = headerRow.indexOf("F401"); 
        const f402Index = headerRow.indexOf("F402"); 
        const f403Index = headerRow.indexOf("F403"); 
        const f404Index = headerRow.indexOf("F404"); 
        const f405Index = headerRow.indexOf("F405"); 
        const f406Index = headerRow.indexOf("F406"); 
        const f407Index = headerRow.indexOf("F407"); 
        const f408Index = headerRow.indexOf("F408"); 
        const f409Index = headerRow.indexOf("F409"); 
        const f410Index = headerRow.indexOf("F410"); 
        const f411Index = headerRow.indexOf("F411"); 
        const f412Index = headerRow.indexOf("F412"); 
        const f413Index = headerRow.indexOf("F413"); 
        const f414Index = headerRow.indexOf("F414"); 
        const f415Index = headerRow.indexOf("F415"); 
        const f6Index = headerRow.indexOf("F6"); 
        const f7Index = headerRow.indexOf("F7"); 
        const f7aIndex = headerRow.indexOf("F7a"); 
        const f1001Index = headerRow.indexOf("F1001"); 
        const f1601Index = headerRow.indexOf("F1601"); 
        const f1602Index = headerRow.indexOf("F1602"); 
        const f1603Index = headerRow.indexOf("F1603"); 
        const f1604Index = headerRow.indexOf("F1604"); 
        const f1605Index = headerRow.indexOf("F1605"); 
        const f1606Index = headerRow.indexOf("F1606"); 
        const f1607Index = headerRow.indexOf("F1607"); 
        const f1608Index = headerRow.indexOf("F1608"); 
        const f1609Index = headerRow.indexOf("F1609"); 
        const f1610Index = headerRow.indexOf("F1610"); 
        const f1611Index = headerRow.indexOf("F1611"); 
        const f1612Index = headerRow.indexOf("F1612"); 
        const f1613Index = headerRow.indexOf("F1613"); 

        if (statusIndex === -1) {
            alert('Kolom "F8" tidak ditemukan');
            return;
        }

        if (jenisPerusahaanIndex === -1) {
            alert('Kolom "F1101" tidak ditemukan');
            return;
        }

        if (startJobIndex === -1) {
            alert('Kolom "F502" tidak ditemukan');
            return;
        }
        if (incomeIndex === -1) {
            alert('Kolom "F505" tidak ditemukan');
            return;
        }
        if (provinceIndex === -1) {
            alert('Kolom "F5a1" tidak ditemukan');
            return;
        }
        if (kabupatenIndex === -1) {
            alert('Kolom "F5a2" tidak ditemukan');
            return;
        }
        if (companyNameIndex === -1) {
            alert('Kolom "F5b" tidak ditemukan');
            return;
        }
        if (positionIfSelfEmployedIndex === -1) {
            alert('Kolom "F5c" tidak ditemukan');
            return;
        }
        if (levelCompanyIndex === -1) {
            alert('Kolom "F5d" tidak ditemukan');
            return;
        }
        if (sourceOfFundIndex === -1) {
            alert('Kolom "F18a" tidak ditemukan');
            return;
        }
        if (collegeNameIndex === -1) {
            alert('Kolom "F18b" tidak ditemukan');
            return;
        }
        if (programStudyIndex === -1) {
            alert('Kolom "F18c" tidak ditemukan');
            return;
        }
        if (startDateIndex === -1) {
            alert('Kolom "F18d" tidak ditemukan');
            return;
        }
        if (biayaKuliahIndex === -1) {
            alert('Kolom "F18d" tidak ditemukan');
            return;
        }
        

        // Ganti header
        headerRow[statusIndex] = "Jelaskan status Anda saat ini?";
        headerRow[jenisPerusahaanIndex] = "Apa jenis perusahaan/intansi/institusi tempat anda bekerja sekarang?";
        headerRow[startJobIndex] = "Dalam berapa bulan setelah lulus anda memulai pekerjaan ?";
        headerRow[incomeIndex] = "Berapa rata-rata pendapatan Anda per bulan ?";
        headerRow[provinceIndex] = "Di Provinsi mana lokasi tempat Anda bekerja saat ini ?";
        headerRow[kabupatenIndex] = "Di Kota/Kabupaten mana lokasi tempat Anda bekerja saat ini ?";
        headerRow[companyNameIndex] = "Apa nama perusahaan/kantor tempat Anda bekerja ?";
        headerRow[positionIfSelfEmployedIndex] = "Bila berwiraswasta, apa posisi/jabatan Anda saat ini ?";
        headerRow[levelCompanyIndex] = "Apa tingkat tempat kerja Anda ?";
        headerRow[sourceOfFundIndex] = "Sumber biaya";
        headerRow[collegeNameIndex] = "Perguruan Tiggi";
        headerRow[programStudyIndex] = "Program Studi";
        headerRow[startDateIndex] = "Tanggal Masuk";
        headerRow[biayaKuliahIndex] = "Sebutkan sumberdana dalam pembiayaan selama berkuliah di Politeknik Negeri Subang";
        headerRow[hubunganBidangStudiIndex] = "Seberapa erat hubungan bidang studi dengan pekerjaan Anda?";
        headerRow[tingkatPendidikanYangTepatIndex] = "Tingkat pendidikan apa yang paling tepat/sesuai untuk pekerjaan anda saat ini?";
        headerRow[etikaAIndex] = "Etika (A)";
        headerRow[etikaBIndex] = "Etika (B)";
        headerRow[keahlianAIndex] = "Keahlian berdasarkan bidang ilmu (A)";
        headerRow[keahlianBIndex] = "Keahlian berdasarkan bidang ilmu (B)";
        headerRow[englishAIndex] = "Bahasa Inggris (A)";
        headerRow[englishBIndex] = "Bahasa Inggris (B)";
        headerRow[useItAIndex] = "Penggunaan Teknologi Informasi (A)";
        headerRow[useItBIndex] = "Penggunaan Teknologi Informasi (B)";
        headerRow[komunikasiAIndex] = "Komunikasi (A)";
        headerRow[komunikasiBIndex] = "Komunikasi (B)";
        headerRow[kerjaSamaAIndex] = "Kerja Sama (A)";
        headerRow[kerjaSamaBIndex] = "Kerja Sama (B)";
        headerRow[pengembanganAIndex] = "Pengembangan (A)";
        headerRow[pengembanganBIndex] = "Pengembangan (B)";
        headerRow[perkuliahanIndex] = "Perkuliahan";
        headerRow[demonstrasiIndex] = "Demonstrasi";
        headerRow[projekAkhirIndex] = "Partisipasi dalam proyek riset";
        headerRow[magangIndex] = "Magang";
        headerRow[praktikumIndex] = "Praktikum";
        headerRow[kerjaLapanganIndex] = "Kerja Lapangan";
        headerRow[diskusiIndex] = "Diskusi";
        headerRow[f301Index] = "Kapan anda mulai mencari pekerjaan?";
        headerRow[f401Index] = "Melalui iklan di koran/majalah, brosur";
        headerRow[f402Index] = "Melamar ke perusahaan tanpa mengetahui lowongan yang ada";
        headerRow[f403Index] = "Pergikebursa/pamerankerja";
        headerRow[f404Index] = "Mencarilewatinternet/iklanonline/milis";
        headerRow[f405Index] = "Dihubungi oleh perusahaan";
        headerRow[f406Index] = "Menghubungi Kemenakertrans";
        headerRow[f407Index] = "Menghubungi agen tenaga kerja komersial/swasta";
        headerRow[f408Index] = "Memeroleh informasi dari pusat/kantor pengembangan karir fakultas/universitas";
        headerRow[f409Index] = "Menghubungikantorkemahasiswaan/hubunganalumni";
        headerRow[f410Index] = "Membangunjejaring(network)sejakmasihkuliah";
        headerRow[f411Index] = "Melalui relasi (misalnya dosen, orang tua, saudara, teman, dll.)";
        headerRow[f412Index] = "Membangun bisnis sendiri";
        headerRow[f413Index] = "Melalui penempatan kerja atau magang";
        headerRow[f414Index] = "Bekerja di tempat yang sama dengan tempat kerja semasa kuliah";
        headerRow[f415Index] = "Lainnya";
        headerRow[f6Index] = "Berapa perusahaan/instansi/institusi yang sudah anda lamar (lewat surat atau e-mail) sebelum anda memeroleh pekerjaan pertama?";
        headerRow[f7Index] = "Berapa banyak perusahaan/instansi/institusi yang merespons lamaran anda?";
        headerRow[f7aIndex] = "Berapa banyak perusahaan/instansi/institusi yang mengundang anda untuk wawancara?";
        headerRow[f1001Index] = "Apakah anda aktif mencari pekerjaan dalam 4 minggu terakhir?";
        headerRow[f1601Index] = "Pertanyaan tidak sesuai; pekerjaan saya sekarang sudah sesuai dengan pendidikan saya.";
        headerRow[f1602Index] = "Saya belum mendapatkan pekerjaan yang lebih sesuai.";
        headerRow[f1603Index] = "Di pekerjaan ini saya memeroleh prospek karir yang baik.";
        headerRow[f1604Index] = "Saya lebih suka bekerja di area pekerjaan yang tidak ada hubungannya dengan pendidikan saya.";
        headerRow[f1605Index] = "Saya dipromosikan ke posisi yang kurang berhubungan dengan pendidikan saya dibanding posisi sebelumnya.";
        headerRow[f1606Index] = "Saya dapat memeroleh pendapatan yang lebih tinggi di pekerjaan ini.";
        headerRow[f1607Index] = "Pekerjaan saya saat ini lebih aman/terjamin/secure";
        headerRow[f1608Index] = "Pekerjaan saya saat ini lebih menarik";
        headerRow[f1609Index] = "Pekerjaan saya saat ini lebih memungkinkan saya mengambil pekerjaan tambahan/jadwal yang fleksibel, dll.";
        headerRow[f1610Index] = "Pekerjaan saya saat ini lokasinya lebih dekat dari rumah saya.";
        headerRow[f1611Index] = "Pekerjaan saya saat ini dapat lebih menjamin kebutuhan keluarga saya.";
        headerRow[f1612Index] = "Pada awal meniti karir ini, saya harus menerima pekerjaan yang tidak berhubungan dengan pendidikan saya";
        headerRow[f1613Index] = "Lainnya";

        for (let i = 1; i < jsonData.length; i++) { 
            jsonData[i][statusIndex] = parseInt(jsonData[i][statusIndex]);

            // Ubah nilai kolom "Status"
            switch(jsonData[i][statusIndex]) {
                case 1:
                    jsonData[i][statusIndex] = 'Bekerja (full time / part time)';
                    break;
                case 2:
                    jsonData[i][statusIndex] = 'Belum memungkinkan bekerja';
                    break;
                case 3:
                    jsonData[i][statusIndex] = 'Wiraswasta';
                    break;
                case 4:
                    jsonData[i][statusIndex] = 'Melanjutkan Pendidikan';
                    break;
                case 5:
                    jsonData[i][statusIndex] = 'Tidak kerja tetapi sedang mencari kerja';
                    break;
                default:
                    break;
            }

            // Ubah nilai kolom "Jenis Perusahaan"
            if (!isNaN(jsonData[i][jenisPerusahaanIndex])) {
                jsonData[i][jenisPerusahaanIndex] = parseInt(jsonData[i][jenisPerusahaanIndex]);

                // Ubah nilai kolom "Jenis Perusahaan"
                switch(jsonData[i][jenisPerusahaanIndex]) {
                    case 1:
                        jsonData[i][jenisPerusahaanIndex] = 'Instansi pemerintah';
                        break;
                    case 2:
                        jsonData[i][jenisPerusahaanIndex] = 'Organisasi non-profit/Lembaga Swadaya Masyarakat';
                        break;
                    case 3:
                        jsonData[i][jenisPerusahaanIndex] = 'Perusahaan swasta';
                        break;
                    case 4:
                        jsonData[i][jenisPerusahaanIndex] = 'Wiraswasta/perusahaan sendiri';
                        break;
                    case 5:
                        jsonData[i][jenisPerusahaanIndex] = 'Lainnya';
                        break;
                    case 6:
                        jsonData[i][jenisPerusahaanIndex] = 'BUMN/BUMD';
                        break;
                    case 7:
                        jsonData[i][jenisPerusahaanIndex] = 'Institusi/Organisasi Multilateral';
                        break;
                    default:
                        jsonData[i][jenisPerusahaanIndex] = ' ';
                        break;
                }
            }

            switch(jsonData[i][positionIfSelfEmployedIndex]) {
                case 1:
                    jsonData[i][positionIfSelfEmployedIndex] = 'Founder';
                    break;
                case 2:
                    jsonData[i][positionIfSelfEmployedIndex] = 'Co-Founder';
                    break;
                case 3:
                    jsonData[i][positionIfSelfEmployedIndex] = 'Staff';
                    break;
                case 4:
                    jsonData[i][positionIfSelfEmployedIndex] = 'Freelance/Kerja Lepas';
                    break;
                default:
                    break;
            }

            jsonData[i][levelCompanyIndex] = parseInt(jsonData[i][levelCompanyIndex]);

            switch(jsonData[i][levelCompanyIndex]) {
                case 1:
                    jsonData[i][levelCompanyIndex] = 'Lokal/Wilayah/Wiraswasta tidak berbadan hukum';
                    break;
                case 2:
                    jsonData[i][levelCompanyIndex] = 'Nasional/Wiraswasta berbadan hukum';
                    break;
                case 3:
                    jsonData[i][levelCompanyIndex] = 'Multinasional/Internasional';
                    break;
                default:
                    jsonData[i][levelCompanyIndex] = ' ';
                    break;
            }

            jsonData[i][biayaKuliahIndex] = parseInt(jsonData[i][biayaKuliahIndex]);

            switch(jsonData[i][biayaKuliahIndex]) {
                case 1:
                    jsonData[i][biayaKuliahIndex] = 'Biaya Sendiri/Keluarga';
                    break;
                case 2:
                    jsonData[i][biayaKuliahIndex] = 'Beasiswa ADIK';
                    break;
                case 3:
                    jsonData[i][biayaKuliahIndex] = 'Beasiswa BIDIKMISI';
                    break;
                case 4:
                    jsonData[i][biayaKuliahIndex] = 'Beasiswa PPA';
                    break;
                case 5:
                    jsonData[i][biayaKuliahIndex] = 'Beasiswa AFIRMASI';
                    break;
                case 6:
                    jsonData[i][biayaKuliahIndex] = 'Beasiswa Perusahaan/Swasta';
                    break;
                case 7:
                    jsonData[i][biayaKuliahIndex] = 'Lainnya';
                    break;
                default:
                    break;
            }

            jsonData[i][hubunganBidangStudiIndex] = parseInt(jsonData[i][hubunganBidangStudiIndex]);

            switch(jsonData[i][hubunganBidangStudiIndex]) {
                case 1:
                    jsonData[i][hubunganBidangStudiIndex] = 'Sangat Erat';
                    break;
                case 2:
                    jsonData[i][hubunganBidangStudiIndex] = 'Erat';
                    break;
                case 3:
                    jsonData[i][hubunganBidangStudiIndex] = 'Cukup Erat';
                    break;
                case 4:
                    jsonData[i][hubunganBidangStudiIndex] = 'Kurang Erat';
                    break;
                case 5:
                    jsonData[i][hubunganBidangStudiIndex] = 'Tidak Sama Sekali';
                    break;
                default:
                    jsonData[i][hubunganBidangStudiIndex] = ' ';
                    break;
            }

            jsonData[i][tingkatPendidikanYangTepatIndex] = parseInt(jsonData[i][tingkatPendidikanYangTepatIndex]);

            switch(jsonData[i][tingkatPendidikanYangTepatIndex]) {
                case 1:
                    jsonData[i][tingkatPendidikanYangTepatIndex] = 'Setingkat Lebih Tinggi';
                    break;
                case 2:
                    jsonData[i][tingkatPendidikanYangTepatIndex] = 'Tingkat yang Sama';
                    break;
                case 3:
                    jsonData[i][tingkatPendidikanYangTepatIndex] = 'Setingkat Lebih Rendah';
                    break;
                case 4:
                    jsonData[i][tingkatPendidikanYangTepatIndex] = 'Tidak Perlu Pendidikan Tinggi';
                    break;
                default:
                    jsonData[i][tingkatPendidikanYangTepatIndex] = ' ';
                    break;
            }

            const indices = [perkuliahanIndex, demonstrasiIndex, projekAkhirIndex, magangIndex, praktikumIndex, kerjaLapanganIndex, diskusiIndex];
            indices.forEach(index => {
                if (jsonData[i][index] !== undefined && jsonData[i][index] !== null) {
                    jsonData[i][index] = parseInt(jsonData[i][index]);
                }
            });
            
            indices.forEach(index => {
                switch (jsonData[i][index]) {
                    case 1:
                        jsonData[i][index] = 'Sangat Besar';
                        break;
                    case 2:
                        jsonData[i][index] = 'Besar';
                        break;
                    case 3:
                        jsonData[i][index] = 'Cukup Besar';
                        break;
                    case 4:
                        jsonData[i][index] = 'Kurang Besar';
                        break;
                    case 5:
                        jsonData[i][index] = 'Tidak Sama Sekali';
                        break;
                    default:
                        break;
                }
            });

            jsonData[i][f301Index] = parseInt(jsonData[i][f301Index]);

            switch (jsonData[i][f301Index]) {
                case 1:
                    jsonData[i][f301Index] = `Kira-kira ${jsonData[i][f302Index]} bulan sebelum lulus`;
                    break;
                case 2:
                    jsonData[i][f301Index] = `Kira-kira ${jsonData[i][f303Index]} bulan setelah lulus`;
                    break;
                case 3:
                    jsonData[i][f301Index] = "Saya tidak mencari kerja";
                    break;
                default:
                    jsonData[i][f301Index] = " ";
                    break;
            }

            const findJob = [f401Index, f402Index, f403Index, f404Index, f405Index, f406Index, f407Index, f408Index, f409Index, f410Index, f411Index, f412Index, f413Index, f414Index, f415Index];
            findJob.forEach(index => {
                if (jsonData[i][index] !== undefined && jsonData[i][index] !== null) {
                    jsonData[i][index] = parseInt(jsonData[i][index]);
                }
            });
            
            findJob.forEach(index => {
                switch (jsonData[i][index]) {
                    case 0:
                        jsonData[i][index] = 'Tidak';
                        break;
                    case 1:
                        jsonData[i][index] = 'Ya';
                        break;
                    default:
                        break;
                }
            });

            jsonData[i][f1001Index] = parseInt(jsonData[i][f1001Index]);

            switch (jsonData[i][f1001Index]) {
                case 1:
                    jsonData[i][f1001Index] = "Tidak";
                    break;
                case 2:
                    jsonData[i][f1001Index] = "Tidak, tapi saya sedang menunggu hasil lamaran kerja";
                    break;
                case 3:
                    jsonData[i][f1001Index] = "Ya, saya akan mulai bekerja dalam 2 minggu ke depan";
                    break;
                case 4:
                    jsonData[i][f1001Index] = "Ya, tapi saya belum pasti akan bekerja dalam 2 minggu ke depan";
                    break;
                case 5:
                    jsonData[i][f1001Index] = "Lainnya";
                    break;
                default:
                    jsonData[i][f1001Index] = " ";
                    break;
            }

            const kesesuaianPekerjaan = [f1601Index, f1602Index, f1603Index, f1604Index, f1604Index, f1605Index, f1606Index, f1607Index, f1608Index, f1609Index, f1610Index, f1611Index, f1612Index, f1613Index];
            kesesuaianPekerjaan.forEach(index => {
                if (jsonData[i][index] !== undefined && jsonData[i][index] !== null) {
                    jsonData[i][index] = parseInt(jsonData[i][index]);
                }
            });
            
            kesesuaianPekerjaan.forEach(index => {
                switch (jsonData[i][index]) {
                    case 0:
                        jsonData[i][index] = 'Tidak';
                        break;
                    case 1:
                        jsonData[i][index] = 'Ya';
                        break;
                    default:
                        break;
                }
            });
        }

        const newWorksheet = XLSX.utils.aoa_to_sheet(jsonData);

        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

        const newExcelFile = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });

        const downloadButton = document.getElementById('download');
        downloadButton.style.display = 'inline';
        downloadButton.addEventListener('click', function() {
            saveAs(new Blob([newExcelFile], { type: "application/octet-stream" }), 'processed_file.xlsx');
        });
    };

    reader.readAsArrayBuffer(input.files[0]);
});