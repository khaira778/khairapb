/* =========================
ARRAY DATA
========================= */

let transaksi = [];

/* Menyimpan index yang sedang diedit */
let indexEdit = -1;

/* =========================
TANGGAL OTOMATIS
========================= */

window.onload = function () {

let hariIni = new Date();

let tahun = hariIni.getFullYear();

let bulan = String(
    hariIni.getMonth() + 1
).padStart(2, "0");

let hari = String(
    hariIni.getDate()
).padStart(2, "0");

document.getElementById("tanggal").value =
    tahun + "-" + bulan + "-" + hari;


};

/* =========================
TAMBAH / SIMPAN EDIT
========================= */

function tambahData() {

let pelanggan =
    document.getElementById("namaPelanggan").value.trim();

let tanggal =
    document.getElementById("tanggal").value;

let nama =
    document.getElementById("namaBarang").value.trim();

let jumlah =
    Number(
        document.getElementById("jumlah").value
    );

let harga =
    Number(
        document.getElementById("harga").value
    );


/* CEK DATA */

if (
    pelanggan === "" ||
    tanggal === "" ||
    nama === "" ||
    jumlah <= 0 ||
    harga < 0
) {

    alert(
        "Silakan isi semua data terlebih dahulu!"
    );

    return;
}


let total = jumlah * harga;


/* =========================
   JIKA SEDANG EDIT
   ========================= */

if (indexEdit !== -1) {

    transaksi[indexEdit] = {

        pelanggan: pelanggan,
        tanggal: tanggal,
        nama: nama,
        jumlah: jumlah,
        harga: harga,
        total: total

    };

    indexEdit = -1;

    document.getElementById("btnTambah").innerText =
        "+ Tambah Data";

    tampilkanData();

    kosongkanForm();

    return;
}


/* =========================
   DATA BARU
   ========================= */

transaksi.push({

    pelanggan: pelanggan,
    tanggal: tanggal,
    nama: nama,
    jumlah: jumlah,
    harga: harga,
    total: total

});

tampilkanData();

kosongkanBarang();


}

/* =========================
EDIT DATA
========================= */

function editData(index) {

let data = transaksi[index];

document.getElementById("namaPelanggan").value =
    data.pelanggan;

document.getElementById("tanggal").value =
    data.tanggal;

document.getElementById("namaBarang").value =
    data.nama;

document.getElementById("jumlah").value =
    data.jumlah;

document.getElementById("harga").value =
    data.harga;


/* Simpan index yang diedit */

indexEdit = index;


/* Ubah tulisan tombol */

document.getElementById("btnTambah").innerText =
    "💾 Simpan Perubahan";


/* Scroll ke form */

window.scrollTo({
    top: 250,
    behavior: "smooth"
});


}

/* =========================
TAMPILKAN DATA
========================= */

function tampilkanData() {

let tabel =
    document.getElementById("tabelData");

tabel.innerHTML = "";

let totalSemua = 0;


if (transaksi.length === 0) {

    tabel.innerHTML = `

        <tr>
            <td colspan="8" class="kosong">
                Belum ada data
            </td>
        </tr>

    `;
}


transaksi.forEach(function (data, index) {

    totalSemua += data.total;

    let baris =
        document.createElement("tr");


    baris.innerHTML = `

        <td>
            ${index + 1}
        </td>

        <td>
            ${data.pelanggan}
        </td>

        <td>
            ${formatTanggal(data.tanggal)}
        </td>

        <td>
            ${data.nama}
        </td>

        <td>
            ${data.jumlah}
        </td>

        <td>
            ${formatRupiah(data.harga)}
        </td>

        <td>
            ${formatRupiah(data.total)}
        </td>

        <td>

            <div class="aksi">

                <button
                    class="btn-edit"
                    onclick="editData(${index})"
                >
                    ✏️ Edit
                </button>

                <button
                    class="btn-hapus"
                    onclick="hapusData(${index})"
                >
                    🗑️ Hapus
                </button>

            </div>

        </td>

    `;

    tabel.appendChild(baris);

});


document.getElementById("totalSemua").innerText =
    formatRupiah(totalSemua);


}

/* =========================
HAPUS SATU DATA
========================= */

function hapusData(index) {

let yakin =
    confirm(
        "Apakah kamu yakin ingin menghapus data ini?"
    );

if (yakin) {

    transaksi.splice(index, 1);

    tampilkanData();

}


}

/* =========================
KOSONGKAN BARANG
========================= */

function kosongkanBarang() {

document.getElementById("namaBarang").value = "";

document.getElementById("jumlah").value = "";

document.getElementById("harga").value = "";

document.getElementById("namaBarang").focus();


}

/* =========================
KOSONGKAN FORM
========================= */

function kosongkanForm() {

document.getElementById("namaPelanggan").value = "";

document.getElementById("namaBarang").value = "";

document.getElementById("jumlah").value = "";

document.getElementById("harga").value = "";

indexEdit = -1;

document.getElementById("btnTambah").innerText =
    "+ Tambah Data";


}

/* =========================
FORMAT RUPIAH
========================= */

function formatRupiah(angka) {

return "Rp " +
    angka.toLocaleString("id-ID");


}

/* =========================
FORMAT TANGGAL
========================= */

function formatTanggal(tanggal) {

let bagian =
    tanggal.split("-");

return bagian[2] +
    "-" +
    bagian[1] +
    "-" +
    bagian[0];


}

/* =========================
EXPORT EXCEL
========================= */

function exportExcel() {

if (transaksi.length === 0) {

    alert(
        "Belum ada data untuk diekspor!"
    );

    return;
}


let dataExcel = [];


/* HEADER */

dataExcel.push([

    "No",
    "Nama Pelanggan",
    "Tanggal",
    "Nama Barang",
    "Jumlah",
    "Harga Satuan",
    "Total"

]);


/* SEMUA DATA */

transaksi.forEach(function (data, index) {

    dataExcel.push([

        index + 1,

        data.pelanggan,

        formatTanggal(
            data.tanggal
        ),

        data.nama,

        data.jumlah,

        data.harga,

        data.total

    ]);

});


/* TOTAL */

let totalSemua =
    transaksi.reduce(
        function (total, data) {

            return total + data.total;

        },
        0
    );


dataExcel.push([]);

dataExcel.push([

    "",
    "",
    "",
    "",
    "",
    "TOTAL SEMUA",
    totalSemua

]);


/* BUAT EXCEL */

let workbook =
    XLSX.utils.book_new();

let worksheet =
    XLSX.utils.aoa_to_sheet(
        dataExcel
    );


/* LEBAR KOLOM */

worksheet["!cols"] = [

    { wch: 8 },
    { wch: 25 },
    { wch: 15 },
    { wch: 25 },
    { wch: 12 },
    { wch: 18 },
    { wch: 18 }

];


XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Data Kasir"
);


/* DOWNLOAD */

XLSX.writeFile(
    workbook,
    "Data_Kasir.xlsx"
);


}

/* =========================
HAPUS SEMUA
========================= */

function resetData() {

if (transaksi.length === 0) {
    return;
}


let yakin =
    confirm(
        "Apakah kamu yakin ingin menghapus semua data?"
    );


if (yakin) {

    transaksi = [];

    indexEdit = -1;

    tampilkanData();

}


}