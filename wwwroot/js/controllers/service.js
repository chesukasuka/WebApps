var btoastbefore = false;
var toastObj;
function created() {
    toastObj = this;
}
setTimeout(function () {
    toastObj.show({
        target: document.body
    });
}, 500);
function toastbefore(e) {
    cek = btoastbefore;
    if (cek == true) {
        e.cancel = false;
    }
    else {
        e.cancel = true;
    }
}

document.addEventListener('DOMContentLoaded', function () {
    // Ambil semua elemen dengan kelas 'numeric-textbox'
    var numericTextBoxes = document.getElementsByClassName('numeric-textbox');

    Array.from(numericTextBoxes).forEach(function (textBox) {
        // Ambil instance NumericTextBox
        var numericTextBoxInstance = textBox.ej2_instances[0];

        // Tambahkan event listener untuk menangani perubahan input
        numericTextBoxInstance.addEventListener('input', function () {
            // Ambil nilai dari NumericTextBox
            let value = numericTextBoxInstance.value;

            // Hapus semua karakter non-digit
            let numericValue = value.replace(/[^0-9]/g, '');

            // Format angka dengan separator ribuan
            let formattedValue = Number(numericValue).toLocaleString();

            // Atur nilai kembali ke NumericTextBox
            numericTextBoxInstance.value = formattedValue;
        });
    });
});

function jenisChange() {

    var jeniskegiatanusaha = document.getElementById('jeniskegiatanusaha').ej2_instances[0];
    var klasifikasiusaha = document.getElementById('klasifikasiusaha').ej2_instances[0];

    var tempQuery = new ej.data.Query().where('JenisKegiatanUsaha', 'equal', jeniskegiatanusaha.value);
    klasifikasiusaha.query = tempQuery;
    klasifikasiusaha.text = null;
    klasifikasiusaha.dataBind();
};


function klasifikasiChange() {
    var klasifikasiusaha = document.getElementById('klasifikasiusaha').ej2_instances[0];
    var ratio = document.getElementById('ratio').ej2_instances[0];

    var tempQuery = new ej.data.Query().where('KlasifikasiUsaha', 'equal', klasifikasiusaha.value);
    ratio.query = tempQuery;
    ratio.text = null;
    ratio.dataBind();
};


function ratioChange() {
    var games = document.getElementById('ratio').ej2_instances[0];
    //var value = document.getElementById('value');
    //var text = document.getElementById('text');
    //value.innerHTML = games.value === null ? 'null' : games.value.toString();
    //text.innerHTML = games.text === null ? 'null' : games.text.toString();

    let text = document.getElementById('tentang_rasio');
    console.log(games.value)
    if (games.value == 'Cost Plus Methode') {
        text.innerHTML = "Cost Plus Methode (CPM) : Gross Profit / Cost Plus Methode"
        document.getElementById('penjelasan').classList.remove('hidden');
    }
    else if (games.value == 'Net Cost Plus Methode') {
        text.innerHTML = "Net Cost Plus Methode (NCPM) : Operating Profit / (Cost Plus Methode + Operating Expense)"
        document.getElementById('penjelasan').classList.remove('hidden');
    }
    else if (games.value == 'Resale Price Methode') {
        text.innerHTML = "Resale Price Methode (RPM) : Gross Profit / Operating Revenue"
        document.getElementById('penjelasan').classList.remove('hidden');
    }
    else if (games.value == 'Return On Sales') {
        text.innerHTML = "Return On Sales (ROS) : Operating Profit / Operating Revenue"
        document.getElementById('penjelasan').classList.remove('hidden');
    }
    else {
        document.getElementById('penjelasan').classList.add('hidden');
    }


};

function toolbarClick(args) {
    var gridObj = document.getElementById("Grid").ej2_instances[0];
    if (args.item.id === 'Grid_pdfexport') {
        gridObj.showSpinner();
        gridObj.pdfExport();
    }
    else if (args.item.id === 'Grid_excelexport') {
        gridObj.showSpinner();
        gridObj.excelExport();
    }
}
function pdfExportComplete(args) {
    this.hideSpinner();
}
function excelExportComplete(args) {
    this.hideSpinner();
}

document.getElementById('btnHitung').onclick = () => {
    loadCustomers()
};

document.getElementById('btnReset').onclick = () => {
    resetAll()
};

function resetAll() {
    
    // Reset elements with class 'to-hide' and 'to-show'
    const elementsToHide = document.querySelectorAll('.to-hide');
    elementsToHide.forEach(element => {
        element.classList.add('hidden');
    });

    const elementsToShow = document.querySelectorAll('.to-show');
    elementsToShow.forEach(element => {
        element.classList.remove('hidden');
    });

    // Reset Grid data source
    //var grid = document.getElementById("Grid").ej2_instances[0];
    //var grid2 = document.getElementById("Grid2").ej2_instances[0];
    //grid.changeDataSource([]);
    //grid2.changeDataSource([]);
    //grid.refresh();
    //grid2.refresh();

    // Reset dropdowns and inputs
    var rasio = document.getElementById("ratio").ej2_instances[0];
    var klasifikasi = document.getElementById("klasifikasiusaha").ej2_instances[0];
    var jenis = document.getElementById("jeniskegiatanusaha").ej2_instances[0];
    var tahun = document.getElementById("tahun").ej2_instances[0];
    var tahunpajak = document.getElementById("tahunpajak").ej2_instances[0];

    rasio.value = null;
    klasifikasi.value = null;
    jenis.value = null;
    tahun.value = null;
    tahunpajak.value = null;

    // Reset financial inputs
    var penjualan = document.getElementById("penjualan").ej2_instances[0];
    var hargapokokpenjualan = document.getElementById("hargapokokpenjualan").ej2_instances[0];
    var bebanoperasional = document.getElementById("bebanoperasional").ej2_instances[0];
    var labakotor = document.getElementById("labakotor").ej2_instances[0];
    var labaoperasional = document.getElementById("labaoperasional").ej2_instances[0];
    var testedparty = document.getElementById("testedparty").ej2_instances[0];

    penjualan.value = null;
    hargapokokpenjualan.value = null;
    bebanoperasional.value = null;
    labakotor.value = null;
    labaoperasional.value = null;
    testedparty.value = null; 
}
function loadCustomers() {
    btoastbefore = true;
    const elements = document.querySelectorAll('.to-hide'); // Only target elements with the 'to-hide' class
    elements.forEach(element => {
        element.classList.add('hidden');
    });

    const elements2 = document.querySelectorAll('.to-show'); // Only target elements with the 'to-hide' class
    elements2.forEach(elements2 => {
        elements2.classList.remove('hidden');
    });

    //var grid = document.getElementById("Grid").ej2_instances[0];
    var rasio = document.getElementById("ratio").ej2_instances[0];
    var klasifikasi = document.getElementById("klasifikasiusaha").ej2_instances[0];
    var jenis = document.getElementById("jeniskegiatanusaha").ej2_instances[0];
    // var tahunslider = document.getElementById("tahun").ej2_instances[0];
    var tahunpajak = document.getElementById("tahunpajak").ej2_instances[0].value;
    var tahun = document.getElementById("tahun").ej2_instances[0].value;

    var tahun1 = 0;
    var tahun2 = 0;

    if (tahun == "1 Tahun") {
        var tahun1 = tahunpajak - 1;
        var tahun2 = tahunpajak - 1;
    }
    else if (tahun == "3 Tahun") {
        var tahun1 = tahunpajak - 3;
        var tahun2 = tahunpajak - 1;
    }
    else if (tahun == "5 Tahun") {
        var tahun1 = tahunpajak - 5;
        var tahun2 = tahunpajak - 1;
    }
    else {
        var tahun1 = tahunpajak - 1;
        var tahun2 = tahunpajak - 1;
    }

    if (tahun1 >= 2017) {
        // fetch('@Url.Action("Hitung", "Service")' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahunslider.value[0] + '&tahun2=' + tahunslider.value[1])
        fetch('/Service/Hitung' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahun1 + '&tahun2=' + tahun2)
            .then(response => response.json())
            .then(data => {
                //grid.changeDataSource(data);
                //grid.refresh
                $("#grid-header").empty();
                $("#grid-body").empty();

                if (data.length != 0) {
                    let header = data[0];
                    $.each(header, function (key, value){
                        $("#grid-header").append(`<th ${key == 'Nama Perusahaan' || key == 'Negara' ? 'width="40%"' : ''}>${key}</th>`);
                    })
                    data.forEach((items) => {
                        // Buat variabel untuk menyimpan baris tabel
                        let row = "<tr>";

                        // Iterasi melalui item
                        $.each(items, function (key, value) {
                            // Tambahkan setiap nilai sebagai kolom tabel (td) ke dalam baris
                            row += `<td>${value}</td>`;
                        });

                        // Tutup tag <tr>
                        row += "</tr>";

                        // Tambahkan seluruh baris ke dalam tbody tabel
                        $("#grid-body").append(row);
                    })
                    const elements = document.querySelectorAll('.to-hide'); // Only target elements with the 'to-hide' class
                    elements.forEach(element => {
                        element.classList.remove('hidden');
                    });
                    const elements2 = document.querySelectorAll('.to-show'); // Only target elements with the 'to-hide' class
                    elements2.forEach(elements2 => {
                        elements2.classList.add('hidden');
                    });

                }
                else {
                    toastObj.show();
                }
            });

        //var grid2 = document.getElementById("Grid2").ej2_instances[0];
        //fetch('@Url.Action("Hitung2", "Service")' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahunslider.value[0] + '&tahun2=' + tahunslider.value[1])
        fetch('/Service/Hitung2' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahun1 + '&tahun2=' + tahun2)
            .then(response => response.json())
            .then(data => {
                //grid2.changeDataSource(data);
                //grid2.refresh();

                $("#grid2-header").empty();
                $("#grid2-body").empty();

                let header = data[0];
                $.each(header, function (key, value) {
                    $("#grid2-header").append(`<th ${key == 'Keterangan' ? 'width="80%"' : ''}>${key}</th>`);
                })
                data.forEach((items) => {
                    // Buat variabel untuk menyimpan baris tabel
                    let row = "<tr>";

                    // Iterasi melalui item
                    $.each(items, function (key, value) {
                        // Tambahkan setiap nilai sebagai kolom tabel (td) ke dalam baris
                        row += `<td>${value}</td>`;
                    });

                    // Tutup tag <tr>
                    row += "</tr>";

                    // Tambahkan seluruh baris ke dalam tbody tabel
                    $("#grid2-body").append(row);
                })
            });

        var penjualan = document.getElementById("penjualan").ej2_instances[0];
        var hargapokokpenjualan = document.getElementById("hargapokokpenjualan").ej2_instances[0];
        var bebanoperasional = document.getElementById("bebanoperasional").ej2_instances[0];

        var labakotor = document.getElementById("labakotor").ej2_instances[0];
        var labaoperasional = document.getElementById("labaoperasional").ej2_instances[0];

        labakotor.value = penjualan.value - hargapokokpenjualan.value;
        labaoperasional.value = penjualan.value - hargapokokpenjualan.value - bebanoperasional.value;

        var testedparty = document.getElementById("testedparty").ej2_instances[0];
        if (rasio.value == "Resale Price Methode" || rasio.value == "Resale Price Method") {
            testedparty.value = (penjualan.value - hargapokokpenjualan.value) / penjualan.value;
            testedparty.value = Math.round(testedparty.value * 100) /100
        }
        else if (rasio.value == "Cost Plus Methode" || rasio.value == "Cost Plus Method") {
            testedparty.value = (penjualan.value - hargapokokpenjualan.value) / hargapokokpenjualan.value;
            testedparty.value = Math.round(testedparty.value * 100) / 100
        }
        else if (rasio.value == "Net Cost Plus Methode" || rasio.value == "Net Cost Plus Method") {
            testedparty.value = (penjualan.value - hargapokokpenjualan.value - bebanoperasional.value) / (hargapokokpenjualan.value + bebanoperasional.value);
            testedparty.value = Math.round(testedparty.value * 100) / 100
        }
        else if (rasio.value == "Return On Sales") {
            testedparty.value = (penjualan.value - hargapokokpenjualan.value - bebanoperasional.value) / penjualan.value;
            testedparty.value = Math.round(testedparty.value * 100) / 100
        }
        document.getElementById("jeniskegiatanusaharesult").innerHTML = jenis.value;
        document.getElementById("klasifikasiusaharesult").innerHTML = klasifikasi.value;
        document.getElementById("ratioresult").innerHTML = rasio.value;
        document.getElementById("testedpartyresult").innerHTML = testedparty.value + ' %';
        document.getElementById("tahunpajakresult").innerHTML = tahunpajak;
    }
    else {
        toastObj.show();
    }


};

function dataBound(args) {
    this.getColumns()[0].width = 200;
    this.getColumns()[1].width = 150;
    this.refreshColumns();
}
function dataBound2(args) {
    this.getColumns()[0].width = 350;
    this.refreshColumns();
}

