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
    //var games = document.getElementById('Rasio').ej2_instances[0];
    //var value = document.getElementById('value');
    //var text = document.getElementById('text');
    //value.innerHTML = games.value === null ? 'null' : games.value.toString();
    //text.innerHTML = games.text === null ? 'null' : games.text.toString();
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

    var grid = document.getElementById("Grid").ej2_instances[0];
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
                grid.changeDataSource(data);
                grid.refresh

                if (data.length != 0) {
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

        var grid2 = document.getElementById("Grid2").ej2_instances[0];
        //fetch('@Url.Action("Hitung2", "Service")' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahunslider.value[0] + '&tahun2=' + tahunslider.value[1])
        fetch('/Service/Hitung2' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahun1 + '&tahun2=' + tahun2)
            .then(response => response.json())
            .then(data => {
                grid2.changeDataSource(data);
                grid2.refresh();
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
    }
    else {
        toastObj.show();
    }

    //console.log("test");
    //document.getElementById("jeniskegiatanusaharesult").innerHTML = jenis.value;
    //document.getElementById("klasifikasiusaharesult").innerHTML = klasifikasi.value;
    //document.getElementById("ratioresult").innerHTML = rasio.value;
    //document.getElementById("testedpartyresult").innerHTML = testedparty.value + ' %';
    //document.getElementById("tahunpajakresult").innerHTML = tahunpajak;


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

