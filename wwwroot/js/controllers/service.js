
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
    var grid = document.getElementById("Grid").ej2_instances[0];
    var rasio = document.getElementById("ratio").ej2_instances[0];
    var klasifikasi = document.getElementById("klasifikasiusaha").ej2_instances[0];
    var jenis = document.getElementById("jeniskegiatanusaha").ej2_instances[0];
    // var tahunslider = document.getElementById("tahun").ej2_instances[0];
    var tahunpajak = document.getElementById("tahunpajak").ej2_instances[0].value.getFullYear();
    var tahun1 = tahunpajak - 3;
    var tahun2 = tahunpajak - 0;

    // fetch('@Url.Action("Hitung", "Service")' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahunslider.value[0] + '&tahun2=' + tahunslider.value[1])
    fetch('/Service/Hitung' + '?rasio=' + rasio.value + '&jenis=' + jenis.value + '&klasifikasi=' + klasifikasi.value + '&tahun1=' + tahun1 + '&tahun2=' + tahun2)
        .then(response => response.json())
        .then(data => {
            grid.changeDataSource(data);
            grid.refresh();
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
    }
    else if (rasio.value == "Cost Plus Methode" || rasio.value == "Cost Plus Method") {
        testedparty.value = (penjualan.value - hargapokokpenjualan.value) / hargapokokpenjualan.value;
    }
    else if (rasio.value == "Net Cost Plus Methode" || rasio.value == "Net Cost Plus Method") {
        testedparty.value = (penjualan.value - hargapokokpenjualan.value - bebanoperasional.value) / (hargapokokpenjualan.value + bebanoperasional.value);
    }
    else if (rasio.value == "Return On Sales") {
        testedparty.value = (penjualan.value - hargapokokpenjualan.value - bebanoperasional.value) / penjualan.value;
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
