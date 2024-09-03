// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
//document.getElementById('tahunpajak').addEventListener('change', function () {
//    var tahunpajak = document.getElementById("tahunpajak").ej2_instances[0].value;
//    var tahunslider = document.getElementById("tahun").ej2_instances[0];

//    tahunslider.minValue = 2019;
//    tahunslider.maxValue = 2022;

//});


//document.getElementById('tahunpajak').addEventListener('change', function () {
//    var tahunpajak = document.getElementById("tahunpajak").ej2_instances[0].value;
//    var tahunslider = document.getElementById("tahun").ej2_instances[0];

//    tahunslider.value[0] = 2019;
//    tahunslider.value[1] = 2020;

//});

function jenisChange() {

    var selectedCategory = document.getElementById('jeniskegiatanusaha').value;
    //var comboBox = document.getElementById('klasifikasiusaha').ej2_instances[0];

    var comboBox = new ej.dropdowns.ComboBox({
        placeholder: 'Select an item',
        fields: { text: 'KlasifikasiUsaha', value: 'KlasifikasiUsaha' },
        dataSource: [] // Initially empty data source
    });
    comboBox.appendTo('#klasifikasiusaha');

    fetch(`/Service/Klasifikasi?param=${selectedCategory}`)
        .then(response => response.json())
        .then(data => {
            comboBox.dataSource = data;
            comboBox.dataBind();
        })
        .catch(error => console.error('Error:', error));
};