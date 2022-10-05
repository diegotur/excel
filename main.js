let controlSeccionamiento = document.getElementById("controlSeccionamientoH1");
let controlMalUsoDelSube = document.getElementById("controlMalUsoDelSubeH1");
let controlSeccCorrido = document.getElementById("controlSeccCorridoH1");
let controlCortados = document.getElementById("controlCortadosH1");
let controlKmPorRamal = document.getElementById("controlKmPorRamalH1");
let controlMensual = document.getElementById("controlMensualH1");

function VerControl(a, b, c, d, e, f) {
    a.style.visibility = "visible";
    b.style.visibility = "hidden";
    c.style.visibility = "hidden";
    d.style.visibility = "hidden";
    e.style.visibility = "hidden";
    f.style.visibility = "hidden";


}
document.getElementById("controlSeccionamiento").addEventListener("click", () => {
    VerControl(controlSeccionamiento, controlMalUsoDelSube, controlSeccCorrido, controlCortados, controlKmPorRamal, controlMensual)
});
document.getElementById("controlMalUsoDelSube").addEventListener("click", () => {
    VerControl(controlMalUsoDelSube, controlSeccCorrido, controlCortados, controlSeccionamiento, controlKmPorRamal, controlMensual)
});
document.getElementById("controlCortados").addEventListener("click", () => {
    VerControl(controlSeccCorrido, controlMalUsoDelSube, controlSeccionamiento, controlCortados, controlKmPorRamal, controlMensual)
});
document.getElementById("controlSeccCorrido").addEventListener("click", () => {
    VerControl(controlCortados, controlMalUsoDelSube, controlSeccCorrido, controlSeccionamiento, controlKmPorRamal, controlMensual)
});
document.getElementById("controlKmPorRamal").addEventListener("click", () => {
    VerControl(controlKmPorRamal, controlCortados, controlMalUsoDelSube, controlSeccCorrido, controlSeccionamiento, controlMensual)
});
document.getElementById("controlMensual").addEventListener("click", () => {
    VerControl(controlMensual, controlKmPorRamal, controlCortados, controlMalUsoDelSube, controlSeccCorrido, controlSeccionamiento)
});




function upload() {
    var files = document.getElementById('file_upload').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

let result = {};
let roa;

let tableP = document.getElementById("tableP");

const ExcelDateToJSDate2 = (date) => {
    let converted_date = new Date(Math.round((date - 25568) * 864e5));
    converted_date = String(converted_date).slice(4, 15);
    date = converted_date.split(" ");
    let day = date[1];
    let month = date[0];
    month = "JanFebMarAprMayJunJulAugSepOctNovDec".indexOf(month) / 3 + 1;
    if (month.toString().length <= 1) {
        month = '0' + month;
    }
    let year = date[2];
    return String(day + '/' + month + '/' + year);
};

function excelDateToJSDate3(excel_date, time = false) {
    let day_time = excel_date % 1
    let meridiem = "AMPM"
    let hour = Math.floor(day_time * 24)
    let minute = Math.floor(Math.abs(day_time * 24 * 60) % 60)
    let second = Math.floor(Math.abs(day_time * 24 * 60 * 60) % 60)
    hour >= 12 ? meridiem = meridiem.slice(2, 4) : meridiem = meridiem.slice(0, 2)
    hour > 12 ? hour = hour : hour = hour
    hour = hour < 10 ? "0" + hour : hour
    minute = minute < 10 ? "0" + minute : minute
    second = second < 10 ? "0" + second : second
    let daytime = "" + hour + ":" + minute + ":" + second
    return time ? daytime : daytime
};

function excelFileToJSON(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);


                let roaCortos = roa.filter((elem) => elem.Secciones < 8);
                let roaLargos = roa.filter((elem) => elem.Secciones > 8);

                let newArray = roaLargos.filter((elem) => elem.Diferencia > 7);
                let newArray1 = newArray.filter((elem) => elem.Ramal < 316);
                let newArray2 = newArray1.filter((elem) => elem.Kms > 30);
                let newArray6 = roaCortos.filter((elem) => elem.Kms > 12);
                let newArray7 = newArray6.filter((elem) => elem.Diferencia > 3);

                roaCortos = newArray7;

                for (const elem of newArray2) {
                    roaCortos.push(elem);
                }

                let newArray3 = [];
                for (const elem of roaCortos) {

                    let result = (({
                        Interno,
                        Ramal,
                        Legajo,
                        Kms,
                        Secciones,
                        Chofer,
                        Diferencia
                    }) => ({
                        Interno,
                        Ramal,
                        Legajo,
                        Kms,
                        Secciones,
                        Chofer,
                        Diferencia
                    }))(elem);

                    newArray3.push(result);
                }
                roa = newArray3;

                roa.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);




                for (const elem of roa) {
                    const node = document.createElement("tr");
                    const subNode = document.createElement("td");
                    const subNode1 = document.createElement("td");
                    const subNode2 = document.createElement("td");
                    const subNode3 = document.createElement("td");
                    const subNode4 = document.createElement("td");
                    const subNode5 = document.createElement("td");
                    const subNode6 = document.createElement("td");

                    const textnode = document.createTextNode(elem.Interno);
                    const textnode1 = document.createTextNode(elem.Ramal);
                    const textnode2 = document.createTextNode(elem.Legajo);
                    const textnode3 = document.createTextNode(elem.Kms);
                    const textnode4 = document.createTextNode(elem.Secciones);
                    const textnode5 = document.createTextNode(elem.Chofer);
                    const textnode6 = document.createTextNode(elem.Diferencia);
                    subNode.appendChild(textnode);
                    subNode1.appendChild(textnode1);
                    subNode2.appendChild(textnode2);
                    subNode3.appendChild(textnode3);
                    subNode4.appendChild(textnode4);
                    subNode5.appendChild(textnode5);
                    subNode6.appendChild(textnode6);
                    node.appendChild(subNode);
                    node.appendChild(subNode1);
                    node.appendChild(subNode2);
                    node.appendChild(subNode3);
                    node.appendChild(subNode4);
                    node.appendChild(subNode5);
                    node.appendChild(subNode6);
                    tableP.appendChild(node);
                }
                if (roa.length > 0) {

                    result[sheetName] = roa;
                }
            });

        }
    } catch (e) {
        console.error(e);
    }
}


function upload2() {
    var files = document.getElementById('file_upload2').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON2(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

let result2 = {};
let roa2;

let tableP2 = document.getElementById("tableP2");

function excelFileToJSON2(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa2 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

                for (const elem of roa2) {

                    elem.FechaInicio = ExcelDateToJSDate2(elem.FechaInicio);
                    elem.FechaFin = ExcelDateToJSDate2(elem.FechaFin);
                }

                for (const elem of roa2) {

                    elem.HoraInicio = excelDateToJSDate3(elem.HoraInicio);
                    elem.HoraFin = excelDateToJSDate3(elem.HoraFin);
                }

                let fechasDisp = [];

                for (const elem of roa2) {
                    if (fechasDisp.some((n) => n == elem.FechaInicio) == false) {
                        fechasDisp.push(elem.FechaInicio);
                    }
                }


                fechasDisp.sort((a, b) => (a > b) ? 1 : -1);



                let drop = document.getElementsByClassName("dropdown-item");
                let dropIndex = 0;
                let dropAr = [];

                for (const elem of fechasDisp) {

                    drop[dropIndex].innerText = elem;
                    dropAr.push("drop" + dropIndex);
                    dropIndex++;
                }


                for (i = 0; i < fechasDisp.length; i++) {
                    let w = document.getElementById(dropAr[i])
                    w.addEventListener("click", () => {
                        Write(w.textContent)
                    });
                }

                let new6 = [];

                function Write(a) {
                    let infoP2 = document.getElementsByClassName("infoP2");

                    if (infoP2.length > 0) {
                        console.log(infoP2);
                        do {
                            tableP2.removeChild(infoP2[0]);
                            console.log(infoP2);

                        } while (infoP2.length != 0);
                    }
                    console.log(roa2);

                    let newArray10 = roa2.filter((elem) => a == elem.FechaInicio);
                    let newPP = newArray10.filter((elem) => a !== elem.FechaFin);
                    let newPPP = newPP.filter((elem) => elem.HoraFin > "02:00:00");
                    let newArray11 = newPPP.filter((elem) => elem.kms > 15);



                    let newArray20 = newArray10.filter((elem) => a == elem.FechaFin);
                    let kmDeMas63 = newArray20.filter((elem) => elem.kms > "63");
                    let pcoMasde48 = newArray20.filter((elem) => elem.kms > "48" && elem.Recorrido == "PCO COM");
                    let pcoRapMasde48 = newArray20.filter((elem) => elem.kms > "48" && elem.Recorrido == "PCO RAP");
                    let fonMasde54 = newArray20.filter((elem) => elem.kms > "54" && elem.Recorrido == "FON COM");
                    let fonRapMasde54 = newArray20.filter((elem) => elem.kms > "54" && elem.Recorrido == "FON RAP");

                    let new2 = newArray11.concat(kmDeMas63);
                    let new3 = new2.concat(pcoMasde48);
                    let new4 = new3.concat(pcoRapMasde48);
                    let new5 = new4.concat(fonMasde54);
                    new6 = new5.concat(fonRapMasde54);

                    new6.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);


                    for (const elem of new6) {
                        const node = document.createElement("tr");
                        node.classList.add("infoP2");
                        const subNode = document.createElement("td");
                        const subNode1 = document.createElement("td");
                        const subNode2 = document.createElement("td");
                        const subNode3 = document.createElement("td");
                        const subNode4 = document.createElement("td");
                        const subNode5 = document.createElement("td");
                        const subNode6 = document.createElement("td");
                        const subNode7 = document.createElement("td");

                        const textnode = document.createTextNode(elem.Legajo);
                        const textnode1 = document.createTextNode(elem.Interno);
                        const textnode2 = document.createTextNode(elem.FechaInicio);
                        const textnode3 = document.createTextNode(elem.FechaFin);
                        const textnode4 = document.createTextNode(elem.HoraInicio);
                        const textnode5 = document.createTextNode(elem.HoraFin);
                        const textnode6 = document.createTextNode(elem.Recorrido);
                        const textnode7 = document.createTextNode(elem.kms);
                        subNode.appendChild(textnode);
                        subNode1.appendChild(textnode1);
                        subNode2.appendChild(textnode2);
                        subNode3.appendChild(textnode3);
                        subNode4.appendChild(textnode4);
                        subNode5.appendChild(textnode5);
                        subNode6.appendChild(textnode6);
                        subNode7.appendChild(textnode7);
                        node.appendChild(subNode);
                        node.appendChild(subNode1);
                        node.appendChild(subNode2);
                        node.appendChild(subNode3);
                        node.appendChild(subNode4);
                        node.appendChild(subNode5);
                        node.appendChild(subNode6);
                        node.appendChild(subNode7);
                        tableP2.appendChild(node);

                    }
                }
                if (roa2.length > 0) {
                    result[sheetName] = roa2;
                }
            });
        }
    } catch (e) {
        console.error(e);
    }
}

function upload3() {
    var files = document.getElementById('file_upload3').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON3(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}
let roa3;

function excelFileToJSON3(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa3 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

                for (const elem of roa3) {


                    elem.interno = elem.__EMPTY;
                    elem.fechaInicio = elem.__EMPTY_1;
                    elem.horaInicio = elem.__EMPTY_2;
                    elem.recorrido = elem.__EMPTY_4;
                    elem.ramal = elem.__EMPTY_5;
                    elem.direccion = elem.__EMPTY_6;
                    elem.seccion = elem.__EMPTY_7;
                    elem.legajo = elem.__EMPTY_8;
                    elem.tipoDeMarca = elem.__EMPTY_9;

                    delete elem.__EMPTY;
                    delete elem.__EMPTY_1;
                    delete elem.__EMPTY_2;
                    delete elem.__EMPTY_3;
                    delete elem.__EMPTY_4;
                    delete elem.__EMPTY_5;
                    delete elem.__EMPTY_6;
                    delete elem.__EMPTY_7;
                    delete elem.__EMPTY_8;
                    delete elem.__EMPTY_9;
                    delete elem.__EMPTY_10;
                    delete elem.__EMPTY_11;
                    delete elem.__EMPTY_12;
                }


                let soloCambio = roa3.filter((elem) => elem.tipoDeMarca == "Cambio Seccion");
                let soloCambio2 = soloCambio.filter((elem) => elem.legajo != "0");
                let porInterno = [];


                for (const elem of soloCambio2) {
                    elem.fechaInicio = ExcelDateToJSDate2(elem.fechaInicio);
                    elem.horaInicio = excelDateToJSDate3(elem.horaInicio);
                    let e = elem.horaInicio.split(':');
                    elem.horaSinSec = (`${e[0]}:${e[1]}`);

                    if (porInterno.some((n) => n == elem.interno) == false) {
                        porInterno.push(elem.interno);
                    }
                }

                let x2 = [];
                let x3 = [];
                let x4 = [];
                let intArr = [];

                for (const elem of porInterno) {
                    let x = soloCambio2.filter((elem2) => elem2.interno == elem);
                    x2.push(x);
                }

                for (i = 0; i < porInterno.length; i++) {
                    x3[i] = soloCambio2.filter((elem) => elem.interno == porInterno[i]);
                }

                x4 = [...x3];
                for (i = 0; i < x4.length; i++) {
                    for (ii = 0; ii < x4[i].length; ii++) {
                        let m = x4[i][ii].horaSinSec;
                        intArr[i] = x4[i][ii].interno;
                        x4[i][ii] = m;
                    }
                }
                for (f = 0; f < 3; f++) {
                    for (i = 0; i < x4.length; i++) {
                        const dup = x4[i].filter((item, index) => x4[i].indexOf(item) !== index);
                        x4[i] = dup;
                    }
                }

                for (i = 0; i < x4.length; i++) {
                    let dup = x4[i].filter((item, index) => x4[i].indexOf(item) == index);
                    x4[i] = dup;
                }

                let finalArr2 = [];


                for (i = 0; i < x4.length; i++) {
                    for (ii = 0; ii < x4[i].length; ii++) {
                        let x = soloCambio2.filter((elem) => elem.interno == intArr[i] && elem.horaSinSec == x4[i][ii]);
                        for (const elem of x) {
                            finalArr2.push(elem);
                        }
                    }
                }

                finalArr2.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);

                for (const elem of finalArr2) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP3");
                    const subNode = document.createElement("td");
                    const subNode1 = document.createElement("td");
                    const subNode2 = document.createElement("td");
                    const subNode3 = document.createElement("td");
                    const subNode4 = document.createElement("td");
                    const subNode5 = document.createElement("td");
                    const subNode6 = document.createElement("td");
                    const subNode7 = document.createElement("td");
                    const subNode8 = document.createElement("td");

                    const textnode = document.createTextNode(elem.interno);
                    const textnode1 = document.createTextNode(elem.legajo);
                    const textnode2 = document.createTextNode(elem.fechaInicio);
                    const textnode3 = document.createTextNode(elem.horaInicio);
                    const textnode4 = document.createTextNode(elem.recorrido);
                    const textnode5 = document.createTextNode(elem.ramal);
                    const textnode6 = document.createTextNode(elem.direccion);
                    const textnode7 = document.createTextNode(elem.seccion);
                    const textnode8 = document.createTextNode(elem.tipoDeMarca);
                    subNode.appendChild(textnode);
                    subNode1.appendChild(textnode1);
                    subNode2.appendChild(textnode2);
                    subNode3.appendChild(textnode3);
                    subNode4.appendChild(textnode4);
                    subNode5.appendChild(textnode5);
                    subNode6.appendChild(textnode6);
                    subNode7.appendChild(textnode7);
                    subNode8.appendChild(textnode8);
                    node.appendChild(subNode);
                    node.appendChild(subNode1);
                    node.appendChild(subNode2);
                    node.appendChild(subNode3);
                    node.appendChild(subNode4);
                    node.appendChild(subNode5);
                    node.appendChild(subNode6);
                    node.appendChild(subNode7);
                    node.appendChild(subNode8);
                    tableP3.appendChild(node);

                }


                if (roa3.length > 0) {

                    result[sheetName] = roa3;
                }
            });
        }
    } catch (e) {
        console.error(e);
    }
}

function upload4() {
    var files = document.getElementById('file_upload4').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON4(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}
let roa4;

let tableP4 = document.getElementById("tableP4");

function excelFileToJSON4(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa4 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);


                for (const elem of roa4) {

                    elem.Fecha = ExcelDateToJSDate2(elem.Fecha);
                    elem.Hora = excelDateToJSDate3(elem.Hora);
                }
                let temp1 = roa4.filter((elem) => elem.Cabecera != "Estación Benavidez");
                let temp2 = temp1.filter((elem) => elem.MotivoCorte != "VUELTA ANULADA");

                let fechasDisp2 = [];

                for (const elem of temp2) {
                    if (fechasDisp2.some((n) => n == elem.Fecha) == false) {
                        fechasDisp2.push(elem.Fecha);
                    }
                }
                fechasDisp2.sort((a, b) => (a > b) ? 1 : -1);

                let drop2 = document.getElementsByClassName("dropdown-item2");
                let dropIndex2 = 0;
                let dropAr2 = [];

                for (const elem of fechasDisp2) {

                    drop2[dropIndex2].innerText = elem;
                    dropAr2.push("cortadosDrop" + dropIndex2);
                    dropIndex2++;
                }

                for (i = 0; i < fechasDisp2.length; i++) {
                    let j = document.getElementById(dropAr2[i])
                    j.addEventListener("click", () => {
                        Write2(j.textContent)
                    });
                }

                function Write2(a) {
                    let infoP4 = document.getElementsByClassName("infoP4");

                    if (infoP4.length > 0) {
                        console.log(infoP4);
                        do {
                            tableP4.removeChild(infoP4[0]);
                            console.log(infoP4);

                        } while (infoP4.length != 0);
                    }

                    let soloFecha = temp2.filter((elem) => a == elem.Fecha);

                    soloFecha.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);

                    for (const elem of soloFecha) {
                        const node = document.createElement("tr");
                        node.classList.add("infoP4");
                        const subNode = document.createElement("td");
                        const subNode1 = document.createElement("td");
                        const subNode2 = document.createElement("td");
                        const subNode3 = document.createElement("td");
                        const subNode4 = document.createElement("td");
                        const subNode5 = document.createElement("td");


                        const textnode = document.createTextNode(elem.Coche);
                        const textnode1 = document.createTextNode(elem.Legajo);
                        const textnode5 = document.createTextNode(elem.Apellido);
                        const textnode2 = document.createTextNode(elem.Hora);
                        const textnode3 = document.createTextNode(elem.Recorrido);
                        const textnode4 = document.createTextNode(elem.MotivoCorte);


                        subNode.appendChild(textnode);
                        subNode1.appendChild(textnode1);
                        subNode2.appendChild(textnode2);
                        subNode3.appendChild(textnode3);
                        subNode4.appendChild(textnode4);
                        subNode5.appendChild(textnode5);

                        node.appendChild(subNode);
                        node.appendChild(subNode1);
                        node.appendChild(subNode5);
                        node.appendChild(subNode2);
                        node.appendChild(subNode3);
                        node.appendChild(subNode4);

                        tableP4.appendChild(node);

                    }
                }



                if (roa4.length > 0) {

                    result[sheetName] = roa4;
                }
            });
        }
    } catch (e) {
        console.error(e);
    }
}


function upload5() {
    var files = document.getElementById('file_upload5').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON5(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}
let roa5;
let benIda;
let benVta;
let benRelIda;
let pcoRelIda;
let fonRelIda;
let pcoIda;
let pcoVta;
let fonIda;
let fonVta;
let ida197;
let vta197;
let ida202;
let vta202;
let rivIda;
let rivVta;
let bcasIda;
let bcasVta;
let talarABen;
let talarAPco;
let talarAFon;
let benATalar;
let pcoATalar;
let fonATalar;


function excelFileToJSON5(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa5 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

                for (const elem of roa5){
                    let x = elem.RecIda.split("");
                    x.shift();
                    elem.RecIda = x.join("");
                    if (elem.RecVuelta!=undefined){
                        let xx = elem.RecVuelta.split("");
                        xx.shift();
                        elem.RecVuelta = xx.join("");
                    }
                }
                

                
                let temp1 = roa5.filter((elem) => elem.Cabecera != "Estación Benavidez ");
                let temp2 = temp1.filter((elem) => elem.Tipo != "C ");
                let temp3 = temp2.filter((elem) => elem.Tipo != "CE ");
                let temp4 = temp3.filter((elem) => elem.Legajo != "2743");
                console.log(temp4);

                let benIda = temp4.filter((elem) => elem.RecIda === "BEN LaV COM IDA"|| elem.RecIda ==="BEN SAB COM IDA" || elem.RecIda ==="BEN DOM IDA" || elem.RecIda ==="BEN FER IDA"|| elem.RecIda ==="BEN LaV RAP IDA" || elem.RecIda ==="BEN SAB RAP IDA");
                let pcoIda = temp4.filter((elem) => elem.RecIda === "PCO LaV COM IDA"|| elem.RecIda ==="PCO SAB COM IDA" || elem.RecIda ==="PCO DOM IDA" || elem.RecIda ==="PCO FER IDA"|| elem.RecIda ==="PCO LaV RAP IDA" || elem.RecIda ==="PCO SAB RAP IDA");
                let fonIda = temp4.filter((elem) => elem.RecIda === "FON LaV COM IDA"|| elem.RecIda ==="FON SAB COM IDA" || elem.RecIda ==="FON DOM IDA" || elem.RecIda ==="FON FER IDA"|| elem.RecIda ==="FON LaV RAP IDA" || elem.RecIda ==="FON SAB RAP IDA");
                let fonVta = temp4.filter((elem) => elem.RecVuelta === "FON LaV COM VTA" || elem.RecVuelta ==="FON SAB COM VTA" || elem.RecVuelta ==="FON DOM VTA" || elem.RecVuelta ==="FON FER VTA"|| elem.RecVuelta ==="FON LaV RAP VTA" || elem.RecVuelta ==="FON SAB RAP VTA");
                let pcoVta = temp4.filter((elem) => elem.RecVuelta === "PCO LaV COM VTA" || elem.RecVuelta ==="PCO SAB COM VTA" || elem.RecVuelta ==="PCO DOM VTA" || elem.RecVuelta ==="PCO FER VTA"|| elem.RecVuelta ==="PCO LaV RAP VTA" || elem.RecVuelta ==="PCO SAB RAP VTA"|| elem.RecVuelta ==="PCO LaV R S/202 VTA");
                let benVta = temp4.filter((elem) => elem.RecVuelta === "BEN LaV COM VTA" || elem.RecVuelta ==="BEN SAB COM VTA" || elem.RecVuelta ==="BEN DOM VTA" || elem.RecVuelta ==="BEN FER VTA"|| elem.RecVuelta ==="BEN LaV RAP VTA" || elem.RecVuelta ==="BEN SAB RAP VTA"|| elem.RecVuelta ==="BEN SAB R S/202 VTA"|| elem.RecVuelta ==="BEN LaV R DIR VTA"|| elem.RecVuelta ==="BEN LaV R S/202 VTA");
                let ida197 = temp4.filter((elem) => elem.RecIda === "197 LaV COM IDA"|| elem.RecIda ==="197 SAB COM IDA" || elem.RecIda ==="197 DOM IDA" || elem.RecIda ==="197 FER IDA"|| elem.RecIda ==="197 LaV RAP IDA" || elem.RecIda ==="197 SAB RAP IDA"); 
                let ida202 = temp4.filter((elem) => elem.RecIda === "202 LaV COM IDA"|| elem.RecIda ==="202 SAB COM IDA" || elem.RecIda ==="202 DOM IDA" || elem.RecIda ==="202 FER IDA"|| elem.RecIda ==="202 LaV RAP IDA" || elem.RecIda ==="202 SAB RAP IDA");  
                let rivIda = temp4.filter((elem) => elem.RecIda === "RIV LaV IDA"|| elem.RecIda ==="RIV SAB IDA" || elem.RecIda ==="RIV DOM IDA" || elem.RecIda ==="RIV FER IDA");
                let bcasIda = temp4.filter((elem) => elem.RecIda === "BCAS LaV IDA"|| elem.RecIda ==="BCAS SAB IDA" || elem.RecIda ==="BCAS DOM IDA" || elem.RecIda ==="BCAS FER IDA"); 
                let vta197 = temp4.filter((elem) => elem.RecVuelta === "197 LaV COM VTA"|| elem.RecVuelta ==="197 SAB COM VTA" || elem.RecVuelta ==="197 DOM VTA" || elem.RecVuelta ==="197 FER VTA"|| elem.RecVuelta ==="197 LaV RAP VTA" || elem.RecVuelta ==="197 SAB RAP VTA"|| elem.RecVuelta ==="197 LaV R S/202 VTA");  
                let vta202 = temp4.filter((elem) => elem.RecVuelta === "202 LaV COM VTA"|| elem.RecVuelta ==="202 SAB COM VTA" || elem.RecVuelta ==="202 DOM VTA" || elem.RecVuelta ==="202 FER VTA"|| elem.RecVuelta ==="202 LaV RAP VTA" || elem.RecVuelta ==="202 SAB RAP VTA"); 
                let rivVta = temp4.filter((elem) => elem.RecVuelta === "RIV LaV VTA"|| elem.RecVuelta ==="RIV SAB VTA" || elem.RecVuelta ==="RIV DOM VTA" || elem.RecVuelta ==="RIV FER VTA"); 
                let bcasVta = temp4.filter((elem) => elem.RecVuelta === "BCAS LaV VTA"|| elem.RecVuelta ==="BCAS SAB VTA" || elem.RecVuelta ==="BCAS DOM VTA" || elem.RecVuelta ==="BCAS FER VTA"); 
                let benRelIda = temp4.filter((elem) => elem.RecIda === "BEN LaV REL COM IDA"|| elem.RecIda ==="BEN SAB REL COM IDA" || elem.RecIda ==="BEN DOM REL IDA" || elem.RecIda ==="BEN FER REL IDA"|| elem.RecIda ==="BEN LaV REL RAP IDA" || elem.RecIda ==="BEN SAB REL RAP IDA");
                let pcoRelIda = temp4.filter((elem) => elem.RecIda === "PCO LaV REL COM IDA"|| elem.RecIda ==="PCO SAB REL COM IDA" || elem.RecIda ==="PCO DOM REL IDA" || elem.RecIda ==="PCO FER REL IDA"|| elem.RecIda ==="PCO LaV REL RAP IDA" || elem.RecIda ==="PCO SAB REL RAP IDA");
                let fonRelIda = temp4.filter((elem) => elem.RecIda === "FON LaV REL COM IDA"|| elem.RecIda ==="FON SAB REL COM IDA" || elem.RecIda ==="FON DOM REL IDA" || elem.RecIda ==="FON FER REL IDA"|| elem.RecIda ==="FON LaV REL RAP IDA" || elem.RecIda ==="FON SAB REL RAP IDA");
                let talarABen = temp4.filter((elem) => elem.RecIda === "TALAR A BEN IDA 2"||elem.RecIda === "TALAR A BEN IDA");
                let talarAPco = temp4.filter((elem) => elem.RecIda === "TALAR A PCO IDA 2"||elem.RecIda === "TALAR A PCO IDA");
                let talarAFon = temp4.filter((elem) => elem.RecIda === "TALAR A FON IDA 2"||elem.RecIda === "TALAR A FON IDA");
                let benATalar = temp4.filter((elem) => elem.RecVuelta === "BEN A TALAR VTA");
                let pcoATalar = temp4.filter((elem) => elem.RecVuelta === "PCO A TALAR VTA");
                let fonATalar = temp4.filter((elem) => elem.RecVuelta === "FON A TALAR VTA");

                let arrayRec=[];

                arrayRec.push(benIda);
                arrayRec.push(benVta);
                arrayRec.push(benRelIda);
                arrayRec.push(talarABen);
                arrayRec.push(benATalar);
                arrayRec.push(pcoIda);
                arrayRec.push(pcoVta);
                arrayRec.push(pcoRelIda);
                arrayRec.push(talarAPco);
                arrayRec.push(pcoATalar);
                arrayRec.push(fonIda);
                arrayRec.push(fonVta);
                arrayRec.push(fonRelIda);
                arrayRec.push(talarAFon);
                arrayRec.push(fonATalar);
                arrayRec.push(ida197);
                arrayRec.push(vta197);
                arrayRec.push(ida202);
                arrayRec.push(vta202);
                arrayRec.push(rivIda);
                arrayRec.push(rivVta);
                arrayRec.push(bcasIda);
                arrayRec.push(bcasVta);

                console.log(arrayRec);


                let tableP5 = document.getElementById("tableP5");

                for (const elem of arrayRec) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP5");
                    const subNode = document.createElement("td");

                    const textnode = document.createTextNode(elem.length);
                    subNode.appendChild(textnode);
                    node.appendChild(subNode);
                    tableP5.appendChild(node);

                }


                if (roa5.length > 0) {

                    result[sheetName] = roa5;
                }
            });
        }
    } catch (e) {
        console.error(e);
    }
}

function upload6() {
    var files = document.getElementById('file_upload6').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON6(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

let result6 = {};
let roa6;
let choferesSeccionamiento=[];
let tempChoferesSeccionamiento=[];
let choferesMalUso=[];
let tempChoferesMalUso=[];
let choferesCortados=[];
let tempChoferesCortados=[];
let corridos=[];
let tempCorridos=[];
let corridosCant=[];

let tableP6 = document.getElementById("tableP6");

function excelFileToJSON6(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });


            workbook.SheetNames.forEach(function(sheetName) {
                roa6 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

            //console.log (roa6);

            let numberMalUso = roa6.findIndex((elem) =>elem.props == "CONTROL MAL USO DEL SUBE");
            let numberCortados = roa6.findIndex((elem) =>elem.props == "SERVICIOS CORTADOS");
            let numberSecCorridos = roa6.findIndex((elem) =>elem.props == "CONTROL SECCIONAMIENTO CORRIDO");


            let arraySeccionamiento=[];
            let arrayMalUso=[];
            let arrayCortados=[];
            let arrayCorridos=[];

            for (i=2; i<numberMalUso; i++){
                arraySeccionamiento.push(roa6[i]);
            }
            for (i=numberMalUso+2; i<numberCortados; i++){
                arrayMalUso.push(roa6[i]);
            }
            for (i=numberCortados+2; i<numberSecCorridos; i++){
                arrayCortados.push(roa6[i]);
            }
            for (i=numberSecCorridos+2; i<roa6.length; i++){
                arrayCorridos.push(roa6[i]);
            }
            

            for (const elem of arraySeccionamiento){
                xx = elem.__EMPTY_1;
                tempChoferesSeccionamiento.push({legajo:xx});
            }
            for (const elem of arrayMalUso){
                xx = elem.props;
                tempChoferesMalUso.push({legajo:xx});
            }
            for (const elem of arrayCortados){
                xx = elem.__EMPTY;
                tempChoferesCortados.push({legajo:xx});
            }
            for (const elem of arrayCorridos) {

                elem.fecha = ExcelDateToJSDate2(elem.__EMPTY_1);
                elem.hora = excelDateToJSDate3(elem.__EMPTY_2);
            }
            let pija=[];
            for (const elem of arrayCorridos){
                xx = elem.__EMPTY;
                tempCorridos.push(xx);
            }
            for (const elem of tempCorridos){
                pija.push(elem);
            }
            for (const el of pija){
            if (corridosCant.includes(el)==false){
                corridosCant.push(el);
            }
        }


            //console.log(corridosCant);

           //console.log(arrayCortados);
            //console.log(tempChoferesSeccionamiento);

                /* for (const elem of roa2) {

                    elem.FechaInicio = ExcelDateToJSDate2(elem.FechaInicio);
                    elem.FechaFin = ExcelDateToJSDate2(elem.FechaFin);
                }

                for (const elem of roa2) {

                    elem.HoraInicio = excelDateToJSDate3(elem.HoraInicio);
                    elem.HoraFin = excelDateToJSDate3(elem.HoraFin);
                }

                let fechasDisp = [];

                for (const elem of roa2) {
                    if (fechasDisp.some((n) => n == elem.FechaInicio) == false) {
                        fechasDisp.push(elem.FechaInicio);
                    }
                }


                fechasDisp.sort((a, b) => (a > b) ? 1 : -1);



                let drop = document.getElementsByClassName("dropdown-item");
                let dropIndex = 0;
                let dropAr = [];

                for (const elem of fechasDisp) {

                    drop[dropIndex].innerText = elem;
                    dropAr.push("drop" + dropIndex);
                    dropIndex++;
                }


                for (i = 0; i < fechasDisp.length; i++) {
                    let w = document.getElementById(dropAr[i])
                    w.addEventListener("click", () => {
                        Write(w.textContent)
                    });
                }

                let new6 = [];

                function Write(a) {
                    let infoP2 = document.getElementsByClassName("infoP2");

                    if (infoP2.length > 0) {
                        console.log(infoP2);
                        do {
                            tableP2.removeChild(infoP2[0]);
                            console.log(infoP2);

                        } while (infoP2.length != 0);
                    }
                    console.log(roa2);

                    let newArray10 = roa2.filter((elem) => a == elem.FechaInicio);
                    let newPP = newArray10.filter((elem) => a !== elem.FechaFin);
                    let newPPP = newPP.filter((elem) => elem.HoraFin > "02:00:00");
                    let newArray11 = newPPP.filter((elem) => elem.kms > 15);



                    let newArray20 = newArray10.filter((elem) => a == elem.FechaFin);
                    let kmDeMas63 = newArray20.filter((elem) => elem.kms > "63");
                    let pcoMasde48 = newArray20.filter((elem) => elem.kms > "48" && elem.Recorrido == "PCO COM");
                    let pcoRapMasde48 = newArray20.filter((elem) => elem.kms > "48" && elem.Recorrido == "PCO RAP");
                    let fonMasde54 = newArray20.filter((elem) => elem.kms > "54" && elem.Recorrido == "FON COM");
                    let fonRapMasde54 = newArray20.filter((elem) => elem.kms > "54" && elem.Recorrido == "FON RAP");

                    let new2 = newArray11.concat(kmDeMas63);
                    let new3 = new2.concat(pcoMasde48);
                    let new4 = new3.concat(pcoRapMasde48);
                    let new5 = new4.concat(fonMasde54);
                    new6 = new5.concat(fonRapMasde54);

                    new6.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);


                    for (const elem of new6) {
                        const node = document.createElement("tr");
                        node.classList.add("infoP2");
                        const subNode = document.createElement("td");
                        const subNode1 = document.createElement("td");
                        const subNode2 = document.createElement("td");
                        const subNode3 = document.createElement("td");
                        const subNode4 = document.createElement("td");
                        const subNode5 = document.createElement("td");
                        const subNode6 = document.createElement("td");
                        const subNode7 = document.createElement("td");

                        const textnode = document.createTextNode(elem.Legajo);
                        const textnode1 = document.createTextNode(elem.Interno);
                        const textnode2 = document.createTextNode(elem.FechaInicio);
                        const textnode3 = document.createTextNode(elem.FechaFin);
                        const textnode4 = document.createTextNode(elem.HoraInicio);
                        const textnode5 = document.createTextNode(elem.HoraFin);
                        const textnode6 = document.createTextNode(elem.Recorrido);
                        const textnode7 = document.createTextNode(elem.kms);
                        subNode.appendChild(textnode);
                        subNode1.appendChild(textnode1);
                        subNode2.appendChild(textnode2);
                        subNode3.appendChild(textnode3);
                        subNode4.appendChild(textnode4);
                        subNode5.appendChild(textnode5);
                        subNode6.appendChild(textnode6);
                        subNode7.appendChild(textnode7);
                        node.appendChild(subNode);
                        node.appendChild(subNode1);
                        node.appendChild(subNode2);
                        node.appendChild(subNode3);
                        node.appendChild(subNode4);
                        node.appendChild(subNode5);
                        node.appendChild(subNode6);
                        node.appendChild(subNode7);
                        tableP2.appendChild(node);

                    }
                } */
            
                if (roa6.length > 0) {
                    result[sheetName] = roa6;
                }
            });


            
            for (const elem of tempChoferesSeccionamiento){
                xx = elem.legajo;
                let x = tempChoferesSeccionamiento.filter((el)=>el.legajo==elem.legajo);
                choferesSeccionamiento.push({legajo:xx, seccionamiento:x.length});
                }

                choferesSeccionamiento = choferesSeccionamiento.filter((e)=>e.seccionamiento>9);

            for (const elem of tempChoferesMalUso){
                xx = elem.legajo;
                let x = tempChoferesMalUso.filter((el)=>el.legajo==elem.legajo);
                choferesMalUso.push({legajo:xx, malUso:x.length});
                }

                choferesMalUso = choferesMalUso.filter((e)=>e.malUso>9);

            for (const elem of tempChoferesCortados){
                xx = elem.legajo;
                let x = tempChoferesCortados.filter((el)=>el.legajo==elem.legajo);
                choferesCortados.push({legajo:xx, cortes:x.length});
                }

                choferesCortados = choferesCortados.filter((e)=>e.cortes>2);

            for (const elem of corridosCant){
                let x = tempCorridos.filter((el)=>el==elem);
                //console.log(x.length);
                corridos.push({legajo:elem,seccCorrido:x.length});
                }
                corridos = corridos.filter((e)=>e.seccCorrido>40);
                
                

                let pene=[];
                let pene2=[];
                let pene3=[];
                let pene4=[];
                let pene5=[];
                let pene6=[];
                let pene7=[];
                let pene8=[];
                let pene9=[];

                for (i=0;i<choferesSeccionamiento.length;i++){
                        pene.push(choferesSeccionamiento[i].legajo);
                }

                for (const el of pene){
                    if (pene2.includes(el)==false){
                        pene2.push(el);
                    }
                }
                for(const elem of pene2){
                    x = tempChoferesSeccionamiento.filter((el)=>el.legajo==elem);
                    pene3.push({legajo:elem, seccionamiento:x.length});
                }
                choferesSeccionamiento = pene3;

                
                for (i=0;i<choferesMalUso.length;i++){
                    pene4.push(choferesMalUso[i].legajo);
                }

                for (const el of pene4){
                    if (pene5.includes(el)==false){
                        pene5.push(el);
                    }
                }
                for(const elem of pene5){
                    x = tempChoferesMalUso.filter((el)=>el.legajo==elem);
                    pene6.push({legajo:elem, malUso:x.length});
                }
                choferesMalUso = pene6;
            



                for (i=0;i<choferesCortados.length;i++){
                    pene7.push(choferesCortados[i].legajo);
                }
                for (const el of pene7){
                    if (pene8.includes(el)==false){
                        pene8.push(el);
                    }
                }
                for(const elem of pene8){
                    x = tempChoferesCortados.filter((el)=>el.legajo==elem);
                    pene9.push({legajo:elem, cortes:x.length});
                }
                choferesCortados = pene9;
                
                let soloChoferes=[];
                
                for (i=0;i<choferesSeccionamiento.length;i++){
                        soloChoferes.push({
                            legajo: choferesSeccionamiento[i].legajo,
                            seccionamiento: choferesSeccionamiento[i].seccionamiento,
                            malUso: 0,
                            cortes:0,
                            seccCorridos:0,
                        });
                }
                /* for (const elem of soloChoferes){
                    for(i=0;i<choferesMalUso.length;i++){

                        if (elem.legajo===choferesMalUso[i].legajo==true){
                            
                            elem.malUso = choferesMalUso[i].malUso;
                        } else{
                            
                        soloChoferes.push({
                            legajo: choferesMalUso[i].legajo,
                            seccionamiento: 0,
                            malUso: choferesMalUso[i].malUso,
                            cortes:0,
                            seccCorridos:0,
                        });
                    }
                }
            } */

                console.log(soloChoferes);






        }
    } catch (e) {
        console.error(e);
    }
}

//PARA CUANDO SE ARREGLE EL TEMA DE LA DIFERENCIA DE HORARIO
//  let masHoras = [];
//for (const elem of newArray20){

//  let s = elem.HoraInicio.split(':');
//let e = elem.HoraFin.split(':');

//let tiempo = e[0] - s[0];

// if (tiempo > "6"==true){ masHoras.push(elem);}