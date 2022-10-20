let l1 = document.getElementById("controlSeccionamientoH1");
let l2 = document.getElementById("controlMalUsoDelSubeH1");
let l3 = document.getElementById("controlSeccCorridoH1");
let l4 = document.getElementById("controlCortadosH1");
let l5 = document.getElementById("controlKmPorRamalH1");
let l6 = document.getElementById("controlMensualH1");
let l7 = document.getElementById("controlFrancosH1");
let l8 = document.getElementById("controlPlanillasH1");
let l9 = document.getElementById("controlKMSubeH1");

function VerControl(a, b, c, d, e, f, g, h, i) {
    a.style.visibility = "visible";
    b.style.visibility = "hidden";
    c.style.visibility = "hidden";
    d.style.visibility = "hidden";
    e.style.visibility = "hidden";
    f.style.visibility = "hidden";
    g.style.visibility = "hidden";
    h.style.visibility = "hidden";
    i.style.visibility = "hidden";


}

let arrLinks = [l1,l2,l3,l4,l5,l6,l7,l8,l9];

document.getElementById("controlSeccionamiento").addEventListener("click", () => {
    VerControl(l1, l2, l3, l4, l5, l6, l7, l8, l9)});
document.getElementById("controlMalUsoDelSube").addEventListener("click", () => {
    VerControl(l2, l1, l3, l4, l5, l6, l7, l8, l9)});
document.getElementById("controlCortados").addEventListener("click", () => {
    VerControl(l3, l2, l1, l4, l5, l6, l7, l8, l9)});
document.getElementById("controlSeccCorrido").addEventListener("click", () => {
    VerControl(l4, l2, l3, l1, l5, l6, l7, l8, l9)});
document.getElementById("controlKmPorRamal").addEventListener("click", () => {
    VerControl(l5, l2, l3, l4, l1, l6, l7, l8, l9)});
document.getElementById("controlMensual").addEventListener("click", () => {
    VerControl(l6, l2, l3, l4, l5, l1, l7, l8, l9)});
document.getElementById("controlFrancos").addEventListener("click", () => {
    VerControl(l7, l2, l3, l4, l5, l6, l1, l8, l9)});
document.getElementById("controlPlanillas").addEventListener("click", () => {
    VerControl(l8, l2, l3, l4, l5, l6, l7, l1, l9)});
document.getElementById("controlKMSube").addEventListener("click", () => {
    VerControl(l9, l2, l3, l4, l5, l6, l7, l8, l1)});


    let arrayKM4=[];


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

                roa2.shift();


                for (const elem of roa2) {
                    elem.Legajo = elem.__EMPTY;
                    elem.Interno = elem.__EMPTY_1;
                    elem.HoraInicio = elem.__EMPTY_4;
                    elem.FechaInicio = elem.__EMPTY_3;
                    elem.FechaFin = elem.__EMPTY_5;
                    elem.HoraFin = elem.__EMPTY_6;
                    elem.Ramal = elem.__EMPTY_7;
                    elem.Recorrido = elem.__EMPTY_9;
                    elem.Kms = elem.__EMPTY_10;

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
                    fechasDisp.push(elem.FechaInicio);
                }
                
                const fechasDisp3 = [...new Set(fechasDisp)];
                
                fechasDisp = fechasDisp3;

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
                        Write(w.textContent);
                    });
                }

                    
                function Write(a) {
                    let infoP2 = document.getElementsByClassName("infoP2");

                    if (infoP2.length > 0) {
                        do {
                            tableP2.removeChild(infoP2[0]);

                        } while (infoP2.length != 0);
                    }
                    
                    console.log (roa2)
                    let newArray10 = roa2.filter((elem) => a == elem.FechaInicio);
                    

                    let newPP = newArray10.filter((elem) => a !== elem.FechaFin);
                    let newPPP = newPP.filter((elem) => elem.HoraFin > "02:00:00");
                    let newArray11 = newPPP.filter((elem) => elem.Kms > 15);



                    let newArray20 = newArray10.filter((elem) => a == elem.FechaFin);
                    let kmDeMas63 = newArray20.filter((elem) => elem.Kms > "63" && elem.Recorrido.includes("BEN"));
                    let pcoMasde48 = newArray20.filter((elem) => elem.Kms > "48" && elem.Recorrido == "PCO COM");
                    let pcoRapMasde48 = newArray20.filter((elem) => elem.Kms > "48" && elem.Recorrido == "PCO RAP");
                    let fonMasde54 = newArray20.filter((elem) => elem.Kms > "54" && elem.Recorrido == "FON COM");
                    let fonRapMasde54 = newArray20.filter((elem) => elem.Kms > "54" && elem.Recorrido == "FON RAP");

                    let new2 = newArray11.concat(kmDeMas63);
                    let new3 = new2.concat(pcoMasde48);

                    let new4 = new3.concat(pcoRapMasde48);
                    let new5 = new4.concat(fonMasde54);
                    let new6 = new5.concat(fonRapMasde54);

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
                        const textnode7 = document.createTextNode(elem.Kms);
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

                let finalArr3 = [];
                let finalArr4 = [];


                for (const s of finalArr2) {
                    finalArr3.push(s.legajo);
                    finalArr4.push(s.horaSinSec);

                }

                function Duplicados(a) {
                    return [...new Set(a)];
                }

                let finalArr5 = Duplicados(finalArr3);

                let finalArr6 = [];
                for (const e of finalArr5) {
                    let x = finalArr2.filter((el) => el.legajo === e);

                    h = [];
                    for (const e of x) {
                        h.push(e.horaSinSec);
                    }

                    let xx = Duplicados(h);

                    finalArr6.push({
                        legajo: e,
                        cantCorridos: xx.length
                    })
                }

                for (const elem of finalArr6) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP3");
                    const subNode = document.createElement("td");
                    const subNode1 = document.createElement("td");

                    const textnode = document.createTextNode(elem.legajo);
                    const textnode1 = document.createTextNode(elem.cantCorridos);
                    subNode.appendChild(textnode);
                    subNode1.appendChild(textnode1);
                    node.appendChild(subNode);
                    node.appendChild(subNode1);
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
                        do {
                            tableP4.removeChild(infoP4[0]);

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

                for (const elem of roa5) {
                    let x = elem.RecIda.split("");
                    x.shift();
                    elem.RecIda = x.join("");
                    if (elem.RecVuelta != undefined) {
                        let xx = elem.RecVuelta.split("");
                        xx.shift();
                        elem.RecVuelta = xx.join("");
                    }
                }



                let temp1 = roa5.filter((elem) => elem.Cabecera != "Estación Benavidez ");
                let temp2 = temp1.filter((elem) => elem.Tipo != "C ");
                let temp3 = temp2.filter((elem) => elem.Tipo != "CE ");
                let temp4 = temp3.filter((elem) => elem.Legajo != "2743");

                let benIda = temp4.filter((elem) => elem.RecIda === "BEN LaV COM IDA" || elem.RecIda === "BEN SAB COM IDA" || elem.RecIda === "BEN DOM IDA" || elem.RecIda === "BEN FER IDA" || elem.RecIda === "BEN LaV RAP IDA" || elem.RecIda === "BEN SAB RAP IDA");
                let pcoIda = temp4.filter((elem) => elem.RecIda === "PCO LaV COM IDA" || elem.RecIda === "PCO SAB COM IDA" || elem.RecIda === "PCO DOM IDA" || elem.RecIda === "PCO FER IDA" || elem.RecIda === "PCO LaV RAP IDA" || elem.RecIda === "PCO SAB RAP IDA");
                let fonIda = temp4.filter((elem) => elem.RecIda === "FON LaV COM IDA" || elem.RecIda === "FON SAB COM IDA" || elem.RecIda === "FON DOM IDA" || elem.RecIda === "FON FER IDA" || elem.RecIda === "FON LaV RAP IDA" || elem.RecIda === "FON SAB RAP IDA");
                let fonVta = temp4.filter((elem) => elem.RecVuelta === "FON LaV COM VTA" || elem.RecVuelta === "FON SAB COM VTA" || elem.RecVuelta === "FON DOM VTA" || elem.RecVuelta === "FON FER VTA" || elem.RecVuelta === "FON LaV RAP VTA" || elem.RecVuelta === "FON SAB RAP VTA");
                let pcoVta = temp4.filter((elem) => elem.RecVuelta === "PCO LaV COM VTA" || elem.RecVuelta === "PCO SAB COM VTA" || elem.RecVuelta === "PCO DOM VTA" || elem.RecVuelta === "PCO FER VTA" || elem.RecVuelta === "PCO LaV RAP VTA" || elem.RecVuelta === "PCO SAB RAP VTA" || elem.RecVuelta === "PCO LaV R S/202 VTA");
                let benVta = temp4.filter((elem) => elem.RecVuelta === "BEN LaV COM VTA" || elem.RecVuelta === "BEN SAB COM VTA" || elem.RecVuelta === "BEN DOM VTA" || elem.RecVuelta === "BEN FER VTA" || elem.RecVuelta === "BEN LaV RAP VTA" || elem.RecVuelta === "BEN SAB RAP VTA" || elem.RecVuelta === "BEN SAB R S/202 VTA" || elem.RecVuelta === "BEN LaV R DIR VTA" || elem.RecVuelta === "BEN LaV R S/202 VTA");
                let ida197 = temp4.filter((elem) => elem.RecIda === "197 LaV COM IDA" || elem.RecIda === "197 SAB COM IDA" || elem.RecIda === "197 DOM IDA" || elem.RecIda === "197 FER IDA" || elem.RecIda === "197 LaV RAP IDA" || elem.RecIda === "197 SAB RAP IDA");
                let ida202 = temp4.filter((elem) => elem.RecIda === "202 LaV COM IDA" || elem.RecIda === "202 SAB COM IDA" || elem.RecIda === "202 DOM IDA" || elem.RecIda === "202 FER IDA" || elem.RecIda === "202 LaV RAP IDA" || elem.RecIda === "202 SAB RAP IDA");
                let rivIda = temp4.filter((elem) => elem.RecIda === "RIV LaV IDA" || elem.RecIda === "RIV SAB IDA" || elem.RecIda === "RIV DOM IDA" || elem.RecIda === "RIV FER IDA");
                let bcasIda = temp4.filter((elem) => elem.RecIda === "BCAS LaV IDA" || elem.RecIda === "BCAS SAB IDA" || elem.RecIda === "BCAS DOM IDA" || elem.RecIda === "BCAS FER IDA");
                let vta197 = temp4.filter((elem) => elem.RecVuelta === "197 LaV COM VTA" || elem.RecVuelta === "197 SAB COM VTA" || elem.RecVuelta === "197 DOM VTA" || elem.RecVuelta === "197 FER VTA" || elem.RecVuelta === "197 LaV RAP VTA" || elem.RecVuelta === "197 SAB RAP VTA" || elem.RecVuelta === "197 LaV R S/202 VTA");
                let vta202 = temp4.filter((elem) => elem.RecVuelta === "202 LaV COM VTA" || elem.RecVuelta === "202 SAB COM VTA" || elem.RecVuelta === "202 DOM VTA" || elem.RecVuelta === "202 FER VTA" || elem.RecVuelta === "202 LaV RAP VTA" || elem.RecVuelta === "202 SAB RAP VTA");
                let rivVta = temp4.filter((elem) => elem.RecVuelta === "RIV LaV VTA" || elem.RecVuelta === "RIV SAB VTA" || elem.RecVuelta === "RIV DOM VTA" || elem.RecVuelta === "RIV FER VTA");
                let bcasVta = temp4.filter((elem) => elem.RecVuelta === "BCAS LaV VTA" || elem.RecVuelta === "BCAS SAB VTA" || elem.RecVuelta === "BCAS DOM VTA" || elem.RecVuelta === "BCAS FER VTA");
                let benRelIda = temp4.filter((elem) => elem.RecIda === "BEN LaV REL COM IDA" || elem.RecIda === "BEN SAB REL COM IDA" || elem.RecIda === "BEN DOM REL IDA" || elem.RecIda === "BEN FER REL IDA" || elem.RecIda === "BEN LaV REL RAP IDA" || elem.RecIda === "BEN SAB REL RAP IDA");
                let pcoRelIda = temp4.filter((elem) => elem.RecIda === "PCO LaV REL COM IDA" || elem.RecIda === "PCO SAB REL COM IDA" || elem.RecIda === "PCO DOM REL IDA" || elem.RecIda === "PCO FER REL IDA" || elem.RecIda === "PCO LaV REL RAP IDA" || elem.RecIda === "PCO SAB REL RAP IDA");
                let fonRelIda = temp4.filter((elem) => elem.RecIda === "FON LaV REL COM IDA" || elem.RecIda === "FON SAB REL COM IDA" || elem.RecIda === "FON DOM REL IDA" || elem.RecIda === "FON FER REL IDA" || elem.RecIda === "FON LaV REL RAP IDA" || elem.RecIda === "FON SAB REL RAP IDA");
                let talarABen = temp4.filter((elem) => elem.RecIda === "TALAR A BEN IDA 2" || elem.RecIda === "TALAR A BEN IDA");
                let talarAPco = temp4.filter((elem) => elem.RecIda === "TALAR A PCO IDA 2" || elem.RecIda === "TALAR A PCO IDA");
                let talarAFon = temp4.filter((elem) => elem.RecIda === "TALAR A FON IDA 2" || elem.RecIda === "TALAR A FON IDA");
                let benATalar = temp4.filter((elem) => elem.RecVuelta === "BEN A TALAR VTA");
                let pcoATalar = temp4.filter((elem) => elem.RecVuelta === "PCO A TALAR VTA");
                let fonATalar = temp4.filter((elem) => elem.RecVuelta === "FON A TALAR VTA");

                let arrayRec = [];

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
let choferesSeccionamiento = [];
let tempChoferesSeccionamiento = [];
let choferesMalUso = [];
let tempChoferesMalUso = [];
let choferesCortados = [];
let tempChoferesCortados = [];
let corridos = [];
let tempCorridos = [];

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


                let numberMalUso = roa6.findIndex((elem) => elem.props == "CONTROL MAL USO DEL SUBE");
                let numberCortados = roa6.findIndex((elem) => elem.props == "SERVICIOS CORTADOS");
                let numberSecCorridos = roa6.findIndex((elem) => elem.props == "CONTROL SECCIONAMIENTO CORRIDO");



                let arraySeccionamiento = [];
                let arrayMalUso = [];
                let arrayCortados = [];
                let arrayCorridos = [];

                for (i = 2; i < numberMalUso; i++) {
                    arraySeccionamiento.push(roa6[i]);
                }
                for (i = numberMalUso + 2; i < numberCortados; i++) {
                    arrayMalUso.push(roa6[i]);
                }
                for (i = numberCortados + 2; i < numberSecCorridos; i++) {
                    arrayCortados.push(roa6[i]);
                }
                for (i = numberSecCorridos + 2; i < roa6.length; i++) {
                    arrayCorridos.push(roa6[i]);
                }

                for (const elem of arraySeccionamiento) {
                    xx = elem.__EMPTY_1;
                    tempChoferesSeccionamiento.push({
                        legajo: xx
                    });
                }
                for (const elem of arrayMalUso) {
                    xx = elem.props;
                    tempChoferesMalUso.push({
                        legajo: xx
                    });
                }
                for (const elem of arrayCortados) {
                    xx = elem.__EMPTY;
                    tempChoferesCortados.push({
                        legajo: xx
                    });
                }

                let pija = [];

                for (const elem of arrayCorridos) {
                    let leg = elem.props,
                        xx = elem.__EMPTY;
                    tempCorridos.push({
                        leg,
                        xx
                    });
                }



                if (roa6.length > 0) {
                    result[sheetName] = roa6;
                }
            });


            let corr = [];
            let corr3 = [];


            for (const el of tempCorridos) {
                corr.push(el.leg);
            }

            const corr2 = [...new Set(corr)];

            for (const el of corr2) {
                let x = tempCorridos.filter((e) => e.leg == el);

                let initialValue = 0

                let ff = x.reduce(function(accumulator, curValue) {

                    return accumulator + curValue.xx

                }, initialValue);

                corr3.push({
                    legajo: el,
                    seccCorridos: ff
                });

            }
            corr3.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);


            for (const elem of tempChoferesSeccionamiento) {
                xx = elem.legajo;
                let x = tempChoferesSeccionamiento.filter((el) => el.legajo == elem.legajo);
                choferesSeccionamiento.push({
                    legajo: xx,
                    seccionamiento: x.length
                });
            }

            choferesSeccionamiento = choferesSeccionamiento.filter((e) => e.seccionamiento > 9);

            for (const elem of tempChoferesMalUso) {
                xx = elem.legajo;
                let x = tempChoferesMalUso.filter((el) => el.legajo == elem.legajo);
                choferesMalUso.push({
                    legajo: xx,
                    malUso: x.length
                });
            }

            choferesMalUso = choferesMalUso.filter((e) => e.malUso > 5);

            for (const elem of tempChoferesCortados) {
                xx = elem.legajo;
                let x = tempChoferesCortados.filter((el) => el.legajo == elem.legajo);
                choferesCortados.push({
                    legajo: xx,
                    cortes: x.length
                });
            }

            choferesCortados = choferesCortados.filter((e) => e.cortes > 2);

            let corr4 = corr3.filter((e) => e.seccCorridos > 5);

            corridos = corr4;



            let pene = [];
            let pene2 = [];
            let pene3 = [];
            let pene4 = [];
            let pene5 = [];
            let pene6 = [];
            let pene7 = [];
            let pene8 = [];
            let pene9 = [];

            for (i = 0; i < choferesSeccionamiento.length; i++) {
                pene.push(choferesSeccionamiento[i].legajo);
            }

            for (const el of pene) {
                if (pene2.includes(el) == false) {
                    pene2.push(el);
                }
            }
            for (const elem of pene2) {
                x = tempChoferesSeccionamiento.filter((el) => el.legajo == elem);
                pene3.push({
                    legajo: elem,
                    seccionamiento: x.length
                });
            }
            choferesSeccionamiento = pene3;


            for (i = 0; i < choferesMalUso.length; i++) {
                pene4.push(choferesMalUso[i].legajo);
            }

            for (const el of pene4) {
                if (pene5.includes(el) == false) {
                    pene5.push(el);
                }
            }
            for (const elem of pene5) {
                x = tempChoferesMalUso.filter((el) => el.legajo == elem);
                pene6.push({
                    legajo: elem,
                    malUso: x.length
                });
            }
            choferesMalUso = pene6;




            for (i = 0; i < choferesCortados.length; i++) {
                pene7.push(choferesCortados[i].legajo);
            }
            for (const el of pene7) {
                if (pene8.includes(el) == false) {
                    pene8.push(el);
                }
            }
            for (const elem of pene8) {
                x = tempChoferesCortados.filter((el) => el.legajo == elem);
                pene9.push({
                    legajo: elem,
                    cortes: x.length
                });
            }
            choferesCortados = pene9;

            for (const e of corridos) {
                x = choferesSeccionamiento.filter((n) => n.legajo == e.legajo);
                if (x.length > 0 == true) {
                    e.seccionamiento = x[0].seccionamiento;
                } else {
                    e.seccionamiento = 0;
                }
            }

            for (const e of choferesSeccionamiento) {
                x = corridos.filter((n) => n.legajo == e.legajo);
                if (x.length < 1 == true) {
                    corridos.push({
                        ...e,
                        seccCorridos: 0,
                    })
                }
            }

            for (const e of corridos) {
                x = choferesMalUso.filter((n) => n.legajo == e.legajo);
                if (x.length > 0 == true) {
                    e.malUso = x[0].malUso;
                } else {
                    e.malUso = 0;
                }
            }
            for (const e of choferesMalUso) {
                x = corridos.filter((n) => n.legajo == e.legajo);
                if (x.length < 1 == true) {
                    corridos.push({
                        ...e,
                        seccCorridos: 0,
                        seccionamiento: 0,
                    })
                }
            }

            for (const e of corridos) {
                x = choferesCortados.filter((n) => n.legajo == e.legajo);
                if (x.length > 0 == true) {
                    e.cortes = x[0].cortes;
                } else {
                    e.cortes = 0;
                }
            }
            for (const e of choferesCortados) {
                x = corridos.filter((n) => n.legajo == e.legajo);
                if (x.length < 1 == true) {
                    corridos.push({
                        ...e,
                        seccCorridos: 0,
                        seccionamiento: 0,
                        malUso: 0,
                    })
                }
            }

            corridos.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);

            for (i=0;i<corridos.length;i++){
                let x = corridos[i].seccionamiento + corridos[i].cortes + corridos[i].malUso + corridos[i].seccCorridos;
                corridos[i] = {...corridos[i], total:x};
                }

            console.log(corridos);

            for (const elem of corridos) {
                const node = document.createElement("tr");
                node.classList.add("infoP6");
                const subNode = document.createElement("td");
                const subNode1 = document.createElement("td");
                const subNode2 = document.createElement("td");
                const subNode3 = document.createElement("td");
                const subNode4 = document.createElement("td");
                const subNode5 = document.createElement("td");

                const textnode = document.createTextNode(elem.legajo);
                const textnode1 = document.createTextNode(elem.seccionamiento);
                const textnode2 = document.createTextNode(elem.malUso);
                const textnode3 = document.createTextNode(elem.cortes);
                const textnode4 = document.createTextNode(elem.seccCorridos);
                const textnode5 = document.createTextNode(elem.total);
                subNode.appendChild(textnode);
                subNode1.appendChild(textnode1);
                subNode2.appendChild(textnode2);
                subNode3.appendChild(textnode3);
                subNode4.appendChild(textnode4);
                subNode5.appendChild(textnode5);
                node.appendChild(subNode);
                node.appendChild(subNode1);
                node.appendChild(subNode2);
                node.appendChild(subNode3);
                node.appendChild(subNode4);
                node.appendChild(subNode5);
                tableP6.appendChild(node);

            }

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

function upload7() {
    var files = document.getElementById('file_upload7').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON7(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}
let roa7;

let tableP7 = document.getElementById("tableP7");
let francosMoreTR = document.getElementById("francosDeMas");
let francosLessTR = document.getElementById("francosDeMenos");

function excelFileToJSON7(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa7 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

                na1 = roa7.filter((el) => el.__EMPTY_1 > 0);

                let arrayPresent = roa7;


                for (const elem of roa7) {

                    elem.interno = elem.__EMPTY;
                    elem.legajo = elem.__EMPTY_1;
                    elem.chofer = elem.__EMPTY_2;
                    elem.xx1 = elem.__EMPTY_3;
                    elem.xx2 = elem.__EMPTY_4;
                    elem.xx3 = elem.__EMPTY_5;
                    elem.xx4 = elem.__EMPTY_6;
                    elem.xx5 = elem.__EMPTY_7;
                    elem.xx6 = elem.__EMPTY_8;
                    elem.xx7 = elem.__EMPTY_9;
                    elem.xx8 = elem.__EMPTY_10;
                    elem.xx9 = elem.__EMPTY_11;
                    elem.xx10 = elem.__EMPTY_12;
                    elem.xx11 = elem.__EMPTY_13;
                    elem.xx12 = elem.__EMPTY_14;
                    elem.xx13 = elem.__EMPTY_15;
                    elem.xx14 = elem.__EMPTY_16;
                    elem.xx15 = elem.__EMPTY_17;
                    elem.xx16 = elem.__EMPTY_18;
                    elem.xx17 = elem.__EMPTY_19;
                    elem.xx18 = elem.__EMPTY_20;
                    elem.xx19 = elem.__EMPTY_21;
                    elem.xx20 = elem.__EMPTY_22;
                    elem.xx21 = elem.__EMPTY_23;
                    elem.xx22 = elem.__EMPTY_24;
                    elem.xx23 = elem.__EMPTY_25;
                    elem.xx24 = elem.__EMPTY_26;
                    elem.xx25 = elem.__EMPTY_27;
                    elem.xx26 = elem.__EMPTY_28;
                    elem.xx27 = elem.__EMPTY_29;
                    elem.xx28 = elem.__EMPTY_30;
                    elem.xx29 = elem.__EMPTY_31;
                    elem.xx30 = elem.__EMPTY_32;


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
                    delete elem.__EMPTY_13;
                    delete elem.__EMPTY_14;
                    delete elem.__EMPTY_15;
                    delete elem.__EMPTY_16;
                    delete elem.__EMPTY_17;
                    delete elem.__EMPTY_18;
                    delete elem.__EMPTY_19;
                    delete elem.__EMPTY_20;
                    delete elem.__EMPTY_21;
                    delete elem.__EMPTY_22;
                    delete elem.__EMPTY_23;
                    delete elem.__EMPTY_24;
                    delete elem.__EMPTY_25;
                    delete elem.__EMPTY_26;
                    delete elem.__EMPTY_27;
                    delete elem.__EMPTY_28;
                    delete elem.__EMPTY_29;
                    delete elem.__EMPTY_30;
                    delete elem.__EMPTY_31;
                    delete elem.__EMPTY_32;
                    delete elem.__EMPTY_33;
                    delete elem.__EMPTY_34;
                    delete elem.interno;

                }
                na1.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);

                for (const elem of na1) {
                    let borrarCaracs = elem.chofer.split(' ');

                    let x = borrarCaracs.filter((m) => m.length > 2);

                    elem.chofer = x[0];
                }

                let francos = [];

                for (i = 0; i < na1.length; i++) {

                    let pr = [];

                    Object.entries(na1[i]).forEach(pair => {

                        let fdf = pair[0].split("");

                        if (fdf[0] === "x") {
                            fdf.shift();
                            fdf.shift();

                            let d = fdf.join("");

                            pair[0] = d;
                        }

                        let key = pair[0];
                        let value = pair[1];

                        if (value == "F*") {
                            value = "F";
                        }
                        if (value == "FV") {
                            value = "F";
                        }
                        if (value == " F") {
                            value = "F";
                        }
                        if (value == " F ") {
                            value = "F";
                        }
                        if (value == "F ") {
                            value = "F";
                        }
                        if (value == "F *") {
                            value = "F";
                        }


                        if (value != undefined && value != "6" && value != "7" && value != "8" && value != "**" && value != "*" && value != "V" && value != " " && value != "9" && value != "10") {

                            pr.push(key, value);

                            let x = pr.filter((d) => d != "F" && d != "legajo" && d != "chofer");
                            pr = x;


                        }
                        francos[i] = pr;
                    });
                }

                arrayPresent = arrayPresent.filter((e) => e.legajo > 1000);

                let pr2 = [];
                for (const e of arrayPresent){
                    e.xx31 = e.__EMPTY_33;
                
                }

                arrayPresent.sort((a,b)=>(a.legajo > b.legajo)? 1 : -1);

                for (i = 0; i < arrayPresent.length; i++) {


                    let pr3 = [];
                    Object.entries(arrayPresent[i]).forEach(pair => {
                       pr3.push(pair[1]);

                        });
                        pr2.push(pr3);
                    }

                for (const e of pr2){
                    for (i=0;i<e.length;i++){
                        if (e[i] == undefined || e[i] == "V" || e[i] == "*"){
                            e[i] = "";
                        }
                        if (e[i] == "FV" || e[i] == "F*" || e[i] == " F*"){
                            e[i] = "";
                        }
                
                    }
                }


                let dropFR = document.getElementsByClassName("dropdown-itemFR");

                let dropArFR = [0,1,2,3];
                let dropArFR2 = ["francosDrop0","francosDrop1","francosDrop2","francosDrop3"];

                for (i = 0; i < dropArFR.length; i++) {
                    dropFR[i].innerText = dropArFR[i];
                }

                for (i = 0; i < dropArFR.length; i++) {
                    let j = document.getElementById(dropArFR2[i])
                    j.addEventListener("click", () => {
                        WriteFR(j.textContent)
                    });
                }

                function WriteFR(a) {

                let frDM = document.getElementsByClassName("tableCH");

                for (const el of frDM){
                    el.style.visibility = "visible";
                }

                let francosLess = francos.filter((e) => e.length < 8);
                let francosMore = francos.filter((e) => e.length > 8 + a);
                

                for (const elem of francosMore) {
                    const node = document.createElement("tr");
                    const subNode = document.createElement("td");
                    node.classList.add("detailP7");
                    const textnode = document.createTextNode(elem[0]);
                    subNode.appendChild(textnode);
                    node.appendChild(subNode);
                    francosMoreTR.appendChild(node);

                }
                for (const elem of francosLess) {
                    const node = document.createElement("tr");
                    const subNode = document.createElement("td");
                    node.classList.add("detail2P7");
                    const textnode = document.createTextNode(elem[0]);
                    subNode.appendChild(textnode);
                    node.appendChild(subNode);
                    francosLessTR.appendChild(node);

                }


                for (const elem of francos) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP7");
                    for (const e of elem) {
                        const subNode = document.createElement("td");
                        const textnode = document.createTextNode(e);
                        subNode.appendChild(textnode);
                        node.appendChild(subNode);
                        tableP7.appendChild(node);
                    }

                }

            }

    let presentismo = document.getElementById("presentismo");
    
    presentismo.addEventListener('click', ()=>{upload7Bis()});

    function upload7Bis(){
        let borrarFR = document.getElementsByClassName("detailP7");
        let borrarFR2 = document.getElementsByClassName("infoP7");
        let borrarFR1 = document.getElementsByClassName("detail2P7");

                    if (borrarFR2.length > 0) {
                        do {
                            tableP7.removeChild(borrarFR2[0]);

                        } while (borrarFR2.length != 0);
                    }
                    if (borrarFR.length > 0) {
                        do {
                            francosMoreTR.removeChild(borrarFR[0]);

                        } while (borrarFR.length != 0);
                    }
                    if (borrarFR1.length > 0) {
                        do {
                            francosLessTR.removeChild(borrarFR1[0]);

                        } while (borrarFR1.length != 0);
                    }
                    let frDM = document.getElementsByClassName("tableCH");

                    for (const el of frDM){
                        el.style.visibility = "hidden";
                    }

                    let tablaPresentismo = document.getElementById("tablaPresentismo");


                    tablaPresentismo.style.visibility = "visible";


                    for (const elem of pr2) {
                        const node = document.createElement("tr");
                        node.classList.add("presentP7");
                        for (const e of elem) {
                            const subNode = document.createElement("td");
                            const textnode = document.createTextNode(e);
                            subNode.appendChild(textnode);
                            node.appendChild(subNode);
                            tablaPresentismo.appendChild(node);
                        }
    
                    }
        

    }


                if (roa7.length > 0) {

                    result[sheetName] = roa7;
                }
            });
        }
    } catch (e) {
        console.error(e);
    }
}



function upload8() {
    var files = document.getElementById('file_upload8').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX'|| extension == '.DBF') {
        excelFileToJSON8(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

let roa8;

let tableP8 = document.getElementById("tableP8");

function excelFileToJSON8(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa8 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);


                for (const elem of roa8) {

                    elem.FECHA = ExcelDateToJSDate2(elem.FECHA);
                    delete elem.RECA;
                }
            
                
                
                
                let fechasDispPL = [];

                for (const elem of roa8) {
                    fechasDispPL.push(elem.FECHA);
                }
                
                const fechasDispPL1 = [...new Set(fechasDispPL)];
                

                fechasDispPL1.sort((a, b) => (a > b) ? 1 : -1);

                
                
                let dropPL = document.getElementsByClassName("dropdown-itemPL");
                let dropIndexPL = 0;
                let dropArPL = [];

                for (const elem of fechasDispPL1) {

                    dropPL[dropIndexPL].innerText = elem;
                    dropArPL.push("planillasDrop" + dropIndexPL);
                    dropIndexPL++;
                }

                for (i = 0; i < fechasDispPL1.length; i++) {
                    let wPL = document.getElementById(dropArPL[i])
                    wPL.addEventListener("click", () => {
                        Write(wPL.textContent);
                    });
                }

                    
                function Write(a) {
                    let infoP8 = document.getElementsByClassName("infoP8");

                    if (infoP8.length > 0) {
                        do {
                            tableP8.removeChild(infoP8[0]);

                        } while (infoP8.length != 0);
                    }
                    
                    let arrayPL = roa8.filter((elem) => a == elem.FECHA);
                    
                    arrayPL = arrayPL.sort((a, b)=>(a.COCHE > b.COCHE)? 1:-1);

                    let cochesPL = [];

                    for (const el of arrayPL){
                        cochesPL.push(el.COCHE);
                    }

                    cochesPL = [... new Set(cochesPL)];

                    let tempPL=[];

                    for (const el of cochesPL){
                        tx = arrayPL.filter((e)=> e.COCHE == el);
                        tx  = tx.sort((a,b)=>(a.HSALE > b.HSALE)? 1:-1);
                        tempPL.push(tx);
                    }
                    
                    for (i=0;i<tempPL.length;i++){
                        if (i & 1 == true){
                        for (const el of tempPL[i]){
                                    el.COLOR = "red";
                            }
                        } 
                    }
                    finalPL = [];

                    for (i=0;i<tempPL.length;i++){
                        for (const el of tempPL[i]){
                            finalPL.push(el);        
                        }
                    } 
                    console.log(finalPL);

                



                    for (const elem of finalPL) {
                        const node = document.createElement("tr");
                        node.classList.add("infoP8");
                        const subNode = document.createElement("td");
                        const subNode1 = document.createElement("td");
                        const subNode2 = document.createElement("td");
                        const subNode3 = document.createElement("td");
                        const subNode4 = document.createElement("td");
                        const subNode5 = document.createElement("td");
                        const subNode6 = document.createElement("td");

                        const textnode = document.createTextNode(elem.COCHE);
                        const textnode1 = document.createTextNode(elem.LEGAJO);
                        const textnode2 = document.createTextNode(elem.HSALE);
                        const textnode3 = document.createTextNode(elem.HLLEGA);
                        const textnode4 = document.createTextNode(elem.HCITA);
                        const textnode5 = document.createTextNode(elem.KMTS);
                        const textnode6 = document.createTextNode(elem.COLOR);
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
                        tableP8.appendChild(node);

                    } 
                }
                
                
                

                
                if (roa8.length > 0) {
                    
                    result[sheetName] = roa8;
                }
        })
    }
        
            
    } catch (e) {
        console.error(e);
    }
}
function upload9() {
    var files = document.getElementById('file_upload9').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX'|| extension == '.DBF') {
        excelFileToJSON9(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

let roa9;


function excelFileToJSON9(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa9 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);


                for (const elem of roa9) {

                    elem.FECHA = ExcelDateToJSDate2(elem.FECHA);
                    delete elem.RECA;
                }
            
                
                
                
                let fechasDispPL = [];

                for (const elem of roa9) {
                    fechasDispPL.push(elem.FECHA);
                }
                
                const fechasDispPL1 = [...new Set(fechasDispPL)];
                

                fechasDispPL1.sort((a, b) => (a > b) ? 1 : -1);

                
                
                let dropPL = document.getElementsByClassName("dropdown-itemPL");
                let dropIndexPL = 0;
                let dropArPL = [];

                for (const elem of fechasDispPL1) {

                    dropPL[dropIndexPL].innerText = elem;
                    dropArPL.push("planillasDrop" + dropIndexPL);
                    dropIndexPL++;
                }

                for (i = 0; i < fechasDispPL1.length; i++) {
                    let wPL = document.getElementById(dropArPL[i])
                    wPL.addEventListener("click", () => {
                        Write(wPL.textContent);
                    });
                }

                    
                function Write(a) {
                    let infoP9 = document.getElementsByClassName("infoP9");

                    if (infoP9.length > 0) {
                        do {
                            tableP8.removeChild(infoP9[0]);

                        } while (infoP9.length != 0);
                    }
                    
                    let arrayPL = roa9.filter((elem) => a == elem.FECHA);
                    
                    arrayPL = arrayPL.sort((a, b)=>(a.COCHE > b.COCHE)? 1:-1);

                    let cochesPL = [];

                    for (const el of arrayPL){
                        cochesPL.push(el.COCHE);
                    }

                    cochesPL = [... new Set(cochesPL)];

                    let tempPL=[];

                    for (const el of cochesPL){
                        tx = arrayPL.filter((e)=> e.COCHE == el);
                        tx  = tx.sort((a,b)=>(a.HSALE > b.HSALE)? 1:-1);
                        tempPL.push(tx);
                    }
                    
                    
                    finalPL = [];

                    for (i=0;i<tempPL.length;i++){
                        for (const el of tempPL[i]){
                            finalPL.push(el);        
                        }
                    } 

                    finalPL = finalPL.filter((el)=> el.KMTS != 0);
                    console.log(finalPL);

                    let initialValue = 0

                    let kmSum = finalPL.reduce((acc, value)=> acc + value.KMTS, initialValue);


                    for (const elem of finalPL) {
                        const node = document.createElement("tr");
                        node.classList.add("infoP9");
                        const subNode = document.createElement("td");
                        const subNode1 = document.createElement("td");
                        const subNode2 = document.createElement("td");
                        const subNode3 = document.createElement("td");
                        const subNode4 = document.createElement("td");
                        const subNode5 = document.createElement("td");

                        const textnode = document.createTextNode(elem.COCHE);
                        const textnode1 = document.createTextNode(elem.LEGAJO);
                        const textnode2 = document.createTextNode(elem.HSALE);
                        const textnode3 = document.createTextNode(elem.HLLEGA);
                        const textnode4 = document.createTextNode(elem.HCITA);
                        const textnode5 = document.createTextNode(elem.KMTS);
                        subNode.appendChild(textnode);
                        subNode1.appendChild(textnode1);
                        subNode2.appendChild(textnode2);
                        subNode3.appendChild(textnode3);
                        subNode4.appendChild(textnode4);
                        subNode5.appendChild(textnode5);
                        node.appendChild(subNode);
                        node.appendChild(subNode1);
                        node.appendChild(subNode2);
                        node.appendChild(subNode3);
                        node.appendChild(subNode4);
                        node.appendChild(subNode5);
                        tableP8.appendChild(node);

                    } 
                    const kmPL = document.createElement("td");
                    const kmPL1 = document.createElement("td");
                    const kmPL2 = document.createElement("td");
                    const kmPL3 = document.createElement("td");
                    const kmPL4 = document.createElement("td");
                    const kmPL5 = document.createElement("td");

                    const kmPLTextNode = document.createTextNode(kmSum);
                    const kmPL1TextNode = document.createTextNode("TOTAL");
                    const kmPL2TextNode = document.createTextNode("");
                    const kmPL3TextNode = document.createTextNode("");
                    const kmPL4TextNode = document.createTextNode("");
                    const kmPL5TextNode = document.createTextNode("");
                    kmPL.classList.add("infoP9");
                    kmPL1.classList.add("infoP9");
                    kmPL2.classList.add("infoP9");
                    kmPL3.classList.add("infoP9");
                    kmPL4.classList.add("infoP9");
                    kmPL5.classList.add("infoP9");
                    kmPL.appendChild(kmPLTextNode);
                    kmPL1.appendChild(kmPL1TextNode);
                    kmPL2.appendChild(kmPL2TextNode);
                    kmPL3.appendChild(kmPL3TextNode);
                    kmPL4.appendChild(kmPL4TextNode);
                    kmPL5.appendChild(kmPL5TextNode);
                    tableP8.appendChild(kmPL1);
                    tableP8.appendChild(kmPL2);
                    tableP8.appendChild(kmPL3);
                    tableP8.appendChild(kmPL4);
                    tableP8.appendChild(kmPL5);
                    tableP8.appendChild(kmPL);

                }
                if (roa9.length > 0) {
                    
                    result[sheetName] = roa9;
                }
        })
    }
            
    } catch (e) {
        console.error(e);
    }
}

function upload10() {
    var files = document.getElementById('file_upload10').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX'|| extension == '.DBF') {
        excelFileToJSON10(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

let roa10;

let tableP10 = document.getElementById("tableP10");

function excelFileToJSON10(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa10 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);


                for (const elem of roa10) {

                    elem.FECHA = ExcelDateToJSDate2(elem.FECHA);
                    delete elem.HSALE;
                    delete elem.HLLEGA;
                    delete elem.HCITA;
                    delete elem.LEGAJO;
                    delete elem.RECA;
                }

                let fechasDispKM = [];

                for (const elem of roa10) {
                    fechasDispKM.push(elem.FECHA);
                }
                
                const fechasDispKM1 = [...new Set(fechasDispKM)];
                

                fechasDispKM1.sort((a, b) => (a > b) ? 1 : -1);

                
                
                let dropKM = document.getElementsByClassName("dropdown-itemKM");
                let dropIndexKM = 0;
                let dropArKM = [];

                for (const elem of fechasDispKM1) {

                    dropKM[dropIndexKM].innerText = elem;
                    dropArKM.push("kmSUbeDrop" + dropIndexKM);
                    dropIndexKM++;
                }

                for (i = 0; i < fechasDispKM1.length; i++) {
                    let wKM = document.getElementById(dropArKM[i])
                    wKM.addEventListener("click", () => {
                        Write(wKM.textContent);
                    });
                }

                    
                function Write(a) {
                    let infoP10 = document.getElementsByClassName("infoP10");

                    if (infoP10.length > 0) {
                        do {
                            tableP10.removeChild(infoP10[0]);

                        } while (infoP10.length != 0);
                    }
                    
                    let arrayKM = roa10.filter((elem) => a == elem.FECHA);

                    let arrayKM2=[];

                    for (const e of arrayKM){
                        arrayKM2.push(e.COCHE);
                    }


                    arrayKM2 = [...new Set (arrayKM2)];

                    let arrayKM3 = arrayKM2.sort((a,b)=>(a > b)? 1 : -1);


                    for (const e of arrayKM3){
                        x = arrayKM.filter((el)=>el.COCHE == e);
                        iv = 0;
                        xx = x.reduce((elem,value)=> elem + value.KMTS, iv);
                        arrayKM4.push({coche:e, kms:xx})

                    }


                }
                   
                if (roa10.length > 0) {
                    
                    result[sheetName] = roa10;
                }
        })
    }
            
    } catch (e) {
        console.error(e);
    }
}
function upload11() {
    var files = document.getElementById('file_upload11').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX'|| extension == '.DBF') {
        excelFileToJSON11(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

let roa11;

let tableP11 = document.getElementById("tableP11");

function excelFileToJSON11(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa11 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            
                for (const e of roa11){
                    e.coche = e.__EMPTY_2;
                    e.kms = e.__EMPTY_4;

                    delete e.__EMPTY;
                    delete e.__EMPTY_1;
                    delete e.__EMPTY_2;
                    delete e.__EMPTY_3;
                    delete e.__EMPTY_4;
                }
                roa11.shift();


                let kms2 = arrayKM4;

                for (const e of kms2){
                    x = roa11.filter((el)=> el.coche == e.coche);
                    if (x.length>0){
                    e.kmSube = x[0].kms;
                    }else{
                        e.kmSube = "0";
                        e.nov = "*";
                    } 
                }
                for (const e of roa11){
                    x = kms2.filter((el)=> el.coche == e.coche);
                    if (x.length<1 && e.kms>15){
                        kms2.push({coche:e.coche, kms:"0",kmSube:e.kms, nov:"x"});
                    }
                } 

                for (const e of kms2){
                    x = e.kms - e.kmSube;
                    e.dif = x.toFixed(2);
                }

                console.log(kms2);

                for (const elem of kms2) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP10");
                    const subNode = document.createElement("td");
                    const subNode1 = document.createElement("td");
                    const subNode2 = document.createElement("td");
                    const subNode3 = document.createElement("td");
                    const subNode4 = document.createElement("td");

                    const textnode = document.createTextNode(elem.coche);
                    const textnode1 = document.createTextNode(elem.kms);
                    const textnode2 = document.createTextNode(elem.kmSube);
                    const textnode3 = document.createTextNode(elem.dif);
                    const textnode4 = document.createTextNode(elem.nov);
                    subNode.appendChild(textnode);
                    subNode1.appendChild(textnode1);
                    subNode2.appendChild(textnode2);
                    subNode3.appendChild(textnode3);
                    subNode4.appendChild(textnode4);
                    node.appendChild(subNode);
                    node.appendChild(subNode1);
                    node.appendChild(subNode2);
                    node.appendChild(subNode3);
                    node.appendChild(subNode4);
                    tableP10.appendChild(node);

                } 
            }
        
    )};
                
            
    } catch (e) {
        console.error(e);
    }
}