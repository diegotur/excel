let linkVis = document.getElementsByClassName("linkVis"); 
let elementsById = document.getElementsByClassName("dropdown-item3"); 

let arrLinks = [];

for (i=0;i<linkVis.length;i++){
    arrLinks.push([elementsById[i],linkVis[i]]);
}
for (const elem of arrLinks){
    elem[0].addEventListener("click", () => {
        for (ii=0; ii<linkVis.length; ii++)
        {linkVis[ii].style.visibility = "hidden";}
        elem[1].style.visibility = "visible";
    })
    }

function WriteTable (elem, htmlTable, node){

    for(i=0; i<Object.keys(elem).length;i++){
        const subNode = document.createElement("td");
        let textnode;
        textnode = document.createTextNode(elem[i]);
        subNode.appendChild(textnode);
        node.appendChild(subNode);
        htmlTable.appendChild(node);
    }
}

const cambioFecha = (date, quantity) => {
    let converted_date = new Date(Math.round((date - quantity) * 864e5));
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

function cambioHora(excel_date, time = false) {
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

function Limpiar (a, b){
    if (a.length > 0) {
        do {
            b.removeChild(a[0]);

        } while (a.length != 0);
    }
}

function TitleList (info, titleList, table ){
    const nodeP = document.createElement("tr");
    nodeP.classList.add(info);
            
        for (i=0; i<titleList.length; i++){
            let subNode = document.createElement("th");
            let textnode = document.createTextNode(titleList[i]);
            subNode.appendChild(textnode);
            nodeP.appendChild(subNode);
            }
        table.appendChild(nodeP);
}



function upload(source, func) {
    var files = document.getElementById(source).files;
    if (files.length == 0) {
        alert("Seleccione Archivo");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX'|| extension == '.TXT') {
        func(files[0]);
    } else {
        alert("No se puede leer el archivo");
    }
}

let roa;

let tableP = document.getElementsByClassName("tabla");

let choferesPorLegajo = [];


function Func0(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {type: 'binary'});
            workbook.SheetNames.forEach(function(sheetName) {
            roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

            if (roa.length > 0){
                for (const elem of roa){
                    choferesPorLegajo.push({legajo: elem.LEGAJO, chofer: elem.APELLIDO});
                }
                console.log(choferesPorLegajo);
            };
        })    

    }
    } catch (e) {
        console.error(e);
    }
}


function Func1(file) {
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

            if (roa.length > 0){
            
            let inter = 0;
            console.log(choferesPorLegajo);

            for (const elem of roa) {
                elem.Interno = roa[inter][Object.keys(roa[inter])[0]];
                elem.Ramal = elem.__EMPTY_2;
                elem.Legajo = elem.__EMPTY_6;
                elem.Kms = elem.__EMPTY_9;
                elem.Secciones = elem.__EMPTY_10;
                elem.Chofer = elem.__EMPTY_11;
                elem.Diferencia = elem.__EMPTY_12;
                inter++;
            }

            let roaCortos = roa.filter((elem) => elem.Secciones < 8 && elem.Kms > 12 && elem.Diferencia > 3);
            let roaLargos = roa.filter((elem) => elem.Secciones > 8 && elem.Diferencia > 7 && elem.Ramal < 316 && elem.Kms > 30);

            for (const elem of roaLargos) {
                roaCortos.push(elem);
            }

            roaCortos.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);

            let roaFinal=[];

            let infoP = document.getElementsByClassName("infoP");

            Limpiar(infoP, tableP[0]);

            let titleList = ["Interno", "Ramal", "Legajo", "Kms", "Secciones", "Chofer", "Diferencia"];
            
            TitleList("infoP", titleList, tableP[0]);

            for(const elem of roaCortos){
                
                let {Interno, Ramal, Legajo, Kms, Secciones, Chofer, Diferencia} = elem;
                
                roaFinal.push([Interno, Ramal, Legajo, Kms, Secciones, Chofer, Diferencia]);
                }
                
            for (const elem of roaFinal){
                const node = document.createElement("tr");
                node.classList.add("infoP");
                WriteTable(elem, tableP[0], node);
            }
            } 
        });
    }
    } catch (e) {
        console.error(e);
    }
}

function Func2(file) {
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


                if (roa.length > 0){

                    let inter = 0;

    
                    for (const elem of roa) {
                        elem.Legajo = roa[inter][Object.keys(roa[inter])[0]];
                        elem.Interno = elem.__EMPTY;
                        elem.HoraInicio = elem.__EMPTY_3;
                        elem.FechaInicio = elem.__EMPTY_2;
                        elem.FechaFin = elem.__EMPTY_4;
                        elem.HoraFin = elem.__EMPTY_5;
                        elem.Ramal = elem.__EMPTY_6;
                        elem.Recorrido = elem.__EMPTY_8;
                        elem.Kms = elem.__EMPTY_9;
                        inter++;
                    }

                    roa.shift();

                    
                    for (const elem of roa) {
    
                        elem.FechaInicio = cambioFecha(elem.FechaInicio, 25568);
                        elem.FechaFin = cambioFecha(elem.FechaFin, 25568);
                        elem.HoraInicio = cambioHora(elem.HoraInicio);
                        elem.HoraFin = cambioHora(elem.HoraFin);
                    }
                    
                    let fechasDisp = [];
    
                    for (const elem of roa) {
                        fechasDisp.push(elem.FechaInicio);
                    }
                    
                    fechasDisp = [...new Set(fechasDisp)];
                    
                    fechasDisp.sort((a, b) => (a > b) ? 1 : -1);
                    
                    const roaBis = roa;
                    
                    
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
                        
                        Limpiar(infoP2, tableP[1]);
                        
                        let newArray10 = roaBis.filter((elem) => a == elem.FechaInicio);
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

                    let roaFinal=[];

                    let titleList = ["Legajo", "Interno", "FechaInicio", "FechaFin", "HoraInicio", "HoraFin", "Recorrido", "Kms"];
            
                    TitleList("infoP2", titleList, tableP[1]);

                    for(const elem of new6){
                    
                        let {Legajo, Interno, FechaInicio, FechaFin, HoraInicio, HoraFin, Recorrido, Kms} = elem;
                    
                        roaFinal.push([Legajo, Interno, FechaInicio, FechaFin, HoraInicio, HoraFin, Recorrido, Kms]);
                    }
                    for (const elem of roaFinal){
                        const node = document.createElement("tr");
                        node.classList.add("infoP2");
                        WriteTable(elem, tableP[1], node);
                    }
    
                    }
                }
        });
    }
    } catch (e) {
        console.error(e);
    }
}

function Func3(file) {
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

                if (roa.length > 0){

                    for (const elem of roa) {
                        elem.interno = elem.__EMPTY;
                        elem.fechaInicio = elem.__EMPTY_1;
                        elem.horaInicio = elem.__EMPTY_2;
                        elem.recorrido = elem.__EMPTY_4;
                        elem.ramal = elem.__EMPTY_5;
                        elem.direccion = elem.__EMPTY_6;
                        elem.seccion = elem.__EMPTY_7;
                        elem.legajo = elem.__EMPTY_8;
                        elem.tipoDeMarca = elem.__EMPTY_9;
                }
                
                
                let soloCambio = roa.filter((elem) => elem.tipoDeMarca == "Cambio Seccion");
                let soloCambio2 = soloCambio.filter((elem) => elem.legajo != "0");
                let porInterno = [];
                
                
                for (const elem of soloCambio2) {
                    elem.fechaInicio = cambioFecha(elem.fechaInicio, 25568);
                    elem.horaInicio = cambioHora(elem.horaInicio, 25568);
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
                
                let roaFinal=[];

                let infoP3 = document.getElementsByClassName("infoP3");

                Limpiar(infoP3, tableP[2]);
                
                let titleList = ["Legajo", "Secc. De Corrido"];
                
                TitleList("infoP3", titleList, tableP[2]);
                
                for(const elem of finalArr6){
                    
                    let {legajo, cantCorridos} = elem;
                    
                    roaFinal.push([legajo, cantCorridos]);
                }
                for (const elem of roaFinal){
                    const node = document.createElement("tr");
                    node.classList.add("infoP3");
                    WriteTable(elem, tableP[2], node);
                }
            }
            });
        }
    } catch (e) {
        console.error(e);
    }
}

function Func4(file) {
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

                if (roa.length > 0){

                for (const elem of roa) {

                    elem.Fecha = cambioFecha(elem.Fecha, 25568);
                    elem.Hora = cambioHora(elem.Hora);
                }
                let temp1 = roa.filter((elem) => elem.Cabecera != "Estación Benavidez");
                let temp2 = temp1.filter((elem) => elem.MotivoCorte != "VUELTA ANULADA");

                let fechasDisp2 = [];

                for (const elem of temp2) {
                    if (fechasDisp2.some((n) => n == elem.Fecha) == false) {
                        fechasDisp2.push(elem.Fecha);
                    }
                }

                fechasDisp2.sort((a, b) => (a > b) ? 1 : -1);
                console.log(fechasDisp2);


                let drop2 = document.getElementsByClassName("dropdown-item2");
                let dropIndex2 = 0;
                let dropAr2 = [];

                for (const elem of fechasDisp2) {

                    drop2[dropIndex2].innerText = elem;
                    dropAr2.push("cortadosDrop" + dropIndex2);
                    dropIndex2++;
                }

                for (i = 0; i < dropAr2.length; i++) {
                    let j = document.getElementById(dropAr2[i])
                    j.addEventListener("click", () => {
                        Write2(j.textContent)
                    });
                }
                

                function Write2(a) {
                    let infoP4 = document.getElementsByClassName("infoP4");

                    Limpiar(infoP4, tableP[3]);

                    let soloFecha = temp2.filter((elem) => a == elem.Fecha);

                    soloFecha.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);

                    let soloFechaFinal = [];

                    let titleList = ["Coche", "Legajo", "Apellido", "Hora", "Recorrido", "Motivo De Corte"];
                
                    TitleList("infoP4", titleList, tableP[3]);

                    for(const elem of soloFecha){
                    
                    let {Coche, Legajo, Apellido, Hora, Recorrido, MotivoCorte} = elem;
                    
                    soloFechaFinal.push([Coche, Legajo, Apellido, Hora, Recorrido, MotivoCorte]);
                    }
                    for (const elem of soloFechaFinal){
                        const node = document.createElement("tr");
                        node.classList.add("infoP4");
                        WriteTable(elem, tableP[3], node);
                    }
                }
            }
        });
    }
    } catch (e) {
        console.error(e);
    }
}

function Func5(file) {
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

                if (roa.length>0){

                    for (const elem of roa) {
                        let x = elem.RecIda.split("");
                        x.shift();
                    elem.RecIda = x.join("");
                    if (elem.RecVuelta != undefined) {
                        let xx = elem.RecVuelta.split("");
                        xx.shift();
                        elem.RecVuelta = xx.join("");
                    }
                }
                
                let temp4 = roa.filter((elem) => elem.Cabecera != "Estación Benavidez " && elem.Tipo != "C " && elem.Tipo != "CE " && elem.Legajo != "2743");

                
                let arrayRec = [];

                function FillArray(var1, var2, var3, var4, var5, var6){
                    arrayRec.push(temp4.filter((elem) => elem.RecIda === var1 || elem.RecIda === var2 || elem.RecIda === var3 || elem.RecIda === var4 || elem.RecIda === var5 || elem.RecIda === var6));
                }
                
                function FillArrayVta(var1, var2, var3, var4, var5, var6, var7, var8, var9){
                    arrayRec.push(temp4.filter((elem) => elem.RecVuelta === var1 || elem.RecVuelta === var2 || elem.RecVuelta === var3 || elem.RecVuelta === var4 || elem.RecVuelta === var5 || elem.RecVuelta === var6 || elem.RecVuelta === var7 || elem.RecVuelta === var8 || elem.RecVuelta === var9));
                }

                FillArray("BEN LaV COM IDA", "BEN SAB COM IDA", "BEN DOM IDA", "BEN FER IDA", "BEN LaV RAP IDA", "BEN SAB RAP IDA");
                FillArrayVta("BEN LaV COM VTA","BEN SAB COM VTA","BEN DOM VTA","BEN FER VTA","BEN LaV RAP VTA","BEN SAB RAP VTA","BEN SAB R S/202 VTA","BEN LaV R DIR VTA","BEN LaV R S/202 VTA");
                FillArray("BEN LaV REL COM IDA","BEN SAB REL COM IDA","BEN DOM REL IDA","BEN FER REL IDA","BEN LaV REL RAP IDA","BEN SAB REL RAP IDA");
                FillArray("TALAR A BEN IDA 2","TALAR A BEN IDA", "null", "null", "null", "null");
                FillArrayVta( "BEN A TALAR VTA", "null", "null", "null", "null", "null", "null", "null", "null");
                FillArray("PCO LaV COM IDA" ,"PCO SAB COM IDA" ,"PCO DOM IDA" ,"PCO FER IDA" ,"PCO LaV RAP IDA" ,"PCO SAB RAP IDA");
                FillArrayVta("PCO LaV COM VTA","PCO SAB COM VTA","PCO DOM VTA","PCO FER VTA","PCO LaV RAP VTA","PCO SAB RAP VTA","PCO LaV R S/202 VTA", "null", "null");
                FillArray("PCO LaV REL COM IDA","PCO SAB REL COM IDA","PCO DOM REL IDA","PCO FER REL IDA","PCO LaV REL RAP IDA","PCO SAB REL RAP IDA");
                FillArray("TALAR A PCO IDA 2", "null", "null", "null", "null", "null");
                FillArrayVta( "PCO A TALAR VTA", "null","null","null","null","null","null","null","null");
                FillArray("FON LaV COM IDA" ,"FON SAB COM IDA" ,"FON DOM IDA" ,"FON FER IDA" ,"FON LaV RAP IDA" ,"FON SAB RAP IDA");
                FillArrayVta("FON LaV COM VTA" ,"FON SAB COM VTA" ,"FON DOM VTA" ,"FON FER VTA" ,"FON LaV RAP VTA" ,"FON SAB RAP VTA", "null", "null", "null");
                FillArray("FON LaV REL COM IDA","FON SAB REL COM IDA","FON DOM REL IDA","FON FER REL IDA","FON LaV REL RAP IDA","FON SAB REL RAP IDA");
                FillArray("TALAR A FON IDA 2","TALAR A FON IDA", "null", "null", "null", "null");
                FillArrayVta( "FON A TALAR VTA", "null", "null", "null", "null", "null", "null", "null", "null");
                FillArray("197 LaV COM IDA" ,"197 SAB COM IDA" ,"197 DOM IDA" ,"197 FER IDA" ,"197 LaV RAP IDA" ,"197 SAB RAP IDA");
                FillArrayVta( "197 LaV COM VTA" ,"197 SAB COM VTA" ,"197 DOM VTA" ,"197 FER VTA" ,"197 LaV RAP VTA" ,"197 SAB RAP VTA" ,"197 LaV R S/202 VTA", "null", "null");
                FillArray("202 LaV COM IDA" ,"202 SAB COM IDA" ,"202 DOM IDA" ,"202 FER IDA" ,"202 LaV RAP IDA" ,"202 SAB RAP IDA");
                FillArrayVta( "202 LaV COM VTA" ,"202 SAB COM VTA" ,"202 DOM VTA" ,"202 FER VTA" ,"202 LaV RAP VTA" ,"202 SAB RAP VTA", "null", "null", "null");
                FillArray("RIV LaV IDA" ,"RIV SAB IDA" ,"RIV DOM IDA" ,"RIV FER IDA", "null", "null");
                FillArrayVta( "RIV LaV VTA","RIV SAB VTA","RIV DOM VTA","RIV FER VTA", "null", "null", "null", "null", "null");
                FillArray("BCAS LaV IDA" ,"BCAS SAB IDA" ,"BCAS DOM IDA" ,"BCAS FER IDA", "null", "null");
                FillArrayVta( "BCAS LaV VTA","BCAS SAB VTA","BCAS DOM VTA","BCAS FER VTA", "null", "null", "null", "null", "null");

                
                for (const elem of arrayRec) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP5");
                    const subNode = document.createElement("td");
                    
                    const textnode = document.createTextNode(elem.length);
                    subNode.appendChild(textnode);
                    node.appendChild(subNode);
                    tableP[4].appendChild(node);
                }
            }
            });
        }
    } catch (e) {
        console.error(e);
    }
}

let choferesSeccionamiento = [];
let tempChoferesSeccionamiento = [];
let choferesMalUso = [];
let tempChoferesMalUso = [];
let choferesCortados = [];
let tempChoferesCortados = [];
let corridos = [];
let tempCorridos = [];
let tempChoferesSpeed = [];
let tempChoferesEsperas = [];
let choferesSpeed = [];
let choferesEsperas = [];

function Func6(file) {
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

                if (roa.length > 0 ){

                    let numberMalUso = roa.findIndex((elem) => elem.props == "CONTROL MAL USO DEL SUBE");
                    let numberCortados = roa.findIndex((elem) => elem.props == "SERVICIOS CORTADOS");
                    let numberSecCorridos = roa.findIndex((elem) => elem.props == "CONTROL SECCIONAMIENTO CORRIDO");
                    let numberSpeed = roa.findIndex((elem) => elem.props == "CONTROL EXCESO DE VELOCIDAD");
                    let numberEsperas = roa.findIndex((elem) => elem.props == "ESPERAS EXCESIVAS");
                    
                    
                    let arraySeccionamiento = [];
                    let arrayMalUso = [];
                    let arrayCortados = [];
                    let arrayCorridos = [];
                    let arraySpeed = [];
                    let arrayEsperas = [];
                    
                    
                    for (i = 2; i < numberMalUso; i++) {
                        arraySeccionamiento.push(roa[i]);
                    }
                    for (i = numberMalUso + 2; i < numberCortados; i++) {
                        arrayMalUso.push(roa[i]);
                    }
                    for (i = numberCortados + 2; i < numberSecCorridos; i++) {
                        arrayCortados.push(roa[i]);
                    }
                    for (i = numberSecCorridos + 2; i < numberSpeed; i++) {
                        arrayCorridos.push(roa[i]);
                    }
                    for (i = numberSpeed + 2; i < numberEsperas; i++) {
                        arraySpeed.push(roa[i]);
                    }
                    for (i = numberEsperas + 2; i < roa.length; i++) {
                        arrayEsperas.push(roa[i]);
                    }
                    for (const elem of arraySeccionamiento) {
                        tempChoferesSeccionamiento.push({legajo: elem.__EMPTY_1});
                    }
                    for (const elem of arrayMalUso) {
                        tempChoferesMalUso.push({legajo: elem.props});
                    }
                    for (const elem of arrayCortados) {
                        tempChoferesCortados.push({legajo: elem.__EMPTY});
                    }
                    for (const elem of arrayCorridos) {
                        tempCorridos.push({leg: elem.props, xx: elem.__EMPTY});
                    }
                    for (const elem of arraySpeed) {
                        tempChoferesSpeed.push({legajo: elem.__EMPTY_3});
                    }
                    for (const elem of arrayEsperas) {
                        tempChoferesEsperas.push({legajo: parseInt(elem.__EMPTY)});
                    }
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
            
            function FilterArr (var1, var2, var3){
                for (const elem of var1) {
                let x = var1.filter((el) => el.legajo == elem.legajo);
                var2.push({legajo: elem.legajo, [var3]: x.length});
                }
            }

            FilterArr(tempChoferesSeccionamiento, choferesSeccionamiento, "seccionamiento");
            FilterArr(tempChoferesMalUso, choferesMalUso, "malUso");
            FilterArr(tempChoferesSpeed, choferesSpeed, "velocidad");
            FilterArr(tempChoferesEsperas, choferesEsperas, "espera");
            FilterArr(tempChoferesCortados, choferesCortados, "cortes");

            corridos = corr3.filter((e) => e.seccCorridos > 0);
            
            function Deal(var1, var2, var3) {
                let deal1 = [];
                let deal2 = [];

                for (i = 0; i < var1.length; i++) {
                    deal1.push(var1[i].legajo);
                }
                deal1 = [...new Set(deal1)];
                
                for (const elem of deal1) {
                    x = var2.filter((el) => el.legajo == elem);
                    deal2.push({legajo: elem,[var3]: x.length});
                }
                a = deal2;
            }

            Deal(choferesSeccionamiento, tempChoferesSeccionamiento, "seccionamiento");
            Deal(choferesMalUso, tempChoferesMalUso, "malUso");
            Deal(choferesSpeed, tempChoferesSpeed, "velocidad");
            Deal(choferesEsperas, tempChoferesEsperas, "espera");
            Deal(choferesCortados, tempChoferesCortados, "cortes");

            function FillCorridos(var1, var2){
                for (const e of corridos) {
                    x = var1.filter((n) => n.legajo == e.legajo);
                    if (x.length > 0 == true) {
                        e[var2] = x[0][var2];
                    } else {
                        e[var2] = 0;
                    }
                }
            }
            
            FillCorridos(choferesSeccionamiento, "seccionamiento");

            for (const e of choferesSeccionamiento) {
                x = corridos.filter((n) => n.legajo == e.legajo);
                if (x.length < 1 == true) {
                    corridos.push({
                        ...e,
                        seccCorridos: 0,
                    })
                }
            }
            FillCorridos(choferesMalUso, "malUso");
            
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
            FillCorridos(choferesCortados, "cortes");

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
            FillCorridos(choferesSpeed, "velocidad");

            for (const e of choferesSpeed) {
                x = corridos.filter((n) => n.legajo == e.legajo);
                if (x.length < 1 == true) {
                    corridos.push({
                        ...e,
                        seccCorridos: 0,
                        seccionamiento: 0,
                        malUso: 0,
                        cortes: 0,
                    })
                }
            }
            FillCorridos(choferesEsperas, "espera");

            for (const e of choferesEsperas) {
                x = corridos.filter((n) => n.legajo == e.legajo);
                if (x.length < 1 == true) {
                    corridos.push({
                        ...e,
                        seccCorridos: 0,
                        seccionamiento: 0,
                        malUso: 0,
                        cortes: 0,
                        velocidad: 0,
                    })
                }
            }

            corridos.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);

            for (i = 0; i < corridos.length; i++) {
                let x = corridos[i].seccionamiento + corridos[i].cortes + corridos[i].malUso + corridos[i].seccCorridos + corridos[i].velocidad + corridos[i].espera;
                corridos[i] = {
                    ...corridos[i],
                    total: x
                };
            }

            let titleList = ["Legajo", "Seccionamiento", "Mal Uso Del Sube", "Servicios Cortados", "Secc. Corrido", "Exceso De Velocidad", "Esperas Excesivas", "Total Por Chofer"];

            let roaFinal = [];

            TitleList("infoP6", titleList, tableP[5]);

                for(const elem of corridos){
                    
                    let {legajo, seccionamiento, malUso, cortes, seccCorridos, velocidad, espera, total} = elem;
                    
                    roaFinal.push([legajo, seccionamiento, malUso, cortes, seccCorridos, velocidad, espera, total]);
                }
                for (const elem of roaFinal){
                    const node = document.createElement("tr");
                    node.classList.add("infoP6");
                    WriteTable(elem, tableP[5], node);
                }
        }
    } catch (e) {
        console.error(e);
    }
}

let francosMoreTR = document.getElementById("francosDeMas");
let francosLessTR = document.getElementById("francosDeMenos");

function Func7(file) {
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

                
                let roa1 = [];
                
                for (const elem of roa){
                    const {chofer, coche, __EMPTY, __EMPTY_1, ...rest} = elem;
                    
                    let x = choferesPorLegajo.filter((n)=> n.legajo == elem.legajo);
                    //console.log(elem,x);
                    if (x.length>0){
                        roa1.push({...rest, chofer: x[0].chofer});
                    } 
                }
                
                for (i=0; i<choferesPorLegajo.length;i++){
                    if (roa1.filter((n)=> n.legajo == choferesPorLegajo[i].legajo)== false){
                        roa1.push(choferesPorLegajo[i]);
                    }
                }
                
                roa1 = roa1.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);

                roa1.forEach(function(elem) {
                    if (elem[20] == "F" || elem[20] == "F" || elem[20] == "v"|| elem[20] == "f"){
                    elem[20] = "FN";
                    };
                    if (elem[21] == "F" || elem[21] == "F" || elem[21] == "v"|| elem[21] == "f"){
                        elem[21] = "FN";
                        console.log(elem[21]);
                    };
                });



                console.log(roa1);
               /*  let gs = []; 
                let gs2 = []; 

                for (const elem of roa1){
                    if (choferesPorLegajo.some((x => x.legajo == elem.legajo))!=false){
                            gs.push(elem); 
                }
            }

            
            for (i=0; i<choferesPorLegajo.length;i++){
                let x = gs.filter((n)=> n.legajo == choferesPorLegajo[i].legajo);
                let xx = {...choferesPorLegajo[i], ...x}
                choferesPorLegajo[i] = xx;
            }
            
            console.log(choferesPorLegajo);
 */
           /*  for (const elem of gs){
                for (i=0;i<elem.length;i++){

                    if (elem[i][1]=="FV" || elem[i][1]=="F " || elem[i][1]==" F" || elem[i][1]==" F"|| elem[i][1]==" F*"  || elem[i][1]=="F*" || elem[i][1]=="FV*" || elem[i][1]=="f" || elem[i][1]=="F *"){
                        elem[i][1]="F" ;
                    } else if (elem[i][1]=="V" || elem[i][1]==" V" || elem[i][1]=="VF"|| elem[i][1]=="  "|| elem[i][1]=="e"|| elem[i][1]=="V "|| elem[i][1]==" " || elem[i][1]=="*" || elem[i][1]==" *" || elem[i][1]=="V*"){
                        elem[i].shift();
                        elem[i].shift();
                    }
                }
            }
            for (const elem of gs){

                gs2
                .push(elem.filter(n => n.length > 0));
            }

            for (const elem of gs2){
            for (i=0; i<elem.length; i++){

                if (elem[i][1] == "F"){
                    elem[i].pop()
                } else if(elem[i][0]=="legajo"){
                    elem[i].shift();
                }
                if (elem[i].length>1){
                    console.log(elem[i]);
                }
            }
            }
            let frFinal = [];

            for (const elem of gs2){
                let x;
                for (i=0;i<elem.length;i++){
                    x = {...x, [i]:elem[i][0]}
                    }
                    frFinal.push(x);
                }
            for (const elem of choferesPorLegajo){
                const {chofer} = elem;
                let x;
                for (i=0;i<frFinal.length;i++){

                    x = gs2[i].findIndex((n)=> n === chofer);
                }
                console.log(x);
            }
 */
          /*   for (const elem of choferesPorLegajo){
                let at = gs2.findIndex((n)=> parseInt(n) == elem.legajo);
                frFinal.push(at);
            } */
            //console.log(frFinal);
                
            /* let arrayPresent = roa;

            
            let pr2 = [];
            
            console.log(arrayPresent);

            arrayPresent.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);

            for (i = 0; i < arrayPresent.length; i++) {


                let pr3 = [];
                Object.entries(arrayPresent[i]).forEach(pair => {
                    pr3.push(pair[1]);

                });
                pr2.push(pr3);
            }

            for (const e of pr2) {
                for (i = 0; i < e.length; i++) {
                    if (e[i] == undefined || e[i] == "V" || e[i] == "V*" || e[i] == "e" || e[i] == "*" || e[i] == " *" || e[i] == "* ") {
                        e[i] = "";
                    }
                    if (e[i] == "FV" || e[i] == "F*" || e[i] == "VF" || e[i] == "F *" || e[i] == " F*" || e[i] == "FV*" || e[i] == "F* ") {
                        e[i] = "F";
                    }

                }
            }


                       

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


                    let francosMasL = [];

                    for (const e of francos) {
                        francosMasL.push(e);
                    }

                    francosMasL.push([1625, "GIMENEZ"]);
                    francosMasL.push([1677, "SOTO"]);
                    francosMasL.push([1986, "RODRIGUEZ"]);
                    francosMasL.push([2079, "ROBERT"]);
                    francosMasL.push([2443, "DIAZ"]);
                    francosMasL.push([2448, "BENINATTI"]);
                    francosMasL.push([2599, "ALIZ"]);
                    francosMasL.push([2678, "PALACIO"]);
                    francosMasL.push([2680, "MUNTADA"]);
                    francosMasL.push([2682, "FAZZARI"]);
                    francosMasL.push([2752, "OLIVIERI"]);
                    francosMasL.push([2975, "PERESON"]);

                    for (const el of francosMasL) {
                        if (el.length < amountFR) {
                            do {
                                el.push("");
                            } while (el.length != amountFR);
                        }
                    }

                    console.log(francosMasL);

                    francosMasL.sort((a, b) => (a > b) ? 1 : -1);



                    if (francos.length > 25) {

                        for (const elem of francosMasL) {
                            const node = document.createElement("tr");
                            node.classList.add("infoP7");
                            for (const e of elem) {
                                const subNode = document.createElement("td");
                                const textnode = document.createTextNode(e);
                                subNode.appendChild(textnode);
                                node.appendChild(subNode);
                                tableP[6].appendChild(node);
                            }

                        }
                    }

                }

                let presentismo = document.getElementById("presentismo");

                presentismo.addEventListener('click', () => {
                    upload7Bis()
                }); */

                function upload7Bis() {
                    let borrarFR = document.getElementsByClassName("detailP7");
                    let borrarFR2 = document.getElementsByClassName("infoP7");
                    let borrarFR1 = document.getElementsByClassName("detail2P7");

                    if (borrarFR2.length > 0) {
                        do {
                            tableP[6].removeChild(borrarFR2[0]);

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

                    for (const el of frDM) {
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
            });
        }
    } catch (e) {
        console.error(e);
    }
}
let roa8;

let tableP8 = document.getElementById("tableP8");

function Func8(file) {
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

            if (roa8.length > 0){

                for (const elem of roa8) {

                    elem.FECHA = cambioFecha(elem.FECHA, 25568);
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
                    console.log(dropArPL[i]);
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

                arrayPL = arrayPL.sort((a, b) => (a.COCHE > b.COCHE) ? 1 : -1);

                let cochesPL = [];

                for (const el of arrayPL) {
                cochesPL.push(el.COCHE);
                }

                cochesPL = [...new Set(cochesPL)];

                let tempPL = [];

                for (const el of cochesPL) {
                    tx = arrayPL.filter((e) => e.COCHE == el);
                    tx = tx.sort((a, b) => (a.HSALE > b.HSALE) ? 1 : -1);
                    tempPL.push(tx);
                }

                for (i = 0; i < tempPL.length; i++) {
                    if (i & 1 == true) {
                        for (const el of tempPL[i]) {
                            el.COLOR = "red";
                        }
                    }
                }
                finalPL = [];

                for (i = 0; i < tempPL.length; i++) {
                    for (const el of tempPL[i]) {
                        finalPL.push(el);
                    }
                }
                let roaFinal= [];

                for(const elem of finalPL){
                    
                    let {COCHE, LEGAJO, HSALE, HLLEGA, HCITA, KMTS, COLOR} = elem;
                    
                    roaFinal.push([COCHE, LEGAJO, HSALE, HLLEGA, HCITA, KMTS, COLOR]);
                }


                for (const elem of roaFinal) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP8");
                    WriteTable(elem, tableP8, node);

                }
            }
        }
    })
}
    } catch (e) {
        console.error(e);
    }
}

let roa9;


function Func9(file) {
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

                    elem.FECHA = cambioFecha(elem.FECHA, 25568);
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

                    arrayPL = arrayPL.sort((a, b) => (a.COCHE > b.COCHE) ? 1 : -1);

                    let cochesPL = [];

                    for (const el of arrayPL) {
                        cochesPL.push(el.COCHE);
                    }

                    cochesPL = [...new Set(cochesPL)];

                    let tempPL = [];

                    for (const el of cochesPL) {
                        tx = arrayPL.filter((e) => e.COCHE == el);
                        tx = tx.sort((a, b) => (a.HSALE > b.HSALE) ? 1 : -1);
                        tempPL.push(tx);
                    }


                    finalPL = [];

                    for (i = 0; i < tempPL.length; i++) {
                        for (const el of tempPL[i]) {
                            finalPL.push(el);
                        }
                    }

                    finalPL = finalPL.filter((el) => el.KMTS != 0);
                    console.log(finalPL);

                    let initialValue = 0

                    let kmSum = finalPL.reduce((acc, value) => acc + value.KMTS, initialValue);


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

let roa10;

let tableP10 = document.getElementById("tableP10");

let arrayKM4=[];

function Func10(file) {
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

                    elem.FECHA = cambioFecha(elem.FECHA, 25568);
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

                    let arrayKM2 = [];

                    for (const e of arrayKM) {
                        arrayKM2.push(e.COCHE);
                    }


                    arrayKM2 = [...new Set(arrayKM2)];

                    let arrayKM3 = arrayKM2.sort((a, b) => (a > b) ? 1 : -1);

                    

                    for (const e of arrayKM3) {
                        x = arrayKM.filter((el) => el.COCHE == e);
                        iv = 0;
                        xx = x.reduce((elem, value) => elem + value.KMTS, iv);
                        arrayKM4.push({
                            coche: e,
                            kms: xx
                        })

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

let roa11;

let tableP11 = document.getElementById("tableP11");

function Func11(file) {
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

                    if (roa11.length > 0){

                        for (const e of roa11) {
                            e.coche = e.__EMPTY_1;
                            e.kms = e.__EMPTY_3;
                            
                    }
                     roa11.shift(); 
                    
                    
                    let kms2 = arrayKM4;
                    
                    
                    
                    for (const e of kms2) {
                        x = roa11.filter((el) => el.coche == e.coche);
                        if (x.length > 0) {
                            e.kmSube = x[0].kms;
                            e.nov = "";
                        } else {
                            e.kmSube = "0";
                            e.nov = "*";
                        }
                    }
                    for (const e of roa11) {
                        x = kms2.filter((el) => el.coche == e.coche);
                        if (x.length < 1 && e.kms > 15) {
                            kms2.push({
                                coche: e.coche,
                                kms: "0",
                                kmSube: e.kms,
                                nov: "x"
                            });
                        }
                    }
                    
                    for (const e of kms2) {
                        x = e.kms - e.kmSube;
                        if (x < 10 && x >-10){
                            delete e;
                        } else{
                            e.dif = x.toFixed(2);
                        }
                    }

                    let kms3 = [];

                    for (const e of kms2) {
                        if (e.dif != undefined){
                            kms3.push(e);
                        }
                    }

                kms3 = kms3.sort((a, b) => (parseInt(a.dif) > parseInt(b.dif)) ? -1 : 1);

                console.log(kms3);
                    

                let titleList = ["Coche", "KM Tráfico", "KM Sube", "Diferencia", "Novedad"];
               
                TitleList("infoP12", titleList, tableP10);
                
                for (const elem of kms3) {
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
                    if (roa11.length > 0) {

                        result[sheetName] = roa11;
                    }
                }
                }

            )
        };


    } catch (e) {
        console.error(e);
    }
}

let roa12;

let tableP12 = document.getElementById("tableP12");

function Func12(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function(sheetName) {
                roa12 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

                if (roa12.length>0){

                console.log(roa12);

                    for (const elem of roa12) {
                        
                        
                        elem.coche = elem.ESPERAS;
                        elem.legajo = elem.__EMPTY;
                        elem.chofer = elem.__EMPTY_1;
                        elem.recorrido = elem.__EMPTY_9;
                        elem.espera = cambioHora(elem.__EMPTY_10);
                        elem.horaEntrada = cambioHora(elem.__EMPTY_3);
                        elem.fechaEntrada = cambioFecha(elem.__EMPTY_3, 25569);
                        elem.horaSalida = cambioHora(elem.__EMPTY_7);
                        elem.fechaSalida = cambioFecha(elem.__EMPTY_7, 25569);
                        
                        delete elem.ESPERAS;
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
                    }

                    let serv = [];

                    for (const elem of roa12){
                        let x = elem.espera.split(":");
                        if (x[0]<1 && x[1]<10){
                            delete elem.espera;
                        }
                    }
                    for (const elem of roa12){
                        
                        if (elem.espera!=undefined){
                            serv.push(elem);
                        }
                    }

                    let servBcas = serv.filter((elem) => elem.recorrido == "BARRANCAS");
                    let servRiv = serv.filter((elem) => elem.recorrido == "EST. RIVADAVIA");
                    let servResto = serv.filter((elem) => elem.recorrido != "EST. RIVADAVIA" && elem.recorrido !="BARRANCAS");
                   

                    for (const elem of servResto){
                        let x = elem.espera.split(":");
                        if (x[0]<1 && x[1]<26){
                            delete elem.espera;
                        }
                    }

                    let servTotal = [];
                    for (const elem of servResto){
                        
                        if (elem.espera!=undefined){
                            servTotal.push(elem);
                        }
                    }
                    for (const elem of servBcas){
                        servTotal.push(elem);
                    }
                    for (const elem of servRiv){
                        servTotal.push(elem);
                    }
                    console.log(servTotal);

                    

                    let titleList = ["Coche", "Legajo", "Chofer", "Fecha", "Entrada", "Salida", "Ramal", "Espera"];

                    const nodeP = document.createElement("tr");
                    nodeP.classList.add("infoP12");

                    for (i=0; i<titleList.length; i++){
                        let subNode = document.createElement("th");
                        let textnode = document.createTextNode(titleList[i]);
                        subNode.appendChild(textnode);
                        nodeP.appendChild(subNode);
                    }
                    
                    tableP12.appendChild(nodeP);


                for (const elem of servTotal) {
                    const node = document.createElement("tr");
                    node.classList.add("infoP12");
                    const subNode = document.createElement("td");
                    const subNode1 = document.createElement("td");
                    const subNode2 = document.createElement("td");
                    const subNode3 = document.createElement("td");
                    const subNode4 = document.createElement("td");
                    const subNode6 = document.createElement("td");
                    const subNode7 = document.createElement("td");
                    const subNode8 = document.createElement("td");
                    
                    const textnode = document.createTextNode(elem.coche);
                    const textnode1 = document.createTextNode(elem.legajo);
                    const textnode2 = document.createTextNode(elem.chofer);
                    const textnode3 = document.createTextNode(elem.fechaEntrada);
                    const textnode4 = document.createTextNode(elem.horaEntrada);
                    const textnode6 = document.createTextNode(elem.horaSalida);
                    const textnode7 = document.createTextNode(elem.recorrido);
                    const textnode8 = document.createTextNode(elem.espera);
                    subNode.appendChild(textnode);
                    subNode1.appendChild(textnode1);
                    subNode2.appendChild(textnode2);
                    subNode3.appendChild(textnode3);
                    subNode4.appendChild(textnode4);
                    subNode6.appendChild(textnode6);
                    subNode7.appendChild(textnode7);
                    subNode8.appendChild(textnode8);
                    node.appendChild(subNode);
                    node.appendChild(subNode1);
                    node.appendChild(subNode2);
                    node.appendChild(subNode3);
                    node.appendChild(subNode4);
                    node.appendChild(subNode6);
                    node.appendChild(subNode7);
                    node.appendChild(subNode8);
                    tableP12.appendChild(node);
            } 
            
        }

                if (roa12.length > 0) {

                    result[sheetName] = roa12;
                }
            });
        }
    } catch (e) {
        console.error(e);
    }
}