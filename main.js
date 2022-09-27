function upload() {
    var files = document.getElementById('file_upload').files;
    if(files.length==0){
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    }else{
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
    if (month.toString().length <= 1){
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

function excelFileToJSON(file){
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type : 'binary'
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

                for (const elem of newArray2){
                    roaCortos.push(elem);
                }

                let newArray3 = [];
                for (const elem of roaCortos){

                    let result = (({ Interno, Ramal, Legajo, Kms, Secciones, Chofer, Diferencia }) => ({ Interno, Ramal, Legajo, Kms, Secciones, Chofer, Diferencia }))(elem);

                    newArray3.push(result);
                }
                roa = newArray3;          

                roa.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);     
                
                for (const elem of roa){
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
    }catch(e){
        console.error(e);
    }
}


function upload2() {
    var files = document.getElementById('file_upload2').files;
    if(files.length==0){
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON2(files[0]);
    }else{
        alert("Please select a valid excel file.");
    }
}

let result2 = {};
let roa2;

let tableP2 = document.getElementById("tableP2");

function excelFileToJSON2(file){
    try {
    var reader = new FileReader();
    reader.readAsBinaryString(file);
    reader.onload = function(e) {

    var data = e.target.result;
    var workbook = XLSX.read(data, {
        type : 'binary'
    });
    workbook.SheetNames.forEach(function(sheetName) {
        roa2 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

        for (const elem of roa2){
            
            elem.FechaInicio  = ExcelDateToJSDate2(elem.FechaInicio);
            elem.FechaFin  = ExcelDateToJSDate2(elem.FechaFin);
        }

        for (const elem of roa2){
            
            elem.HoraInicio  = excelDateToJSDate3(elem.HoraInicio);
            elem.HoraFin  = excelDateToJSDate3(elem.HoraFin);
            }
            
        let fechasDisp = [];
            
        for (const elem of roa2){
            if (fechasDisp.some((n) => n == elem.FechaInicio) == false){
                fechasDisp.push(elem.FechaInicio);
            }
        }
        fechasDisp.sort((a, b) => (a > b) ? 1 : -1);

        let drop = document.getElementsByClassName("dropdown-item");
        let dropIndex = 0;
        let dropAr = [];
        
        for (const elem of fechasDisp){

            drop[dropIndex].innerText = elem;
            dropAr.push("drop"+dropIndex);
            dropIndex++;
        }


        for (i=0;i<fechasDisp.length;i++) {
            let j = document.getElementById(dropAr[i])
            j.addEventListener("click", () => {Write(j.textContent)});
        }

        let new6=[];
        function Write(a) {
            let infoP2 = document.getElementsByClassName("infoP2");

            if (infoP2.length>0){
                console.log(infoP2);
                do {
                    tableP2.removeChild(infoP2[0]);
                    console.log(infoP2);

                }while (infoP2.length!=0);
            }
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
            let new3 =new2.concat(pcoMasde48);
            let new4 =new3.concat(pcoRapMasde48);
            let new5 =new4.concat(fonMasde54);
            new6 =new5.concat(fonRapMasde54);

            new6.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);     

            for (const elem of new6){
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
        if (roa2.length > 0) {result[sheetName] = roa2;}
    });
}
}catch(e){
console.error(e);
}
}

function upload3() {
    var files = document.getElementById('file_upload3').files;
    if(files.length==0){
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON3(files[0]);
    }else{
        alert("Please select a valid excel file.");
    }
}
let roa3;

function excelFileToJSON3(file){
    try {
    var reader = new FileReader();
    reader.readAsBinaryString(file);
    reader.onload = function(e) {

        var data = e.target.result;
        var workbook = XLSX.read(data, {
            type : 'binary'
        });
        workbook.SheetNames.forEach(function(sheetName) {
            roa3 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            
            for (const elem of roa3){
                
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

            
            for (const elem of soloCambio2){
                elem.fechaInicio  = ExcelDateToJSDate2(elem.fechaInicio);
                elem.horaInicio  = excelDateToJSDate3(elem.horaInicio);
                let e = elem.horaInicio.split(':');
                elem.horaSinSec = (`${e[0]}:${e[1]}`);

                if (porInterno.some((n) => n == elem.interno) == false){
                    porInterno.push(elem.interno);
                }
            } 

            let x2 = [];
            let x3 = [];
            let x4 = [];
            let intArr = [];

            for (const elem of porInterno){
                let x = soloCambio2.filter((elem2) => elem2.interno == elem);
                x2.push(x);
            }

            for (i=0;i<porInterno.length;i++){
                x3[i] = soloCambio2.filter((elem) => elem.interno == porInterno[i]);
            }
            
            x4 = [ ...x3 ];
            for (i=0;i<x4.length;i++){
                for (ii=0;ii<x4[i].length;ii++){
                    let m = x4[i][ii].horaSinSec;
                    intArr[i] = x4[i][ii].interno;
                    x4[i][ii] = m;
                } 
            }
            for (f=0;f<3;f++){
                for (i=0;i<x4.length;i++){
                    const dup = x4[i].filter((item, index) => x4[i].indexOf(item) !== index);
                    x4[i] = dup;
                }
            }
                
            for (i=0;i<x4.length;i++){
                let dup = x4[i].filter((item, index) => x4[i].indexOf(item) == index);
                x4[i] = dup;
            }

            let finalArr2 = [];

                
                for (i=0;i<x4.length;i++){
                    for (ii=0;ii<x4[i].length;ii++){
                        let x = soloCambio2.filter((elem)=> elem.interno == intArr[i] && elem.horaSinSec == x4[i][ii]);
                        for (const elem of x){
                            finalArr2.push(elem);
                        }
                    } 
                }
               
                finalArr2.sort((a, b) => (a.legajo > b.legajo) ? 1 : -1);

                for (const elem of finalArr2){
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
    }catch(e){
        console.error(e);
    }
}

function upload4() {
    var files = document.getElementById('file_upload4').files;
    if(files.length==0){
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON4(files[0]);
    }else{
        alert("Please select a valid excel file.");
    }
}
let roa4;

let tableP4 = document.getElementById("tableP4");

function excelFileToJSON4(file){
    try {
    var reader = new FileReader();
    reader.readAsBinaryString(file);
    reader.onload = function(e) {

        var data = e.target.result;
        var workbook = XLSX.read(data, {
            type : 'binary'
        });
        workbook.SheetNames.forEach(function(sheetName) {
            roa4 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            
            
            for (const elem of roa4){
                
                elem.Fecha  = ExcelDateToJSDate2(elem.Fecha);
                elem.Hora  = excelDateToJSDate3(elem.Hora);
            }
            let temp1 = roa4.filter((elem) => elem.Cabecera != "Estación Benavidez");
            let temp2 = temp1.filter((elem) => elem.MotivoCorte != "VUELTA ANULADA");

            let fechasDisp2 = [];
            
            for (const elem of temp2){
                if (fechasDisp2.some((n) => n == elem.Fecha) == false){
                    fechasDisp2.push(elem.Fecha);
                }
            }
            fechasDisp2.sort((a, b) => (a > b) ? 1 : -1);
    
            let drop2 = document.getElementsByClassName("dropdown-item2");
            let dropIndex2 = 0;
            let dropAr2 = [];
            
            for (const elem of fechasDisp2){
    
                drop2[dropIndex2].innerText = elem;
                dropAr2.push("cortadosDrop"+dropIndex2);
                dropIndex2++;
            }
    
            for (i=0;i<fechasDisp2.length;i++) {
                let j = document.getElementById(dropAr2[i])
                j.addEventListener("click", () => {Write2(j.textContent)});
            }
            
        function Write2(a) {
            let infoP4 = document.getElementsByClassName("infoP4");

            if (infoP4.length>0){
                console.log(infoP4);
                do {
                    tableP4.removeChild(infoP4[0]);
                    console.log(infoP4);

                }while (infoP4.length!=0);
            }
            
            let soloFecha = temp2.filter((elem) => a == elem.Fecha);

            soloFecha.sort((a, b) => (a.Legajo > b.Legajo) ? 1 : -1);     

            for (const elem of soloFecha){
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
    }catch(e){
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
    
   