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

//let result3 = {};
let roa3;

//let tableP3 = document.getElementById("tableP3");

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
            
           // let soloCambio3 = [];
            
            let x2 = [];
            let x3 = [];
            for (const elem of porInterno){
                let x = soloCambio2.filter((elem2) => elem2.interno == elem);

                console.log(x);

                //NO FUNCIONA NADA

                /* 
                for (const elem3 of x){
                    //let p = x.findIndex((n)=> n == elem3);
                    let h = x;

                    h.shift();
                    
                    if (h.some((n)=>n.horaSinSec == elem3.horaSinSec) == false){
                        delete elem3;
                    } */
                

                   // for (i=0;i<x;i++){
                     //   console.log(elem3, x[i]);
                        //if (x[i] !== elem3 == true){
                            /* if (x[i].horaSinSec == elem3.horaSinSec == true){
                                x2.push(x[i]);
                            }
                            if (x2.length > 2 == true){
                                x3.push(elem3);
                                x3.push(x2);
                            } */

                       // }
                    }
                    
                   console.log(x);
                
            }
                
        
            //console.log(soloCambio3);
            
            //let malSecc = soloCambio.filter((elem) => elem.horaInicio  == "Cambio Seccion");
            
            
                if (roa3.length > 0) {

                    result[sheetName] = roa3;
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
    
       /* console.log(elem3.horaSinSec);
                    for (i=0;i<x;i++){
                        if (elem3.horaSinSec == x[i].horaSinSec == true){
                            elemRep.push(elem3);
                            elemRep.push(x[i]);
                        }
                        if (elemRep.length>3 ==true){
                            for (const el of elemRep){
                                soloCambio3.push(el); */