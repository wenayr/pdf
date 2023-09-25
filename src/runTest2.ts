import fs from "fs";

import { exec,spawn } from 'child_process';

import unoconv from "./unoconv_my";

////@ts-ignore
//import unoconv from 'node-unoconv';



//const unoconvRunCommand = `python "c:/OSPanel/domains/pdf/src/unoconv.py"`;//./unoconv.py`;



//const unoconv = require('node-unoconv');

// Путь к исходному файлу XLSX
const inputFilePath = './resource/excel/4pgruz_key.xlsx';

// Путь для сохранения результирующего PDF
const outputFilePath = './resource/excel/4pgruz_keyNew.pdf';

// Функция для конвертации XLSX в PDF с использованием unoconv
async function convertXlsxToPdf(inputPath :string|Buffer, outputPath :string) {
    return new Promise((resolve, reject)=>{
        unoconv.convert(inputPath, 'pdf', (err, result) => {
          if (err) {
            console.error(`Ошибка конвертации: ${err}`);
            reject(err);
          } else {
            fs.writeFileSync(outputPath, result!); //.pipe(fs.createWriteStream(outputPath)
            console.log(`Документ успешно сконвертирован в PDF: ${outputPath}`);
            resolve(result);
          }
        });
    });
    //.pipe(fs.createWriteStream(outputPath));
}

let prefix= "--";
console.time(prefix);



let res= unoconv.listen()
    .on("message",(code,signal)=>{ console.timeLog(prefix,"message ",signal); })
    .on("spawn",()=>{ console.timeLog(prefix, "spawn.  Connected=",res.connected); })


if (1) {
    let inputBuf= fs.readFileSync(inputFilePath);
    await convertXlsxToPdf(inputBuf, outputFilePath);
}
else
    await convertXlsxToPdf(inputFilePath, outputFilePath);

//console.timeLog();
console.timeLog(prefix);//,"Connected:",res.connected);

//setTimeout(()=>console.timeLog(prefix, "timeout.  Connected=",res.connected), 5000)
// unoconv.convert(inputFilePath, 'pdf', function (err :any, result :any) {
// 	// result is returned as a Buffer
// 	if (result) {
// 	    fs.writeFileSync(outputFilePath, result);
//         console.log("Written file")
//     }
//     else console.error(err);
//     console.log("finish");
// });

// unoconv.listen();
// console.log(fs.existsSync(inputFilePath));
// let buf= await unoconv.convert(inputFilePath, {format: "PDF", output: "outputFilePath"}) as Buffer;
// console.log(buf.length);
// fs.writeFileSync(outputFilePath, buf)


// Запустить LibreOffice в режиме работы по соксету  () port=2001,  8100
 //const childProcess = spawn('soffice', ["--headless", "--invisible", "--accept=socket,host=127.0.0.1,port=2001;urp;StarOffice.ComponentContext"]);
//"c:\Program Files\LibreOffice\program\soffice.com"
if (0) {
    exec('"c:\\Program Files\\LibreOffice\\program\\soffice.com" --headless --invisible --accept=socket,host=127.0.0.1,port=2001;urp;StarOffice.ComponentContext', (error, stdout, stderr) => {
      if (error) {
        console.error(`Ошибка запуска LibreOffice: ${error}`);
      } else {
        console.log('LibreOffice успешно запущен.');

        // Вызов функции для конвертации XLSX в PDF
        convertXlsxToPdf(inputFilePath, outputFilePath);
      }
    });
    setTimeout(()=>convertXlsxToPdf(inputFilePath, outputFilePath), 5000);
    console.log("finish");
}

