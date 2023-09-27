import fs from "fs";

import { exec,spawn } from 'child_process';

import unoconv from "./unoconv_my";

////@ts-ignore
//import unoconv from 'node-unoconv';

import childProcess , {ChildProcess} from 'child_process'; import {execAndWaitForSpawn, isProcessRunning,waitForProcessRunAsync} from "./processHelper";

//const unoconvRunCommand = `python "c:/OSPanel/domains/pdf/src/unoconv.py"`;//./unoconv.py`;



//isRunning('myprocess.exe', 'myprocess', 'myprocess').then((v) => console.log(v))


let prefix= "--";
console.time(prefix);

const defaultConsoleLog= console.log;
// const defaultConsoleTimeLog= console.timeLog;
//
let _timeLogRunning= false;

console.log= (...args :any[])=> _timeLogRunning ? defaultConsoleLog(...args) : (()=>{ _timeLogRunning= true;  console.timeLog(prefix, ...args);  _timeLogRunning=false; })();
//console.timeLog= (prefix, ...args :any[])=> _timeLogRunning ? defaultConsoleTimeLog

//await unoconv.convertAsync("resource/excel/test.xlsx", "pdf");
//unoconv.listen();
//console.timeLog(prefix, "finish");

async function check(id :string) {
    console.timeLog(prefix,id);
    return 0;
    let is1= await isProcessRunning('soffice.exe');
    //let is2= await isRunning('soffice.bin');
    console.timeLog(prefix,id, "exe:",is1);//, "bin:",is2);
}


// await check("1");
// await check("2");
//await waitForSpawnExec("python unoconv.py --listener");
//console.timeLog(prefix,"finish");

if (0) {
    //let proc= childProcess.spawn("python unoconv.py -fpdf resource/excel/test.xlsx", { shell: true })
    let proc= childProcess.spawn("python unoconv.py --listener", { shell: true })
        .on("message",(code,signal)=>{ check("message"); })
        .on("spawn",()=>{ check("spawn"); })
        .on("error",()=>{ check("error"); })
        .on("exit",()=>{ check("exit"); })
        .on("close",()=>{ check("close"); })

    check("first");
}

//childProcess.spawn("python unoconv.py --listener", { shell: true });

//import terminate from "terminate";

// let proc= childProcess.spawn("python unoconv.py -l", {shell: true, /*detached: true*/ });
//
// waitForProcessRunAsync("soffice.exe").then(()=>console.log("running!"));

await execAndWaitForSpawn("python unoconv.py -l");
console.log("Spawned");
import process from "process"
if (0)
await setTimeout(()=>{
    //proc.kill("SIGTERM");
    //proc.kill("SIGINT");
    //proc.kill();
    //proc.send()
    //if (0) process.kill(-proc.pid!);//, "SIGHUP");//, 'SIGINT');
    //if (1) process.exit(0);
    //proc.stdout.push(null);
    //proc.
    // if (0)
    // if (proc.pid)
    //     terminate(proc.pid);
    //proc.disconnect();
    console.log("killed");
}, 11000);


if (0) {
    //import ps from "ps-node";

    let ok= await isProcessRunning("soffice.exe");
    console.log(ok);

    //ps.lookup({command: "soffice.exe"}, (err,list)=> err ? console.log("not found") : console.log("ok ",list));
}

async function f() {

    await waitForProcessRunAsync('soffice.exe');
//console.timeLog(prefix,"running");//
}



if (0) {

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

}
