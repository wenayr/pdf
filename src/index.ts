
//import * as fs from "fs";
import fs from "fs";
//import * as bodyParser from "body-parser";
import bodyParser from "body-parser";

// при включённом esModuleInterop импорт пишется без *

//import * as cors from 'cors';
import cors from 'cors';

//import * as express from 'express';
import express from 'express';

import {tRequest,tRequestAddTemplate,tRequestAddTemplateByBuffer} from "./interface";

import {aExcel, aFont, aImage, aResult} from "./address";

import {test} from "./test";

import {fApi} from "./API"

const api = fApi();


function getInputConfig() {
    //console.log("Args: ",process.argv);
    let config : { port?: number } = { }
    for(let arg of process.argv.slice(2)) {
        let pair= arg.split("=");
        let key= pair[0].toLowerCase();
        if (key=="port" || key=="-port" || key=="--port") {
            let port= parseInt(pair[1]);
            //console.log(port);
            if (! isNaN(port)) {
                console.log("got argument: port=",port);
                config.port= port;
            }
        }
    }
    return config;
}

export const HOST = '0.0.0.0';
export const PORT: number = getInputConfig().port ?? 4051; ////+process.env.PORT

import http from 'http';




export function start() {

    const app = express();
    //const server = require('http').createServer(app);
    const server= http.createServer(app);
    app.use(cors({credentials: true, origin: true}))
    // app.use(bodyParser.urlencoded({extended: true}))
    // app.use(bodyParser.json())

    app.use(bodyParser.json({limit: "50mb"}));
    app.use(bodyParser.urlencoded({limit: "50mb", extended: true, parameterLimit:50000}));

    app.use(cors({credentials: true, origin: true}))

    /*
      точка addTemplateExcel
      принимает объект:
      {
        name: string  - название  шаблона
        excel: string|Buffer  - файл шаблон эксель с ключами (либо буфер)
        excelSimple: string|Buffer|undefined  - файл шаблон эксель без ключей (либо буфер)
      }
     */
    app.post('/addTemplateExcel', async (req, res) => {
        const data = req.body as tRequestAddTemplate;
        let str = {
            name: data.name,
            excel: data.excel && typeof(data.excel)=="object" ? "buffer" : data.excel,
            excelSimple: data.excelSimple && typeof(data.excelSimple)=="object" ? "buffer" : data.excelSimple
        } satisfies { [key in keyof tRequestAddTemplate] : string };
        console.log("addTemplateExcel",str);

         async function toFileBuffer(val :unknown, key :keyof tRequestAddTemplate) {
            if (typeof(val)=="string") return await fs.promises.readFile(aExcel + val);
            if (typeof(val)=="object" && val) return val as Buffer;
            throw "wrong data.key "+key;
        }
        try {
            let data2 : tRequestAddTemplateByBuffer = {
                name: data.name,
                excel: await toFileBuffer(data.excel, "excel"),
                excelSimple: data.excelSimple ? await toFileBuffer(data.excelSimple, "excelSimple") : undefined
            }
            await api.addTemplateExcel(data2)
            res.status(200)
                .json({status: "ok"})
        } catch (e) {
            res.status(404)
                .json({status: e})
        }
    }, )

    /*
     точка dataToPDF
     принимает объект типа  {[p: string]:  {[p: string]: string | Buffer}[]}
     Пример:
     { имя шаблона 1 :
        [
            {
             имя ключа 1 : значение или текст
                или обьект с полями для картинки  {name: string,  x?: number, y?: number, wight?: number, height?: number}
                или обьект с полями текста
                    text: string,
                    x?: number,
                    y?: number,
                    width?: number,
                    height?: number,
                    pageIndex?: number
                    size?: number,
                    font?: "origin" | "bold" | "boldItalic" | "italic",
                    maxWidth?: number,
             имя ключа 2 : значение или текст или ...,
             },
            {
             имя ключа 1 : значение или текст или буфер картинки,
             имя ключа 2 : значение или текст или ...,
             }
         ],
      имя шаблона 2 :
        [
            {
             имя ключа 1 : значение или текст или ...,
             имя ключа 2 : значение или текст или ...,
             имя ключа 3 : значение или текст или ...,
             }
         ]
     }

     на выходе пришлет
        {
        result : буфер ПДФ файла,
        status: "ok"
        }
        или ошибку если что-то не так
     */
    app.post('/dataToPDF2', async (req, res) => {
        const data = req.body as tRequest; // {[p: string]:  {[p: string]: string | Buffer}[]}
        try {
            const results = await api.dataToPDFMulti(data)
            res.status(200)
                .json({status: "ok", result: results[0]})
        } catch (e) {
            res.status(404)
                .json({status: e});
            console.error(e);
        }
    }, )
    app.post('/dataToPDF', async (req, res) => {
        const data = req.body as tRequest // {[p: string]:  {[p: string]: string | Buffer}[]}
        try {
            console.log("dataToPDF")
            console.time("11")
            const results = await api.dataToPDFMulti(data)
                .catch((e)=>{
                    throw " dataToPDF " + e
                });
            console.timeEnd("11");
            for(let [i,result] of results.entries()) {
                const name = String(Date.now()) + ".pdf"
                const arrBytes = await result.save()
                    .catch((e)=>{
                        throw " result.save: "+JSON.stringify(e)
                    });
                await fs.promises.writeFile(aResult + name, arrBytes)
                    .catch((e)=>{
                        throw "writeFile: "+JSON.stringify(e)
                    });
                console.log(`status: "ok", fileName: ${name}, address: ${aResult + name}`);
                if (i==0)
                    res.status(200).json({status: "ok", fileName: name})
            }
            // res.status(200)
            //     .json({status: "ok", nameFile: name})
        } catch (e) {
            res.status(404)
                .json({status: e + " " }) // ООО «СИСТЕМА»\nОГРН(ОГРНИП) 5177746289804\nИНН 7734409110 NaN26.206NaN560.403NaN6.49NaN[object Object]NaN7.4635NaN164
            console.error(e);
        }
    }, )
    // вернет все ранее загруженные шаблоны - не требует параметров,
    // вернет обьект - ключом которого является название шаблона, значением - файл эксель который имееется
    app.post('/getExcel', (req, res) => {
        res.status(200)
            .json(api.getExcel())
    }, )

    app.post('/status', async (req, res) => {
        try {
            const result =  "ok"
            res.status(200)
                .json({status: "ok", result})
        } catch (e) {
            res.status(404)
                .json({status: e});
            console.error(e);
        }
    }, )

    app.get('/s', async (req, res) => {
        try {
            const result =  "ok"
            res.status(200)
                .json({status: "ok", result})
        } catch (e) {
            res.status(404)
                .json({status: e});
            console.error(e);
        }
    }, )

    app.get("/stop", async(req, res)=> {
        res.status(200)
            .json({status: "ok", result: "server closed"});
        server.close();
        api.disconnect();
    })

    server.listen(PORT, HOST, () => {
        console.log(`Server has been started on port:${PORT}`);
        //test()
    })
}

start();
//test();

