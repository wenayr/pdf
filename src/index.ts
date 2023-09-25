
//import * as fs from "fs";
import fs from "fs";
//import * as bodyParser from "body-parser";
import bodyParser from "body-parser";

// при включённом esModuleInterop импорт пишется без *

//import * as cors from 'cors';
import cors from 'cors';

//import * as express from 'express';
import express from 'express';

import {tRequest} from "./interface";

import {aExcel, aFont, aImage, aResult} from "./address";

import {test} from "./test";

import {fApi} from "./API"

const api = fApi()

export const HOST = '0.0.0.0';
export const PORT: number =  4051//+process.env.PORT

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
        excel: Buffer  - файл шаблон эксель с ключами
        excelSimple: Buffer  - файл шаблон эксель без ключей
      }
     */
    app.post('/addTemplateExcel2', async (req, res) => {

        const data = req.body as {excel: Buffer, name: string, excelSimple?: Buffer}
        try {
            console.log("add excel2 " + data.excel + " " + data.name)
            await api.addTemplateExcel(data)
            res.status(200)
                .json({status: "ok"})
        } catch (e) {
            res.status(404)
                .json({status: e})
        }
    }, )
    app.post('/addTemplateExcel', async (req, res) => {
        const data = req.body as {excel: string, name: string, excelSimple?: string}
        try {
            console.log("add excel " + data.excel + " " + data.name)
            let data2 = {
                name: data.name,
                excel: await fs.promises.readFile(aExcel + data.excel),
                excelSimple: data.excelSimple ? await fs.promises.readFile(aExcel + data.excelSimple) : undefined
            }
            console.log("ok");
            await api.addTemplateExcel(data2)

            res.status(200)
                .json({status: "ok"})
        } catch (e) {
            res.status(404)
                .json({status: e})
            console.error(e);
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
    })

    server.listen(PORT, HOST, () => {
        console.log(`Server has been started on port:${PORT}`);
        //test()
    })
}

start();
//test();

