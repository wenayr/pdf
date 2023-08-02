import * as ExcelJS from "exceljs";
import {Workbook} from "exceljs";
import * as fs from "fs";
import * as bodyParser from "body-parser";
import * as cors from 'cors';

import * as express from 'express';
import {test} from "./iTest";
import {createPDF} from "./iCreatPDF";
import {tCellInfo, tMapExcel, tMapPDF, tObjImage, tPFD, tRequest} from "./inteface";
import {TF} from "wenay-common";
import {PDFImage} from "pdf-lib";
import {aExcel, aFont, aImage, aResult} from "./addres";
// @ts-ignore
// import type {Buffer} from "exceljs/index";

const pdf = require('pdf-parse');

const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);

async function render_page(pageData:any) {
    let render_options = {
        normalizeWhitespace: false,
        disableCombineTextItems: false
    }
    const textContent = await pageData.getTextContent(render_options)

    const obj: { [key: string]: tPFD } = {}
    for (let item of textContent.items) {
        // надо удалить все переносы строк если такие есть
        const str2 = item.str.replace(/\n/g, '')

        if (str2.includes('key_')) {
            obj[item.str] = {
                transform: item.transform,
                pageIndex: pageData.pageIndex,
                pageView: pageData.pageInfo.view,
                fontName: item.fontName,
                width: item.width,
                height: item.height
            } as tPFD
        }
    }
    return obj
}


async function ExcelToMapCell(file: Buffer) {
    // чтение стиля из эксель
    const workbook:Workbook = new ExcelJS.Workbook();
    const w = await workbook.xlsx.load(file);
    const firstSheet = w.getWorksheet(1)

    const cellInfo: tCellInfo = {}
    const tt = TF.M15
    let a = false
    firstSheet.eachRow((row,rowNumber)=>{
        row.eachCell((cell:any, colNumber)=> {
            if(cell.value?.includes('key_')){
                const style = cell.master ? cell.master.style : cell.style
                //  {
                //     a = true;
                //     let rangeX= [cell.master?._column._number ?? cell._column._number,cell._column._number];
                //     let rangeY= [cell.master?._row._number ?? cell._row._number,cell._row._number];
                //     console.log(rangeX);
                //     console.log(rangeY);
                // }
                cellInfo[cell.value]={
                    rangeX: [cell.master?._column._number ?? cell._column._number,cell._column._number],
                    rangeY: [cell.master?._row._number ?? cell._row._number,cell._row._number],
                    font: {
                        name: style.font.name,
                        style: !style.font.bold && !style.font.italic ? 'origin' : style.font.bold && style.font.italic ? 'boldItalic' : style.font.bold ? 'bold' : 'italic'
                    },
                    alignment: {
                        vertical: style.alignment?.vertical ?? 'top',
                        horizontal: style.alignment?.horizontal ?? 'left',
                    }
                }
            }
        });
    })

    const row1=firstSheet.getRow(1)
    for(const [key,value] of Object.entries(cellInfo)){
        let w = 0
        let h = 0
        for(let i=value.rangeX[0];i<=value.rangeX[1];i++){
            const x:any = row1.getCell(i)
            w += Math.round(x._column.width+5); // 6 ширина символа шрифта (проверить надо точную !!)  , 5 - padding (тоже примерно) // 6*
            // console.log(x._column)
        }
        console.log(w)
        for(let i=value.rangeY[0];i<=value.rangeY[1];i++){
            const x:any = firstSheet.getRow(i)
            h += x.height;
        }
        cellInfo[key].width=w
        cellInfo[key].height=h
    }
    return cellInfo
}


async function convertExcToPDF(excel: Buffer) {
    return await libre.convertAsync(excel, '.pdf', undefined) as Buffer
}

let _fonts: {origin: Buffer, italic: Buffer, bold: Buffer, boldItalic: Buffer} = undefined

async function getFonts() {
    if (!_fonts) _fonts = {
        origin: await (fs.promises.readFile(aFont + 'arial.ttf')),
        italic: await (fs.promises.readFile(aFont + 'ariali.ttf')),
        bold: await (fs.promises.readFile(aFont + 'arialbd.ttf')),
        boldItalic: await (fs.promises.readFile(aFont + 'arialbi.ttf')),
    }
    return _fonts
}

function fApi() {
    const mapExcelStyle: tMapExcel = {}
    const mapPDFKey: tMapPDF = {}
    const mapPDFKeyMap: {[k: string]: {[p: string]: tPFD}}  = {}
    const mapPDF: tMapPDF = {}
    /// addTemplateExcel    req: {excel: Buffer, name: string, excelSimple: Buffer}  res: {status:"ok"}
    const addTemplateExcel = async ({excelSimple, excel, name}: {excel: Buffer, name: string, excelSimple: Buffer}) => {
        const xcl = await ExcelToMapCell(excel)
        mapExcelStyle[name] = xcl;
        /// надо конвертировать excel в пдф
        const resKey = await convertExcToPDF(excel)
        const res = await convertExcToPDF(excelSimple)
        const pdfMapKey = await PDFToMapKey(resKey);

        mapPDFKey[name] = resKey;
        mapPDF[name] = res;
        mapPDFKeyMap[name] = pdfMapKey;


        // для проверки сохранит промежуточные PDF
        // await fs.promises.writeFile("test"+name+".pdf", res)
        // await fs.promises.writeFile("testKey"+name+".pdf", resKey)

        return true;
    }

    const PDFToMapKey = (pdfBuffer: Buffer) => {
        return new Promise<{[p: string]: tPFD}>((resolve, reject)=>{
            pdf(pdfBuffer, {
                pagerender: async (data: any )=>{
                    resolve(await render_page(data))
                }})
        })
    }

    /// dataToPDF    req:  {[p: string]: string | Buffer} res: {status:"ok"}
    const dataToPDF = async (data: tRequest) => {
        const arrPDF = []
        const fonts = await getFonts()
            .catch((e)=>{
                console.log("error")
                throw " font " + e
            })

        const objImageB: tObjImage = {}
        const ff = async (name: string) => objImageB[name] ??= await fs.promises.readFile(aImage + name)
            .catch((e)=>{
                console.log(" error ")
                throw " cannot read " + name + e
            })



        for (const [name, dataKey] of Object.entries(data)) {
            // файл с метками
            const pdfKey = mapPDFKey[ name ]// открытие буфера пдф по имени
            const pdf = mapPDF[ name ]// открытие буфера пдф по имени3
            const excelKey = mapExcelStyle[ name ]
            const pdfMapKey = mapPDFKeyMap[ name ]
            if (!pdfMapKey) throw "не создан pdfMapKey с ключами для шаблона " + name
            if (!pdfKey) throw "не создан pdf с ключами для шаблона " + name
            if (!pdf) throw "не создан чистый pdf для шаблона " + name
            if (!excelKey) throw "не создана карта ключей и стилей по excel для шаблона " + name

            const objImage: tObjImage = {}
            {
                const arr2: Promise<any>[] = []
                for (const data of dataKey)
                    for (const value of Object.values(data))
                        if (typeof value == "object" && value.name) arr2.push(ff(value.name).then(e=>objImage[value.name] = e))

                await Promise.all(arr2)
                    .catch((e)=>{
                        console.log("error 555")
                        throw " error promise all reading  " + e
                    })
            }

            const result = await createPDF(pdf, pdfMapKey, dataKey, excelKey, fonts, objImage, name)
            return result;
        }
        // return arrPDF
        /*
const pdfBytes = await pdfDoc.save();
fs.writeFileSync('example.pdf', pdfBytes);*/
    }
    return {
        // добавить шаблон
        addTemplateExcel,
        // конвертировать данные в пдф
        dataToPDF,
        // получить все текущие шаблоны
        getExcel: ()=> mapExcelStyle
    }
}

const api = fApi()

export const HOST = '0.0.0.0';
export const PORT: number =  4051//+process.env.PORT



function start() {

    const app = express();
    const server = require('http').createServer(app)
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

        const data = req.body as {excel: Buffer, name: string, excelSimple: Buffer}
        try {
            await api.addTemplateExcel(data)
            res.status(200)
                .json({status: "ok"})
        } catch (e) {
            res.status(404)
                .json({status: e})
        }
    }, )
    app.post('/addTemplateExcel', async (req, res) => {
        const data = req.body as {excel: string, name: string, excelSimple: string}
        try {
            console.log("add excel " + data.excel + " " + data.name)
            let data2 = {name: data.name, excel: await fs.promises.readFile(aExcel + data.excel), excelSimple: await fs.promises.readFile(aExcel + data.excelSimple)}
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
        const data = req.body as tRequest // {[p: string]:  {[p: string]: string | Buffer}[]}
        try {
            const result = await api.dataToPDF(data)
            res.status(200)
                .json({status: "ok", result})
        } catch (e) {
            res.status(404)
                .json({status: e})
        }
    }, )
    app.post('/dataToPDF', async (req, res) => {
        const data = req.body as tRequest // {[p: string]:  {[p: string]: string | Buffer}[]}
        try {
            console.log("dataToPDF")
            console.time("11")
            const result = await api.dataToPDF(data)
                .catch((e)=>{
                    throw " dataToPDFe " + e
                })
            console.timeEnd("11")
            const name = String(Date.now()) + ".pdf"
            const arrBaits = await result.save()
                .catch((e)=>{
                    throw "result.save"
                })
            await fs.promises.writeFile(aResult + name, arrBaits)
                .catch((e)=>{
                    throw "writeFile"
                })
            res.status(200)
                .json({status: "ok", nameFile: name})
        } catch (e) {
            res.status(404)
                .json({status: e})
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
                .json({status: e})
        }
    }, )

    app.get('/s', async (req, res) => {
        try {
            const result =  "ok"
            res.status(200)
                .json({status: "ok", result})
        } catch (e) {
            res.status(404)
                .json({status: e})
        }
    }, )

    server.listen(PORT, HOST, () => {
        console.log(`Server has been started on port:${PORT}`);
            test()
    })
}
start()


