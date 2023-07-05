import * as ExcelJS from "exceljs";
import {Workbook} from "exceljs";
import * as fs from "fs";
import * as cors from 'cors';
import * as path from "path";
import * as fontkit from '@pdf-lib/fontkit';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'

import * as express from 'express';
// @ts-ignore
import type {Buffer} from "exceljs/index";

const pdf = require('pdf-parse');

const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);

// картинка qrcode 75 на 75

type tPFD = {
    transform: [sizeFont: number, t1:number, t2:number, t3:number, x: number, y: number],
    pageIndex: number,
    pageView: [x: number, y: number, w: number, h: number],
    fontName: string,
    width:number,
    height:number
}

async function render_page(pageData:any) {
    let render_options = {
        normalizeWhitespace: false,
        disableCombineTextItems: false
    }
    const textContent = await pageData.getTextContent(render_options)

    const obj: { [key: string]: tPFD } = {}
    for (let item of textContent.items) {
        if (item.str.includes('key_')) {
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

type tCellInfo = {[key:string]: {
        rangeX:[number,number],
        rangeY:[number,number],
        font: {name:string,style:'origin'|'bold'|'italic'|'boldItalic'},
        alignment:{
            vertical:'top'|'bottom'|'middle'|'distributed'|'justify',
            horizontal: 'left'|'right'|'center'|'fill'|'justify'|'centerContinuous'|'distributed'
        },
        width?:number,
        height?:number
    }}


type tExcel =  tCellInfo
type tMapExcel = {[key: string]: tExcel}
type tMapPDF = {[key: string]: Buffer}

async function ExcelToMapCell(file: Buffer) {
    // чтение стиля из эксель
    const workbook:Workbook = new ExcelJS.Workbook();
    const w = await workbook.xlsx.load(file);
    const firstSheet = w.getWorksheet(1)

    const cellInfo: tCellInfo = {}

    firstSheet.eachRow((row,rowNumber)=>{
        row.eachCell((cell:any, colNumber)=> {
            if(cell.value?.includes('key_')){
                const style = cell.master ? cell.master.style : cell.style
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
            w += Math.round(6*x._column.width+5); // 6 ширина символа шрифта (проверить надо точную !!)  , 5 - padding (тоже примерно)
        }
        for(let i=value.rangeY[0];i<=value.rangeY[1];i++){
            const x:any = firstSheet.getRow(i)
            h += x.height;
        }
        cellInfo[key].width=w
        cellInfo[key].height=h
    }
    return cellInfo
}

type tDataKey = {
    [key: string]: string | Buffer
}
type tRequest = {
    // тия шаблона
    [nameTemplate: string]: tDataKey[]
}


async function convertExcToPDF(excel: Buffer) {
    return await libre.convertAsync(excel, '.pdf', undefined) as Buffer
}

async function createPDF(pdfSimple:PDFDocument, keyMap:  {[key: string]: tPFD}, dataKey: tDataKey[], excelKey: tCellInfo) {
    const length = pdfSimple.getPages().length
    const pdfDocCopy = await pdfSimple.copy()
    const arr:number[] = (new Array(length)).map((v,i)=>i)
    const data = await pdfDocCopy.copyPages(pdfSimple,arr) //
    for (let arrElement of dataKey) {
        for (let i = 0; i < length; i++) {
            pdfDocCopy.addPage(data[i])
        }
    }

    pdfDocCopy.registerFontkit(fontkit);
    const customFont = {
        origin: await pdfDocCopy.embedFont(fs.readFileSync('./fonts/arial.ttf')),
        italic: await pdfDocCopy.embedFont(fs.readFileSync('./fonts/ariali.ttf')),
        bold: await pdfDocCopy.embedFont(fs.readFileSync('./fonts/arialbd.ttf')),
        boldItalic: await pdfDocCopy.embedFont(fs.readFileSync('./fonts/arialbi.ttf')),
    }

    const pages = pdfDocCopy.getPages()
    for (let i = 0; i < dataKey.length; i++) {
        const data = dataKey[i]
        for (const [key,value] of Object.entries(data)) {
            // console.log(name)
            const tt = keyMap[key]
            if (!tt) continue;
            if (typeof value == "string") {
                pages[tt.pageIndex + i*length]
                    .drawText(value,{
                        x: tt.transform[4],
                        y: tt.transform[5],
                        size: tt.transform[0],
                        font: customFont[excelKey[key]?.font.style ?? "origin"],
                        lineHeight: tt.transform[0] * 1.15,
                        // color:
                    })
            }
            else {
                // тут код для вставки картинки
            }
        }
    }
    return pdfDocCopy
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
        const resKey = await convertExcToPDF(xcl)
        const res = await convertExcToPDF(excelSimple)
        const pdfMapKey = await PDFToMapKey(resKey);

        mapPDFKey[name] = resKey;
        mapPDF[name] = res;
        mapPDFKeyMap[name] = pdfMapKey;
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

        for (const [name, value] of Object.entries(data)) {
            // файл с метками
            const pdfKey = mapPDFKey[ name ]// открытие буфера пдф по имени
            const pdf = mapPDF[ name ]// открытие буфера пдф по имени
            const excelKey = mapExcelStyle[ name ]
            const pdfMapKey = mapPDFKeyMap[ name ]
            if (!pdfMapKey) throw "не создан pdfMapKey с ключами для шаблона " + name
            if (!pdfKey) throw "не создан pdf с ключами для шаблона " + name
            if (!pdf) throw "не создан чистый pdf для шаблона " + name
            if (!excelKey) throw "не создана карта ключей и стилей по ecelдля шаблона " + name
            const result = await createPDF(pdf, pdfMapKey, value, excelKey)
            return result; // тут ошибка
        }
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

function start() {

    const HOST = '0.0.0.0';
    const PORT: number =  4051//+process.env.PORT

    const app = express();

    const server = require('http').createServer(app)

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

        const data2 = {name: data.name, excel: await fs.promises.readFile(data.excel), excelSimple: await fs.promises.readFile(data.excelSimple),}

        try {
            await api.addTemplateExcel( data2)
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
             имя ключа 1 : значение или текст или буфер картинки,
             имя ключа 2 : значение или текст или буфер картинки,
             },
            {
             имя ключа 1 : значение или текст или буфер картинки,
             имя ключа 2 : значение или текст или буфер картинки,
             }
         ],
      имя шаблона 2 :
        [
            {
             имя ключа 1 : значение или текст или буфер картинки,
             имя ключа 2 : значение или текст или буфер картинки,
             имя ключа 3 : значение или текст или буфер картинки,
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
    app.post('/dataToPDF', async (req, res) => {
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

    app.get('/statusGet', async (req, res) => {
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
    })
}
start()