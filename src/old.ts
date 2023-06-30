import * as ExcelJS from "exceljs";
import {Workbook} from "exceljs";
import * as fs from "fs";
import * as path from "path";
import * as fontkit from '@pdf-lib/fontkit';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'
import {Buffer} from "exceljs/index";

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
            } satisfies tPFD
        }
    }
    return obj
}

function parsingPDF() {

}

const customFont = async (pdfDoc: PDFDocument)=>({
    origin: await pdfDoc.embedFont(fs.readFileSync('./fonts/arial.ttf')),
    italic: await pdfDoc.embedFont(fs.readFileSync('./fonts/ariali.ttf')),
    bold: await pdfDoc.embedFont(fs.readFileSync('./fonts/arialbd.ttf')),
    boldItalic: await pdfDoc.embedFont(fs.readFileSync('./fonts/arialbi.ttf')),
});

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

async function start() {
    console.time('pdf')

    const nameFIle = '4pgruz.pdf'


    let dataBuffer = fs.readFileSync(nameFIle);
    const resultPDF: {[key: string]: tPFD} = {}
    await pdf(dataBuffer, {pagerender: (data: any )=>(render_page(data, resultPDF))})
    console.log("result ", resultPDF)

    const pdfDoc = await PDFDocument.load(dataBuffer)


    pdfDoc.registerFontkit(fontkit);



    // чтение стиля из эксель
    const workbook:Workbook = new ExcelJS.Workbook();
    const w = await workbook.xlsx.readFile("4pgruz.xlsx");
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
    // console.log(JSON.stringify(cellInfo,null,2));
    // {
    //   "key_orgInfo": {
    //     "rangeX": [
    //       9,
    //       87
    //     ],
    //     "rangeY": [
    //       4,
    //       4
    //     ],
    //     "font": {
    //       "name": "Arial Cyr",
    //       "style": "origin"
    //     },
    //     "alignment": {
    //       "vertical": "middle",
    //       "horizontal": "left"
    //     },
    //     "width": 136.0546875,
    //     "height": 13
    //   }
    // }



    //////// Проверка областей - отрисовка красных прямоугольников и добавление qr code
    // const zoom = 0.65 // в ПДФ меньше масштаб (от балды поставил)
    // const pngImage = await pdfDoc.embedPng(fs.readFileSync('qr.png'))
    // const pages = pdfDoc.getPages()
    // Object.keys(resultPDF).forEach(key=>{
    //     const excelInfo = cellInfo[key]
    //     const pdfInfo = resultPDF[key]
    //     if(key=='key_image'){
    //         cellInfo['key_image'].height=75*zoom
    //         cellInfo['key_image'].width=75*zoom
    //         pages[0].drawImage(pngImage, {
    //             x: pdfInfo.transform[4],
    //             y: pdfInfo.transform[5]-(excelInfo.height!*zoom)+pdfInfo.height,
    //             width: excelInfo.width!*zoom,
    //             height: excelInfo.height!*zoom,
    //         });
    //     }
    //     else {
    //         if(excelInfo.alignment.horizontal != "right"){
    //             pages[0].drawRectangle({
    //                 x: pdfInfo.transform[4],
    //                 y: pdfInfo.transform[5]-(excelInfo.height!*zoom)+pdfInfo.height,
    //                 width: excelInfo.width!*zoom,
    //                 height: excelInfo.height!*zoom,
    //                 borderColor: rgb(1, 0, 0),
    //             })
    //         }
    //         else {
    //             pages[0].drawRectangle({
    //                 x: pdfInfo.transform[4] - excelInfo.width!*zoom + pdfInfo.width,
    //                 y: pdfInfo.transform[5]-(excelInfo.height!*zoom)+pdfInfo.height,
    //                 width: excelInfo.width!*zoom,
    //                 height: excelInfo.height!*zoom,
    //                 borderColor: rgb(1, 0, 0),
    //             })
    //         }
    //     }
    // })








    const text = 'This is text in an embedded font!'
    const textSize = 35
    const convertFunc = async (pdfDoc: PDFDocument)=> ({
        "g_d0_f1" : (await customFont(pdfDoc)).bold,
        "g_d0_f2" : (await customFont(pdfDoc)).origin,
        "g_d0_f3" : (await customFont(pdfDoc)).boldItalic,
        "g_d0_f4" : (await customFont(pdfDoc)).italic,
    })

    const convert = convertFunc(pdfDoc)

    ////////

    // const pdfDocCopy = await pdfDoc.copy()
    const pdfDocCopy = await pdfDoc.copy()/// pdfDoc// await pdfDoc.copy()// await PDFDocument.create() // await pdfDoc.copy()
    pdfDocCopy.registerFontkit(fontkit);

    const Data = {
        arr: <{}[]>[1,2,3]
    }
    Data.arr.length=5000

    // const pdfBytes = await pdfDoc.save()


    const data = await pdfDocCopy.copyPages(pdfDoc,[0,1]) //
    for (let arrElement of Data.arr) {
        pdfDocCopy.addPage(data[0])
        pdfDocCopy.addPage(data[1])
    }
    // const arr = await  pdfDocCopy.save()

    console.timeEnd("112")
    console.time("11")

    const pages = pdfDocCopy.getPages()
    for (let i = 0; i < pages.length; i++) {
        // if (i%100==1) console.log(Math.round(i*100/pages.length) + "%");
        if (i%2==1) continue;
        for (const [key,value] of Object.entries(resultPDF)) {
            const name =key+"! sdasdasd! " + i
            // console.log(name)
            pages[value.pageIndex + i]
                .drawText(name,{
                    x: value.transform[4],
                    y: value.transform[5],
                    size: value.transform[0],
                    font: convert[value.fontName],
                    lineHeight: value.transform[0] * 1.15,
                    // color:
                })
        }
    }

    console.timeEnd("11")


    const pdfBytes = await pdfDocCopy.save()
    fs.writeFileSync('test.pdf',pdfBytes)

    return

}


start()


type tExcel =  tCellInfo
type tMapExcel = {[key: string]: tExcel}

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

function ffff3(pdf: any, pdfTemplate: any, cell: any) {

}
type tDataKey = {
    [key: string]: string | Buffer
}
type tRequest = {
    // тия шаблона
    [nameTemplate: string]: tDataKey[]
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

            }
        }
    }

}

function fApi() {
    const mapExcelStyle: tMapExcel = {}
    const addTemplateExcel = async ({excelSimple, excel, name}: {excel: Buffer, name: string, excelSimple?: Buffer}) => {
        const xcl = await ExcelToMapCell(excel)
        mapExcelStyle[name] = xcl;
        /// надо конвертнуть экссель в пдф
        return true;
    }

    const PDFToMapKey = async (pdfBuffer: Buffer) => {
        return new Promise<{[p: string]: tPFD}>((resolve, reject)=>{
            pdf(pdfBuffer, {
                pagerender: async (data: any )=>{
                    resolve(await render_page(data))
                }})
        })
    }

    const dataToPDF = (data: tRequest) => {
        Object.entries(data).forEach(([nameTemplate, value])=>{
            // файл с метками
            const pdf = {}// Open(nameTemplate) // открытие буфера пдф по имени
            const pdfTemplate = {}
            const pdfMapKey = PDFToMapKey(pdf)
            const excelKey = mapExcelStyle[nameTemplate]
            createPDF(pdf, pdfMapKey, value, excelKey)

        })
    }
    return {
        addTemplateExcel
    }
}
