///??? <reference path="../pdf-parse.d.ts" />
//import * as ExcelJS from "exceljs";
import ExcelJS from "exceljs";
import {Workbook} from "exceljs";

//import * as fs from "fs";
import fs from "fs";

import {createPDF} from "./createPDF";
import {tCellInfo,tKeyData,tExcel, tMapExcel, tMapPDF, tObjImage, tPFD, tRequest,tPDFInfo} from "./interface";
import {PDFDocument, PDFImage} from "pdf-lib";
import {aFont, aImage,aResult} from "./address";


//globalThis.
//document.body
// import * as PdfParse_ from 'pdf-parse';
// import type PdfParse from '../pdf-parse';
// //import type PdfParse from '../pdf-parse';
// const pdfParse= PdfParse_ as typeof PdfParse;
//

//import * as PdfParse from 'pdf-parse-debugging-disabled;
//const pdfParse= PdfParse.default;

import PdfParse from 'pdf-parse-debugging-disabled';
const pdfParse= PdfParse;



//import {PageData} from "pdf-parse";
//import type {PageData} from "pdf-parse";

//const libre = require('libreoffice-convert');
//libre.convertAsync = require('util').promisify(libre.convert);

import * as libre from 'libreoffice-convert';
//import libre2 from 'libreoffice-convert';
import * as util from 'util';
const libreConvertAsync = util.promisify(libre.convert);



async function render_page(pageData :PdfParse.PageData)
{
    let render_options : PdfParse.RenderOptions = {
        normalizeWhitespace: false,
        disableCombineTextItems: false
    }
    const textContent : PdfParse.TextContent = await pageData.getTextContent(render_options);
    //console.log("textContent for page #",pageData.pageIndex,"\n",textContent);
    const obj: { [key: string]: tPFD } = {}
    for (let item of textContent.items satisfies PdfParse.PageItem[]) {
        // надо удалить все переносы строк если такие есть
        const str2 = item.str.replace(/\n/g, '');
        //type PageItem = PdfParse.PageItem;

        if (str2.includes('key_')) {
            //console.log("item:",str2, item);
            obj[item.str] = {
                transform: item.transform,
                pageIndex: pageData.pageIndex,
                pageView: pageData.pageInfo.view,
                fontName: item.fontName,
                width: item.width,
                height: item.height
            };// satisfies tPFD
        }
    }
    return obj
}


type tCellPrivate = ExcelJS.Cell & { _column: ExcelJS.Column, _row: ExcelJS.Row, master: tCellPrivate};





function picels2chars(px :number) {

    // there are two magic numbers: 1/12 and 1/7 that found experimentally
    // convertion assumes that Normal style has default font of "Calibri 11pt"
    // and display is 96ppi
    // !!!note: px must be integer
    let dpi = 120;
    if (px < 12) {
        return Math.round(px / 12 * 100) / 100 * 96/dpi
    }
    else {
        return Math.round((1 + (px - 12) / 7) * 100) / 100 * 96/dpi
    }
}
function chars2pixels(ch :number) {
    return ch * 4;
    let dpi = 120;
    if (ch < 1) {
        return Math.round(ch * 12) * dpi/96
    }
    else {
        return Math.round((ch - 1) * 7 + 12) * dpi/96
    }
}


async function ExcelToMapCell(file: Buffer) {
    // чтение стиля из эксель
    const workbook:Workbook = new ExcelJS.Workbook();
    const w = await workbook.xlsx.load(file);
    const firstSheet = w.getWorksheet(1);
    //await workbook.xlsx.writeFile("myFile.xlsx");

    const cellsInfo: tExcel = {}
    //const tt = TF.M15
    let a = false
    firstSheet.eachRow((row,rowNumber)=>{
        row.eachCell((cell_, colNumber)=> {
            let cell = cell_ as tCellPrivate;
            //if (cell==cell.master)
            if(typeof(cell.value)=="string" && cell.value.includes('key_')) {
                let cellInfo= cellsInfo[cell.value];
                if (! cellInfo) {
                    let masterCell= cell.master;
                    const style = masterCell.style;
                    const font= style.font;
                    // console.log(row.getCell(cell._column._number));
                    //
                    // console.log("!!!!!");
                    cellInfo = {
                        left: 0,
                        rangeX: [-1,-1],//[masterCell._column.number, cell._column.number],
                        rangeY: [-1,-1],//[masterCell._row.number, cell._row.number],
                        font: {
                            name: font?.name ?? "",
                            style: !font?.bold && !font?.italic ? 'origin' : font.bold && font.italic ? 'boldItalic' : font.bold ? 'bold' : 'italic'
                        },
                        alignment: {
                            vertical: style.alignment?.vertical ?? 'top',
                            horizontal: style.alignment?.horizontal ?? 'left',
                        },
                        width: 0,
                        height: 0
                    }
                    //console.log("excel cell: ",cell.value, cell, "->", cellsInfo[cell.value]);
                }
                if (cell._column.number > cellInfo.rangeX[1]) {
                  cellInfo.width += cell._column.width??0;
                  if (cellInfo.rangeX[0]==-1) cellInfo.rangeX[0]= cell._column.number;
                  cellInfo.rangeX[1]= cell._column.number;
                }
                if (cell._row.number > cellInfo.rangeY[1]) {
                    cellInfo.height += cell._row.height??0;
                    if (cellInfo.rangeY[0]==-1) cellInfo.rangeY[0]= cell._row.number;
                    cellInfo.rangeY[1]= cell._row.number;
                }
                cellsInfo[cell.value]= cellInfo;
                //if (cell.value=="key_periodInfo") console.log("+",ch2px(cell._column.width??0));
                //cellsInfo[cell.value].rangeX[1]=
                //if (removeKeys) cell.value="";
            }
        });
    })

    for(let [key,cell] of Object.entries(cellsInfo)) {
        cell.width= Math.round(chars2pixels(cell.width) +0);
        cell.height= Math.round(chars2pixels(cell.height) +0);
        //console.log(key, cell);
    }

    //console.log(cellsInfo);
    // const row1= firstSheet.getRow(1)
    // for(const [key,value] of Object.entries(cellsInfo)){
    //     let w = 0;
    //     let h = 0;
    //     let xx = 0;
    //     for(let i=1; i<=value.rangeX[0]; i++) {
    //         const cell = row1.getCell(i) as tCellPrivate;
    //         xx += Math.round(cell._column.width??0 +5); // 6 ширина символа шрифта (проверить надо точную !!)  , 5 - padding (тоже примерно) // 6*
    //         // console.log(x._column)
    //     }
    //     for(let i=value.rangeX[0]; i<=value.rangeX[1]; i++){
    //         const cell = row1.getCell(i) as tCellPrivate;
    //         w += Math.round(cell._column.width??0 +5); // 6 ширина символа шрифта (проверить надо точную !!)  , 5 - padding (тоже примерно) // 6*
    //         // console.log(x._column)
    //     }
    //     // console.log(w)
    //     for(let i=value.rangeY[0]; i<=value.rangeY[1]; i++){
    //         const row = firstSheet.getRow(i)
    //         h += row.height;
    //     }
    //     cellsInfo[key]!.width= w;
    //     cellsInfo[key]!.height= h;
    // }
    return cellsInfo;
}



async function convertExcToPDF(excel: Buffer) {
    return await libreConvertAsync(excel, '.pdf', undefined) as Buffer
}

let _fonts: {origin: Buffer, italic: Buffer, bold: Buffer, boldItalic: Buffer} | undefined;

async function getFonts() {
    if (!_fonts) _fonts = {
        origin: await (fs.promises.readFile(aFont + 'arial.ttf')),
        italic: await (fs.promises.readFile(aFont + 'ariali.ttf')),
        bold: await (fs.promises.readFile(aFont + 'arialbd.ttf')),
        boldItalic: await (fs.promises.readFile(aFont + 'arialbi.ttf')),
    }
    return _fonts
}


type TemplateData = {
    name : string,
    excelData : tExcel,
    pdfData : tPDFInfo,
    pdfBuffer : Buffer;
};

type TempSaveInfoData = { excelData: tExcel, pdfData : tPDFInfo };



const TEMPL_FILE_DIR = "dist/templates/";

async function saveTemplate(template : TemplateData) {
    const dirPath= TEMPL_FILE_DIR;
    if (! fs.existsSync(dirPath)) console.log("Создание папки ",dirPath);
    await fs.promises.mkdir(dirPath, {recursive: true});
    let t= template;
    let infoFile= TEMPL_FILE_DIR+"templateInfo_"+t.name+".json";
    let pdfFile= TEMPL_FILE_DIR+"template_"+t.name+".pdf";
    await fs.promises.writeFile(infoFile, JSON.stringify({excelData: t.excelData, pdfData: t.pdfData} satisfies TempSaveInfoData))
        .then(()=>console.log("Записан файл",infoFile));
    await fs.promises.writeFile(pdfFile, t.pdfBuffer)
        .then(()=>console.log("Записан файл",pdfFile));;
}

async function loadTemplate(templateName :string) : Promise<TemplateData|null> {
    let infoFile= TEMPL_FILE_DIR+"templateInfo_"+templateName+".json";
    let pdfFile= TEMPL_FILE_DIR+"template_"+templateName+".pdf";
    if (! fs.existsSync(infoFile)) return null;
    if (! fs.existsSync(pdfFile)) return null;
    let infoBuf= await fs.promises.readFile(infoFile).catch((e)=>{console.error("Failed to read file",infoFile,":",e); });
    if (! infoBuf) return null;
    let infoJSON : TempSaveInfoData|{}= JSON.parse(infoBuf.toString());
    if (!("excelData" in infoJSON && "pdfData" in infoJSON)) {
        console.error("Wrong template info file:",infoFile); return null;
    }
    let info= infoJSON;
    let buffer= await fs.promises.readFile(pdfFile).catch((e)=>{console.error("Failed to read file",pdfFile,":",e); });
    if (! buffer) return null;
    return {
        name: templateName,
        excelData: info.excelData,
        pdfData: info.pdfData,
        pdfBuffer: buffer
    };
}


export function fApi()
{
    const _map : {[k: string] : TemplateData} = { };

    const _mapExcelStyle: tMapExcel = {}
    //const _mapPDFKeyMap: {[k: string]: {[p: string]: tPFD}}  = {}
    //const _mapPDF: tMapPDF = {}
    //const _mapPDFKey: tMapPDF = {}

    const addTemplateExcel = async ({excelSimple, excel, name}: {excel: Buffer, name: string, excelSimple: Buffer}) => {
        console.log("! 1");
        let t= Date.now();
        function timerTick() { let delta= Date.now()-t;  t= Date.now();  return delta; }

        const excelInfo = await ExcelToMapCell(excel); console.log("! 2:", timerTick(),"ms");
        //_mapExcelStyle[name] = excelInfo;
    //if (1) return;
        /// надо конвертировать excel в пдф
        const pdfKeyBuf = await convertExcToPDF(excel); console.log("! 3:", timerTick(),"ms");
        //
        //excelSimple= await ExcelRemoveKeys(excel);  console.log("! 4:", timerTick(),"ms");

        const pdfCleanBuf = await convertExcToPDF(excelSimple); console.log("! 5:", timerTick(),"ms");

        //const pdfCleanBuf= await removeKeysFromPDF(pdfKeyBuf);
        //await fs.promises.writeFile(aResult +"testNew"+name+".pdf", pdfCleanBuf)
        //return;
        const pdfKeyMap = await PDFToMapKey(pdfKeyBuf); console.log("! 6:", timerTick(),"ms");

        //_mapPDFKey[name] = pdfKeyBuf;
        //_mapPDF[name] = pdfCleanBuf;
        //_mapPDFKeyMap[name] = pdfKeyMap;

        let data : TemplateData = {
            name,
            excelData : excelInfo,
            pdfData : pdfKeyMap,
            pdfBuffer : pdfCleanBuf
        };

        _map[name] = data;

        //console.log("bufferSimple:",pdfCleanBuf);
        //console.log("buffer:",pdfKeyBuf);
        //console.log("pdfMapKey:\n",pdfKeyMap);

        // для проверки сохранит промежуточные PDF
        if (0) {
            fs.promises.writeFile(aResult +"test"+name+".pdf", pdfCleanBuf)
            fs.promises.writeFile(aResult +"testKey"+name+".pdf", pdfKeyBuf)
        }

        await saveTemplate(data).catch(e=>console.error("Failed to save template "+name+":",e));

        return true;
    }

    const PDFToMapKey = (pdfBuffer: Buffer) => {
        return new Promise<{[p: string]: tPFD}>((resolve, reject)=>{
            let obj: { [key: string]: tPFD } = {};
            pdfParse(pdfBuffer, {
                pagerender: async(data)=>{
                    obj = Object.assign(obj, await render_page(data));
                    // obj = {...obj, ...await render_page(data)}
                    return undefined;
                }
            }).then(()=>resolve(obj));
            return;
            // let numpages = 0
            // // =) костыль, по другому непонятно как заранее узнать количество страниц
            // pdfParse(pdfBuffer, {pagerender:()=>undefined}).then(result=>{
            //     numpages = result.numpages;
            //     let obj: { [key: string]: tPFD } = {};
            //     //console.log("!!");
            //     pdfParse(pdfBuffer, {
            //         pagerender: async(data)=>{
            //             obj = Object.assign(obj, await render_page(data));
            //             numpages--;
            //             // obj = {...obj, ...await render_page(data)}
            //             if (numpages <=0) { resolve(obj);  obj= { }; }
            //             return undefined;
            //         }
            //     })
            // })
        })
    }

    /// dataToPDF    req:  {[p: string]: string | Buffer} res: {status:"ok"}

    const dataToPDFMulti = async (data: tRequest) => {
        let documents : PDFDocument[] = [];
        for (const [name, keyData] of Object.entries(data))
            documents.push(await dataToPDF(name, Array.isArray(keyData) ? keyData : [keyData]));
        return documents;
    }

    const dataToPDF = async (name: string, keyDatas: readonly tKeyData[]) => {

        let templateData : TemplateData|null = _map[name];
        if (! templateData) {
            templateData= await loadTemplate(name);
            if (templateData) { _map[name]= templateData; console.log("Прочитан шаблон",name,"из файла"); }
        }
        if (! templateData)
            throw "отсутствуют данные для шаблона " + name;


        const fonts = await getFonts()
            .catch((e)=>{
                console.error("error")
                throw " font " + e
            });

        const objImageB: tObjImage = {}
        async function loadImage(name: string) {
            return objImageB[name] ??= await fs.promises.readFile(aImage + name)
            .catch((e)=>{
                console.error("readFile error ");
                throw " cannot read " + name + e;
            })
        }

        // файл с метками
        //const pdfKey = mapPDFKey[ name ]// открытие буфера пдф по имени
        const pdf = templateData.pdfBuffer; // открытие буфера пдф по имени3
        const excelKeyMap = templateData.excelData;
        const pdfKeyMap = templateData.pdfData;
        if (!pdfKeyMap) throw "не создан pdfMapKey с ключами для шаблона " + name
        //if (!pdfKey) throw "не создан pdf с ключами для шаблона " + name
        if (!pdf) throw "не создан чистый pdf для шаблона " + name
        if (!excelKeyMap) throw "не создана карта ключей и стилей по excel для шаблона " + name

        const objImage: tObjImage = {}
        {
            const arr2: Promise<unknown>[] = []
            for (const data of keyDatas)
                for (const value of Object.values(data))
                    if (typeof(value)=="object" && value?.name!=null)
                        arr2.push(
                            loadImage(value.name).then(result => objImage[value.name] = result)
                        );

            await Promise.all(arr2)
                .catch((e)=>{
                    console.error("failed to load images")
                    throw " error promise all reading " + e
                })
        }

        const result = await createPDF(pdf, pdfKeyMap, keyDatas, excelKeyMap, fonts, objImage, name)
            .catch((e)=>{
                throw " createPDF " + e
            })
        return result;
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
        dataToPDFMulti,
        // получить все текущие шаблоны
        getExcel: ()=> _mapExcelStyle
    }
}



