///??? <reference path="../pdf-parse.d.ts" />
//import * as ExcelJS from "exceljs";

//import {Workbook} from "exceljs";

//import * as fs from "fs";
import fs from "fs";

import {createPDF} from "./createPDF";
import {tCellInfo,tKeyData,tExcel, tMapExcel, tMapPDF, tObjImage, tPFD, tRequest,tPDFInfo,tRequestAddTemplateByBuffer} from "./interface";
import {PDFDocument, PDFImage} from "pdf-lib";
import {aFont, aImage,aResult} from "./address";


//import * as PdfParse from 'pdf-parse-debugging-disabled;
//const pdfParse= PdfParse.default;

import PdfParse from 'pdf-parse-debugging-disabled';
const pdfParse= PdfParse;

//import {PageData} from "pdf-parse";
//import type {PageData} from "pdf-parse";

//const libre = require('libreoffice-convert');
//libre.convertAsync = require('util').promisify(libre.convert);

import unoconv from "./unoconv_my";

import XlsxPopulate from 'xlsx-populate';
//import node_convert from '../my_modules/node-convert_my/src';



async function render_page(pageData :PdfParse.PageData)
{
    let render_options : PdfParse.RenderOptions = {
        normalizeWhitespace: false,
        disableCombineTextItems: false
    };
    const textContent : PdfParse.TextContent = await pageData.getTextContent(render_options);
    //console.log("textContent for page #",pageData.pageIndex,"\n",textContent);
    const obj: { [key: string]: tPFD[] } = {}
    for (let item of textContent.items satisfies PdfParse.PageItem[]) {
        // надо удалить все переносы строк если такие есть
        const str2 = item.str.replace(/\n/g, '');
        //type PageItem = PdfParse.PageItem;

        if (str2.includes('key_')) {
            //console.log("item:",str2, item);
            let data : tPFD = {
                transform: item.transform,
                pageIndex: pageData.pageIndex,
                pageView: pageData.pageInfo.view,
                fontName: item.fontName,
                width: item.width,
                height: item.height
            };// satisfies tPFD
            (obj[item.str] ??= []).push(data);
        }
    }
    return obj
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

import {ChildProcessWithoutNullStreams} from "node:child_process";
import {Excel_removeKeys} from "./exclude/helper_exclude";




async function parseExcel(buffer :Buffer) {
    //console.log("!");
    let workBook= await XlsxPopulate.fromDataAsync(buffer);
    //console.log("!!");
    const sheet = workBook.sheet(0);
    //await fs.promises.writeFile(aResult+"new.txt", await JSON.stringify(workBook));
    //let obj : {[key: string] : number} = {};
    const cellsInfo: tExcel = {};
    let _keyCells : XlsxPopulate.Cell[]= [];
    let _keyCellValues : any[]= [];
    let _isKeysRemoved= false;
    //console.log((workBook.properties() as any)._properties , (workBook.properties() as any)._node);

    let mergesObj= (sheet as any)._mergeCells;

    for(let [rangeName, mergeInfo] of Object.entries(mergesObj) as [string, {attributes :{ref:string}}][]) {//XlsxPopulate.Range][]) {
        let rangeRef= mergeInfo.attributes.ref;
        let range= sheet.range(rangeRef); //attributes

        let firstCell= range.startCell();
        let lastCell= range.endCell();
        let cell= firstCell;
        let value= cell.value()?.toString() ?? "";
        if (value.includes("key_")) {
        //if (value.includes("key_Z") || value=="key_A1") {
            let cols= lastCell.columnNumber()-firstCell.columnNumber()+1;
            let rows= lastCell.rowNumber()-firstCell.rowNumber()+1;

            let width= 0;  for(let i=0; i<cols; i++) width+= sheet.column(firstCell.columnNumber()+i).width()??0;
            let height= 0; for(let i=0; i<rows; i++) height+= sheet.row(firstCell.rowNumber()+i).height()??0;

            let keys= ["fontFamily", "fontSize", "bold", "italic", "underline", "strikethrough", "horizontalAlignment", "verticalAlignment", "fontColor"] as const;
            let vals= keys.map(key=>{ let val="unknown"; try { val= cell.style(key); } catch(e) { throw e; }  return [key,val] });
            //let style= cell.style();
            let styleMap_= vals.reduce( (obj, [key,val])=>{obj[key]=val;  return obj; }, {} as {[k:string]:unknown} );
            let style= styleMap_ as {[key in typeof keys[number]] : unknown};

            if (0)
            console.log("cell",value, { range: rangeName, cellAddr:cell.address() }, cols+"x"+rows, {width, height}, style);

            (cellsInfo[value] ??=[]).push({
                //left: 0,
                rangeX: [firstCell.columnNumber(), lastCell.columnNumber()], // range in cells
                rangeY: [firstCell.rowNumber(), lastCell.rowNumber()], // range in cells
                font: {
                    name: style.fontFamily as string,
                    style: style.bold && style.italic ? 'boldItalic' : style.bold ? 'bold' : style.italic ? 'italic' : 'origin',
                    strikeThrough: style.strikethrough as boolean,
                    color: style.fontColor===undefined || typeof(style.fontColor)=="string" ? style.fontColor : "#"+(style.fontColor as XlsxPopulate.Color).rgb
                },
                alignment: {
                    vertical: style.verticalAlignment=="center" ? "middle" : style.verticalAlignment as any ?? "bottom", //'top' | 'bottom' | 'middle' | 'distributed' | 'justify',
                    horizontal: style.horizontalAlignment as any,
                },
                width: width*4,
                height: height*0.75
            });
            _keyCells.push(cell);
            _keyCellValues.push(value);
            //cell.value("");
        }
    }

    return {
        keysInfo: cellsInfo,
        async export(removeKeys=false) {
            if (removeKeys!=_isKeysRemoved) for(let [i,key] of _keyCellValues.entries()) _keyCells[i].value(removeKeys ? "" : key);
            _isKeysRemoved= removeKeys;
            return await workBook.outputAsync(Buffer) as Buffer
        }
    } as const;
}



export function fApi()
{
    const _map : {[k: string] : TemplateData} = { };

    let _excelInfoMap : tMapExcel|undefined;// = {}

    unoconv.listen();
    //const _mapPDFKeyMap: {[k: string]: {[p: string]: tPFD}}  = {}
    //const _mapPDF: tMapPDF = {}
    //const _mapPDFKey: tMapPDF = {}

    const addTemplateExcel = async ({excelSimple, excel, name}: tRequestAddTemplateByBuffer) => {
        console.log("! 0");
        let t= Date.now();
        function timerTick() { let delta= Date.now()-t;  t= Date.now();  return delta; }


        let bookInfo= await parseExcel(excel);  console.log("! 1:", timerTick(),"ms");
        let fontsOk= Object.values(bookInfo.keysInfo).length==0 ||  Object.entries(bookInfo.keysInfo).some(([key,infos])=>infos.find(info=> !info.font.strikeThrough));
        // дополнительная переконвертация, т.к. может быть некорректный формат файла (неправильные шрифты)
        if (!fontsOk) {
            excel= await unoconv.convertAsync(excel, 'xlsx');  console.log("! 1.5:", timerTick(),"ms");
            bookInfo= await parseExcel(excel);  console.log("! 2:", timerTick(),"ms");
        }
        //let excelInfo0 = await ExcelToMapCell(excel);  console.log("! 1:", timerTick(),"ms");
        //console.log(excelInfo0);

        //if (1) return true;
        //excel= await bookInfo.export(true);
        //await fs.promises.writeFile(aResult+"pop "+name+".xlsx", excel);  console.log("! 3:", timerTick(),"ms");
        //console.log("exit");  return true;

        let excelInfo= bookInfo.keysInfo;

        excelSimple ??= await bookInfo.export(true);  console.log("! 3:", timerTick(),"ms");

        let pdfKeyBuf : Buffer= await unoconv.convertAsync(excel, 'pdf');  console.log("! 4:", timerTick(),"ms");
        //
        //excelSimple= await ExcelRemoveKeys(excel);  console.log("! 4:", timerTick(),"ms");

        const pdfCleanBuf = await unoconv.convertAsync(excelSimple, 'pdf'); console.log("! 5:", timerTick(),"ms");

        //const pdfCleanBuf= await removeKeysFromPDF(pdfKeyBuf);
        //await fs.promises.writeFile(aResult +"testNew"+name+".pdf", pdfCleanBuf)
        //return;
        const pdfKeyMap = await PDFToMapKey(pdfKeyBuf);  console.log("! 6:", timerTick(),"ms");

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
        _excelInfoMap= undefined;
        //console.log("bufferSimple:",pdfCleanBuf);
        //console.log("buffer:",pdfKeyBuf);
        //console.log("pdfMapKey:\n",pdfKeyMap);

        // для проверки сохранит промежуточные PDF
        if (1) {
            try {
                let task1= fs.promises.writeFile(aResult +"test"+name+".pdf", pdfCleanBuf)
                let task2= fs.promises.writeFile(aResult +"testKey"+name+".pdf", pdfKeyBuf)
            }
            catch(e) {
                console.error(e);
            }
        }

        await saveTemplate(data).catch(e=>console.error("Failed to save template "+name+":",e));

        return true;
    }

    const PDFToMapKey = (pdfBuffer: Buffer) => {
        return new Promise<{[p: string]: tPFD[]}>((resolve, reject)=>{
            let obj: { [key: string]: tPFD[] } = {};
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
            if (templateData) {
                _map[name]= templateData;
                _excelInfoMap= undefined;
                console.log("Прочитан шаблон",name,"из файла");
            }
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
        getExcel: ()=> _excelInfoMap ??= (()=>{
            let obj :tMapExcel= {};
            for(let [key, data] of Object.entries(_map)) obj[key]= data.excelData;
            return obj;
        })()
    }
}



