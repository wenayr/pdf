//////<reference lib="dom" />
//import * as Excel from "excel4node";
//import Excel from "types/excel4node";
import {Workbook} from "exceljs";
import ExcelJS from "exceljs"; import {aResult} from "../address";
import fs from "fs";
//import node_convert from "../../my_modules/node-convert_my/src";
import unoconv from "../unoconv_my";
import {tExcel} from "../interface";

//import * as PDFJS_dist from "pdfjs-dist"; // Версия не старше 3.7.107, иначе м.б. ошибка компиляции // requires 'dom'

//import * as xlsx from "xlsx";

import * as libre from 'libreoffice-convert';
//import libre2 from 'libreoffice-convert';
import * as util from 'util';
const libreConvertAsync = util.promisify(libre.convert);

async function convertExcToPDF(excel: Buffer) {
    return await libreConvertAsync(excel, '.pdf', undefined) as Buffer
}



// Не используем эту функцию, т.к. в создаваемом буфере exceljs сбиты настройки страницы относительно оригинала.

export async function Excel_removeKeys(buffer: Buffer) : Promise<Buffer> {

    //let wb = new Excel.Workbook();
    // чтение стиля из эксель
    const workbook:Workbook = new ExcelJS.Workbook();
    const w = await workbook.xlsx.load(buffer);
    const firstSheet = w.getWorksheet(1);

    firstSheet.eachRow((row,rowNumber)=>{
        row.eachCell((cell, colNumber)=> {
            if(typeof(cell.value)=="string" && cell.value.includes('key_')){
                cell.value="";
            }
        });
    })
    let buf= await workbook.xlsx.writeBuffer({useStyles: true});
    return buf as Buffer;
    return Buffer.from(buf);
    // let aa : Uint8Array = buf.slice(0, buf.byteLength);
    // return aa;
    //return buf;
}

// Не используем эту функцию, т.к. PDFJS_dist работает только на чтение данных. Менять содержимое объекта невозможно.

// async function PDF_removeKeys(pdfBuffer :Buffer)
// {
//     let doc= await PDFJS_dist.getDocument(pdfBuffer.buffer).promise;
//     //let ddd= PDFJS_dist.;
//     console.log(doc);
//     console.log("!!!");
//     console.log("numpages: ",doc.numPages);
//     const storage = doc.annotationStorage;
//     if (1)
//     for(let iPage=1; iPage<=doc.numPages; iPage++) {
//         console.log("page #"+iPage);
//         let page= await doc.getPage(iPage);
//         //console.log("page1:",page);
//         let cont= await page.getTextContent();
//         if (1)
//         for(let [i,item] of cont.items.entries()) {
//             if ("str" in item)
//             if (item.str.includes("key_")) {
//                 console.log("remove key: ",item.str);
//                 item.str="";
//                 cont.items.splice(i,i+1);  // delete item
//                 //console.log((await page.getTextContent()).items[i]);
//                 //storage.remove(item.str);
//             }
//            // if (i<30)
//            //     console.log(item);
//         }
//     }
//     //doc.cleanup();
//
//     console.log(storage);
//     // storage.resetModified();
//    // for (const id in fillIn) {
//    //    storage.setValue(id, fillIn[id]);
//    // }
//     let outBuf= await doc.getData();//await doc.saveDocument();
//     //let outBuf= await doc.saveDocument();
//     console.log("bufSize:",outBuf.byteLength);
//     return Buffer.from(outBuf);
// }



function pixels2chars(px :number) {

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



type tCellPrivate = ExcelJS.Cell & { _column: ExcelJS.Column, _row: ExcelJS.Row, master: tCellPrivate};



async function ExcelToMapCell(file: Buffer) {
    // чтение стиля из эксель
    const workbook= new ExcelJS.Workbook();
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
                let cellInfo= cellsInfo[cell.value]?.[0];
                if (! cellInfo) {
                    let masterCell= cell.master;
                    const style = masterCell.style;
                    const font= style.font;
                    // console.log(row.getCell(cell._column._number));
                    //
                    // console.log("!!!!!");
                    cellInfo = {
                        //left: 0,
                        rangeX: [-1,-1],//[masterCell._column.number, cell._column.number],
                        rangeY: [-1,-1],//[masterCell._row.number, cell._row.number],
                        font: {
                            name: font?.name ?? "",
                            style: !font?.bold && !font?.italic ? 'origin' : font.bold && font.italic ? 'boldItalic' : font.bold ? 'bold' : 'italic',
                            strikeThrough: font?.strike,
                            color: font?.color?.argb
                        },
                        alignment: {
                            vertical: style.alignment?.vertical ?? 'top',
                            horizontal: style.alignment?.horizontal ?? 'left',
                        },
                        width: 0,
                        height: 0
                    };
                    if (cell.value=="key_Z2" || cell.value=="key_Z3")
                        console.log("excel cell: ",cell.value, cell, "->", cellsInfo[cell.value],"  master:",{value: masterCell.value, font:font, column: masterCell._column});
                }
                //if (cell.value=="key_Z2")
                    //console.log("internal cell: ",cell.value,{font: cell.style.font, column: cell._column});
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
                (cellsInfo[cell.value]??=[])[0]= cellInfo;
                //(cellsInfo[cell.value] ??=[]).push(cellInfo);
                //cellsInfo[cell.value]= cellInfo;
                //if (cell.value=="key_periodInfo") console.log("+",ch2px(cell._column.width??0));
                //cellsInfo[cell.value].rangeX[1]=
                //if (removeKeys) cell.value="";
            }
        });
    })

    for(let [key,[cell]] of Object.entries(cellsInfo)) {
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





    // if (0) {
    //
    //     let isNew= false;
    //     let tempFileName= aResult +"temp"+name+".xlsx";
    //     await fs.promises.writeFile(tempFileName, excel);
    //     let mode= 3;
    //     let excelNew=   //console.log("! 1:", timerTick(),"ms");
    //         (mode==1) ? await libreConvertAsync(excel, '.xlsx', undefined) as Buffer
    //       : (mode==2) ? await node_convert(tempFileName, {type: "xlsx"}) as Buffer
    //                   : await unoconv.convertPromise(excel, 'xlsx');
    //     console.log("! 1:", timerTick(),"ms")
    //
    //      fs.promises.writeFile(aResult +"new"+name+".xlsx", excelNew);
    //
    //     const excelInfo = await ExcelToMapCell(excel); console.log("! 2:", timerTick(),"ms");
    //
    //     {
    //         let cleanExcelBuf= await Excel_removeKeys(excel);  console.log("! 2.3:", timerTick(),"ms");
    //         fs.promises.writeFile(aResult +"cleanOld"+name+".xlsx", cleanExcelBuf);
    //     }
    //
    //     let cleanExcelBuf= await Excel_removeKeys(excelNew);  console.log("! 2.5:", timerTick(),"ms");
    //     fs.promises.writeFile(aResult +"clean"+name+".xlsx", cleanExcelBuf);
    //
    //   //console.log("exit");  return true;
    //     //_mapExcelStyle[name] = excelInfo;
    // //if (1) return;
    //     /// надо конвертировать excel в пдф
    //
    //     let pdfKeyBuf : Buffer = //isNew ? await node_convert(tempFileName, {type: "pdf"}) as Buffer  //console.log("! 3:", timerTick(),"ms")
    //         //: await convertExcToPDF(excel);
    //         (mode==1) ? await libreConvertAsync(excel, '.pdf', undefined)
    //       : (mode==2) ? await node_convert(tempFileName, {type: "pdf"}) as Buffer
    //                   : await unoconv.convertPromise(excel, 'pdf');
    //     console.log("! 3:", timerTick(),"ms");
    //
    //     fs.promises.writeFile(aResult +"testKey"+name+".pdf", pdfKeyBuf)
    //     console.log("exit");  return true;
    //
    // }

