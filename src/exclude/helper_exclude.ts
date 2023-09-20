//////<reference lib="dom" />
//import * as Excel from "excel4node";
//import Excel from "types/excel4node";
import {Workbook} from "exceljs";
import * as ExcelJS from "exceljs";
import * as PDFJS_dist from "pdfjs-dist"; // Версия не старше 3.7.107, иначе м.б. ошибка компиляции // requires 'dom'

//import * as xlsx from "xlsx";


// Не используем эту функцию, т.к. в создаваемом буфере exceljs сбиты настройки страницы относительно оригинала.

async function Excel_removeKeys(buffer: Buffer) : Promise<Buffer> {

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

async function PDF_removeKeys(pdfBuffer :Buffer)
{
    let doc= await PDFJS_dist.getDocument(pdfBuffer.buffer).promise;
    //let ddd= PDFJS_dist.;
    console.log(doc);
    console.log("!!!");
    console.log("numpages: ",doc.numPages);
    const storage = doc.annotationStorage;
    if (1)
    for(let iPage=1; iPage<=doc.numPages; iPage++) {
        console.log("page #"+iPage);
        let page= await doc.getPage(iPage);
        //console.log("page1:",page);
        let cont= await page.getTextContent();
        if (1)
        for(let [i,item] of cont.items.entries()) {
            if ("str" in item)
            if (item.str.includes("key_")) {
                console.log("remove key: ",item.str);
                item.str="";
                cont.items.splice(i,i+1);  // delete item
                //console.log((await page.getTextContent()).items[i]);
                //storage.remove(item.str);
            }
           // if (i<30)
           //     console.log(item);
        }
    }
    //doc.cleanup();

    console.log(storage);
    // storage.resetModified();
   // for (const id in fillIn) {
   //    storage.setValue(id, fillIn[id]);
   // }
    let outBuf= await doc.getData();//await doc.saveDocument();
    //let outBuf= await doc.saveDocument();
    console.log("bufSize:",outBuf.byteLength);
    return Buffer.from(outBuf);
}
