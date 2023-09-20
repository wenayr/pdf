import {PDFDocument, PDFImage, PDFObject,PDFTextField, drawLinesOfText,PDFFont} from "pdf-lib";
//import * as fontkit from "@pdf-lib/fontkit";
import fontkit from "@pdf-lib/fontkit";

import {tCellInfo, tKeyData,tExcel, tFonts, tObjectString, tObjImage, tPFD} from "./interface";



export async function createPDF(
    pdfSimpleBuf: Buffer,
    keyMap: {[key: string]: tPFD},
    keyDatas: readonly tKeyData[],
    excelKeys: tExcel,
    fonts: tFonts,
    objImage: tObjImage,
    tmplName: string
) : Promise<PDFDocument>
{
    const pdfDoc = await PDFDocument.load(pdfSimpleBuf)
        .catch((e) => {
            throw " PDFDocument.load error"
        })

    const pageCount = pdfDoc.getPages().length;
    const pdfDocCopy = await pdfDoc.copy();
    const arr: number[] = new Array(pageCount);

    console.log("!");
    const fields = pdfDoc.getForm().getFields();
    for(let field of fields) {
      const type = field.constructor.name;
      const name = field.getName();
      //if (field instanceof PDFTextField)
        console.log(`field:  ${type}: ${name}`);
    }
    console.log("!!");

    //let d= new PDFJS.Document();
    //let doc1= new PDFJS.ExternalDocument(pdfSimpleBuf);

    //console.log("!!!!");
    if (0)
    for(let [i,p] of pdfDoc.getPages().entries())
        for(let [name, object] of p.node.entries())
            console.log("page #",i," entry:  ",name," ",object);

    //if(1) return pdfDoc;

    for (let i=0; i<pageCount; i++)
        arr[i] = i;

    for (let j=1; j<keyDatas.length; j++) {
        const data = await pdfDocCopy.copyPages(pdfDoc, arr);
        for (let i=0; i<pageCount; i++) //   for (let i = 1; i < length; i++)
            pdfDocCopy.addPage(data[i]);
    }

    pdfDocCopy.registerFontkit(fontkit);
    const customFont = {
        origin: await pdfDocCopy.embedFont(fonts.origin.buffer),
        italic: await pdfDocCopy.embedFont(fonts.italic.buffer),//await pdfDocCopy.embedFont(fonts.italic),
        bold: await pdfDocCopy.embedFont(fonts.bold.buffer),
        boldItalic:  await pdfDocCopy.embedFont(fonts.boldItalic.buffer),//await pdfDocCopy.embedFont(fonts.boldItalic),
    };
    type kCustomFont = keyof typeof customFont


    const pages = pdfDocCopy.getPages()

    // оптимизированная версия pngImage: PDFImage
    const objImage2 : {[key: string]: PDFImage} = {}
    for (const [k, v] of Object.entries(objImage))
        objImage2[k] = await pdfDocCopy.embedPng(v)

    let index = 0
    for (let iPage=0; iPage < keyDatas.length; iPage++) {
        const data = keyDatas[iPage]
        for (const [key, value] of Object.entries(data)) {
            const tt = keyMap[key]
            if (value == null) continue;
            if (typeof value == "string" || (typeof value == "object" && value.text)) {
                let text: string | undefined
                let obj: tObjectString | undefined

                //let cellWidth = 100;
                const cellInfo= excelKeys[key];
                if (!cellInfo) { console.warn("Wrong key: ",key);  continue; }

                let cellWidth = cellInfo.width;// /2;

                // if (cellInfo.width && cellInfo.width > 0) {
                //     if (nameTmpl == "madi.xlsx" || nameTmpl == "form3.xlsx") {
                //         cellWidth = cellInfo.width * 2.0;
                //     } else if (nameTmpl == "form4.xlsx") {
                //         cellWidth = cellInfo.width * 2.2;
                //     }
                //     else {
                //         cellWidth = cellInfo.width;
                //     }
                // }

                if (typeof value == "object")  obj = value as tObjectString
                else text = value
                if (!tt && text) continue;
                const cellFont= customFont[cellInfo?.font.style ?? "origin"];
                const cellFontSize= tt.transform[0];
                const objFont = obj?.font ? customFont[obj.font] : cellFont;
                const fontSize= obj?.size ?? cellFontSize;
                console.log("key:",key, obj?.font, objFont, cellWidth);
                text ??= obj?.text ?? "none";
                let x= obj?.x;// number | undefined = undefined;
                if (x==null) {
                    x= tt.transform[4];
                    let textWidth= objFont.widthOfTextAtSize(text, fontSize);
                    textWidth= Math.min(textWidth, cellWidth);
                    let keyTextWidth= tt.width; //cellFont.widthOfTextAtSize(key, cellFontSize);
                    const hAlign = cellInfo.alignment.horizontal;
                    //console.log("align:",hAlign);
                    if (hAlign=="center" || hAlign=="centerContinuous") {
                        //x = (obj?.x ?? tt.transform[4])+ x + 0.5 * cellWidth
                        //x = x + cellWidth*0.5 - tt.width * 0.5;
                        x += keyTextWidth/2 - textWidth/2;
                    }
                    if (hAlign=="right")
                        x += keyTextWidth - textWidth;
                }

                try {
                    // pdfDoc.createContentStream(
                    //     drawLinesOfText([],{
                    //         x: x, //obj?.x ?? tt.transform[4], // x ??
                    //         y: obj?.y ?? tt.transform[5],
                    //         size: fontSize,
                    //         font: objFont.name,
                    //         lineHeight: 50, //tt.transform[0] * 1.15,
                    //     })
                    // )
                    pages[tt.pageIndex + iPage * pageCount] // i * length
                        .drawText(text, {
                            x: x, //obj?.x ?? tt.transform[4], // x ??
                            y: obj?.y ?? tt.transform[5],
                            size: fontSize,
                            font: objFont,
                            lineHeight: tt.transform[0] * 1.15,
                            maxWidth: cellWidth,
                        })
                } catch (e) {
                    throw (" drawText error for key=" +key+":  text: "+(text ?? obj?.text ?? "none") + " "
                        +"   i: "+             (tt.pageIndex + iPage * pageCount)
                        +"   x: "+             (x ?? obj?.x ?? tt?.transform[4])
                        +"   y: "+             (obj?.y ?? tt?.transform[5])
                        +"   size: "+          (    obj?.size ?? tt?.transform[0])
                        +"   font: "+          (    objFont)
                        +"   lineHeight: "+    (            tt?.transform[0] * 1.15)
                        +"   maxWidth: "+      (        cellWidth ?? excelKeys[key]?.width ?? 100))
                }
            } else if (typeof value == "object") {
                // тут код для вставки картинки
                try {
                    //if (value.name) {
                        const img = objImage2[value.name!]; //await ff(value.name)
                        pages[(value.pageIndex ?? (tt?.pageIndex ?? 0)) + iPage * pageCount ] //  i * length
                            .drawImage(img, {
                                    x: value.x ?? tt?.transform[4],
                                    y: value.y ?? tt?.transform[5],
                                    height: value.height,
                                    width: value.width
                                }
                            )
                    //} else console.warn("undefined image name for key: ",key);
                } catch (e) {
                    throw "drawImage error for key="+key+" : " + JSON.stringify(e)
                }

            }
        }
    }
    return pdfDocCopy
}

