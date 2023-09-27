import {PDFDocument, PDFImage, PDFObject,PDFTextField, drawLinesOfText,PDFFont,lineSplit,cleanText,breakTextIntoLines} from "pdf-lib";
//import * as fontkit from "@pdf-lib/fontkit";
import fontkit from "@pdf-lib/fontkit";

import {tCellInfo, tKeyData,tExcel, tFonts, tObjectString, tObjImage, tPFD,tObjectImage} from "./interface";



export async function createPDF(
    pdfSimpleBuf: Buffer,
    keyMap: {[key: string]: tPFD[]},
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

    // console.log("!");
    // const fields = pdfDoc.getForm().getFields();
    // for(let field of fields) {
    //   const type = field.constructor.name;
    //   const name = field.getName();
    //   //if (field instanceof PDFTextField)
    //     console.log(`field:  ${type}: ${name}`);
    // }
    // console.log("!!");

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
        origin: await pdfDocCopy.embedFont(fonts.origin.buffer,{ subset: true }),
        italic: await pdfDocCopy.embedFont(fonts.italic.buffer,{ subset: true }),//await pdfDocCopy.embedFont(fonts.italic),
        bold: await pdfDocCopy.embedFont(fonts.bold.buffer,{ subset: true }),
        boldItalic:  await pdfDocCopy.embedFont(fonts.boldItalic.buffer,{ subset: true }),//await pdfDocCopy.embedFont(fonts.boldItalic),
    };//
    type kCustomFont = keyof typeof customFont


    const pages = pdfDocCopy.getPages()

    // оптимизированная версия pngImage: PDFImage
    const objImage2 : {[key: string]: PDFImage} = {}
    for (const [k, v] of Object.entries(objImage))
        objImage2[k] = await pdfDocCopy.embedPng(v)

    for (let iPage=0; iPage < keyDatas.length; iPage++) {
        const data = keyDatas[iPage]
        for (const [key, value] of Object.entries(data)) {

            if (value == null) continue;
            console.log(key, value);
            let isText = typeof value == "string" || (typeof value == "object" && value.text);

            let keyCells= excelKeys[key];
            if (!keyCells && isText) { console.warn("Wrong key: ",key);  continue; }
            if (!keyCells)
                if (typeof value == "object")
                    drawImage(value);

            for(let [i,cellInfo] of keyCells?.entries() ?? [])
            {
                const tt = keyMap[key][i];
                if (isText) {
                    let text: string | undefined
                    let obj: tObjectString | undefined

                    //let cellWidth = 100;

                    let cellWidth = cellInfo.width;// /2;
                    console.log("key",key+": ",cellInfo, tt);
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

                    let page= pages[tt.pageIndex + iPage * pageCount];

                    let x= obj?.x;// number | undefined = undefined;
                    let y= obj?.y;
                    let lineDelta= 0.15;
                    let linesInfo : { x: number, y :number, text :string }[] = [];

                    if (x!=null && y!=null)
                        linesInfo= [{ x, y, text}];
                    else { //if (x==null || y==null) {
                        let fontHeight= objFont.heightAtSize(fontSize);

                        const wordBreaks = page.doc.defaultWordBreaks; //options.wordBreaks || defaultWordBreaks;
                        const getTextWidth = (t: string) => objFont.widthOfTextAtSize(t, fontSize);
                        let maxWidth= cellWidth - fontHeight*0.2;
                        const lines =
                            maxWidth === undefined
                             ? lineSplit(cleanText(text))
                             : breakTextIntoLines(text, wordBreaks, maxWidth, getTextWidth);

                        //let textWidth=0;
                        //for(let line of lines) textWidth= Math.max(textWidth, getTextWidth(line));


                        let textHeight= fontHeight * lines.length + fontHeight * lineDelta * (lines.length-1);

                        let keyTextHeight= tt.height;
                        let y= tt.transform[5];
                        let vAlign= cellInfo.alignment.vertical;

                        if (vAlign=="distributed")
                            if (lines.length==1) vAlign="middle";
                            else vAlign= "justify";

                        if (vAlign=="justify") {
                            vAlign= "top";
                            //(cellInfo.height - lines.length * fontHeight) / (lines.length-1) / fontHeight;
                            if (lines.length>1) lineDelta= (cellInfo.height/fontHeight - lines.length) / (lines.length-1);
                            lineDelta= Math.max(lineDelta, 0.15);
                        }

                        if (vAlign=="middle") {
                            //x = (obj?.x ?? tt.transform[4])+ x + 0.5 * cellWidth
                            //x = x + cellWidth*0.5 - tt.width * 0.5;
                            y -= keyTextHeight/2 - textHeight/2;
                        }
                        if (vAlign=="bottom")
                            y -= keyTextHeight - textHeight;

                        for(let line of lines)
                        {
                            let x= tt.transform[4];

                            let textWidth= getTextWidth(line)
                            // let textWidth= objFont.widthOfTextAtSize(text, fontSize);
                            // let textHeight= objFont.heightAtSize(fontSize);
                            // if (textWidth>cellWidth) {
                            //     let rows= Math.ceil(textWidth / cellWidth);
                            //     textHeight= textHeight * rows + textHeight * 0.15 * (rows-1);
                            //     textWidth= cellWidth;
                            // }
                            //textWidth= Math.min(textWidth, cellWidth);
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

                            linesInfo.push({text: line, x, y});

                            y -= fontHeight * (1 + lineDelta);
                        }
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

                        //pages[tt.pageIndex + iPage * pageCount].setFontColor();

                        for(let line of linesInfo) {
                            page.drawText(line.text, {
                                x: line.x, //obj?.x ?? tt.transform[4], // x ??
                                y: line.y,
                                size: fontSize,
                                font: objFont,
                                lineHeight: fontSize * (1 + lineDelta), // 1.15
                                maxWidth: cellWidth,
                                //color:
                            })
                        }
                    } catch (e) {
                        throw (" drawText error for key=" +key+":  text: "+(text ?? obj?.text ?? "none") + " "
                            +"   i: "+             (tt.pageIndex + iPage * pageCount)
                            +"   x: "+             (x ?? obj?.x ?? tt?.transform[4])
                            +"   y: "+             (obj?.y ?? tt?.transform[5])
                            +"   size: "+          (    obj?.size ?? tt?.transform[0])
                            +"   font: "+          (    objFont)
                            +"   lineHeight: "+    (            tt?.transform[0] * 1.15)
                            +"   maxWidth: "+      (        cellWidth ?? keyCells[0]?.width ?? 100))
                    }
                } else if (typeof(value)=="object") {
                    // тут код для вставки картинки
                    drawImage(value, tt);
                }
            }
            function drawImage(value :tObjectImage|tObjectString, pfdInfo? :tPFD) {
                try {
                       console.log("image ",value);
                    //if (value.name) {
                        const img = objImage2[value.name!]; //await ff(value.name)
                        pages[(value.pageIndex ?? (pfdInfo?.pageIndex ?? 0)) + iPage * pageCount ] //  i * length
                            .drawImage(img, {
                                    x: value.x ?? pfdInfo?.transform[4],
                                    y: value.y ?? pfdInfo?.transform[5],
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
    return pdfDocCopy;
}

