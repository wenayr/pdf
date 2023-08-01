import {PDFDocument, PDFImage} from "pdf-lib";
import * as fontkit from "@pdf-lib/fontkit";
import fs from "fs";

import {tCellInfo, tDataKey, tFonts, tObjectString, tObjImage, tPFD} from "./inteface";



export async function createPDF(_pdfSimple: Buffer, keyMap: {[key: string]: tPFD }, dataKey: tDataKey[], excelKey: tCellInfo, fonts: tFonts, objImage: tObjImage) {
    const pdfSimple = await PDFDocument.load(_pdfSimple)
        .catch((e) => {
            throw " PDFDocument.load error"
        })

    const length = pdfSimple.getPages().length
    const pdfDocCopy = await pdfSimple.copy()
    const arr: number[] = (new Array(length))
    for (let i = 0; i < length; i++)
        arr[i] = i;

    for (let arrElement of dataKey) {
        const data = await pdfDocCopy.copyPages(pdfSimple, arr)
        for (let i = 1; i < length; i++)
            pdfDocCopy.addPage(data[i])
    }

    pdfDocCopy.registerFontkit(fontkit);
    const customFont = ({
        origin: await pdfDocCopy.embedFont(fonts.origin),
        italic: await pdfDocCopy.embedFont(fonts.italic),
        bold: await pdfDocCopy.embedFont(fonts.bold),
        boldItalic: await pdfDocCopy.embedFont(fonts.boldItalic),
    })
    type kCustomFont = keyof typeof customFont


    const pages = pdfDocCopy.getPages()

    // оптимизированная версия pngImage: PDFImage
    const objImage2 : {[key: string]: PDFImage} = {}
    for (const [k, v] of Object.entries(objImage))
        objImage2[k] = await pdfDocCopy.embedPng(v)


    for (let i = 0; i < dataKey.length; i++) {
        const data = dataKey[i]

        for (const [key, value] of Object.entries(data)) {
            const tt = keyMap[key]
            if (typeof value == "string" || (typeof value == "object" && value.text)) {
                let text: string | undefined
                let obj: tObjectString | undefined

                if (typeof value == "object")  obj = value as tObjectString
                else text = value
                if (!tt && text) continue;
                const objFont = obj?.font && customFont[obj?.font]
                const horizontal = excelKey[key].alignment.horizontal
                if (horizontal == "center" || horizontal == "centerContinuous") {

                }
                try {
                    pages[tt.pageIndex + i * length]
                        .drawText(text ?? obj?.text ?? "none", {
                            x: obj?.x ?? tt.transform[4],
                            y: obj?.y ?? tt.transform[5],
                            size: obj?.size ?? tt.transform[0],
                            font: objFont ?? customFont[excelKey[key]?.font.style ?? "origin"],
                            lineHeight: tt.transform[0] * 1.15,
                            maxWidth: excelKey[key]?.width ?? 100,
                        })
                } catch (e) {
                    throw "drawText error " + JSON.stringify(e)
                }
            } else if (typeof value == "object" && value != null) {
                // тут код для вставки картинки
                try {
                    const img = objImage2[value.name] //await ff(value.name)
                    pages[(value.pageIndex ?? (tt?.pageIndex ?? 0)) + i * length]
                        .drawImage(img, {
                                x: value.x ?? tt?.transform[4],
                                y: value.y ?? tt?.transform[5],
                                height: value.height,
                                width: value.width
                            }
                        )
                } catch (e) {
                    throw "drawImage error " + JSON.stringify(e)
                }


            }
        }
    }
    return pdfDocCopy
}