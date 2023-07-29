import {PDFDocument, PDFImage} from "pdf-lib";
import * as fontkit from "@pdf-lib/fontkit";
import fs from "fs";

import {tCellInfo, tDataKey, tPFD} from "./inteface";

const fonts = {
    origin: (fs.readFileSync('./fonts/arial.ttf')),
    italic: (fs.readFileSync('./fonts/ariali.ttf')),
    bold: (fs.readFileSync('./fonts/arialbd.ttf')),
    boldItalic: (fs.readFileSync('./fonts/arialbi.ttf')),
}

export async function createPDF(_pdfSimple: Buffer, keyMap: {[key: string]: tPFD }, dataKey: tDataKey[], excelKey: tCellInfo) {

    const pdfSimple = await PDFDocument.load(_pdfSimple)
        .catch((e) => {
            throw "PDFDocument.load error"
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


    const pages = pdfDocCopy.getPages()

    // оптимизированная версия pngImage: PDFImage
    const objImage: { [key: string]: PDFImage } = {}
    const ff = async (name: string) => objImage[name] ??= await fs.promises.readFile(name).then(async e => await pdfDocCopy.embedPng(e))

    {
        const arr2: Promise<any>[] = []
        for (const data of dataKey)
            for (const value of Object.values(data))
                if (typeof value == "object") arr2.push(ff(value.name))
        await Promise.all(arr2)
    }

    for (let i = 0; i < dataKey.length; i++) {
        const data = dataKey[i]

        for (const [key, value] of Object.entries(data)) {
            const tt = keyMap[key]
            if (typeof value == "string") {
                if (!tt) continue;
                try {
                    pages[tt.pageIndex + i * length]
                        .drawText(value, {
                            x: tt.transform[4],
                            y: tt.transform[5],
                            size: tt.transform[0],
                            font: customFont[excelKey[key]?.font.style ?? "origin"],
                            lineHeight: tt.transform[0] * 1.15,
                            maxWidth: excelKey[key]?.width ?? 100,
                        })
                } catch (e) {
                    throw "drawText error " + JSON.stringify(e)
                }
            } else if (typeof value == "object" && value != null) {
                // тут код для вставки картинки
                try {
                    const img = objImage[value.name] //await ff(value.name)
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