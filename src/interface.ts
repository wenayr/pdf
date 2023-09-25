//import {PDFImage} from "pdf-lib";

export type tPFD = {
    transform: [fontSize: number, t1: number, t2: number, t3: number, x: number, y: number],
    pageIndex: number,
    pageView: [x: number, y: number, w: number, h: number],
    fontName: string,
    width: number,
    height: number
}

export type tCellInfo = {
    //left: number,
    rangeX: [number, number], // range in cells
    rangeY: [number, number], // range in cells
    font: { name: string, style: 'origin' | 'bold' | 'italic' | 'boldItalic', strikeThrough: boolean|undefined, color: string|undefined },
    alignment: {
        vertical: 'top' | 'bottom' | 'middle' | 'distributed' | 'justify',  // justify - по высоте, distributed - распределённый
        horizontal: 'left' | 'right' | 'center' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
    },
    width: number,
    height: number
}

export type tExcel = { [key: string]: tCellInfo[] };  // для каждого ключа может быть несколько ячеек

export type tPDFInfo = { [key: string]: tPFD[] };

export type tMapExcel = { [key: string]: tExcel }
export type tMapPDF = { [key: string]: Buffer }
export type tObjectImage = {
    text?: undefined, name: string, x?: number, y?: number, width?: number, height?: number, pageIndex?: number
}

export type tObjectString = {
    name?: undefined,
    text: string,
    x?: number,
    y?: number,
    width?: number,
    height?: number,
    pageIndex?: number
    size?: number,
    font?: "origin" | "bold" | "boldItalic" | "italic",
    maxWidth?: number,
}
export type tKeyData = {
    [key: string]: string | tObjectImage | tObjectString
}
export type tRequest = {
    // тия шаблона
    [nameTemplate: string]: tKeyData[]|tKeyData
}
export type tFonts = {
    origin: Buffer,
    italic: Buffer,
    bold: Buffer,
    boldItalic: Buffer
}
export type tObjImage = {
    [key: string]: Buffer
}