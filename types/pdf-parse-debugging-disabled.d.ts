// Type definitions for pdf-parse 1.1
// Project: https://gitlab.com/autokent/pdf-parse
// Definitions by: Philipp Katz <https://github.com/qqilihq>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped

export = PdfParse;

declare function PdfParse(dataBuffer: Buffer, options?: PdfParse.Options): Promise<PdfParse.Result>;

declare namespace PdfParse {
    type Version = 'default' | 'v1.9.426' | 'v1.10.100' | 'v1.10.88' | 'v2.0.550';
    interface Result {
        numpages: number;
        numrender: number;
        info: any;
        metadata: any;
        version: Version;
        text: string;
    }

    interface RenderOptions {
        normalizeWhitespace?: boolean;
        disableCombineTextItems?: boolean;
    }

    type PageItem = {
        str: string,
        fontName: string,
        transform: [sizeFont: number, t1: number, t2: number, t3: number, x: number, y: number],
        width: number,
        height: number,
    }

    type TextContent = { items : PageItem[]; };

    interface PageData {
        pageIndex :number;
        pageInfo : { view: [x: number, y: number, w: number, h: number]; };
        getTextContent(render_options: RenderOptions) : Promise<TextContent>;
    }

    interface Options {
        pagerender?: ((pageData: PageData) => string|undefined|Promise<string|undefined>);
        max?: number | undefined;
        version?: Version | undefined;
    }
}
