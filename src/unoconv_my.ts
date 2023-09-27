
import childProcess , {ChildProcess} from 'child_process';
import mime from "mime";

import fs from "fs";
import * as os from "os";
import path from "path";
import * as console from "console"; import {waitForProcessRunAsync} from "./processHelper";



namespace unoconv {

    export function getLibreOfficePaths() : string[] {
        switch (process.platform) {
            case 'darwin': return ['/Applications/LibreOffice.app/Contents/MacOS'];
            case 'linux': return ['/usr/bin/libreoffice', '/usr/bin', '/snap/bin'];
            case 'win32': return [
                path.join(process.env["PROGRAMFILES"] ?? "", 'LibreOffice/program'),
                path.join(process.env['PROGRAMFILES(X86)'] ?? "", 'LIBREO~1/program'),
                path.join(process.env['PROGRAMFILES(X86)'] ?? "", 'LibreOffice/program'),
            ];
            default:
                throw new Error(`Operating system not yet supported: ${process.platform}`);
        }
    }

    const isWindows= process.platform=="win32"; //os.platform()=='win32';

    if (isWindows) {
        let paths = unoconv.getLibreOfficePaths();

        process.env["PATH"]+= ";"+ paths.join(";");
    }
    //console.log(process.env["PATH"]);


    const defaultRunCommand =  isWindows ? `python unoconv.py` : 'python3 unoconv.py'; //"c:/OSPanel/domains/pdf/src/unoconv.py`;//./unoconv.py`;

    export type Callback = (error: Error|null, output?: Buffer|undefined) => void;

    export type Options = Partial<{
        runCommand :string;  // path (command) for run unoconv
        port :number;
        server :string;
    }>;


    /**
    * Convert a document.
    */
    export function convert(fileOrBuffer :string|Buffer, outputFormat :string, options :Options, callback :Callback) : ChildProcess;
    export function convert(fileOrBuffer :string|Buffer, outputFormat :string, callback :Callback) : ChildProcess;

    export function convert(fileOrBuffer :string|Buffer, outputFormat :string, optionsOrCallback :Options|Callback, callback_? :Callback) {

        let stdout : Uint8Array[] = [];
        let stderr : Uint8Array[] = [];
        let options : Options;
        if (typeof optionsOrCallback=="function") {
            callback_ = optionsOrCallback;
            options = {};
        }
        else { options= optionsOrCallback;  callback_= callback_!; }
        const callback = callback_ satisfies Callback;

        let args = [
            '-f' + outputFormat,
            '--stdout',
            //'--stdin'
        ];

        if (options.port!=null) {
            args.push('-p' + options.port);
        }
        if (options.server!=null)
            args.push('-s' + options.server);

        let file : string;
        if (typeof(fileOrBuffer)=="object") {
            let tmpDir= os.tmpdir();
            let tmpFile= path.join(tmpDir, "unoconv_tempFile_"+Date.now());
            //console.log("Creating temp file: ",tmpFile);
            try {
                fs.writeFileSync(tmpFile, fileOrBuffer);
            } catch(e) { throw "failed to create temp file: "+tmpFile; }
            console.debug("Create temp file: ",tmpFile);
            file= tmpFile;
        }
        else file= fileOrBuffer;

        args.push(file);

        let bin = options.runCommand ?? defaultRunCommand;

        //let buf= fs.readFileSync(file);
        //console.log("! 1");
        let child = childProcess.spawn(bin, args, { shell: true /*, stdio: buf*/ });//, function (err, stdout, stderr) {
        //let child= childProcess.exec(bin+" "+args.join(" "));
        //child.stdin.write(buf);
        //console.log("! 2");
        child.stdout!.on('data', function (data) { //console.log("! stdout.data");
            stdout.push(data);
        });

        child.stderr!.on('data', function (data) { //console.log("! stderr.data",data.toString());
            stderr.push(data);
        });

        child.on('exit', function () { //console.log("! exit");

            if (stderr.length) {
                let str= Buffer.concat(stderr).toString();
                if (str.includes("DeprecationWarning")) {
                    console.warn(str);
                    if (str.includes("Error: Unable to connect or start own listener. Aborting.")) {
                        let optionsExt= { recurs: 0, ...options };
                        if (optionsExt.recurs++ < 2) {  // повторяем до двух раз
                            console.log("Repeat run exec");
                            return convert(fileOrBuffer, outputFormat, options, callback);
                        }
                    }
                }
                else {
                    //console.log("Error str on exit:", str);
                    return callback?.(new Error(str));
                }
            }

            callback?.(null, Buffer.concat(stdout));

            if (typeof(fileOrBuffer)=="object") {
                console.log("removing temp file");
                fs.unlinkSync(file);
                console.log("ok");
            }
        });
        child.on('error', (err) => { console.log("! error ",err);  callback?.(err); });
        return child;
    }

    let _conversionTask = Promise.resolve(Buffer.from([])); //: Promise<void>|undefined;

    export async function convertAsync(fileOrBuffer :string|Buffer, outputFormat :string, options? :Options) : Promise<Buffer> {
        await _startListenerTask;
        // ждём завершения прошлой задачи, т.к. многопоточные запросы невозможны (soffice выдаёт ошибку)
        return _conversionTask= _conversionTask.then( ()=>new Promise((resolve, reject)=>{
           //console.log("!!!!! in")
           convert(fileOrBuffer, outputFormat, options ?? {},
           (err, output)=> err ? reject(err) : resolve(output!)
        )}));
    }

    let _startListenerTask : Promise<void>|undefined;
    let _listener : ChildProcess|undefined;
    /**
    * Start a listener.
    */
    export function listen(options? : Pick<Options,"port"|"runCommand">) : ChildProcess {

        let args = [ '--listener' ];

        if (options?.port!=null) {
            args.push('-p' + options.port);
        }

        let bin= options?.runCommand ?? defaultRunCommand;
        //let python= options?.pythonPath ?? "python";
        //setTimeout.__promisify__(100).
        //
        // then(waitForProcessRunAsync('soffice.exe');
        let _resolve : ()=>void, _reject: (e :Error)=>void;
        let waitingTask= new Promise<void>((resolve, reject)=>{ _resolve= resolve; _reject=reject; });
        // после запуска листенера будем ждать появления процесса soffice
        if (1)
        _startListenerTask= (async()=>{
            await waitingTask;
            console.log("start waiting for 'soffice.bin'");
            let res= await waitForProcessRunAsync('soffice.bin', 1000, 10000);
            console.log("finish waiting for 'soffice.bin'.  Result="+res);
            _startListenerTask=undefined;
        })();

        console.log("Start unoconv listener");
        //_task= new Promise((resolve, reject)=>{ _resolve=resolve; _reject=reject; });
        //_task= new Promise())
        return _listener= childProcess.spawn(bin, args, { shell: true })
            .on("spawn", ()=>_resolve())
            .on("error", (e)=>_reject(new Error("listener error: "+e.message)))
    }
    // Остановка листенера (не работает при завершении из текущего потока)
    export function stopListener() { _listener?.kill("SIGINT");  _listener=undefined; } //_listener?.kill();


    export type Format = {
        format : string,
        extension: string,
        description: string,
        mime: string
    };

    export type Formats = {
        document: Format[];
        graphics: Format[];
        presentation: Format[];
        spreadsheet: Format[];
    }


    type OptionsShort = Pick<Options, "runCommand">;

    /**
    * Detect supported conversion formats.
    */
    export function detectSupportedFormats(options :OptionsShort,  callback : (error :Error|null, formats? :Formats)=>void) : void;
    export function detectSupportedFormats(callback : (error :Error|null, formats? :Formats)=>void) : void;
    export function detectSupportedFormats(optionsOrCallback :OptionsShort|Function,  callbackOrNull? : (error :Error|null, formats? :Formats)=>void) {
        return typeof optionsOrCallback=="function"
            ? _detectSupportedFormats({}, (optionsOrCallback as typeof callbackOrNull)!)
            : _detectSupportedFormats(optionsOrCallback, callbackOrNull!);
    }

    export async function detectSupportedFormatsPromise(options? :OptionsShort) : Promise<Formats> {
        return new Promise((resolve, reject)=>{
           detectSupportedFormats(options ?? {}, (err, output)=> err ? reject(err) : resolve(output!))
        });
    }

    function _detectSupportedFormats(options :OptionsShort,  callback : (error :Error|null, formats? :Formats)=>void) {

        let detectedFormats : Formats = {
            document: [],
            graphics: [],
            presentation: [],
            spreadsheet: []
        };

        let bin= options?.runCommand ?? defaultRunCommand;

        childProcess.execFile(bin, [ '--show' ], { shell: true }, function (err, stdout, stderr) {
            if (err) {
                return callback(err);
            }
            // For some reason --show outputs to stderr instead of stdout
            let lines = stderr.split('\n');

            let docType : keyof typeof detectedFormats | undefined;

            for(let line of lines) {
                if (line === 'The following list of document formats are currently available:') {
                    docType = 'document';
                } else if (line === 'The following list of graphics formats are currently available:') {
                    docType = 'graphics';
                } else if (line === 'The following list of presentation formats are currently available:') {
                    docType = 'presentation';
                } else if (line === 'The following list of spreadsheet formats are currently available:') {
                    docType = 'spreadsheet';
                } else {
                    let formatsMatch = line.match(/^(.*)-/);

                    let format= formatsMatch?.[1]?.trim();

                    let extensionMatch = line.match(/\[(.*)\]/);

                    let extension = extensionMatch?.[1]?.trim().replace('.', '');

                    let descriptioMatch = line.match(/-(.*)\[/);

                    let description = descriptioMatch?.[1]?.trim();

                    if (format && extension && description) {
                        if (!docType) { console.warn("docType is not defined!"); continue; }
                        detectedFormats[docType].push({
                            'format': format,
                            'extension': extension,
                            'description': description,
                            'mime': mime.lookup(extension)
                        });
                    }
                }
            }

            if (detectedFormats.document.length < 1 &&
                detectedFormats.graphics.length < 1 &&
                detectedFormats.presentation.length < 1 &&
                detectedFormats.spreadsheet.length < 1) {
                return callback(new Error('Unable to detect supported formats'));
            }

            callback(null, detectedFormats);
        });
    }

}

export default unoconv;