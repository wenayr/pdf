import childProcess from "child_process";

let _runTaskCache : { [name :string] : Promise<boolean>|undefined } = { };



export async function isProcessRunning(win :string, mac?: string, linux?: string) : Promise<boolean> {

    mac ??= win.substring(0,win.lastIndexOf("."));
    linux ??= win.substring(0,win.lastIndexOf("."));
    const plat = process.platform;
    const proc = plat=='win32' ? win : plat=='darwin' ? mac : plat=='linux' ? linux : undefined;
    const cmd =
        plat=='win32' ? `tasklist /FI "IMAGENAME eq ${proc}"`
        : plat=='darwin' ? 'ps -ax | grep '+proc
        : plat=='linux' ? 'ps -C '+proc // 'ps -A'
        : '';
    if(cmd=='' || !proc)
        return false;
    let task= _runTaskCache[proc];
    if (task) return task;

    return _runTaskCache[proc]= new Promise(function(resolve, reject) {
        // if (1) return false;
        childProcess.exec(cmd, function(err, stdout, stderr) {
            if (err) console.log("Error: ",stderr);
            delete _runTaskCache[proc];
            resolve(stdout.toLowerCase().indexOf(proc.toLowerCase()) > -1)
        });
        // let process= spawn(cmd, {shell: true});
        // process.stdout
        //     .on("data", (data)=> {
        //         console.timeLog(prefix,"spawn:",data.toString());
        //         //resolve(stdout.toLowerCase().indexOf(proc.toLowerCase()) > -1)
        //     });
        // process.on("exit",()=>console.timeLog(prefix, "exit tasks"))
        // process.on("close",()=>console.timeLog(prefix, "close tasks"))
    })
}



export async function waitForProcessRunAsync(name :string, step_ms=500, timeout_ms= 10000) {
    if (process.platform !="win32")
        name = name.substring(0,name.lastIndexOf("."));
    return new Promise((resolve,reject)=>{
        let startTime= Date.now();
        async function runCheck() {
            console.log(`check for running "${name}" ...`);
            if (await isProcessRunning(name)) {
                console.log(`is running: "${name}" `);
                resolve(true);
            }
            else
            if (Date.now() - startTime + step_ms < timeout_ms) {
                console.log(`is not running: "${name}"`);
                setTimeout(runCheck, step_ms);
            }
            else resolve(false); //reject();
        }
        runCheck();
    });
}


export async function execAndWaitForSpawn(command :string) {
    return new Promise((resolve, reject)=>{
        childProcess.spawn(command, { shell: true })
            .on("spawn", ()=>resolve(true))
            .on("error", ()=>reject("error"))
            .on("exit", ()=>reject("exit"))
            .on("close", ()=>reject("close"))
    });
}