import {PORT} from "./index";
import {tDataKey, tRequest} from "./inteface";

export async function test() {
    const status = await fetch("http://localhost:" + PORT + "/s",)
        .then(response => response.json())
    console.log("status !! ", status)

    const r = await fetch("http://localhost:" + PORT + "/addTemplateExcel", {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            excel: "4pgruz.xlsx",
            excelSimple: "4pgruzNo.xlsx",
            name: "4g"
        } as { excel: string, name: string, excelSimple: string })
    })
        .then(response => response.json())
        .catch(e => {
            console.log("error ", e)
        })
    console.log("RRR", r);


    const datum: tRequest = {}
    const tempKey1 = "key_periodInfo"
    console.log("11")

    for (let i = 0; i < 1; i++) {
        const obj: tDataKey = {};
        (datum["4g"] ??= []).push(obj)
        obj[tempKey1] = "test11111111111 11111111111111111111111 11111111111111111111111111 111111111111111111111111111111 11111111111111111111111111111 111111111111111111111111 1111111111111111111111 " + String(i);

        obj["newImage"] = {
            width: 80,
            height: 80,
            x: 400,
            y: 400,
            name: "qr.png"
        }
        for (let j = 0; j < 500; j++) {
            obj["newImage" + j] = {
                width: 80,
                height: 80,
                x: 400 + j,
                y: 400 + j,
                name: "qr.png"
            }
        }

        for (let j = 0; j < 1; j++) {
            obj[tempKey1 + j] = "test1111111 11111111111111111 111111111111111111111 11111111111111 11111111111111111111111111111 1111111111111111111 11111111111111111111 1111111111111111111111111111111 1111111 " + String(i);
        }
    }

    console.log("12")
    const r2 = await fetch("http://localhost:" + PORT + "/dataToPDF", {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(datum)
    })
        .then(response => response.json())
        .catch(e => {
            console.log("error ", e)
        })
    console.log(r2);

    // const r2 = await fetch("http://localhost:" + PORT + "/addTemplateExcel", {
    //     method: 'POST',
    //     headers: {
    //         'Content-Type': 'application/json'
    //     },
    //     body: JSON.stringify({
    //         excel: "4pgruz.xlsx",
    //         excelSimple: "4pgruzNo.xlsx",
    //         name: "4g"
    //     } as {excel: string, name: string, excelSimple: string})
    // })
    //     .then(response => response.json())
    //     .catch(e=>{
    //         console.log("error ",e)
    //     })
    // console.log("RRR",r);


    // api.dataToPDF(data)
}