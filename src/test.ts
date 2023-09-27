import {PORT} from "./index";
import {tKeyData, tRequest,tRequestAddTemplate} from "./interface";
import fetch from "node-fetch";
import * as addr from "./address";
//import * as fetchh from "node-fetch";
//const fetch= fetchh.default;

let _time= Date.now();

//let PORT= 4051;


export async function test() {

    addr.setLocal();

    let url= "http://localhost:" + PORT;

    const status = await fetch(url + "/s",)
        .then(response => response.json())
    console.log("status !! ", status);

    console.log("Отправка запроса на добавление шаблона");

    if (1) {
        const r = await fetch(url + "/addTemplateExcel", {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                //excel: "test4pgruz.xlsx",
                //excelSimple: "test4pgruzNo.xlsx",
                //excel: "4pgruz_key_.xlsx",
                excel: "test.xlsx",
                excelSimple: undefined, //"4pgruz.xlsx",
                name: "4p"
            } satisfies tRequestAddTemplate)
        })
            .then(response => response.json())
            .catch(e => {
                console.error("error ", e)
            })
        console.log("Ответ от сервера:", r);
    }
    // console.log("exit");  await fetch(url + "/stop"); if (1) return;

    const request: tRequest = {}

    console.log("Заполнение tRequest")

    for (let i = 0; i < 1; i++) {

        let obj: tKeyData = {};

        obj= {
            //key_A1: "КлючA1 жирный",
            key_Z2: "КлючZ2 жирный",
            key_Z3: "КлючZ3 жирный курсив",
            key_A1: "КлючА1 (какой-то текст) 11111111111 111 1 2 3 4 4 5 1 1 1 2 3 4 5  1 1 2 3 4 5 6 3 2 1 3 4 5 6 7 8 9 0 1 2 3 4 5 "
            // key_CZ4: "КлючCZ4 по центру",
            // key_BD1: "КлючBD1 по прав.краю",
            // key_D12: "КлючD12: Фамилия Имя Отчество"
        };

        /*
        obj["key_firm"]= "моя фирма";
        obj["key_periodInfo"] = "period11111111111 11111111111111111111111 11111111111111111111111111 111111111111111111111111111111 11111111111111111111111111111 111111111111111111111111 1111111111111111111111 " + String(i);
        obj["key_series"]= "123";
        obj["key_autoInfo"] = "auto222 222 222 22 22 22 22 " + String(i);

        obj["key_image"] = {
            width: 40,
            height: 40,
            name: "qr.png"
        };
        */
        obj["newImage"] = {
            width: 80,
            height: 80,
            x: 400,
            y: 400,
            name: "qr.png"
        };


        // for (let j = 0; j < 5; j++) {
        //     obj["newImage" + j] = {
        //         width: 80,
        //         height: 80,
        //         x: 400 + j,
        //         y: 400 + j,
        //         name: "qr.png"
        //     }
        // }

        // for (let j = 0; j < 1; j++) {
        //     obj[tempKey1 + j] = "test1111111 11111111111111111 111111111111111111111 11111111111111 11111111111111111111111111111 1111111111111111111 11111111111111111111 1111111111111111111111111111111 1111111 " + String(i);
        // }
        //(request["4p"] ??= []).push(obj);
        request["4p"] = obj;
    }

    console.log("Отправка запроса на получение PDF");

    const r2 = await fetch(url + "/dataToPDF", {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(request)
    })
        .then(response => response.json())
        .catch(e => {
            console.log("error ", e)
        })
    console.log("Ответ от сервера:",r2);

    const r3= await fetch(url + "/stop");

    console.log("stop: ",r3.ok, "   elapsed total ",((Date.now()-_time)/1000).toFixed(1)," s");

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