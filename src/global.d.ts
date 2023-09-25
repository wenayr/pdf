/////<reference lib="dom" />
//import type {blob} from "typescript/lib/lib.dom"
export {};
// Это делаем для того, чтобы не подключать dom (node-fetch требует эти типы)
declare global {
  type FormData = unknown; //import('formdata-node').FormData;
  type File = unknown;//import('formdata-node').File;

  var File: File;
  //var Blob: Blob;
  //const console : typeof console2.Console;
}

//type console= typeof console2;



//export console from "console";