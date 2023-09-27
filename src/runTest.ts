/////<reference lib="dom" />
import process from "process";

import {start} from "./index"
import {test} from "./test";

await test();

process.exit(0);

