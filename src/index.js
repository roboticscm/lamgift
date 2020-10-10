import { checkSourceFile } from './input';
import { parse } from './parse';

checkSourceFile().then((res) => {
    if (res.result) {
        // console.log(res.message);
    }
}).catch((res) => {
    console.log(res.message);
})