import { checkSourceFile } from './input';

export const useSecuredMethod = process.argv[2];
export const hideZeroRow = process.argv[3];

checkSourceFile().then((res) => {
    if (res.result) {
        console.log(res.message);
    }
}).catch((res) => {
    console.log(res.message);
})

