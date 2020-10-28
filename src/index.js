import { checkSourceFile } from './input';

export const useSecuredMethod = process.argv[2] === '1';
export const hideZeroRow = process.argv[3] === '1';
export const showAmountColumn = process.argv[4] === '1';

checkSourceFile().then((res) => {
    if (res.result) {
        console.log(res.message);
    }
}).catch((res) => {
    console.log(res.message);
})

