import { checkSourceFile } from './input';
const {PDFNet} = require('@pdftron/pdfnet-node');
const path = require('path');

export const useSecuredMethod = process.argv[2] === '1';
export const hideZeroRow = process.argv[3] === '1';
export const showAmountColumn = process.argv[4] === '1';

checkSourceFile().then((res) => {
    if (res.result) {
        console.log(res.message);

        // const express = require('express');
        // const app = express();
        // const start = new Date().getTime();
        // PDFNet.runWithCleanup(convertToPdf).then(() => {
        //     const end = new Date().getTime();
        //     console.log(end - start);
        // }).catch((error) => {
        //     console.error(error);
        // });

        // const port = process.env.PORT ||  3000;
        // app.listen(port, console.log(`Server is running on port ${port}`));
    }
}).catch((res) => {
    console.log(res.message);
});


const convertToPdf = async () => {
    const inputPath = 'dist/document/khai.xlsx';
    const outputPath = 'dist/document/resume.pdf';
    const pdfdoc = await PDFNet.PDFDoc.create();
    await pdfdoc.initSecurityHandler();
    await PDFNet.Convert.toPdf(pdfdoc, inputPath);
    pdfdoc.save(outputPath, PDFNet.SDFDoc.SaveOptions.e_linearized);
}