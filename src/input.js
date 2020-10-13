import { detailsParse, summaryParse, exportProduct, productDoctorParse, exportPrSummary } from './parse';
import { sourcePath, doctorProductPath, productDoctorPath, productDate, doctorFullName, qty, productId, productName } from './constant';
import { sourceToJson } from './lib';
import { exportDoctor } from './doctor';

export const checkSourceFile = () => {
    return new Promise((resolve, reject) => {
        const fs = require('fs');
        fs.readdir(sourcePath, (error, files) => {
            const sourceFiles = files.filter(f => f !== '.DS_Store' && !f.startsWith('~'));
            const totalFiles = sourceFiles.length;
            if (totalFiles === 0) {
                reject({
                    message: `Ban chua copy file goc vao thu muc: ${sourcePath}`,
                    result: false
                });
            } else if (totalFiles > 1) {
                reject({
                    message: `Co qua nhieu file trong thu muc: ${sourcePath}`,
                    result: false
                });
            } else {
                // source
                const source = sourceToJson(`${sourcePath}/${sourceFiles[0]}`);
                // export product
                exportProduct(source);
                exportDoctor(source);

                // doctor -> product
                fs.readdir(doctorProductPath, (error, files) => {
                    const configFiles = files.filter(f => f !== '.DS_Store' && !f.startsWith('~'));
                    const totalFiles = configFiles.length;
                    const prSummary = [];
                    // pharmaceutical representatives
                    for (let file of configFiles) {
                        const prJson = pharmaceuticalRepresentativesToJson(`${doctorProductPath}/${file}`);
                        detailsParse(source, prJson);
                        const result = summaryParse(source, prJson);
                        prSummary.push(result);
                    }


                    // product -> doctor
                    fs.readdir(productDoctorPath, (error, files) => {
                        const configFiles = files.filter(f => f !== '.DS_Store' && !f.startsWith('~'));
                        const totalFiles = configFiles.length;
                        // pharmaceutical representatives
                        for (let file of configFiles) {
                            const prJson = productDoctorToJson(`${productDoctorPath}/${file}`);

                            const productDoctorList = [];
                            let prevProduct;
                            for (let row of prJson.data) {
                                if (row.ProductId || row.ProductName) {
                                    prevProduct = row;
                                    productDoctorList.push(prevProduct);
                                } else {
                                    if (!prevProduct.doctors) {
                                        prevProduct.doctors = [];
                                    }
                                    prevProduct.doctors.push(row);
                                }
                            }


                            productDoctorList.sort((a, b) => {
                                if (a.ProductName > b.ProductName) {
                                    return 1;
                                } else if (a.ProductName < b.ProductName) {
                                    return -1;
                                } else {
                                    return 0;
                                }
                            });

                            const result = productDoctorParse(source, productDoctorList,
                                prJson.pharmaceuticalRepresentatives
                            );

                            prSummary.push(result);
                        }

                        prSummary.sort((a, b) => {
                            if (a[0] > b[0]) {
                                return 1;
                            } else if (a[0] < b[0]) {
                                return -1;
                            } else {
                                return 0;
                            }
                        })
                        exportPrSummary(source[0][productDate], prSummary);
                    });



                    resolve({
                        result: true,
                        message: 'Success',
                    })
                });
            }

        });
    });
}



const pharmaceuticalRepresentativesToJson = (sourceFileName) => {
    const xlsx = require('xlsx');

    const workbook = xlsx.readFile(sourceFileName);
    if (workbook.SheetNames.length === 0) {
        return undefined;
    }

    const result = [];
    let pharmaceuticalRepresentatives;
    for (let sheetName in workbook.Sheets) {
        if (sheetName !== 'Info') {
            const json = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
            result.push({
                doctor: sheetName,
                data: json
            });
        } else {
            const json = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
            if (json && json.length > 0) {
                pharmaceuticalRepresentatives = {
                    fullName: json[0].FullName,
                    phoneNumber: json[0].PhoneNumber,
                    email: json[0].Email,
                };
            }

        }

    }

    return {
        pharmaceuticalRepresentatives,
        data: result
    };
}

const productDoctorToJson = (sourceFileName) => {
    const xlsx = require('xlsx');

    const workbook = xlsx.readFile(sourceFileName);
    if (workbook.SheetNames.length === 0) {
        return undefined;
    }
    let pharmaceuticalRepresentatives;

    const prJson = xlsx.utils.sheet_to_json(workbook.Sheets['Info']);
    if (prJson && prJson.length > 0) {
        pharmaceuticalRepresentatives = {
            fullName: prJson[0].FullName,
            phoneNumber: prJson[0].PhoneNumber,
            email: prJson[0].Email,
        };
    }
    return {
        pharmaceuticalRepresentatives,
        data: xlsx.utils.sheet_to_json(workbook.Sheets["Index"])
    };
}

// const getColumnHeaders = (worksheet) => {
//     const columnHeaders = [];
//     for (let key in worksheet) {
//         let regEx = new RegExp("^\(\\w\)\(1\){1}$");
//         if (regEx.test(key) == true) {
//             columnHeaders.push({ id: key, name: worksheet[key].v });
//         }
//     }

//     return columnHeaders;
// }

// const getColumnId = (columns, name) => {
//     const found = columns.find((it) => it.name === name);
//     if (found) {
//         return found.id[0];
//     } else {
//         return undefined;
//     }
// }
