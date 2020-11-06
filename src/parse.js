import { prTitle, doctorTitle, priceListPath, sourcePath, doctorProductPath, destPath, pName, pPrice, pDiscount, pDoctorDiscount, productDate, productId, doctorFullName, productName, qty as quantity, pId, copyright } from './constant';
import { unaccentVietnamese, sourceToJson, saveExcelFile, borderExcel, formatCellHeader, formatReportTitle, formatReportFooter } from './lib';
import { from, of, zip } from 'rxjs';
import { groupBy, mergeMap, reduce, toArray, map } from 'rxjs/operators';
import { useSecuredMethod, hideZeroRow, showAmountColumn } from './index';

export const summaryParse = (source, prJson) => {
    const fs = require('fs');
    const excelFileName = unaccentVietnamese(prJson.pharmaceuticalRepresentatives.fullName);
    fs.mkdirSync(getDestPath(), { recursive: true });

    const sheetData = [];
    const detailsSheetData = [];
    sheetData.push([
        `${prJson.pharmaceuticalRepresentatives.fullName}`
    ]);

    sheetData.push(['BẢNG TỔNG HỢP']);
    sheetData.push(showAmountColumn ?  [undefined, undefined, undefined, undefined, undefined, undefined, source[0][productDate]] : [undefined, undefined, undefined, undefined, undefined, source[0][productDate]]);

    detailsSheetData.push([
        `${prJson.pharmaceuticalRepresentatives.fullName}`
    ]);
    detailsSheetData.push(['BẢNG CHI TIẾT']);
    detailsSheetData.push([undefined, undefined, undefined, undefined, source[0][productDate]]);
    const rows = [];
    const detailsRows = [];
    for (let doctor of prJson.data) {
        for (let product of doctor.data) {
            const qty = sumQty(source, getDoctorName(doctor.doctor), product[pName], product[pId]);
            if (qty || !hideZeroRow) {
                rows.push(
                    { name: `${product[pName]}(${product[pDiscount]})`, qty: qty ? qty : 0, price: product[pPrice], discount: product[pDiscount] }
                );
                detailsRows.push({
                    doctor: getDoctorName(doctor.doctor), name: product[pName], qty: qty ? qty : 0, price: product[pPrice], discount: product[pDiscount]
                });
            }
        }
    }

    detailsRows.sort((a, b) => {
        if (a.name > b.name) {
            return 1;
        } else if (a.name < b.name) {
            return -1;
        } else {
            return 0;
        }
    });
    const detailsGroupedRows$ = from(detailsRows).pipe(
        groupBy(row => row.name),
        mergeMap(group => zip(of(group.key), group.pipe(reduce((total, cur) => total + cur.qty, 0)), group.pipe(toArray())))
    );

    detailsSheetData.push([
        'STT', 'TÊN/BS', undefined, undefined, 'SL'
    ]);

    let detailsSumTotalQty = 0;
    let productRowCount = 0;
    detailsGroupedRows$.subscribe((z) => {
        detailsSheetData.push([
            ++productRowCount, z[0], undefined, undefined, z[1],
        ]);

        const doctors = z[2];
        doctors.sort(sortFullName);
        let doctorRowCount = 0;
        for (let doctor of doctors) {
            detailsSumTotalQty += doctor.qty;
            const { firstName: fName, lastName: lName } = getFirstAndLastName(doctor.doctor);
            detailsSheetData.push([
                undefined, ++doctorRowCount, lName, fName, doctor.qty
            ]);
        }

    });

    detailsSheetData.push([
        'TỔNG CỘNG', undefined, undefined, undefined, detailsSumTotalQty
    ]);

    rows.sort((a, b) => {
        if (a.name > b.name) {
            return 1;
        } else if (a.name < b.name) {
            return -1;
        } else {
            return 0;
        }
    });

    const groupedRows$ = from(rows).pipe(
        groupBy(row => row.name),
        mergeMap(group => group
            .pipe(
                reduce((total, cur) => {
                    total.qty += cur.qty;
                    total.price = cur.price;
                    total.discount = cur.discount;
                    return total;
                }, {
                    name: group.key, qty: 0, price: undefined, discount: undefined
                })
            )
        ),
        toArray()
    );

    sheetData.push(showAmountColumn ?  [
        'STT', 'TÊN', 'SL', 'ĐƠN GIÁ', 'TT', '%CK', 'CK'
    ] : [
        'STT', 'TÊN', 'SL', 'ĐƠN GIÁ', '%CK', 'CK'
    ]);

    let sumTotalQty = 0;
    let sumTotalAmount = 0;
    let sumTotalDiscountAmount = 0;
    groupedRows$.subscribe((rows) => {
        let rowNum = 0;
        for (let row of rows) {
            const amount = row.qty * row.price;
            const discountAmount = amount * row.discount / 100;
            sumTotalQty += row.qty;
            sumTotalAmount += amount;
            sumTotalDiscountAmount += discountAmount;
            sheetData.push(showAmountColumn ?  [
                ++rowNum, removeDiscount(row.name), row.qty, row.price, amount, row.discount, discountAmount
            ] : [
                ++rowNum, removeDiscount(row.name), row.qty, row.price, row.discount, discountAmount
            ]);
        }

    });

    sheetData.push(showAmountColumn ?  [
        'TỔNG', undefined, sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscountAmount
    ] : [
        'TỔNG', undefined, sumTotalQty, undefined, undefined, sumTotalDiscountAmount
    ]);


    saveExcelFile(`${getDestPath()}/${excelFileName}.xlsx`, sheetData, "Tổng Hợp", summaryParseFormat, false, detailsSheetData, "Chi Tiết", summaryParseFormat2);
    return [
        prJson.pharmaceuticalRepresentatives.fullName, sumTotalQty, sumTotalAmount, sumTotalDiscountAmount
    ];
}

const summaryParseFormat = (worksheet) => {
    worksheet.pageSetup.margins = {
        left: 0.2, right: 0.2,
        top: 0.45, bottom: 0.45,
        header: 0.2, footer: 0.2
    };
    worksheet.pageSetup.horizontalCentered = true;
    if(showAmountColumn) {
        worksheet.pageSetup.orientation = 'landscape';
    } 
   

    const headerRowNumber = 5;
    const allColumns = showAmountColumn ? ['A', 'B', 'C', 'D', 'E', 'F', 'G'] : ['A', 'B', 'C', 'D', 'E', 'F'];
    //format report title
    formatReportTitle(worksheet, 3, allColumns[0], allColumns[allColumns.length - 1]);

    // right align date
    worksheet.getCell(`${allColumns[allColumns.length - 1]}${headerRowNumber - 1}`).alignment = { horizontal: 'right', vertical: 'middle' };
    // center align column header
    allColumns.map(it => {
        formatCellHeader(worksheet, headerRowNumber, it);
    });

    // Report summary
    allColumns.map(it => {
        formatCellHeader(worksheet, worksheet.rowCount - 2, it, false);
    });
    worksheet.mergeCells(`A${worksheet.rowCount - 2}:B${worksheet.rowCount - 2}`);
    worksheet.getCell(`A${worksheet.rowCount - 2}`).style.alignment = { horizontal: 'center', vertical: 'middle' };

    // border content
    borderExcel(worksheet, headerRowNumber, worksheet.rowCount - 2, allColumns);

    // Format Number
    ['C', 'D', 'E', 'F', 'G'].map(it => {
        worksheet.getColumn(it).numFmt = '#,##0';
    });

    // set column width
    (showAmountColumn ? [5, 40, 12, 12, 20, 7, 20] : [7, 35, 10, 12, 8, 15]).map((it, index) => {
        worksheet.getColumn(index + 1).width = it;
    });
}

const summaryParseFormat2 = (worksheet) => {
    const headerRowNumber = 5;
    const allColumns = ['A', 'B', 'C', 'D', 'E'];
    //format report title
    formatReportTitle(worksheet, 3, allColumns[0], allColumns[allColumns.length - 1]);

    // right align date
    worksheet.getCell(`${allColumns[allColumns.length - 1]}${headerRowNumber - 1}`).alignment = { horizontal: 'right', vertical: 'middle' };
    // center align column header
    allColumns.map(it => {
        formatCellHeader(worksheet, headerRowNumber, it);
    });

    worksheet.eachRow((row, rowNumber) => {
        if (row.values[1] && rowNumber > headerRowNumber && rowNumber < worksheet.rowCount - 2) {
            worksheet.mergeCells(`B${rowNumber}:D${rowNumber}`);
            for (let i = 1; i <= allColumns.length; i++) {
                const cell = worksheet.getCell(`${allColumns[i - 1]}${rowNumber}`);
                cell.style.font = { ...cell.style.font, bold: true };
            }

        } else if (rowNumber === headerRowNumber) {
            worksheet.mergeCells(`B${rowNumber}:D${rowNumber}`);
        } else if (rowNumber === worksheet.rowCount - 2) {
            worksheet.mergeCells(`A${rowNumber}:D${rowNumber}`);
        }
    });

    // Report summary
    allColumns.map(it => {
        formatCellHeader(worksheet, worksheet.rowCount - 2, it, false);
    });
    // worksheet.mergeCells(`A${worksheet.rowCount - 2}:B${worksheet.rowCount - 2}`);
    worksheet.getCell(`A${worksheet.rowCount - 2}`).style.alignment = { horizontal: 'center', vertical: 'middle' };

    // border content
    borderExcel(worksheet, headerRowNumber, worksheet.rowCount - 2, allColumns);

    // Format Number
    worksheet.getColumn('E').numFmt = '#,##0';

    // set column width
    [5, 7, 35, 12, 15].map((it, index) => {
        worksheet.getColumn(index + 1).width = it;
    });
}

export const detailsParse = (source, prJson) => {
    const fs = require('fs');
    const excelFileName = unaccentVietnamese(prJson.pharmaceuticalRepresentatives.fullName);
    fs.mkdirSync(`${getDestPath()}/${excelFileName}`, { recursive: true });

    const sheetData = [];
    let sumTotalQty = 0;
    let sumTotalAmount = 0;
    let sumTotalDiscount = 0;
    let sumTotalDoctorDiscount = 0;

    let usedDoctorDiscount = true;
    sheetData.push([
        `${prJson.pharmaceuticalRepresentatives.fullName}`
    ]);

    sheetData.push([
        `BẢNG TÍNH TIỀN PHÍ`
    ]);

    for (let doctor of prJson.data) {
        const doctorSheetData = [];

        let count = 0;
        let totalQty = 0;
        let totalAmount = 0;
        let totalDiscountAmount = 0;
        let totalDoctorDiscountAmount = 0;
        let totalRemain = 0;

        let hasData = false;

        const rows = [];
        const doctorRows = [];

        for (let product of doctor.data) {
            let qty = sumQty(source, getDoctorName(doctor.doctor), product[pName], product[pId]);
            if (qty || !hideZeroRow) {
                if (qty === undefined) {
                    qty = 0;
                }
                const amount = qty * product[pPrice];
                const discountAmout = amount * product[pDiscount] / 100;

                let doctorDiscountAmount;
                let remain;
                if (!product[pDoctorDiscount]) {
                    usedDoctorDiscount = false;
                } else {
                    doctorDiscountAmount = amount * product[pDoctorDiscount] / 100;
                    remain = discountAmout - doctorDiscountAmount;
                }

                if (usedDoctorDiscount) {
                    rows.push([
                        ++count, getShortProductName(product[pName]), qty, product[pPrice], amount, product[pDiscount], discountAmout, product[pDoctorDiscount], doctorDiscountAmount, remain
                    ]);

                    doctorRows.push(showAmountColumn ? [
                        count, getShortProductName(product[pName], useSecuredMethod), qty, product[pPrice], amount, product[pDoctorDiscount], doctorDiscountAmount
                    ] : [
                        count, getShortProductName(product[pName], useSecuredMethod), qty, product[pPrice], product[pDoctorDiscount], doctorDiscountAmount
                    ]);
                } else {
                    rows.push([
                        ++count, getShortProductName(product[pName]), qty, product[pPrice], amount, product[pDiscount], discountAmout
                    ]);
                }

                totalQty += qty;
                totalAmount += amount;
                totalDiscountAmount += discountAmout;
                sumTotalQty += qty;
                sumTotalAmount += amount;
                sumTotalDiscount += discountAmout;

                if (usedDoctorDiscount) {
                    totalDoctorDiscountAmount += doctorDiscountAmount;
                    sumTotalDoctorDiscount += doctorDiscountAmount;
                    totalRemain += remain;
                }
                hasData = true;
            }

        }

        if (hasData) {
            sheetData.push([]);
            if (usedDoctorDiscount) {
                sheetData.push([
                    `${doctorTitle}: ${useSecuredMethod ? getDoctorNickname(doctor.doctor) : doctor.doctor}`, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, source[0][productDate]
                ]);
            } else {
                sheetData.push([
                    `${doctorTitle}: ${useSecuredMethod ? getDoctorNickname(doctor.doctor) : doctor.doctor}`, undefined, undefined, undefined, undefined, undefined, source[0][productDate]
                ]);
            }
            if (usedDoctorDiscount) {
                sheetData.push([
                    'STT', 'SP', 'SL', 'Đơn giá', 'Thành Tiền', '% CK', 'CK', '% CKBácS', 'CKBS', 'Còn lại'
                ]);
            } else {
                sheetData.push([
                    'STT', 'SP', 'SL', 'Đơn giá', 'Thành Tiền', '% CK', 'CK'
                ]);
            }
            sheetData.push(
                ...rows
            );


            // Doctor Summary
            if (usedDoctorDiscount) {
                sheetData.push([
                    'Tổng cộng', undefined, totalQty, undefined, totalAmount, undefined, totalDiscountAmount, undefined, totalDoctorDiscountAmount, totalRemain
                ]);
            } else {
                sheetData.push([
                    'Tổng cộng', undefined, totalQty, undefined, totalAmount, undefined, totalDiscountAmount
                ]);
            }

            if (usedDoctorDiscount) {
                doctorSheetData.push( showAmountColumn ? [
                    'STT', 'SP', 'SL', 'Đơn giá', 'Thành Tiền', '% CK', 'CK'
                ] : [
                    'STT', 'SP', 'SL', 'Đơn giá', '% CK', 'CK'
                ] );
            }
            doctorSheetData.push(
                ...doctorRows
            );


            if (usedDoctorDiscount) {
                doctorSheetData.unshift([
                    `${useSecuredMethod ? 'Kính gửi khách hàng' : 'Kính gửi ' + doctorTitle}: ${useSecuredMethod ? getDoctorNickname(doctor.doctor) : getDoctorName(doctor.doctor)}`
                ]);
            }

            doctorSheetData.unshift([
                source[0][productDate]
            ]);
            doctorSheetData.unshift([
                `${prJson.pharmaceuticalRepresentatives.fullName}`
            ]);

            // Doctor Summary
            if (usedDoctorDiscount) {
                doctorSheetData.push(showAmountColumn ? [
                    'Tổng cộng', undefined, totalQty, undefined, totalAmount, undefined, totalDoctorDiscountAmount
                ]: [
                    'Tổng cộng', undefined, totalQty, undefined, undefined, totalDoctorDiscountAmount
                ]);
            }

            // write doctor discount
            saveExcelFile(`${getDestPath()}/${excelFileName}/${unaccentVietnamese(getDoctorNickname(doctor.doctor))}_${getDoctorShortName(doctor.doctor)}.xlsx`, doctorSheetData, "Index", doctorParseFormat, true);
        }
    }

    // Report Summary
    if (usedDoctorDiscount) {
        sheetData.push([
            'Tổng cộng tất cả', undefined, sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscount, undefined, sumTotalDoctorDiscount, sumTotalDiscount - sumTotalDoctorDiscount
        ]);
    } else {
        sheetData.push([
            'Tổng cộng tất cả', undefined, sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscount
        ]);
    }


    saveExcelFile(`${getDestPath()}/${excelFileName}/index.xlsx`, sheetData, "Index", (worksheet) => detailsParseIndexFormat(worksheet, usedDoctorDiscount));
}

const detailsParseIndexFormat = (worksheet, usedDoctorDiscount) => {
    worksheet.pageSetup.margins = {
        left: 0.2, right: 0.2,
        top: 0.45, bottom: 0.45,
        header: 0.2, footer: 0.2
    };
    worksheet.pageSetup.horizontalCentered = true;
    worksheet.pageSetup.orientation = 'landscape';
    const headerRowNumber = 6;
    const allColumns = usedDoctorDiscount ? ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'] : ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    //format report title
    formatReportTitle(worksheet, 3, allColumns[0], allColumns[allColumns.length - 1]);


    // center align column header
    worksheet.eachRow((row, rowNumber) => {
        if (row.values[1] && `${row.values[1]}`.startsWith(`${doctorTitle}:`)) {
            // right align date
            worksheet.getCell(`${allColumns[allColumns.length - 1]}${rowNumber}`).alignment = { horizontal: 'right', vertical: 'middle' };
        }
        else if (row.values[1] === 'STT') {
            allColumns.map(it => {
                formatCellHeader(worksheet, rowNumber, it);
            });
        } else if (row.values[1] === 'Tổng cộng' || row.values[1] === 'Tổng cộng tất cả') {
            // Report summary
            allColumns.map(it => {
                formatCellHeader(worksheet, rowNumber, it, false);
            });
            worksheet.mergeCells(`A${rowNumber}:B${rowNumber}`);
            worksheet.getCell(`A${rowNumber}`).style.alignment = { horizontal: 'center', vertical: 'middle' };
        }
    });



    // border content
    borderExcel(worksheet, headerRowNumber, worksheet.rowCount - 2, allColumns);

    // Format Number
    ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].map(it => {
        worksheet.getColumn(it).numFmt = '#,##0';
    });

    // set column width
    [5, 15, 10, 10, 18, 7, 17, 7, 17, 15].map((it, index) => {
        worksheet.getColumn(index + 1).width = it;
    });
}

const doctorParseFormat = (worksheet) => {
    const headerRowNumber = 5;
    const allColumns = showAmountColumn ? ['A', 'B', 'C', 'D', 'E', 'F', 'G'] : ['A', 'B', 'C', 'D', 'E', 'F'];
    //format report title
    formatReportTitle(worksheet, 3, allColumns[0], allColumns[allColumns.length - 1]);

    // right align date
    worksheet.getCell(`${allColumns[allColumns.length - 1]}${headerRowNumber - 1}`).alignment = { horizontal: 'right', vertical: 'middle' };
    // center align column header
    allColumns.map(it => {
        formatCellHeader(worksheet, headerRowNumber, it);
    });

    // Report summary
    allColumns.map(it => {
        formatCellHeader(worksheet, worksheet.rowCount - 2, it, false);
    });
    worksheet.mergeCells(`A${worksheet.rowCount - 2}:B${worksheet.rowCount - 2}`);
    worksheet.getCell(`A${worksheet.rowCount - 2}`).style.alignment = { horizontal: 'center', vertical: 'middle' };

    // border content
    borderExcel(worksheet, headerRowNumber, worksheet.rowCount - 2, allColumns);

    // Format Number
    ['C', 'D', 'E', 'F', 'G'].map(it => {
        worksheet.getColumn(it).numFmt = '#,##0';
    });

    // set column width
    (showAmountColumn ? [5, 18, 10, 10, 18, 7, 17] : [5, 23, 12, 13, 10, 20]).map((it, index) => {
        worksheet.getColumn(index + 1).width = it;
    });
}

const sumQty = (source, doctor, name, id) => {
    const filterProducts = source.filter((it) => {
        if (it[doctorFullName] === doctor) {

            if (it[productId] && it[productId].trim().length > 0) {
                return it[productId].trim().toLowerCase() === id.toLowerCase();
            } else {
                return it[productName].toLowerCase().includes(name.toLowerCase());
            }
        } else {
            return false;
        }
    });

    if (filterProducts && filterProducts.length > 0) {
        return filterProducts.map(it => it[quantity]).reduce((x, y) => {
            return x + y;
        });
    } else {
        return undefined;
    }
}

const getCurrentMonthYear = () => {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    return { month, year };
}

export const getDestPath = () => {
    const { month, year } = getCurrentMonthYear();
    return `${destPath}/${year}/${month}`;
}

export const currencyFormat = (xlsx, worksheet, column) => {
    const col = xlsx.utils.decode_col(column); // 1
    const fmt = '#,##0';
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    for (let i = range.s.r + 1; i <= range.e.r; ++i) {
        const ref = xlsx.utils.encode_cell({ r: i, c: col });
        if (!worksheet[ref]) continue;
        if (worksheet[ref].t != 'n') continue;
        worksheet[ref].z = fmt;
    }
}

export const getDoctorNickname = (name) => {
    const s = name.split('-');

    if (s.length > 1) {
        return s[0].trim();
    } else {
        return getDoctorShortName(name);
    }
}

const getDoctorName = (name) => {
    const s = name.split('-');

    if (s.length > 1) {
        return s[1].trim();
    } else {
        return name;
    }
}

const sortFullName = (a, b, field = "doctor") => {
    const { firstName: firstName1, lastName: lastName1 } = getUnaccentcFirstAndLastName(a[field]);
    const { firstName: firstName2, lastName: lastName2 } = getUnaccentcFirstAndLastName(b[field]);

    if (firstName1 > firstName2) {
        return 1;
    } else if (firstName1 < firstName2) {
        return -1;
    } else {
        if (lastName1 > lastName2) {
            return 1;
        } else if (lastName1 < lastName2) {
            return -1;
        } else {
            return 0;
        }
    }
}

export const getUnaccentcFirstAndLastName = (name) => {
    if (!name) {
        return { lastName: undefined, firstName: undefined };
    }
    const unaccentName = unaccentVietnamese(name);
    const s = unaccentName.split(' ');
    if (s.length < 2) {
        return { lastName: undefined, firstName: unaccentName };
    }

    const firstName = s[s.length - 1].trim();
    s.splice(s.length - 1, 1);
    const lastName = s.join(' ').trim();

    return { lastName, firstName };
}

const getFirstAndLastName = (name) => {
    if (!name) {
        return { lastName: undefined, firstName: undefined };
    }
    const s = name.split(' ');
    if (s.length < 2) {
        return { lastName: undefined, firstName: name };
    }

    const firstName = s[s.length - 1].trim();
    s.splice(s.length - 1, 1);
    const lastName = s.join(' ').trim();

    return { lastName, firstName };
}

export const exportProduct = (source) => {
    const pId = 'Mã thuốc';
    const pName = 'Tên thuốc';
    const fs = require('fs');

    let priceList;
    fs.mkdirSync(`${getDestPath()}/products`, { recursive: true });
    fs.readdir(priceListPath, (error, files) => {
        const priceListFiles = files.filter(f => f !== '.DS_Store' && !f.startsWith('~'));
        const totalFiles = priceListFiles.length;
        if (totalFiles === 0) {
            console.log('Khong tim thay file danh muc hang hoa trong thu muc document -> price_list');
            return undefined;
        } else if (totalFiles > 1) {
            console.log('Co qua nhieu file danh muc hang hoa trong thu muc document -> price_list');
            return undefined;
        }
        else {
            priceList = sourceToJson(`${priceListPath}/${priceListFiles[0]}`);
            for (let product of priceList) {
                const list = source.filter((it) => {
                    if (it[productId] && product[pId] && it[productId].trim().toLowerCase() === product[pId].trim().toLowerCase()) {
                        return true;
                    } else {
                        if (it[productName] && product[pName] && it[productName].includes(product[pName])) {
                            return true;
                        } else {
                            return false;
                        }
                    }
                });
                if (list && list.length > 0) {
                    list.sort((a, b) => sortFullName(a, b, doctorFullName));

                    const sheetData = [];
                    let sumQty = 0;
                    let count = 0;
                    for (let row of list) {
                        const { lastName: lName, firstName: fName } = getFirstAndLastName(row[doctorFullName]);
                        sumQty += row[quantity];
                        sheetData.push([
                            ++count, row[productDate], lName, fName, row[quantity]
                        ]);
                    }

                    sheetData.unshift([
                        'STT', 'Ngày', 'Họ lót BS', 'Tên BS', 'SL'
                    ]);


                    sheetData.unshift([
                        undefined, undefined, undefined, undefined, source[0][productDate]
                    ]);

                    sheetData.unshift([
                        `BÁO CÁO TÌNH HÌNH MUA BÁN THUỐC ${getShortProductName(product[pName]).toUpperCase()}`
                    ]);

                    sheetData.unshift([]);
                    sheetData.push([
                        'Tổng cộng', undefined, undefined, undefined, sumQty
                    ]);

                    saveExcelFile(`${getDestPath()}/products/${getShortProductName(product[pName])}.xlsx`, sheetData, "Index", prProductFormat);
                }

            }
        }
    });
}

const prProductFormat = (worksheet) => {
    const headerRowNumber = 5;
    const allColumns = ['A', 'B', 'C', 'D', 'E'];
    //format report title
    formatReportTitle(worksheet, 3, allColumns[0], allColumns[allColumns.length - 1]);

    // right align date
    worksheet.getCell('E4').alignment = { horizontal: 'right', vertical: 'middle' };
    // center align column header
    allColumns.map(it => {
        formatCellHeader(worksheet, headerRowNumber, it);
    });

    // Report summary
    allColumns.map(it => {
        formatCellHeader(worksheet, worksheet.rowCount - 2, it, false);
    });

    worksheet.mergeCells(`A${worksheet.rowCount - 2}:D${worksheet.rowCount - 2}`);
    worksheet.getCell(`A${worksheet.rowCount - 2}`).style.alignment = { horizontal: 'center', vertical: 'middle' };

    // border content
    borderExcel(worksheet, headerRowNumber, worksheet.rowCount - 2, allColumns);

    // Format Number
    worksheet.getColumn('E').numFmt = '#,##0';

    // set column width
    [5, 30, 25, 11, 12].map((it, index) => {
        worksheet.getColumn(index + 1).width = it;
    });
}

export const getShortProductName = (name, useSecuredMethod = false) => {
    name = name.replace('/', '-');
    const shortName = name.replace(/\(.*?\)/, '').trim();
    if (useSecuredMethod) {
        return shortName.substring(0, 4) + '*'.repeat(shortName.length - 4);
    } else {
        return shortName;
    }
}


export const getDoctorShortName = (name) => {
    const sn = name.split('-');
    let fullName;
    if (sn.length > 1) {
        fullName = sn[1];
    } else {
        fullName = name;
    }

    const s = fullName.split(' ');
    const result = [];
    for (let word of s) {
        result.push(word[0]);
    }

    return result.join('').toLowerCase();
}



export const productDoctorParse = (source, prJson, pr) => {
    const fs = require('fs');
    const excelFileName = unaccentVietnamese(pr.fullName);
    fs.mkdirSync(getDestPath(), { recursive: true });

    const sheetData = [];
    const detailsSheetData = [];

    sheetData.push([
        `${pr.fullName}`
    ]);

    sheetData.push(['BẢNG TỔNG HỢP']);
    sheetData.push(showAmountColumn ? 
        [undefined, undefined, undefined, undefined, undefined, undefined, source[0][productDate]] : 
        [undefined, undefined, undefined, undefined, undefined, source[0][productDate]]
        );

    detailsSheetData.push([
        `${pr.fullName}`
    ]);
    detailsSheetData.push(['BẢNG CHI TIẾT']);
    detailsSheetData.push([undefined, undefined, undefined, undefined, source[0][productDate]]);

    const rows = [];
    const detailsRows = [];
    for (let product of prJson) {
        for (let doctor of product.doctors) {
            let qty = sumQty(source, doctor.DoctorName, product.ProductName, product.ProductId);
            if (qty || !hideZeroRow) {
                if (qty === undefined) {
                    qty = 0;
                }
                rows.push(
                    { name: product.ProductName, qty: qty != undefined ? qty : 0, price: product.Price, discount: product.Discount }
                );
                detailsRows.push({
                    doctor: doctor.DoctorName, name: product.ProductName, qty: qty != undefined ? qty : 0, price: product.Price, discount: product.Discount
                });

            }
        }
    }

    detailsRows.sort((a, b) => {
        if (a.name > b.name) {
            return 1;
        } else if (a.name < b.name) {
            return -1;
        } else {
            return 0;
        }
    });
    const detailsGroupedRows$ = from(detailsRows).pipe(
        groupBy(row => row.name),
        mergeMap(group => zip(of(group.key), group.pipe(reduce((total, cur) => total + cur.qty, 0)), group.pipe(toArray())))
    );

    detailsSheetData.push([
        'STT', 'TÊN/BS', undefined, undefined, 'SL'
    ]);

    let detailsSumTotalQty = 0;
    let rowCount = 0;
    detailsGroupedRows$.subscribe((z) => {
        detailsSheetData.push([
            ++rowCount, z[0], undefined, undefined, z[1],
        ]);

        const doctors = z[2];
        doctors.sort(sortFullName);
        let doctorRowCount = 0;
        for (let doctor of doctors) {
            detailsSumTotalQty += doctor.qty;
            const { firstName: fName, lastName: lName } = getFirstAndLastName(doctor.doctor);
            detailsSheetData.push([
                undefined, ++doctorRowCount, lName, fName, doctor.qty
            ]);
        }

    });

    detailsSheetData.push([
        'TỔNG CỘNG', undefined, undefined, undefined, detailsSumTotalQty
    ]);

    rows.sort((a, b) => {
        if (a.name > b.name) {
            return 1;
        } else if (a.name < b.name) {
            return -1;
        } else {
            return 0;
        }
    });

    const groupedRows$ = from(rows).pipe(
        groupBy(row => row.name),
        mergeMap(group => group
            .pipe(
                reduce((total, cur) => {
                    total.qty += cur.qty;
                    total.price = cur.price;
                    total.discount = cur.discount;
                    return total;
                }, {
                    name: group.key, qty: 0, price: undefined, discount: undefined
                })
            )
        ),
        toArray()
    );

    sheetData.push(
        showAmountColumn ? [
        'STT', 'TÊN', 'SL', 'ĐƠN GIÁ', 'TT', 'CK', 'TCK'
    ] : [
        'STT', 'TÊN', 'SL', 'ĐƠN GIÁ', 'CK', 'TCK'
    ]
    );

    let sumTotalQty = 0;
    let sumTotalAmount = 0;
    let sumTotalDiscountAmount = 0;
    groupedRows$.subscribe((rows) => {
        let rowCount = 0;
        for (let row of rows) {
            const amount = row.qty * row.price;
            const discountAmount = amount * row.discount / 100;
            sumTotalQty += row.qty;
            sumTotalAmount += amount;
            sumTotalDiscountAmount += discountAmount;
            sheetData.push(
                showAmountColumn ? [
                ++rowCount, row.name, row.qty, row.price, amount, row.discount, discountAmount
            ] : [
                ++rowCount, row.name, row.qty, row.price, row.discount, discountAmount
            ]);
        }

    });

    sheetData.push( 
        showAmountColumn ? [
        'TỔNG', undefined, sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscountAmount
    ] : [
        'TỔNG', undefined, sumTotalQty, undefined, undefined, sumTotalDiscountAmount
    ]);

    saveExcelFile(`${getDestPath()}/${excelFileName}.xlsx`, sheetData, "Tổng Hợp", summaryParseFormat, false, detailsSheetData, "Chi Tiết", summaryParseFormat2);

    return [
        pr.fullName, sumTotalQty, sumTotalAmount, sumTotalDiscountAmount
    ];
}


export const exportPrSummary = (date, data) => {
    let count = 1;
    let sumQty = 0;
    let sumAmount = 0;
    let sumDiscount = 0;
    data = data.map((it) => {
        sumQty += it[1];
        sumAmount += it[2];
        sumDiscount += it[3];

        it.unshift(count++);
        return it;
    });

    data.unshift([
        'STT', 'TD', 'SL', 'Doanh Số', 'CK', 'Ghi Chú'
    ]);


    data.unshift([
        undefined, undefined, undefined, undefined, undefined, `${date}`
    ]);

    data.unshift([
        'BẢNG TỔNG HỢP'
    ]);

    data.unshift([
    ]);

    data.push([
        'Tổng Cộng', undefined, sumQty, sumAmount, sumDiscount
    ]);

    const fs = require('fs');
    fs.mkdirSync(`${getDestPath()}`, { recursive: true });
    saveExcelFile(`${getDestPath()}/_summary.xlsx`, data, "Index", prSummaryFormat);
}

const prSummaryFormat = (worksheet) => {
    const headerRowNumber = 5;
    const allColumns = ['A', 'B', 'C', 'D', 'E', 'F'];
    //format report title
    formatReportTitle(worksheet, 3, allColumns[0], allColumns[allColumns.length - 1]);

    // right align date
    worksheet.getCell('F4').alignment = { horizontal: 'right', vertical: 'middle' };
    // center align column header
    allColumns.map(it => {
        formatCellHeader(worksheet, headerRowNumber, it);
    });

    // Report summary
    allColumns.map(it => {
        formatCellHeader(worksheet, worksheet.rowCount - 2, it, false);
    });
    worksheet.mergeCells(`A${worksheet.rowCount - 2}:B${worksheet.rowCount - 2}`);
    worksheet.getCell(`A${worksheet.rowCount - 2}`).style.alignment = { horizontal: 'center', vertical: 'middle' };
    for(let row = 6; row < worksheet.rowCount; row ++ ) {
        worksheet.getCell(`D${row}`).font = {
            color: { argb: 'FFAAAAAA' },
            italic: true
          };
    }

    // border content
    borderExcel(worksheet, headerRowNumber, worksheet.rowCount - 2, allColumns);

    // Format Number
    ['C', 'D', 'E'].map(it => {
        worksheet.getColumn(it).numFmt = '#,##0';
    });

    // set column width
    [5, 13, 12, 20, 20, 10].map((it, index) => {
        worksheet.getColumn(index + 1).width = it;
    });
}


const removeDiscount = (name) => {
    return name.replace(/\([0-9]{1,2}\)/, '').trim();
}