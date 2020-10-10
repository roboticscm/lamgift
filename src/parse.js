import { priceListPath, sourcePath, doctorProductPath, destPath, pName, pPrice, pDiscount, pDoctorDiscount, productDate, productId, doctorFullName, productName, qty as quantity, pId } from './constant';
import { unaccentVietnamese, sourceToJson } from './lib';
import { from, of, zip } from 'rxjs';
import { groupBy, mergeMap, reduce, toArray, map } from 'rxjs/operators';

export const summaryParse = (source, prJson) => {
    const xlsx = require('xlsx');

    const workbook = xlsx.utils.book_new();

    workbook.Props = {
        Title: 'LamGift',
        Author: 'Ly Van Khai 0986 409 026',
        Subject: 'Excel Generator',
    }
    workbook.SheetNames.push('Summary');
    workbook.SheetNames.push('Details');

    const fs = require('fs');
    const excelFileName = unaccentVietnamese(prJson.pharmaceuticalRepresentatives.fullName);
    fs.mkdirSync(getDestPath(), { recursive: true });
    // fs.mkdirSync(`${getDestPath()}/${excelFileName}`, { recursive: true });

    const sheetData = [];
    const detailsSheetData = [];
    const merge = [];
    const detailsMerge = [];
    const { month, year } = getCurrentMonthYear();
    sheetData.push([
        `Tháng: ${month}/${year}`
    ]);
    sheetData.push([
        `Trình dược viên: ${prJson.pharmaceuticalRepresentatives.fullName} - ${prJson.pharmaceuticalRepresentatives.phoneNumber || '<Điện thoại>'}  - ${prJson.pharmaceuticalRepresentatives.email || '<Email>'}`
    ]);

    sheetData.push([]);

    detailsSheetData.push([
        `Tháng: ${month}/${year}`
    ]);
    detailsSheetData.push([
        `Trình dược viên: ${prJson.pharmaceuticalRepresentatives.fullName} - ${prJson.pharmaceuticalRepresentatives.phoneNumber || '<Điện thoại>'} - ${prJson.pharmaceuticalRepresentatives.email || '<Email>'}`
    ]);

    detailsSheetData.push([]);
    const rows = [];
    const detailsRows = [];
    for (let doctor of prJson.data) {
        for (let product of doctor.data) {
            const qty = sumQty(source, getDoctorName(doctor.doctor), product[pName], product[pId]);
            if (qty) {
                rows.push(
                    { name: product[pName], qty, price: product[pPrice], discount: product[pDiscount] }
                );
                detailsRows.push({
                    doctor: getDoctorName(doctor.doctor), name: product[pName], qty, price: product[pPrice], discount: product[pDiscount]
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
        'TÊN THUỐC/BS', undefined, undefined, 'SL'
    ]);
    detailsMerge.push(
        { s: { r: detailsSheetData.length - 1, c: 0 }, e: { r: detailsSheetData.length - 1, c: 2 } }
    );

    let detailsSumTotalQty = 0;
    detailsGroupedRows$.subscribe((z) => {
        detailsSheetData.push([
            z[0], undefined, undefined, z[1],
        ]);
        detailsMerge.push(
            { s: { r: detailsSheetData.length - 1, c: 0 }, e: { r: detailsSheetData.length - 1, c: 2 } }
        );

        const doctors = z[2];
        doctors.sort(sortFullName);
        for (let doctor of doctors) {
            detailsSumTotalQty += doctor.qty;
            const { firstName: fName, lastName: lName } = getFirstAndLastName(doctor.doctor);
            detailsSheetData.push([
                undefined, lName, fName, doctor.qty
            ]);
        }

    });

    detailsSheetData.push([
        'TỔNG CỘNG', undefined, undefined, detailsSumTotalQty
    ]);
    detailsMerge.push(
        { s: { r: detailsSheetData.length - 1, c: 0 }, e: { r: detailsSheetData.length - 1, c: 2 } }
    );

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

    sheetData.push([
        'TÊN', 'SL', 'ĐƠN GIÁ', 'TT', 'CK', 'TCK'
    ]);

    let sumTotalQty = 0;
    let sumTotalAmount = 0;
    let sumTotalDiscountAmount = 0;
    groupedRows$.subscribe((rows) => {
        for (let row of rows) {
            const amount = row.qty * row.price;
            const discountAmount = amount * row.discount / 100;
            sumTotalQty += row.qty;
            sumTotalAmount += amount;
            sumTotalDiscountAmount += discountAmount;
            sheetData.push([
                row.name, row.qty, row.price, amount, row.discount, discountAmount
            ]);
        }

    });

    sheetData.push([
        'TỔNG', sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscountAmount
    ]);

    sheetData.push([]);

    // Report Summary
    sheetData.push([]);
    sheetData.push(['File này được sinh ra bởi LamGift']);
    sheetData.push(['Powered by Lý Văn Khải - 0986 409 026 - roboticscm2018@gmail.com']);
    const worksheet = xlsx.utils.aoa_to_sheet(sheetData);

    detailsSheetData.push([]);
    detailsSheetData.push(['File này được sinh ra bởi LamGift']);
    detailsSheetData.push(['Powered by Lý Văn Khải - 0986 409 026 - roboticscm2018@gmail.com']);
    const detailsWorksheet = xlsx.utils.aoa_to_sheet(detailsSheetData);

    currencyFormat(xlsx, worksheet, 'B');
    currencyFormat(xlsx, worksheet, 'C');
    currencyFormat(xlsx, worksheet, 'D');
    currencyFormat(xlsx, worksheet, 'E');
    currencyFormat(xlsx, worksheet, 'F');

    currencyFormat(xlsx, detailsWorksheet, 'D');

    worksheet["!merges"] = merge;
    detailsWorksheet["!merges"] = detailsMerge;

    const wscols = [
        { wch: 30 },
        { wch: 7 },
        { wch: 7 },
        { wch: 15 },
        { wch: 5 },
        { wch: 15 },
    ];

    worksheet['!cols'] = wscols;

    const dwscols = [
        { wch: 5 },
        { wch: 20 },
        { wch: 6 },
        { wch: 10 },
    ];
    detailsWorksheet['!cols'] = dwscols;

    workbook.Sheets['Summary'] = worksheet;
    workbook.Sheets['Details'] = detailsWorksheet;
    xlsx.writeFile(workbook, `${getDestPath()}/${excelFileName}.xlsx`);
}

export const detailsParse = (source, prJson) => {
    const xlsx = require('xlsx');

    const workbook = xlsx.utils.book_new();
    workbook.Props = {
        Title: 'LamGift',
        Author: 'Ly Van Khai 0986 409 026',
        Subject: 'Excel Generator',
    }

    workbook.SheetNames.push('Index');

    const fs = require('fs');
    const excelFileName = unaccentVietnamese(prJson.pharmaceuticalRepresentatives.fullName);
    fs.mkdirSync(getDestPath(), { recursive: true });
    fs.mkdirSync(`${getDestPath()}/${excelFileName}`, { recursive: true });

    const sheetData = [];
    const merge = [];

    sheetData.push([
        `Trình dược viên: ${prJson.pharmaceuticalRepresentatives.fullName} - ${prJson.pharmaceuticalRepresentatives.phoneNumber || '<Điện Thoại>'} - ${prJson.pharmaceuticalRepresentatives.email || '<Email>'}`
    ]);



    let sumTotalQty = 0;
    let sumTotalAmount = 0;
    let sumTotalDiscount = 0;
    let sumTotalDoctorDiscount = 0;

    let usedDoctorDiscount = true;
    for (let doctor of prJson.data) {
        const doctorMerge = [];
        const doctorSheetData = [];
        doctorSheetData.push([
            `Trình dược viên: ${prJson.pharmaceuticalRepresentatives.fullName} - ${prJson.pharmaceuticalRepresentatives.phoneNumber || '<Điện Thoại>'} - ${prJson.pharmaceuticalRepresentatives.email || '<Email>'}`
        ]);

        let count = 1;
        let totalQty = 0;
        let totalAmount = 0;
        let totalDiscountAmount = 0;
        let totalDoctorDiscountAmount = 0;
        let totalRemain = 0;

        let hasData = false;

        const rows = [];
        const doctorRows = [];

        for (let product of doctor.data) {
            const qty = sumQty(source, getDoctorName(doctor.doctor), product[pName], product[pId]);
            if (qty) {
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
                        count, product[pName], qty, product[pPrice], amount, product[pDiscount], discountAmout, product[pDoctorDiscount], doctorDiscountAmount, remain
                    ]);

                    doctorRows.push([
                        count, product[pName], qty, product[pPrice], amount, product[pDoctorDiscount], doctorDiscountAmount
                    ]);
                } else {
                    rows.push([
                        count, product[pName], qty, product[pPrice], amount, product[pDiscount], discountAmout
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
                count++;
                hasData = true;
            }

        }

        if (hasData) {
            sheetData.push([]);
            sheetData.push([
                `BS: ${getDoctorName(doctor.doctor)}`
            ]);
            if (usedDoctorDiscount) {
                sheetData.push([
                    'STT', 'Sản phẩm', 'SL', 'Đơn giá', 'Thành Tiền', '% Chiết khấu', 'Chiết Khấu', '% CKBS', 'CKBS', 'Còn lại'
                ]);
            } else {
                sheetData.push([
                    'STT', 'Sản phẩm', 'SL', 'Đơn giá', 'Thành Tiền', '% Chiết khấu', 'Chiết Khấu'
                ]);
            }
            sheetData.push(
                ...rows
            );

            const lastRow = sheetData.length;
            merge.push(
                { s: { r: lastRow, c: 0 }, e: { r: lastRow, c: 1 } }
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


            // for doctor report
            doctorSheetData.push([]);
            doctorSheetData.push([
                `BS: ${getDoctorNickname(doctor.doctor)}`
            ]);
            if (usedDoctorDiscount) {
                doctorSheetData.push([
                    'STT', 'Sản phẩm', 'SL', 'Đơn giá', 'Thành Tiền', '% CK', 'Chiết Khấu'
                ]);
            }
            doctorSheetData.push(
                ...doctorRows
            );

            doctorMerge.push(
                { s: { r: doctorSheetData.length, c: 0 }, e: { r: doctorSheetData.length, c: 1 } }
            );
            // Doctor Summary
            if (usedDoctorDiscount) {
                doctorSheetData.push([
                    'Tổng cộng', undefined, totalQty, undefined, totalAmount, undefined, totalDoctorDiscountAmount
                ]);
            }

            const doctorWorksheet = xlsx.utils.aoa_to_sheet(doctorSheetData);
            currencyFormat(xlsx, doctorWorksheet, 'C');
            currencyFormat(xlsx, doctorWorksheet, 'D');
            currencyFormat(xlsx, doctorWorksheet, 'E');
            currencyFormat(xlsx, doctorWorksheet, 'F');
            currencyFormat(xlsx, doctorWorksheet, 'G');

            doctorWorksheet["!merges"] = doctorMerge;

            const wscols = [
                { wch: 6 },
                { wch: 30 },
                { wch: 6 },
                { wch: 7 },
                { wch: 15 },
                { wch: 6 },
                { wch: 12 },
            ];

            doctorWorksheet['!cols'] = wscols;

            const doctorWorkbook = xlsx.utils.book_new();
            doctorWorkbook.Props = {
                Title: 'LamGift',
                Author: 'Ly Van Khai 0986 409 026',
                Subject: 'Excel Generator',
            }

            doctorWorkbook.SheetNames.push('Index');

            doctorWorkbook.Sheets['Index'] = doctorWorksheet;

            // write doctor discount
            xlsx.writeFile(doctorWorkbook, `${getDestPath()}/${excelFileName}/${unaccentVietnamese(getDoctorNickname(doctor.doctor))}_${getDoctorShortName(doctor.doctor)}.xlsx`);
        }
    }

    sheetData.push([]);

    // Report Summary
    if (usedDoctorDiscount) {
        sheetData.push([
            'Tổng cộng', undefined, sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscount, undefined, sumTotalDoctorDiscount, sumTotalDiscount - sumTotalDoctorDiscount
        ]);
    } else {
        sheetData.push([
            'Tổng cộng', undefined, sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscount
        ]);
    }
    // if (usedDoctorDiscount) {
    //     sheetData.push([
    //         'Tổng tiền chiết khấu còn lại của Trình dược', undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, sumTotalRemain
    //     ]);

    //     const lastRow = sheetData.length - 1;
    //     merge.push(
    //         { s: { r: lastRow, c: 0 }, e: { r: lastRow, c: 8 } }
    //     );
    // }
    sheetData.push([]);
    sheetData.push(['File này được sinh ra bởi LamGift']);
    sheetData.push(['Powered by Lý Văn Khải - 0986 409 026 - roboticscm2018@gmail.com']);
    const worksheet = xlsx.utils.aoa_to_sheet(sheetData);


    currencyFormat(xlsx, worksheet, 'C');
    currencyFormat(xlsx, worksheet, 'D');
    currencyFormat(xlsx, worksheet, 'E');
    currencyFormat(xlsx, worksheet, 'F');
    currencyFormat(xlsx, worksheet, 'G');
    currencyFormat(xlsx, worksheet, 'H');
    currencyFormat(xlsx, worksheet, 'I');
    currencyFormat(xlsx, worksheet, 'J');


    worksheet["!merges"] = merge;

    const wscols = [
        { wch: 6 },
        { wch: 30 },
        { wch: 10 },
        { wch: 12 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
    ];

    // worksheet["A5"].s = {
    //     font: { sz: 23, bold: true, },
    //     alignment: {
    //         horizontal: "center", vertical: "center", wrap_text: true
    //     }
    // };

    worksheet['!cols'] = wscols;

    workbook.Sheets['Index'] = worksheet;
    xlsx.writeFile(workbook, `${getDestPath()}/${excelFileName}/index.xlsx`);
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

const getDestPath = () => {
    const { month, year } = getCurrentMonthYear();
    return `${destPath}/${year}/${month}`;
}

const currencyFormat = (xlsx, worksheet, column) => {
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

const getDoctorNickname = (name) => {
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

const getUnaccentcFirstAndLastName = (name) => {
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

    const xlsx = require('xlsx');
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

                    const workbook = xlsx.utils.book_new();

                    workbook.Props = {
                        Title: 'LamGift',
                        Author: 'Ly Van Khai 0986 409 026',
                        Subject: 'Excel Generator',
                    }
                    workbook.SheetNames.push('Index');
                    const sheetData = [];

                    sheetData.push([
                        'Ngày', 'Họ lót BS', 'Tên BS', 'Thuốc', 'SL'
                    ]);
                    let sumQty = 0;
                    for (let row of list) {
                        const { lastName: lName, firstName: fName } = getFirstAndLastName(row[doctorFullName]);
                        sumQty += row[quantity];
                        sheetData.push([
                            row[productDate], lName, fName, row[productName], row[quantity]
                        ]);
                    }

                    sheetData.push([
                        'Tổng cộng', undefined, undefined, undefined, sumQty
                    ]);
                    const worksheet = xlsx.utils.aoa_to_sheet(sheetData);
                    const wscols = [
                        { wch: 20 },
                        { wch: 20 },
                        { wch: 10 },
                        { wch: 35 },
                        { wch: 10 },
                    ];
                    worksheet['!cols'] = wscols;

                    const merge = [];
                    merge.push(
                        { s: { r: sheetData.length - 1, c: 0 }, e: { r: sheetData.length - 1, c: 3 } }
                    );
                    currencyFormat(xlsx, worksheet, 'E');
                    worksheet["!merges"] = merge;


                    workbook.Sheets['Index'] = worksheet;
                    xlsx.writeFile(workbook, `${getDestPath()}/products/${getShortProductName(product[pName])}.xlsx`);
                }

            }
        }
    });
}


export const getShortProductName = (name) => {
    return name.replace(/\(.*?\)/, '').trim();
}


const getDoctorShortName = (name) => {
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
    const xlsx = require('xlsx');

    const workbook = xlsx.utils.book_new();

    workbook.Props = {
        Title: 'LamGift',
        Author: 'Ly Van Khai 0986 409 026',
        Subject: 'Excel Generator',
    }
    workbook.SheetNames.push('Summary');
    workbook.SheetNames.push('Details');

    const fs = require('fs');
    const excelFileName = unaccentVietnamese(pr.fullName);
    fs.mkdirSync(getDestPath(), { recursive: true });

    const sheetData = [];
    const detailsSheetData = [];
    const merge = [];
    const detailsMerge = [];
    const { month, year } = getCurrentMonthYear();
    sheetData.push([
        `Tháng: ${month}/${year}`
    ]);
    sheetData.push([
        `Trình dược viên: ${pr.fullName} - ${pr.phoneNumber || '<Điện thoại>'} - ${pr.email || '<Email>'}`
    ]);

    sheetData.push([]);

    detailsSheetData.push([
        `Tháng: ${month}/${year}`
    ]);
    detailsSheetData.push([
        `Trình dược viên: ${pr.fullName} - ${pr.phoneNumber || '<Điện thoại>'} - ${pr.email || '<Email>'}`
    ]);

    detailsSheetData.push([]);
    const rows = [];
    const detailsRows = [];
    for (let product of prJson) {
        for (let doctor of product.doctors) {
            const qty = sumQty(source, doctor.DoctorName, product.ProductName, product.ProductId);
            if (qty) {
                rows.push(
                    { name: product.ProductName, qty, price: product.Price, discount: product.Discount }
                );
                detailsRows.push({
                    doctor: doctor.DoctorName, name: product.ProductName, qty, price: product.Price, discount: product.Discount
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
        'TÊN THUỐC/BS', undefined, undefined, 'SL'
    ]);
    detailsMerge.push(
        { s: { r: detailsSheetData.length - 1, c: 0 }, e: { r: detailsSheetData.length - 1, c: 2 } }
    );

    let detailsSumTotalQty = 0;
    detailsGroupedRows$.subscribe((z) => {
        detailsSheetData.push([
            z[0], undefined, undefined, z[1],
        ]);
        detailsMerge.push(
            { s: { r: detailsSheetData.length - 1, c: 0 }, e: { r: detailsSheetData.length - 1, c: 2 } }
        );

        const doctors = z[2];
        doctors.sort(sortFullName);
        for (let doctor of doctors) {
            detailsSumTotalQty += doctor.qty;
            const { firstName: fName, lastName: lName } = getFirstAndLastName(doctor.doctor);
            detailsSheetData.push([
                undefined, lName, fName, doctor.qty
            ]);
        }

    });

    detailsSheetData.push([
        'TỔNG CỘNG', undefined, undefined, detailsSumTotalQty
    ]);
    detailsMerge.push(
        { s: { r: detailsSheetData.length - 1, c: 0 }, e: { r: detailsSheetData.length - 1, c: 2 } }
    );

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

    sheetData.push([
        'TÊN', 'SL', 'ĐƠN GIÁ', 'TT', 'CK', 'TCK'
    ]);

    let sumTotalQty = 0;
    let sumTotalAmount = 0;
    let sumTotalDiscountAmount = 0;
    groupedRows$.subscribe((rows) => {
        for (let row of rows) {
            const amount = row.qty * row.price;
            const discountAmount = amount * row.discount / 100;
            sumTotalQty += row.qty;
            sumTotalAmount += amount;
            sumTotalDiscountAmount += discountAmount;
            sheetData.push([
                row.name, row.qty, row.price, amount, row.discount, discountAmount
            ]);
        }

    });

    sheetData.push([
        'TỔNG', sumTotalQty, undefined, sumTotalAmount, undefined, sumTotalDiscountAmount
    ]);

    sheetData.push([]);

    // Report Summary
    sheetData.push([]);
    sheetData.push(['File này được sinh ra bởi LamGift']);
    sheetData.push(['Powered by Lý Văn Khải - 0986 409 026 - roboticscm2018@gmail.com']);
    const worksheet = xlsx.utils.aoa_to_sheet(sheetData);

    detailsSheetData.push([]);
    detailsSheetData.push(['File này được sinh ra bởi LamGift']);
    detailsSheetData.push(['Powered by Lý Văn Khải - 0986 409 026 - roboticscm2018@gmail.com']);
    const detailsWorksheet = xlsx.utils.aoa_to_sheet(detailsSheetData);

    currencyFormat(xlsx, worksheet, 'B');
    currencyFormat(xlsx, worksheet, 'C');
    currencyFormat(xlsx, worksheet, 'D');
    currencyFormat(xlsx, worksheet, 'E');
    currencyFormat(xlsx, worksheet, 'F');

    currencyFormat(xlsx, detailsWorksheet, 'D');

    worksheet["!merges"] = merge;
    detailsWorksheet["!merges"] = detailsMerge;

    const wscols = [
        { wch: 30 },
        { wch: 7 },
        { wch: 7 },
        { wch: 15 },
        { wch: 5 },
        { wch: 15 },
    ];

    worksheet['!cols'] = wscols;

    const dwscols = [
        { wch: 5 },
        { wch: 20 },
        { wch: 6 },
        { wch: 10 },
    ];
    detailsWorksheet['!cols'] = dwscols;

    workbook.Sheets['Summary'] = worksheet;
    workbook.Sheets['Details'] = detailsWorksheet;
    xlsx.writeFile(workbook, `${getDestPath()}/${excelFileName}.xlsx`);
}