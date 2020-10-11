import { doctorFullName, qty, productName, amount, doctorPath, productDate } from './constant';
import { getUnaccentcFirstAndLastName, getDoctorShortName, getDestPath, currencyFormat, getDoctorNickname } from './parse';
import { sourceToJson, unaccentVietnamese } from './lib';
import { from, of, zip } from 'rxjs';
import { groupBy, mergeMap, reduce, toArray, filter } from 'rxjs/operators';


export const exportDoctor = (source) => {
    const fs = require('fs');
    fs.readdir(doctorPath, (error, files) => {
        const doctorFiles = files.filter(f => f !== '.DS_Store' && !f.startsWith('~'));
        const totalFiles = doctorFiles.length;
        if (totalFiles > 0) {
            const doctorDirectory = sourceToJson(`${doctorPath}/${doctorFiles[0]}`);

            source.sort(sortDoctorList);

            const groupSource$ = from(source)
                .pipe(
                    filter(it => it[productDate]),
                    groupBy(group => group[doctorFullName]),
                    mergeMap(group => zip(of(group.key), group.pipe(reduce((total, cur) => total + cur[qty], 0)), group.pipe(reduce((total, cur) => total + cur[amount], 0)), group.pipe(toArray())))
                );

            const summaryData = [];
            let sumTotalQty = 0;
            let sumTotalAmount = 0;
            groupSource$.subscribe((z) => {
                sumTotalQty += z[1];
                sumTotalAmount += z[2];
                if (z[0]) {
                    const doctorId = getDoctorIdFromName(doctorDirectory, z[0]);
                    summaryData.push([
                        doctorId, z[0], z[1], z[2]
                    ]);

                    const detailsData = [];
                    z[3].sort((a, b) => {
                        if (a[productName] > b[productName]) {
                            return 1;
                        } else if (a[productName] < b[productName]) {
                            return -1;
                        } else {
                            return 0;
                        }
                    });
                    let sumQty = 0;
                    let sumAmount = 0;
                    for (let product of z[3]) {
                        detailsData.push([
                            product[productName], product[qty], product[amount]
                        ]);
                        sumQty += product[qty];
                        sumAmount += product[amount];
                    }
                    detailsData.push([
                        'Tổng cộng', sumQty, sumAmount
                    ]);
                    writeDetailsFile(doctorId, detailsData);
                } else {
                    summaryData.push([
                        'Bán lẻ', undefined, z[1], z[2]
                    ]);
                }
            });

            summaryData.sort((a, b) => {
                if (a[3] > b[3]) {
                    return -1;
                } else if (a[3] < b[3]) {
                    return 1;
                } else {
                    return 0;
                }
            });
            summaryData.push([
                'Tổng cộng', undefined, sumTotalQty, sumTotalAmount
            ]);
            writeSummaryFile(summaryData);

        } else {
            console.log('Không tìm thấy file danh mục bác sĩ trong thu mục dist -> doctor');
        }
    });


}

const writeSummaryFile = (data) => {
    data.unshift([
        'Mã Bác Sĩ', 'Bác Sĩ', 'SL', 'Tổng Doanh Số Bán Lẻ', 'Hạng'
    ]);

    let count = 0;
    data = data.map((it) => {
        if (it[1]) {
            it.push(count++);
        }
        return it;
    });
    const xlsx = require('xlsx');
    const workbook = xlsx.utils.book_new();
    workbook.Props = {
        Title: 'LamGift',
        Author: 'Ly Van Khai 0986 409 026',
        Subject: 'Excel Generator',
    }
    workbook.SheetNames.push('Index');
    const worksheet = xlsx.utils.aoa_to_sheet(data);


    currencyFormat(xlsx, worksheet, 'C');
    currencyFormat(xlsx, worksheet, 'D');

    const wscols = [

        { wch: 7 },
        { wch: 30 },
        { wch: 12 },
        { wch: 20 },
        { wch: 5 },
    ];

    worksheet['!cols'] = wscols;

    const merge = [];
    merge.push(
        { s: { r: data.length - 1, c: 0 }, e: { r: data.length - 1, c: 1 } }
    );

    worksheet["!merges"] = merge;


    workbook.Sheets['Index'] = worksheet;
    const fs = require('fs');
    fs.mkdirSync(`${getDestPath()}/doctors`, { recursive: true });
    xlsx.writeFile(workbook, `${getDestPath()}/doctors/_summary.xlsx`);
}


const writeDetailsFile = (fileName, data) => {
    data.unshift([
        'Thuốc', 'SL', 'Thành Tiền Bán Lẻ'
    ]);

    const xlsx = require('xlsx');
    const workbook = xlsx.utils.book_new();
    workbook.Props = {
        Title: 'LamGift',
        Author: 'Ly Van Khai 0986 409 026',
        Subject: 'Excel Generator',
    }
    workbook.SheetNames.push('Index');
    const worksheet = xlsx.utils.aoa_to_sheet(data);

    currencyFormat(xlsx, worksheet, 'B');
    currencyFormat(xlsx, worksheet, 'C');

    const wscols = [
        { wch: 40 },
        { wch: 7 },
        { wch: 15 },
    ];

    worksheet['!cols'] = wscols;

    workbook.Sheets['Index'] = worksheet;
    const fs = require('fs');
    fs.mkdirSync(`${getDestPath()}/doctors`, { recursive: true });
    xlsx.writeFile(workbook, `${getDestPath()}/doctors/${fileName}.xlsx`);
}

const sortDoctorList = (a, b) => {
    const { firstName: firstName1, lastName: lastName1 } = getUnaccentcFirstAndLastName(a[doctorFullName]);
    const { firstName: firstName2, lastName: lastName2 } = getUnaccentcFirstAndLastName(b[doctorFullName]);

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
            if (a[productName] > b[productName]) {
                return 1;
            } else if (a[productName] < b[productName]) {
                return -1;
            } else {
                return 0;
            }

        }
    }
}

const getDoctorIdFromName = (doctorDirectory, name) => {
    const found = doctorDirectory.find((it) => it.FullName === name);
    if (found && found.ID) {
        return `${found.ID}`;
    } else {
        return getDoctorShortName(name);
    }

}