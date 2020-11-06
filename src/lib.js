import { copyright, company, shortCompany } from './constant';

export const unaccentVietnamese = (str) => {
    str = str.toLowerCase();
    //     We can also use this instead of from line 11 to line 17
    //     str = str.replace(/\u00E0|\u00E1|\u1EA1|\u1EA3|\u00E3|\u00E2|\u1EA7|\u1EA5|\u1EAD|\u1EA9|\u1EAB|\u0103|\u1EB1|\u1EAF|\u1EB7|\u1EB3|\u1EB5/g, "a");
    //     str = str.replace(/\u00E8|\u00E9|\u1EB9|\u1EBB|\u1EBD|\u00EA|\u1EC1|\u1EBF|\u1EC7|\u1EC3|\u1EC5/g, "e");
    //     str = str.replace(/\u00EC|\u00ED|\u1ECB|\u1EC9|\u0129/g, "i");
    //     str = str.replace(/\u00F2|\u00F3|\u1ECD|\u1ECF|\u00F5|\u00F4|\u1ED3|\u1ED1|\u1ED9|\u1ED5|\u1ED7|\u01A1|\u1EDD|\u1EDB|\u1EE3|\u1EDF|\u1EE1/g, "o");
    //     str = str.replace(/\u00F9|\u00FA|\u1EE5|\u1EE7|\u0169|\u01B0|\u1EEB|\u1EE9|\u1EF1|\u1EED|\u1EEF/g, "u");
    //     str = str.replace(/\u1EF3|\u00FD|\u1EF5|\u1EF7|\u1EF9/g, "y");
    //     str = str.replace(/\u0111/g, "d");
    str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, 'a');
    str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, 'e');
    str = str.replace(/ì|í|ị|ỉ|ĩ/g, 'i');
    str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, 'o');
    str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, 'u');
    str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, 'y');
    str = str.replace(/đ/g, 'd');
    // Some system encode vietnamese combining accent as individual utf-8 characters
    str = str.replace(/\u0300|\u0301|\u0303|\u0309|\u0323/g, ''); // Huyền sắc hỏi ngã nặng
    str = str.replace(/\u02C6|\u0306|\u031B/g, ''); // Â, Ê, Ă, Ơ, Ư
    return str;
};

export const sourceToJson = (sourceFileName) => {
    const xlsx = require('xlsx');

    const workbook = xlsx.readFile(sourceFileName);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!worksheet) {
        return undefined;
    }

    return xlsx.utils.sheet_to_json(worksheet);
}

export const saveExcelFile = async (fileName, data, sheetName = "Index", customFunction, showFullCompany = true, data2, sheetName2, customFunction2) => {
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(sheetName);
    let worksheet2;
    if (sheetName2) {
        worksheet2 = workbook.addWorksheet(sheetName2);
    }
        data.unshift([
            showFullCompany ? company : ''
        ]);
   
    // Report Summary
    data.push([]);
    data.push([copyright]);
    worksheet.addRows(data);

    if (data2 && worksheet2) {
        data2.unshift([
            showFullCompany ? company : ''
        ]);
        data2.push([]);
        data2.push([copyright]);
        worksheet2.addRows(data2);
    }

    setFontAndSizeExcel(worksheet);
    setDefaultRowHeight(worksheet);

    if (worksheet2) {
        setFontAndSizeExcel(worksheet2);
        setDefaultRowHeight(worksheet2);

    }

    if (customFunction) {
        customFunction(worksheet);
    }

    if (customFunction2) {
        customFunction2(worksheet2);
    }
    // foramt report summary
    formatReportFooter(worksheet.getCell(`A${worksheet.rowCount}`));
    if (worksheet2) {
        formatReportFooter(worksheet2.getCell(`A${worksheet2.rowCount}`));
    }

    return await workbook.xlsx.writeFile(fileName);

    // if (fileName.includes('khai/index')) {
    //     const libre = require('libreoffice-convert');

    //     const fs = require('fs');

    //     const extend = '.pdf'

    //     // Read file
    //     const file = fs.readFileSync(fileName);
    //     // Convert it to pdf format with undefined filter (see Libreoffice doc about filter)
    //     libre.convert(file, extend, undefined, (err, done) => {
    //         if (err) {
    //             console.log(`Error converting file: ${err}`);
    //         }

    //         // Here in done you have pdf file which you can save or transfer in another stream
    //         fs.writeFileSync(fileName.replace('.xlsx', extend), done);
    //     });
    // }


}

export const setFontAndSizeExcel = (worksheet, fontFamily = 'Tahoma', fontSize = 13) => {
    for (let col = 1; col <= worksheet.columnCount; col++) {
        worksheet.getColumn(col).eachCell((cell, rowNumber) => {
            cell.style = { font: { name: fontFamily, size: fontSize } }
        });
    }
}


export const borderExcel = (worksheet, fromRow, toRow, columns) => {
    const border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber >= fromRow && rowNumber <= toRow) {
            columns.map((it) => {
                worksheet.getCell(`${it}${rowNumber}`).border = border;
            });
        }
    });
}


export const formatCellHeader = (worksheet, row, column, centerAlignment = true, fillColor = true) => {

    const cell = worksheet.getCell(`${column}${row}`);

    cell.style.font = { name: 'Tahoma', size: 13, bold: true };

    if (centerAlignment) {
        cell.style.alignment = { horizontal: 'center', vertical: 'middle' };
    }

    // increase header height
    worksheet.getRow(row).height = 20;

    // fill gray background
    if (fillColor) {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFe1dbdb' }
        };
    }

}

export const formatReportTitle = (worksheet, row, startColumn, endColumn) => {
    const cell = worksheet.getCell(`A${row}`);
    cell.style.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.style.font = { name: 'Tahoma', size: 16, bold: true };
    worksheet.mergeCells(`${startColumn}${row}:${endColumn}${row}`);

    worksheet.getRow(row).height = 22;
}

export const setDefaultRowHeight = (worksheet, height = 16) => {
    for (let i = 1; i <= worksheet.rowCount; i++) {
        worksheet.getRow(i).height = height;
        worksheet.getRow(i).alignment = { vertical: 'middle' };
    }
}

export const formatReportFooter = (cell) => {
    cell.style.font = { ...cell.style.font, italic: true, color: { argb: 'FF7b7979' } };
}