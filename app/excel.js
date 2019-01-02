var Excel = require('exceljs');

var excel = function () {
    let workbook = new Excel.Workbook();
    workbook.xlsx.readFile('test-file.xlsx').then(function () {
        let worksheet = workbook.getWorksheet(1);
        let worksheet2 = workbook.addWorksheet('test');
        for (let index = 1; index <= worksheet.lastRow.number; index++) {
            let textdata = worksheet.getCell(index, 10).value.toString();
            if (textdata.length <= 3) {
                let textdata2 = textdata + "00";
                worksheet.getCell(index, 10).value = textdata2;
                worksheet2.addRow(worksheet.getRow(index).values).commit();
                workbook.commit;
            } else if (textdata.length <= 4) {
                let textdata3 = textdata + "0";
                worksheet.getCell(index, 10).value = parseInt(textdata3);
                worksheet2.addRow(worksheet.getRow(index).values).commit();
                workbook.commit;
            } else {
                if (textdata.charAt(5) == "-") {
                    let part1 = textdata.substring(0, 5);
                    let part2 = textdata.substring(6, textdata.length);
                    textdata = part1 + part2;
                }
                worksheet.getCell(index, 10).value = textdata;
                worksheet2.addRow(worksheet.getRow(index).values).commit();
                workbook.commit;
            }
        }
        workbook.csv.writeFile('test-file.csv').then(function () {
            console.log("FILE IS WRITTEN!!!!!!!!!!!");
        });
    });
}

var read = function () {
    let workbook = new Excel.Workbook();
    workbook.xlsx.readFile('TU Template Sheet.xlsx').then(function () {
        let worksheet = workbook.getWorksheet(1);
        worksheet.eachRow(function (row, rowNumber) {
            console.log(JSON.stringify(row.values))
            row.eachCell(function (cell, colNumber) {});
        });
    })
}

var write = function () {
    let workbook = new Excel.Workbook();
    let workbook2 = new Excel.Workbook();
    workbook.csv.readFile('TU Template Sheet.csv').then(function () {
        let worksheet = workbook.getWorksheet(1);
        worksheet.getRow(1).values;
        workbook.csv.writeFile('TU Template Sheet.csv').then(function () {
            workbook2.csv.readFile('test-file.csv').then(function () {
                let worksheet2 = workbook2.getWorksheet(1);
                // workbook2.eachSheet(function(worksheet, sheetId) {
                //     console.log(sheetId);
                // });
                for (let index = 2; index <= worksheet2.lastRow.number; index++) {
                    // console.log(worksheet2.getCell(index,2).value);
                    worksheet.getCell(index, 1).value = worksheet2.getCell(index, 2).value;
                    worksheet.getCell(index, 2).value = worksheet2.getCell(index, 3).value;
                    worksheet.getCell(index, 3).value = worksheet2.getCell(index, 11).value;
                    worksheet.getCell(index, 7).value = worksheet2.getCell(index, 5).value.toString() + worksheet2.getCell(index, 6).value.toString();
                    worksheet.getCell(index, 10).value = worksheet2.getCell(index, 8).value;
                    worksheet.getCell(index, 11).value = worksheet2.getCell(index, 9).value;
                    worksheet.getCell(index, 12).value = worksheet2.getCell(index, 10).value;
                    worksheet.addRow(worksheet.getRow(index).values).commit();
                    console.log("FILE WRITTEN!!!")
                }
            })
            .then(function(){
                workbook.csv.writeFile('TU Template Sheet2.csv').then(function () {
                    console.log("THE FILE IS CREATED")
                })
            });
        })
    })
}
module.exports = {
    excel: excel,
    read: read,
    write: write
};