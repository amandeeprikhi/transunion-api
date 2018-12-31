import express from 'express';
var Excel = require('exceljs');
import db from './db/db';
// Set up the express app
const app = express();
// get all todos

var workbook = new Excel.Workbook();
// workbook.xlsx.readFile('test-file.xlsx').then(function () {
//     var worksheet = workbook.getWorksheet(1);
//     worksheet.eachRow(function(row, rowNumber) {
//         // let textdata = worksheet.getCell(rowNumber,10).value.toString();
//         // worksheet.getCell(rowNumber, 10).value = textdata;
//         // console.log(JSON.stringify(row.values));
//     });
//     worksheet.commit;
//     workbook.commit;
//     workbook.csv.writeFile('test-file2.csv').then(function(){
//         var worksheet2 = workbook.addWorksheet('test');
//         var rows = [];
//         for(let index = 1; index <= worksheet.lastRow.number; index ++){
//             // worksheet2.addRow(worksheet.getRow(index).values);
//             let textdata = worksheet.getCell(index,10).value.toString();
//             console.log(JSON.stringify(textdata));
//             worksheet.getCell(index, 10).value = textdata;
//             worksheet2.addRow(worksheet.getRow(index).values).commit();
//             // worksheet2.commit();
//             // console.log(rows);
//         }
        
//         for(let index = 1; index <= worksheet.lastRow.number; index++){
//             worksheet2.getRow(index).commit;
//             // console.log(worksheet2.getRow(index).commit);
//         }
//         // console.log()
//         // console.log(worksheet.properties)
//         // console.log(worksheet.lastRow.values);
//         // console.log("CSV WRITTEN!!!!!")
//         workbook.commit;
//         worksheet2.commit;
//     })

    workbook.csv.readFile('test-file2.csv').then(function(){
        var worksheet3 = workbook.getWorksheet(1);
        worksheet3.eachRow(function(row, rowNumber){
            // console.log(JSON.stringify(row.values));
            row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                // worksheet.getCell(rowNumber, 10).value = cell.text;
                console.log('Cell ' + colNumber + ' = ' + cell.type);
            });
        });
    })
// });



// workbook.xlsx.readFile("test-file2.xlsx").then(function(){
//     var worksheet2 = workbook.getWorksheet(1);
//     worksheet2.eachRow(function(row, rowNumber) {
//         // worksheet.getCell(rowNumber, 10).value = worksheet.getCell(rowNumber, 10).value.toString;
//         console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
//         row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
//             // worksheet.getCell(rowNumber, 10).value = cell.text;
//             console.log('Cell ' + colNumber + ' = ' + cell.type);
//         });
//     });
// });

// var sheet1 = workbook.addWorksheet('Sheet1');
// var row = sheet1.getRow(1);
// var reColumns = [{
//         header: 'FirstName',
//         key: 'firstname'
//     },
//     {
//         header: 'LastName',
//         key: 'lastname'
//     },
//     {
//         header: 'TEST',
//         key: 'othername'
//     }
// ];
// sheet1.columns = reColumns;
// for (let index = 2; index < 5; index++) {
//     sheet1.getCell('B'+index).value = 2014; 
// }
// sheet1.getCell('G6').value='test';
// workbook.xlsx.writeFile("test-file2.xlsx").then(function () {
//     console.log("xlsx file is written.");
//     sheet1.eachRow({ includeEmpty: true }, function(row, rowNumber) {
//         console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
//         row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
//             // cell.value = cell.get;
//             console.log('Cell ' + colNumber + ' = ' + cell.type);
//         });
//     });
//    console.log(sheet1.getCell(6,7).value);
//    console.log(sheet1.rowCount);
// });
app.get('/api', (req, res) => {
    res.status(200).send({
        success: 'true',
        message: 'todos retrieved successfully',
        todos: db
    })
});
const PORT = 5000;

app.listen(PORT, () => {
    console.log(`API running on port ${PORT}`)
});