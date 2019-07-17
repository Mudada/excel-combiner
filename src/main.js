"use strict";
exports.__esModule = true;
var exceljs_1 = require("exceljs");
var workbook = new exceljs_1.Workbook();
workbook.xlsx.readFile("/Users/mudada/Code/Script/excel-combiner/excel/Tableau Carnot TSN-EP-v3.xlsx")
    .then(function () {
    var newWorkbook = new exceljs_1.Workbook();
    workbook.eachSheet(function (worksheet, sheetId) {
        newWorkbook.addWorksheet(worksheet.name);
        console.log(worksheet.getRow(2));
        newWorkbook.getWorksheet(worksheet.name).addRow(worksheet.getRow(2));
        newWorkbook
            .xlsx.writeFile("/Users/mudada/Code/Script/excel-combiner/excel/output/output.xlsx")
            .then(function () { return console.log("done"); });
    });
});
