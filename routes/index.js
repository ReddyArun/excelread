var express = require('express');
var router = express.Router();
var Excel = require("exceljs");
var path = require('path');
var workbook = new Excel.Workbook();

/* GET home page. */
router.get('/', function (req, res, next) {
    workbook.xlsx.readFile(path.join(__dirname, '../', 'excel', 'heba1.xlsx'))
            .then(function () {
                var worksheet = workbook.getWorksheet(1);
                worksheet.eachRow({includeEmpty: true}, function (row, rowNumber) {
                    console.log("Cell " + 1 + " = " + JSON.stringify(row.getCell(1).value));
                    console.log("Cell " + 2 + " = " + JSON.stringify(row.getCell(2).value));
                    console.log("Cell " + 3 + " = " + JSON.stringify(row.getCell(3).value));
                    console.log("Cell " + 4 + " = " + JSON.stringify(row.getCell(4).value));
                    console.log("Cell " + 5 + " = " + JSON.stringify(row.getCell(5).value));
                    console.log("Cell " + 6 + " = " + JSON.stringify(row.getCell(6).value));
                    console.log("Cell " + 7 + " = " + JSON.stringify(row.getCell(7).value));
                    console.log("Cell " + 8 + " = " + JSON.stringify(row.getCell(8).value));
                    console.log("Cell " + 9 + " = " + JSON.stringify(row.getCell(9).value));
                    console.log("Cell " + 10 + " = " + JSON.stringify(row.getCell(10).value));
                    console.log("Cell " + 11 + " = " + JSON.stringify(row.getCell(11).value));
//                    row.eachCell({includeEmpty: true}, function (cell, cellNumber) {
//                        console.log("Cell " + cellNumber + " = " + JSON.stringify(cell.value));
//                    });

                });
            });
    res.render('index', {title: 'Express'});
});

module.exports = router;
