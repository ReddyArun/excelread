var express = require('express');
var router = express.Router();
var Excel = require("exceljs");
var path = require('path');
var workbook = new Excel.Workbook();

/* GET home page. */
router.get('/', function (req, res, next) {
    var pattern = /(\d{2})\/(\d{2})\/(\d{4})/;
    var dateVal = '';
    var data = {}, gender, name, dob, address, mobile, rollnum, cla, fee1, fee2, comb, sslcadd, sslcper;
    var files = ['heba1.xlsx', 'heba2.xlsx', 'pcmb1.xlsx', 'pcmb2.xlsx']
    files.forEach(function (file) {
        workbook.xlsx.readFile(path.join(__dirname, '../', 'excel', file))
                .then(function () {
                    var worksheet = workbook.getWorksheet(1);
                    worksheet.eachRow({includeEmpty: true}, function (row) {

                        if (row.getCell(2).type === 3) {
                            //Gender
                            if (row.getCell(1).type === 3) {
                                if (row.getCell(1).value === 'm')
                                    gender = 'male';
                                else if (row.getCell(1).value === 'f')
                                    gender = 'female';
                                else
                                    gender = '';

                            } else {
                                gender = '';
                            }
                            //Name
                            name = row.getCell(2).value;
                            //DOB
                            if (row.getCell(3).type === 3) { //String type
                                dateVal = row.getCell(3).value;
                                dob = new Date(dateVal.replace(pattern, '$3-$2-$1'));
                            } else if (row.getCell(3).type === 4) {
                                dob = new Date(row.getCell(3).value);
                            } else {
                                dob = '';
                            }
                            //Address
                            if (row.getCell(4).type === 3) {
                                address = row.getCell(4).value;
                            } else {
                                address = '';
                            }
                            //Mobile
                            if (row.getCell(5).type === 2) {
                                mobile = row.getCell(5).value;
                            } else {
                                mobile = '';
                            }
                            //Rollnum
                            if (row.getCell(6).type === 3) {
                                rollnum = row.getCell(6).value;
                            } else {
                                rollnum = '';
                            }
                            //Class
                            if (row.getCell(7).type === 3) {
                                if (row.getCell(7).value === '1st PUC')
                                    cla = 'PUC1';
                                else
                                    cla = 'PUC2';
                            } else {
                                cla = '';
                            }
                            //Combination
                            if (row.getCell(9).type === 3) {
                                comb = row.getCell(9).value;
                            } else {
                                comb = '';
                            }
                            //Fee
                            if (row.getCell(8).type === 2) {
                                if (cla === 'PUC1') {
                                    fee1 = row.getCell(8).value;
                                    fee2 = 0;
                                } else if (cla === 'PUC2') {
                                    fee1 = 0;
                                    fee2 = row.getCell(8).value;
                                } else {
                                    fee1 = 0;
                                    fee2 = 0;
                                }
                            } else {
                                fee1 = 0;
                                fee2 = 0;
                            }
                            //SslcAddress
                            if (row.getCell(10).type === 3) {
                                sslcadd = row.getCell(10).value;
                            } else {
                                sslcadd = '';
                            }
                            //Sslcper
                            if (row.getCell(11).type === 2) {
                                sslcper = row.getCell(11).value;
                            } else {
                                sslcper = '';
                            }
                            data = {
                                name: name,
                                dob: dob,
                                rollnumber: rollnum,
                                class: cla,
                                caste: '',
                                gender: gender,
                                puc1fee: fee1,
                                puc2fee: fee2,
                                mobile: mobile,
                                address: address,
                                fee: [],
                                combination: comb,
                                sslcschooladdress: sslcadd,
                                sslcpercentage: sslcper,
                                createduser: 'root',
                                updateduser: 'root',
                                image: '',
                                updateddate: Date.now(),
                                search: [name, rollnum, mobile]
                            };
                            req.app.db.models.Student.create(data);
                        }
                    });
                });
    });
    res.render('index', {title: 'Express'});
});

module.exports = router;
