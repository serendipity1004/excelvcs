const Excel = require('exceljs');
const path = require('path');
const util = require('util');

const excelParserTool = function(){
    return new Promise((resolve, reject) => {
        let oldFilePath = path.join(appRoot, 'resources', 'testSpreadsheetOriginal.xlsx');
        let newFilePath = path.join(appRoot, 'resources', 'testSpreadsheetNew.xlsx');

        let oldExcelArr = [];
        let newExcelArr = [];
        let excelReadPromises = [];

        let oldExcel = new Excel.Workbook();
        let newExcel = new Excel.Workbook();

        //Promise for creating excel array
        let oldExcelReadPromise = excelReadPromise(oldExcel, oldFilePath);
        let newExcelReadPromise = excelReadPromise(newExcel, newFilePath);

        excelReadPromises.push(oldExcelReadPromise, newExcelReadPromise);

        Promise.all(excelReadPromises).then((sheetArr) =>{
            //sheetArr[0] = old AND sheetArr[1] = new
            resolve(sheetArr)
        });

        oldExcelReadPromise.then((sheetArr) => {
            console.log(sheetArr)
        });
    });
};

//Parser Promise
const excelReadPromise = function(excelFile, path){
    return new Promise((resolve, reject) => {
        excelFile.xlsx.readFile(path).then(() => {
            // console.log(oldExcel)

            let sheetArr = [];

            //Run through each worksheet and create a excel sheet array
            excelFile.eachSheet((worksheet, sheetId) => {

                for (let i = 1; i < worksheet.rowCount; i++) {
                    let colArr = [];
                    for (let j = 1; j < worksheet.columnCount + 1; j++) {
                        colArr.push(worksheet.getRow(i).values[j])
                    }
                    sheetArr.push(colArr)
                }

                console.log('resolving');
                resolve(sheetArr);
            })
        });
    })
};



module.exports = {
    excelParserTool
};