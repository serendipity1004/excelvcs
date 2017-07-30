const {exec} = require('child_process');
const path = require('path');

const excelDiffTool = function(){
    return new Promise((resolve, reject) => {
        let command = `excel_cmp ${path.join(appRoot, 'resources/testSpreadsheetOriginal.xlsx')} ${path.join(appRoot, 'resources\\testSpreadsheetNew.xlsx')}`;

        //Execute command line to run excel_cmp tool
        exec(command, (err, stdout, stderr) => {
            if (err) {
                console.log(err);
                return;
            }

            let lineArray = stdout.split(/\r?\n/);

            console.log(lineArray);

            let extraArr = [];
            let diffArr = [];

            for (let index in lineArray) {
                let curLine = lineArray[index].replace(/\'/g, '');
                let changeType = curLine.split(" ")[0];

                /*
                * Look for changed position (that starts with 'Sheet')
                * and store position = [0] and changed value = [1]
                * into extraArr
                * */
                if (changeType === "EXTRA") {
                    let wordArr = curLine.split(' ');
                    let changedVal = curLine.split('=>')[1].trim();
                    let position = '';
                    let changedWb = '';

                    for (let index in wordArr) {
                        let eachWord = wordArr[index];

                        if (eachWord.startsWith('Sheet')) {
                            // position - the position of the excel spreadsheet where the value was added
                            position = eachWord.split('!')[1];
                        } else if (eachWord.startsWith('WB')) {
                            changedWb = eachWord.trim();
                        }
                    }

                    position = convertExcelPosToNumb(position)

                    extraArr.push([changedWb, position, changedVal])

                    /*
                    * Do the same thing but for DIFF
                    * */
                } else if (changeType === "DIFF") {
                    let wordArr = curLine.split(' ');
                    let changedVal = curLine.split('=>')[1].trim();
                    let position = '';

                    for (let index in wordArr) {
                        let eachWord = wordArr[index];
                        if (eachWord.startsWith('Sheet')) {
                            // position - the position of the excel spreadsheet where the value was added
                            position = convertExcelPosToNumb(eachWord.split('!')[1]);
                            let changedValArr = changedVal.split('v/s');
                            diffArr.push([position, changedValArr[0].trim(), changedValArr[1].trim()]);
                        }
                    }
                    // diffArr.push(curLine.toString())
                }
            }

            console.log('Extras');
            console.log(extraArr);
            console.log('Diffs');
            console.log(diffArr);

            resolve([extraArr, diffArr]);
        });
    });
};

// Converts input excel position string into number position [x][y]
const convertExcelPosToNumb = (stringToConvert) => {
    let xAxis = '';
    let yAxis = '';

    for (let index in stringToConvert) {
        let eachWord = stringToConvert[index];

        if (/\d/.test(eachWord)) {
            yAxis += eachWord
        } else {
            xAxis += eachWord
        }
    }

    // Converts character position into number positions
    let convertedXAxis = letterToNumbers(xAxis);

    let returnVal = [convertedXAxis, yAxis].join(':');

    return returnVal;
};

//Function that converts character position into number positions
function letterToNumbers(string) {
    string = string.toUpperCase();
    let letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', sum = 0, i;
    for (i = 0; i < string.length; i++) {
        sum += Math.pow(letters.length, i) * (letters.indexOf(string.substr(((i + 1) * -1), 1)) + 1);
    }
    return sum;
}

module.exports = {excelDiffTool};