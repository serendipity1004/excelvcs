let data = JSON.parse(oldArray);
let diffArr = JSON.parse(excelCompResult)[1];
let extraArr = JSON.parse(excelCompResult)[0];

console.log(extraArr)

let container = document.getElementById('example');

let hot = new Handsontable(container, {
    data: data,
    rowHeaders: true,
    colHeaders: true
});

for (let arrIndex in extraArr) {
    console.log(data[0][0])
    let curRow = extraArr[arrIndex];
    console.log(curRow)
    let xyPosition = curRow[1].split(":");
    let addedVal = curRow[2];

    console.log(xyPosition);
    console.log(addedVal);

    data[xyPosition[1] - 1][xyPosition[0] - 1] = '***' + addedVal + '***';
}

for (let arrIndex in diffArr) {
    let curRow = diffArr[arrIndex];
    let xyPosition = curRow[0].split(":");
    let changeFrm = curRow[1];
    let changeTo = curRow[2];

    data[xyPosition[1] - 1][xyPosition[0] - 1] = changeFrm + '=>' + changeTo
}

hot.render();