let data = JSON.parse(oldArray);
let diffArr = JSON.parse(excelCompResult)[1];
let extraArr = JSON.parse(excelCompResult)[0];

let extraPositionsArr = [];
let diffPositionsArr = [];

console.log(extraArr);

let container = document.getElementById('example');

// Setup changed data values
for (let arrIndex in extraArr) {
    let curRow = extraArr[arrIndex];
    let xyPosition = curRow[1].split(":");
    let xPosition = xyPosition[1] - 1;
    let yPosition = xyPosition[0] - 1;
    let addedVal = curRow[2];

    console.log(xyPosition);
    console.log(addedVal);

    data[xPosition][yPosition] = '***' + addedVal + '***';
    extraPositionsArr.push([xPosition, yPosition])
}

for (let arrIndex in diffArr) {
    let curRow = diffArr[arrIndex];
    let xyPosition = curRow[0].split(":");
    let xPosition = xyPosition[1] - 1;
    let yPosition = xyPosition[0] - 1;
    let changeFrm = curRow[1];
    let changeTo = curRow[2];

    data[xPosition][yPosition] = changeFrm + '=>' + changeTo;
    diffPositionsArr.push([xPosition, yPosition]);
}

let hot = new Handsontable(container, {
    data: data,
    rowHeaders: true,
    colHeaders: true,
    cells: function (row, col, prop) {
        let cellProperties = {};

        for (let index in extraPositionsArr){
            let eachPos = extraPositionsArr[index];

            if (row === eachPos[0] && col === eachPos[1]){
                cellProperties.renderer = extraRowRenderer;
            }
        }

        for (let index in diffPositionsArr){
            let eachPos = diffPositionsArr[index];

            if (row === eachPos[0] && col === eachPos[1]){
                cellProperties.renderer = diffRowRenderer;
            }
        }

        return cellProperties
    }
});

function extraRowRenderer(instance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);

    td.style.fontWeight = 'bold';
    td.style.backgroundColor = '#F48FB1';
}

function diffRowRenderer(instance, td, row, col, prop, value, cellProperties){
    Handsontable.renderers.TextRenderer.apply(this, arguments)

    td.style.fontWeight = 'bold';
    td.style.backgroundColor = '#CE93D8';
}

hot.render();