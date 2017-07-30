const express = require('express');
const exphbs = require('express-handlebars');
const hbs = require('handlebars');
const path = require('path');

global.appRoot = path.resolve(__dirname);

const {excelDiffTool} = require('./tools/excel/excelDiffTool');
const {excelParserTool} = require('./tools/excel/excelParserTool');

let app = express();

app.use(express.static('views'));
app.use(express.static('handsontable'));

app.set('views', path.join(__dirname, 'views'));

app.engine('handlebars', exphbs({defaultLayout: 'main'}));
app.set('view engine', 'handlebars');
app.set('port', (process.env.PORT || 8000));

app.get('/', (req, res) => {
    let excelCompResult = '';
    let excelArr = '';
    let rootPromisesArr = [];

    let diffPromise = excelDiffTool().then((compResult) => {
            excelCompResult = compResult
        }
    );

    let parserPromise = excelParserTool().then((sheetArr) => {
        excelArr = sheetArr
    });

    rootPromisesArr.push(diffPromise, parserPromise);

    Promise.all(rootPromisesArr).then(()=>{
        res.render('index', {
            oldArray: encodeURIComponent(JSON.stringify(excelArr[0])),
            newArray: encodeURIComponent(JSON.stringify(excelArr[1])),
            excelCompResult: encodeURIComponent(JSON.stringify(excelCompResult)),
            css: ['index.css'],
            js: ['index.js']
        })
    }
)
});

app.listen(app.get('port'), () => {
    console.log(`application listening at port ${app.get('port')}`)
});

//hbs helpers
hbs.registerHelper('json', (context) => {
    return JSON.stringify(context)
});