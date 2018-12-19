// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const remote = require('electron').remote;
const Excel = require('xlsx');
const { app, dialog } = remote;
const { ReformatJSON, displayJSON, Handlers } = require('./utilities'); //one way of importing ReformatJSON from utilities.js file
//const myUtilities = require('./utilities'); //2nd way of importing ReformatJSON from utilities.js file

//var basepath = app.getAppPath();
//document.getElementById('folderpath').innerHTML = basepath;

var dataDiv = document.getElementById('DataSpace');
var keyDiv = document.getElementById('KeySpace');
var resultDiv = document.getElementById('ResultSpace');
var resultjs = [];
var newJSONData;
var newKeyData;
document.getElementById('browseFile').onclick = Handlers.BrowseFileEvent;

var workbook = Excel.readFile('DataFile.xls');
const datajs = Excel.utils.sheet_to_json(
  workbook.Sheets[workbook.SheetNames[0]]
);
workbook = Excel.readFile('KEY.xls');

const keyjs = Excel.utils.sheet_to_json(
  workbook.Sheets[workbook.SheetNames[0]]
);
console.log(datajs);
//TESTING NEW JSON
newJSONData = ReformatJSON(datajs, 4);
newJSONKey = ReformatJSON(keyjs, 3);
displayJSON(newJSONData, dataDiv);
//displayJSON(datajs, dataDiv);
displayJSON(newJSONKey, keyDiv);

// var tbl = document.createElement('');

document.getElementsByName('toggleView').forEach(t => {
  t.onclick = function(event) {
    if (event.target.value == 'key') {
      dataDiv.style.display = 'none';
      keyDiv.style.display = 'block';
      resultDiv.style.display = 'none';
    } else if (event.target.value == 'Data') {
      dataDiv.style.display = 'block';
      keyDiv.style.display = 'none';
      resultDiv.style.display = 'none';
    } else {
      dataDiv.style.display = 'none';
      keyDiv.style.display = 'none';
      resultDiv.style.display = 'block';
    }
  };
});

document.getElementById('btn_Result').onclick = () => {
  datajs.forEach(row => {
    var colNames = Object.keys(row);
    var totalMarks = 0;

    keyjs.forEach(keyRow => {
      if (keyRow.POST == row.POST) {
        colNames.forEach(colName => {
          if (colName[0] == 'Q') {
            if (row[colName] == keyRow[colName]) {
              totalMarks += 3;
            } else {
              totalMarks -= 1;
            }
          }
        });
      }
    });

    resultjs.push({ RollNo: row.RollNo, TotalMarks: totalMarks });
  });
  console.log(resultjs);
  displayJSON(resultjs, resultDiv);
};

document.getElementById('btn_Export').onclick = () => {
  var wb = Excel.utils.book_new();
  wb.Props = {
    Title: 'RESULT',
    Subject: 'Test',
    Author: 'Sajjad Afzal',
    Manager: 'Zaheen Muhammad',
    CreatedDate: new Date(),
  };
  wb.SheetNames.push('RESULT');
  // var ws = Excel.utils.json_to_sheet(resultjs, { skipHeader: true });
  var ws = Excel.utils.json_to_sheet(resultjs);
  // var ws = Excel.utils.table_to_sheet(resultDiv);
  wb.Sheets['RESULT'] = ws;
  Excel.writeFile(wb, 'RESULT.xlsx');
};
