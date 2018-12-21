const Excel = require('xlsx');
const { remote } = require('electron');

const { dataDiv, keyDiv, resultDiv } = require('./uielements');
const { dialog } = remote;

function toggleView(event) {
  const view = event.target.value;

  if (view === 'key') {
    dataDiv.style.display = 'none';
    keyDiv.style.display = 'block';
    resultDiv.style.display = 'none';
  } else if (view === 'Data') {
    dataDiv.style.display = 'block';
    keyDiv.style.display = 'none';
    resultDiv.style.display = 'none';
  } else {
    dataDiv.style.display = 'none';
    keyDiv.style.display = 'none';
    resultDiv.style.display = 'block';
  }
}

function readExcelToJSON(path) {
  var workbook = Excel.readFile(path);
  const datajs = Excel.utils.sheet_to_json(
    workbook.Sheets[workbook.SheetNames[0]]
  );

  return datajs;
}

function prepareResult() {
  global.resultjs = [];

  // read data file
  const datajs = ReformatJSON(readExcelToJSON('DataFile.xls'), 4);
  const keyjs = ReformatJSON(readExcelToJSON('KEY.xls'), 3);

  var totalMarks = 0;

  datajs.forEach(row => {
    var colNames = Object.keys(row);
    var marksObtained = 0;

    var correct = 0;
    var inCorrect = 0;
    var percentage = 0;
    var nonAttempt = 0;

    keyjs.forEach(keyRow => {
      if (keyRow.POST === row.POST) {
        colNames.forEach(colName => {
          if (keyRow[colName] !== '*' || keyRow[colName] !== '?') {
            if (colName[0] === 'Q') {
              if (row[colName] === '?') {
                nonAttempt++;
              } else {
                if (row[colName] === keyRow[colName]) {
                  marksObtained += 3;
                  correct++;
                } else {
                  marksObtained -= 1;
                  inCorrect++;
                }
              }
            }
          }
        });
      }
    });

    global.resultjs.push({ RollNo: row.RollNo, 'Total Marks Obtained': marksObtained, Correct: correct, 'In Correct': inCorrect, 'Un Attempted': nonAttempt });
  });

  // console.log(resultjs);
  displayJSON(global.resultjs, resultDiv);
}

function exportResult() {
  var wb = Excel.utils.book_new();
  wb.Props = {
    Title: 'RESULT',
    Subject: 'Test',
    Author: 'Sajjad Afzal',
    Manager: 'Zaheen Muhammad',
    CreatedDate: new Date()
  };
  wb.SheetNames.push('RESULT');
  var ws = Excel.utils.json_to_sheet(global.resultjs);
  wb.Sheets['RESULT'] = ws;
  Excel.writeFile(wb, 'RESULT.xlsx');
}

function ReformatJSON(sourceJSON, noOfColShifts) {
  var colNames = Object.keys(sourceJSON[0]);

  // performs a circular shift on an array
  while (noOfColShifts--) {
    colNames.unshift(colNames.pop());
  }

  var dataText = JSON.stringify(sourceJSON, colNames);
  return JSON.parse(dataText);
}

function displayJSON(jsonData, displayDiv) {
  var tbl = document.createElement('table');
  var tr = document.createElement('tr'); // Header row

  // tbl.className = 'table-bordered';

  Object.keys(jsonData[0]).forEach(hdr => {
    var th = document.createElement('th');
    th.innerHTML = hdr;
    tr.append(th);
  });

  tbl.append(tr);

  // Getting data values
  jsonData.forEach(row => {
    var rw = document.createElement('tr');
    Object.values(row).forEach(ent => {
      var cl = document.createElement('td');
      cl.innerHTML = ent;
      rw.append(cl);
    });
    tbl.append(rw);
  });

  displayDiv.innerHTML = '';
  displayDiv.append(tbl);
}

function openFileDialog(event) {
  dialog.showOpenDialog(
    {
      properties: ['openFile'],
      filters: [
        {
          name: '',
          extensions: ['xlsx', 'xls', 'xlsm']
        }
      ]
    },
    files => {
      if (files) {
        event.sender.send('selected-directory', files);
      }
    }
  );
}

module.exports = { ReformatJSON, displayJSON, readExcelToJSON, openFileDialog, exportResult, prepareResult, toggleView };

// var dataText = '[ ';
// var isFirst = true;

// sourceJSON.forEach(r => {
//   if (!isFirst) {
//     dataText = dataText + ',';
//   }
//   isFirst = false;
//   dataText = dataText + '{ "RollNo":';
//   dataText = dataText + r.RollNo;
//   dataText = dataText + ', "POST":';
//   dataText = dataText + r.POST;
//   dataText = dataText + ', "CENTER":';
//   dataText = dataText + r.CENTER;
//   dataText = dataText + ', "TIME":';
//   dataText = dataText + r.TIME;
//   var colNames = Object.keys(r);
//   colNames.forEach(colName => {
//     if ((colName[0] = 'Q')) {
//       dataText = dataText + ',"' + colName + '":';
//       dataText = dataText + '"' + r[colName] + '"';
//     }
//   });
//   dataText = dataText + '}';
// });
// dataText = dataText + ']';
// console.log(dataText);
// return JSON.parse(dataText);
