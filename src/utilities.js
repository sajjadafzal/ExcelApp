const Excel = require('xlsx');
const { remote } = require('electron');

const { dataDiv, keyDiv, resultDiv } = require('./uielements');
const { dialog } = remote;

global.myColors = {
  Red: '#ff9999',
  Green: '#b3ff99',
  Yellow: '#ffff80',
  White: '#ffffff'
};
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

function AddColor(sJSON) {
  const datajs = sJSON;

  datajs.forEach(row => {
    var colNames = Object.keys(row);
    colNames.forEach(col => {
      row[col] = { value: row[col], color: global.myColors.White };
    });
  });
  return datajs;
}
function prepareResult() {
  global.resultjs = [];

  // read data file
  var totalMarks = 0;

  global.datajs.forEach((row, indx) => {
    var scolor = global.myColors.White;
    var colNames = Object.keys(row).filter(cN => { return cN[0] === 'Q'; });
    var marksObtained = 0;

    var correct = 0;
    var inCorrect = 0;
    var percentage = 0;
    var nonAttempt = 0;
    global.keyjs.forEach(keyRow => {
      if (keyRow.POST.value === row.POST.value) {
        colNames.forEach(colName => {
          if (keyRow[colName] !== undefined && (keyRow[colName].value !== '*' && keyRow[colName].value !== '?')) {
            if (colName[0] === 'Q') {
              if (row[colName].value === '?') {
                scolor = global.myColors.White;
                nonAttempt++;
              } else {
                if (row[colName].value === keyRow[colName].value) {
                  marksObtained += 3;
                  scolor = global.myColors.Green;
                  correct++;
                } else {
                  marksObtained -= 1;
                  scolor = global.myColors.Red;
                  inCorrect++;
                }
              }
            }
          } else {
            scolor = global.myColors.Yellow;
          }
          global.datajs[indx][colName].color = scolor;
        });
      }
    });
    global.resultjs.push({ RollNo: row.RollNo, 'Total Marks Obtained': { value: marksObtained, color: global.myColors.White }, Correct: { value: correct, color: global.myColors.White }, 'In Correct': { value: inCorrect, color: global.myColors.White }, 'Un Attempted': { value: nonAttempt, color: global.myColors.White } });
  });

  // console.log(resultjs);
  displayJSON(global.datajs, dataDiv);
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
  // console.log(dataText);
  return JSON.parse(dataText);
}

function displayJSON(jsonData, displayDiv) {
  var tbl = document.createElement('table');
  var tr = document.createElement('tr'); // Header row

  tbl.className = 'fixed_headers';

  Object.keys(jsonData[0]).forEach(hdr => {
    var th = document.createElement('th');
    th.innerHTML = hdr;
    tr.append(th);
  });
  var tblHead = document.createElement('thead');
  // tblHead.style.position = 'fixed';
  var tblBody = document.createElement('tbody');
  tblHead.append(tr);
  tbl.append(tblHead);
  // Getting data values
  jsonData.forEach(row => {
    var rw = document.createElement('tr');
    Object.values(row).forEach(ent => {
      // console.log(ent);
      var cl = document.createElement('td');
      cl.innerHTML = ent.value;
      cl.style.backgroundColor = ent.color;
      rw.append(cl);
    });
    tblBody.append(rw);
  });
  tbl.append(tblBody);
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

module.exports = { ReformatJSON, displayJSON, readExcelToJSON, AddColor, openFileDialog, exportResult, prepareResult, toggleView };

// var dataText = '[ ';
// var isFirst = true;

// sourceJSON.forEach(r => {
//   if (!isFirst) {
//     dataText = dataText + ',';
//   }
//   isFirst = false;
//   dataText = dataText + '{ "RollNo": {"value" : '+ r.RollNo  + '"color" : "#AA0000" }'; '
//   dataText = dataText + ', "POST": {"value" : '+  r.POST + '"color" : "#AA0000" }'; '
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
