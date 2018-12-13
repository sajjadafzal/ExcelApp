// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const remote = require('electron').remote;
const Excel = require('xlsx');
const { app, dialog } = remote;
//var basepath = app.getAppPath();
//document.getElementById('folderpath').innerHTML = basepath;

var dataDiv = document.getElementById('DataSpace');
var keyDiv = document.getElementById('KeySpace');
var resultDiv = document.getElementById('ResultSpace');
var resultjs = [];
var newJSONData;
var newKeyData;
document.getElementById('browseFile').onclick = () => {
  //console.log(dialog);
  dialog.showOpenDialog(
    {
      properties: ['openFile'],
      filters: [
        {
          name: '',
          extensions: ['xlsx', 'xls', 'xlsm'],
        },
      ],
    },
    files => {
      if (files) {
        event.sender.send('selected-directory', files);
      }
    }
  );
};

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
function displayJSON(jsonData, displayDiv) {
  var tbl = document.createElement('table');

  var tr = document.createElement('tr'); //Header row

  Object.keys(jsonData[0]).forEach(hdr => {
    var th = document.createElement('th');
    th.innerHTML = hdr;
    //th.style = 'font-weight: bold; font-style: italic; font-size: 30px;';
    //console.log(c1);
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
  displayDiv.append(tbl);
}

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

function ReformatJSON(sourceJSON, noOfColShifts) {
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
  var colNames = Object.keys(sourceJSON[0]);

  // performs a circular shift on an array
  while (noOfColShifts--) {
    colNames.unshift(colNames.pop());
  }

  dataText = JSON.stringify(sourceJSON, colNames);
  return JSON.parse(dataText);
}
