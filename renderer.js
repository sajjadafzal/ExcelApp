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
displayJSON(datajs, dataDiv);
displayJSON(keyjs, keyDiv);

// var tbl = document.createElement('');
function displayJSON(jsonData, displayDiv) {
  var tbl = document.createElement('table');

  var tr = document.createElement('tr'); //Header row

  Object.keys(jsonData[0]).forEach(hdr => {
    var th = document.createElement('th');
    th.innerHTML = hdr;
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
