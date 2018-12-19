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

const Handlers = {
  BrowseFileEvent: () => {
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
  },
};

module.exports = { ReformatJSON, displayJSON, Handlers };
