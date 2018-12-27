// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const { ipcRenderer } = require('electron');

require('./handlers');

global.resultjs = [];

// read data file

const { dataDiv, keyDiv } = require('./uielements');
const { ReformatJSON, displayJSON, readExcelToJSON } = require('./utilities');

const datajs = readExcelToJSON('DataFile.xls');
const keyjs = readExcelToJSON('KEY.xls');
console.log(datajs);
// TESTING NEW JSON
const newJSONData = ReformatJSON(datajs, 4);
const newJSONKey = ReformatJSON(keyjs, 3);
// console.log(newJSONData);
displayJSON(newJSONData, dataDiv);
displayJSON(newJSONKey, keyDiv);

// recieve event from file menu
ipcRenderer.addListener('files', (err, data) => {
  // console.log('yo', err, data);
});
