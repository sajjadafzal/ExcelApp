// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const { ipcRenderer } = require('electron');

require('./handlers');

global.resultjs = [];
global.datajs = [];
global.keyjs = [];

// read data file

const { dataDiv, keyDiv } = require('./uielements');
const { ReformatJSON, displayJSON, readExcelToJSON, AddColor } = require('./utilities');

const datajs = readExcelToJSON('DataFile.xls');
const keyjs = readExcelToJSON('KEY.xls');

// console.log(JSON.stringify([{ 'Q1': { 'value': 'T' }, 'Q2': 'B', 'Q3': 'C' }, { 'Q1': 'D', 'Q2': 'E', 'Q3': 'F' }], ['Q1', 'Q2']));

// console.og(datajs);
// TESTING NEW JSON
const newJSONData = AddColor(ReformatJSON(datajs, 4), 5);
// console.log(newJSONData);
const newJSONKey = AddColor(ReformatJSON(keyjs, 3));
global.datajs = newJSONData;
global.keyjs = newJSONKey;
displayJSON(global.datajs, dataDiv);
displayJSON(newJSONKey, keyDiv);

// recieve event from file menu
ipcRenderer.addListener('files', (err, data) => {
  // console.log('yo', err, data);
});
