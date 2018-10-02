// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const remote = require('electron').remote;
const { app, dialog } = remote;
//var basepath = app.getAppPath();
//document.getElementById('folderpath').innerHTML = basepath;

document.getElementById('browseFile').onclick = () => {
  console.log(dialog);
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
