const { openFileDialog, prepareResult, exportResult, toggleView } = require('./utilities');

// initialize handlers
document.getElementById('browseFile').onclick = openFileDialog;
document.getElementsByName('toggleView').forEach(t => {
  t.onclick = toggleView;
});
document.getElementById('btn_Result').onclick = prepareResult;
document.getElementById('btn_Export').onclick = exportResult;
