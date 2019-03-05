$("#run").click(() => tryCatch(run));

function run() {
  return Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    range.load("address");
    return context.sync().then(function() {
      console.log('The range address was "' + range.address + '".');
    });
  });
}

/** Default helper for invoking an action and handling errors. */
function tryCatch(callback) {
  Promise.resolve()
    .then(callback)
    .catch(function(error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    });
}
