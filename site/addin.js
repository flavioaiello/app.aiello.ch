// addin.js
(function () {
  Office.onReady(function () {
    // Add-in is initialized, but event handling is defined in the manifest
  });
})();

function onNewEmailHandler(event) {
  // Get the current time
  var currentTime = new Date().toLocaleTimeString();

  // Display a dialog box with the current time
  Office.context.ui.displayDialogAsync(
    'https://yourdomain.com/dialog.html?time=' + encodeURIComponent(currentTime),
    { height: 30, width: 20 },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Dialog failed: " + result.error.message);
      }
    }
  );

  // Signal that the event handling is complete
  event.completed();
}