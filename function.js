(function() {
  "use strict";

  Office.initialize = function(reason) {
    console.log("Add-in initialized: " + reason);
    // Optional: Log environment details
    console.log("Host: " + Office.context.host + ", Platform: " + Office.context.platform);
  };

  // Manual button trigger
  window.showTimeDialog = function(event) {
    console.log("showTimeDialog triggered at " + new Date().toLocaleTimeString());
    try {
      const currentTime = new Date().toLocaleTimeString();
      Office.context.ui.displayDialogAsync(
        "https://app.aiello.ch/dialog.html?time=" + encodeURIComponent(currentTime),
        { height: 20, width: 30, displayInIframe: true },
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Dialog failed: " + asyncResult.error.code + " - " + asyncResult.error.message);
          } else {
            console.log("Dialog opened successfully");
            setTimeout(() => {
              asyncResult.value.close();
              console.log("Dialog closed after 5 seconds");
            }, 5000);
          }
        }
      );
      event.completed();
    } catch (error) {
      console.log("Error in showTimeDialog: " + error.message);
      event.completed();
    }
  };

  // Event-based trigger (for testing)
  window.onNewEmail = function(event) {
    console.log("onNewEmail triggered at " + new Date().toLocaleTimeString());
    try {
      const currentTime = new Date().toLocaleTimeString();
      Office.context.ui.displayDialogAsync(
        "https://app.aiello.ch/dialog.html?time=" + encodeURIComponent(currentTime),
        { height: 20, width: 30, displayInIframe: true },
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Dialog failed in onNewEmail: " + asyncResult.error.message);
          } else {
            console.log("Dialog opened via event");
          }
        }
      );
      event.completed({ allowEvent: true });
    } catch (error) {
      console.log("Error in onNewEmail: " + error.message);
      event.completed({ allowEvent: false });
    }
  };
})();