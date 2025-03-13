(function() {
    "use strict";
  
    Office.initialize = function(reason) {
      // Add-in initialization code if needed
    };
  
    window.onNewEmail = function(event) {
      try {
        const currentTime = new Date().toLocaleTimeString();
        
        Office.context.ui.displayDialogAsync(
          'https://app.aiello.ch/dialog.html?time=' + encodeURIComponent(currentTime),
          { height: 20, width: 30, displayInIframe: true },
          function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log("Dialog failed: " + asyncResult.error.message);
            }
          }
        );
        
        event.completed();
      } catch (error) {
        console.log("Error: " + error);
        event.completed();
      }
    };
  
  })();