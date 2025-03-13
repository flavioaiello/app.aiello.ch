(function() {
    "use strict";
  
    Office.initialize = function(reason) {
      // No specific initialization needed for this add-in
    };
  
    window.onNewEmail = function(event) {
      try {
        const currentTime = new Date().toLocaleTimeString();
        
        Office.context.ui.displayDialogAsync(
          'https://app.aiello.ch/dialog.html?time=' + encodeURIComponent(currentTime),
          { height: 20, width: 30, displayInIframe: true },
          function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              event.completed({ allowEvent: false });
              return;
            }
            // Auto-close dialog after 5 seconds
            setTimeout(() => {
              if (asyncResult.value) {
                asyncResult.value.close();
              }
            }, 5000);
          }
        );
        
        event.completed({ allowEvent: true });
      } catch (error) {
        event.completed({ allowEvent: false });
      }
    };
  })();