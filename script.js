// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in ready");
    }
});

function findReplyEmails() {
    try {
        const mailbox = Office.context.mailbox;
        document.getElementById('results').textContent = 'Searching...';

        // Get access token
        mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const accessToken = result.value;
                const restUrl = Office.context.mailbox.restUrl + 
                    "/v2.0/me/MailFolders/Inbox/messages?$top=100&$select=Subject,InReplyTo";

                // Fetch emails
                fetch(restUrl, {
                    method: 'GET',
                    headers: {
                        'Authorization': 'Bearer ' + accessToken,
                        'Accept': 'application/json'
                    }
                })
                .then(response => response.json())
                .then(data => {
                    const emails = data.value;
                    const replyEmails = emails.filter(email => email.inReplyTo !== null && email.inReplyTo !== undefined);
                    
                    // Prepare dialog options
                    const dialogUrl = 'https://your-domain.com/dialog.html'; // Update to your hosted dialog.html URL
                    const dialogOptions = {
                        width: 30,
                        height: 40,
                        displayInIframe: true
                    };

                    // Open dialog and send email data
                    Office.context.ui.displayDialogAsync(dialogUrl, dialogOptions, function(asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const dialog = asyncResult.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(event) {
                                dialog.close();
                                document.getElementById('results').textContent = 'Dialog closed.';
                            });
                            
                            // Send email data to dialog
                            const emailData = replyEmails.map(email => ({
                                subject: email.subject,
                                inReplyTo: email.inReplyTo
                            }));
                            dialog.messageChild(JSON.stringify({ emails: emailData }));
                            
                            document.getElementById('results').textContent = 
                                `Found ${replyEmails.length} reply emails. Check the dialog.`;
                        } else {
                            console.error('Dialog failed to open:', asyncResult.error);
                            document.getElementById('results').textContent = 'Error: Could not open dialog';
                        }
                    });
                })
                .catch(error => {
                    console.error('Error fetching emails:', error);
                    document.getElementById('results').textContent = 'Error: Could not fetch emails';
                });
            } else {
                console.error('Failed to get access token:', result.error);
            }
        });
    } catch (error) {
        console.error('Error in findReplyEmails:', error);
        document.getElementById('results').textContent = 'Error: Something went wrong';
    }
}