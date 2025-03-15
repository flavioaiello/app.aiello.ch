// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in ready");
    }
});

function iterateEmails() {
    try {
        // Get the mailbox
        const mailbox = Office.context.mailbox;
        
        // Convert to REST API format
        mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const accessToken = result.value;
                
                // Get the REST API endpoint
                const restUrl = Office.context.mailbox.restUrl + 
                    "/v2.0/me/MailFolders/Inbox/messages?$top=100";
                
                // Make the API call
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
                    const resultsDiv = document.getElementById('results');
                    resultsDiv.innerHTML = ''; // Clear previous results
                    
                    // Iterate through emails
                    emails.forEach((email, index) => {
                        console.log(`Email ${index + 1}: ${email.subject}`);
                        const emailDiv = document.createElement('div');
                        emailDiv.textContent = `${index + 1}. ${email.subject}`;
                        resultsDiv.appendChild(emailDiv);
                    });
                    
                    console.log(`Total emails processed: ${emails.length}`);
                })
                .catch(error => {
                    console.error('Error fetching emails:', error);
                    document.getElementById('results').textContent = 
                        'Error: Could not fetch emails';
                });
            }
        });
    } catch (error) {
        console.error('Error in iterateEmails:', error);
        document.getElementById('results').textContent = 
            'Error: Something went wrong';
    }
}