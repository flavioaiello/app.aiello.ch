// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in ready");
    }
});

function iterateEmails() {
    try {
        const mailbox = Office.context.mailbox;
        
        // Get access token for REST API
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
                    const resultsDiv = document.getElementById('results');
                    resultsDiv.innerHTML = ''; // Clear previous results
                    
                    // Filter emails with In-Reply-To header
                    const replyEmails = emails.filter(email => email.inReplyTo !== null && email.inReplyTo !== undefined);
                    
                    if (replyEmails.length === 0) {
                        resultsDiv.textContent = 'No reply emails found in the first 100 inbox messages.';
                        console.log('No reply emails found');
                        return;
                    }
                    
                    // Display results
                    replyEmails.forEach((email, index) => {
                        const inReplyTo = email.inReplyTo;
                        console.log(`Reply Email ${index + 1}: ${email.subject} (In-Reply-To: ${inReplyTo})`);
                        const emailDiv = document.createElement('div');
                        emailDiv.textContent = `${index + 1}. ${email.subject} (In-Reply-To: ${inReplyTo})`;
                        resultsDiv.appendChild(emailDiv);
                    });
                    
                    console.log(`Total reply emails found: ${replyEmails.length}`);
                })
                .catch(error => {
                    console.error('Error fetching emails:', error);
                    document.getElementById('results').textContent = 
                        'Error: Could not fetch emails';
                });
            } else {
                console.error('Failed to get access token:', result.error);
            }
        });
    } catch (error) {
        console.error('Error in iterateEmails:', error);
        document.getElementById('results').textContent = 
            'Error: Something went wrong';
    }
}