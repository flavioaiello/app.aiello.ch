Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        startInboxMonitoring();
    }
});

function startInboxMonitoring() {
    // Initial check
    checkInbox();
    
    // Set up periodic checking (every 5 minutes)
    setInterval(checkInbox, 5 * 60 * 1000);
}

async function checkInbox() {
    try {
        const item = Office.context.mailbox.item;
        console.log("Checking inbox from Home pane...");

        // Get mailbox information
        const mailbox = Office.context.mailbox;
        
        // Using Microsoft Graph API (requires appropriate permissions)
        const accessToken = await getAccessToken();
        if (accessToken) {
            const messages = await getRecentMessages(accessToken);
            processMessages(messages);
        }
    } catch (error) {
        console.error("Error checking inbox:", error);
    }
}

async function getAccessToken() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}

async function getRecentMessages(accessToken) {
    const restUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages?$top=10`;
    
    const response = await fetch(restUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        }
    });
    
    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    return await response.json();
}

function processMessages(messages) {
    if (messages && messages.value) {
        messages.value.forEach(message => {
            console.log(`Subject: ${message.subject}`);
            console.log(`From: ${message.from.emailAddress.address}`);
            console.log(`Received: ${message.receivedDateTime}`);
            console.log('---');
            // Add your custom processing logic here
        });
    }
}

// Optional: Add a manual trigger function for the ribbon button
function manualScan(event) {
    checkInbox();
    event.completed();
}
