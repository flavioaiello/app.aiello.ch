Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        startInboxMonitoring();
    }
});

function startInboxMonitoring() {
    // Initial check
    checkAndArchiveInbox();
    
    // Set up periodic checking (every 5 minutes)
    setInterval(checkAndArchiveInbox, 5 * 60 * 1000);
}

async function checkAndArchiveInbox() {
    try {
        console.log(`${new Date().toISOString()}: Starting inbox check and archive...`);
        
        // Get access token for Graph API
        const accessToken = await getAccessToken();
        if (!accessToken) throw new Error("Failed to get access token");

        // Step 1: Gather messages from Inbox
        const inboxMessages = await getRecentMessages(accessToken, "Inbox");
        const totalMessages = inboxMessages.length;
        
        if (totalMessages === 0) {
            console.log(`${new Date().toISOString()}: No messages found in Inbox.`);
            return;
        }
        console.log(`${new Date().toISOString()}: Found ${totalMessages} messages.`);

        // Step 2: Extract headers and build thread data
        const emailList = [];
        const refsList = [];
        const replyList = [];

        for (const message of inboxMessages) {
            const headers = await getMessageHeaders(accessToken, message.id);
            const msgIdHeader = headers["Message-ID"] || `placeholder-${message.id}`;
            const msgReferencesHeader = headers["References"] || "";
            const msgInReplyToHeader = headers["In-Reply-To"] || "";

            emailList.push({ mailId: message.id, msgIdHeader });
            refsList.push(msgReferencesHeader);
            replyList.push(msgInReplyToHeader);
        }

        // Step 3: Categorize messages for archiving
        const messagesToArchiveIDs = [];
        const combinedRefs = [...refsList, ...replyList].join(" ");

        for (const email of emailList) {
            if (combinedRefs.includes(email.msgIdHeader)) {
                messagesToArchiveIDs.push(email.mailId);
            }
        }
        console.log(`${new Date().toISOString()}: To archive: ${messagesToArchiveIDs.length} messages.`);

        // Step 4: Archive messages
        let archivedCount = 0;
        if (messagesToArchiveIDs.length > 0) {
            const archiveFolderId = await getArchiveFolderId(accessToken);
            if (!archiveFolderId) throw new Error("Archive folder not found");

            for (const msgId of messagesToArchiveIDs) {
                try {
                    await moveMessage(accessToken, msgId, archiveFolderId);
                    archivedCount++;
                } catch (error) {
                    console.error(`${new Date().toISOString()}: Archive error for message ${msgId}: ${error.message}`);
                }
            }
        }
        console.log(`${new Date().toISOString()}: Archived ${archivedCount} messages.`);

    } catch (error) {
        console.error(`${new Date().toISOString()}: Error in checkAndArchiveInbox: ${error.message}`);
    }
}

// Helper: Get access token
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

// Helper: Get messages from a folder
async function getRecentMessages(accessToken, folderName) {
    const folderId = folderName.toLowerCase() === "inbox" ? "inbox" : await getFolderId(accessToken, folderName);
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages?$top=100`;
    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        }
    });
    if (!response.ok) throw new Error(`Failed to fetch messages: ${response.status}`);
    const data = await response.json();
    return data.value || [];
}

// Helper: Get message headers via Graph API
async function getMessageHeaders(accessToken, messageId) {
    const url = `https://graph.microsoft.com/v1.0/me/messages/${messageId}?$select=internetMessageHeaders`;
    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });
    if (!response.ok) throw new Error(`Failed to fetch headers: ${response.status}`);
    const data = await response.json();
    const headers = data.internetMessageHeaders || [];
    return headers.reduce((acc, header) => {
        acc[header.name] = header.value;
        return acc;
    }, {});
}

// Helper: Get folder ID by name
async function getFolderId(accessToken, folderName) {
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName eq '${folderName}'`;
    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });
    if (!response.ok) throw new Error(`Failed to fetch folders: ${response.status}`);
    const data = await response.json();
    const folder = data.value.find(f => f.displayName === folderName);
    return folder ? folder.id : null;
}

// Helper: Get Archive folder ID
async function getArchiveFolderId(accessToken) {
    return await getFolderId(accessToken, "Archive");
}

// Helper: Move a message to a folder
async function moveMessage(accessToken, messageId, destinationFolderId) {
    const url = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/move`;
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            destinationId: destinationFolderId
        })
    });
    if (!response.ok) throw new Error(`Failed to move message: ${response.status}`);
    return await response.json();
}

// Manual trigger for ribbon button
function manualScan(event) {
    checkAndArchiveInbox();
    event.completed();
}