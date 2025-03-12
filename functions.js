// Initialize Office.js
Office.onReady();

async function onNewMessageReceived(event) {
    try {
        const token = await getAccessToken();
        const newMessageId = Office.context.mailbox.item.itemId;
        const inboxFolderId = "inbox";

        // Get the "In-Reply-To" header of the new message
        const newMessageHeaders = await getMessageHeaders(token, newMessageId);
        const inReplyTo = getHeaderValue(newMessageHeaders, "In-Reply-To");
        if (!inReplyTo) {
            event.completed();
            return; // No "In-Reply-To" header, nothing to do
        }

        // Get or create the Archive folder
        const archiveFolderId = await getOrCreateFolderId(token, "Archive");

        // Search Inbox for a message with matching "Message-ID"
        const inboxMessages = await getMessagesInFolder(token, inboxFolderId);
        const matchingMessage = inboxMessages.find(async msg => {
            const headers = await getMessageHeaders(token, msg.id);
            const messageId = getHeaderValue(headers, "Message-ID");
            return messageId === inReplyTo;
        });

        if (matchingMessage) {
            // Move the matching message to Archive
            await moveMessage(token, matchingMessage.id, archiveFolderId);
        }

        event.completed();
    } catch (error) {
        console.error("Error in onNewMessageReceived:", error);
        event.completed();
    }
}

async function getAccessToken() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, result => {
            if (result.status === "succeeded") resolve(result.value);
            else reject(new Error(result.error.message));
        });
    });
}

async function getMessageHeaders(token, messageId) {
    const url = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${messageId}?$select=internetMessageHeaders`;
    const response = await fetch(url, {
        headers: { "Authorization": "Bearer " + token }
    });
    const data = await response.json();
    return data.internetMessageHeaders || [];
}

function getHeaderValue(headers, headerName) {
    const header = headers.find(h => h.name.toLowerCase() === headerName.toLowerCase());
    return header ? header.value : null;
}

async function getOrCreateFolderId(token, folderName) {
    const url = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders?$filter=displayName eq '${folderName}'`;
    const response = await fetch(url, {
        headers: { "Authorization": "Bearer " + token }
    });
    const data = await response.json();

    if (data.value && data.value.length > 0) {
        return data.value[0].id;
    } else {
        const createUrl = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders`;
        const createResponse = await fetch(createUrl, {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ displayName: folderName })
        });
        const createData = await createResponse.json();
        if (createResponse.ok) return createData.id;
        throw new Error("Failed to create folder");
    }
}

async function getMessagesInFolder(token, folderId) {
    const url = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/${folderId}/messages?$select=id`;
    const response = await fetch(url, {
        headers: { "Authorization": "Bearer " + token }
    });
    const data = await response.json();
    return data.value || [];
}

async function moveMessage(token, messageId, destinationFolderId) {
    const url = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${messageId}/move`;
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ destinationId: destinationFolderId })
    });
    if (!response.ok) throw new Error("Failed to move message");
}