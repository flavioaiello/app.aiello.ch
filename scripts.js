Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("start-button").onclick = startleaner;
    }
});

function log(message) {
    const logDiv = document.getElementById("log");
    logDiv.innerHTML += message + "<br>";
}

async function getAccessToken() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
            if (result.status === "succeeded") {
                resolve(result.value);
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

async function getFolderIdByName(token, folderName) {
    const url = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders?$filter=displayName eq '${folderName}'`;
    const response = await fetch(url, {
        headers: { "Authorization": "Bearer " + token }
    });
    const data = await response.json();
    if (data.value && data.value.length > 0) {
        return data.value[0].id;
    } else {
        throw new Error(`Folder '${folderName}' not found.`);
    }
}

async function getMessagesInFolder(token, folderId) {
    const url = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/${folderId}/messages?$select=id,internetMessageId`;
    const response = await fetch(url, {
        headers: { "Authorization": "Bearer " + token }
    });
    const data = await response.json();
    return data.value || [];
}

async function getSubfolders(token, folderId) {
    const url = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/${folderId}/childFolders`;
    const response = await fetch(url, {
        headers: { "Authorization": "Bearer " + token }
    });
    const data = await response.json();
    return data.value || [];
}

async function getAllMessages(token, folderId) {
    let allMessages = await getMessagesInFolder(token, folderId);
    const subfolders = await getSubfolders(token, folderId);
    for (const subfolder of subfolders) {
        const subMessages = await getMessagesInFolder(token, subfolder.id);
        allMessages = allMessages.concat(subMessages);
    }
    return allMessages;
}

async function getHeadersForMessages(token, messageIds) {
    const batchSize = 20;
    const headersMap = new Map();
    for (let i = 0; i < messageIds.length; i += batchSize) {
        const batchIds = messageIds.slice(i, i + batchSize);
        const requests = batchIds.map((id, index) => ({
            id: index.toString(),
            method: "GET",
            url: `/me/messages/${id}?$select=internetMessageHeaders`
        }));
        const batchBody = { requests };
        const response = await fetch(`${Office.context.mailbox.restUrl}/v2.0/$batch`, {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(batchBody)
        });
        const data = await response.json();
        data.responses.forEach(resp => {
            if (resp.status === 200) {
                headersMap.set(batchIds[parseInt(resp.id)], resp.body.internetMessageHeaders);
            }
        });
    }
    return messageIds.map(id => headersMap.get(id) || []);
}

async function moveMessages(token, messageIds, destinationFolderId) {
    const batchSize = 20;
    for (let i = 0; i < messageIds.length; i += batchSize) {
        const batchIds = messageIds.slice(i, i + batchSize);
        const requests = batchIds.map((id, index) => ({
            id: index.toString(),
            method: "POST",
            url: `/me/messages/${id}/move`,
            body: { destinationId: destinationFolderId }
        }));
        const batchBody = { requests };
        await fetch(`${Office.context.mailbox.restUrl}/v2.0/$batch`, {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(batchBody)
        });
    }
}

async function startleaner() {
    try {
        log("Starting leaner process...");
        const token = await getAccessToken();
        const inboxFolderId = "inbox";
        const archiveFolderId = await getFolderIdByName(token, "Archive");
        const messages = await getAllMessages(token, inboxFolderId);
        log(`Found ${messages.length} messages.`);
        if (messages.length === 0) {
            log("No messages to process.");
            return;
        }
        const headersList = await getHeadersForMessages(token, messages.map(m => m.id));
        const referencedIds = new Set();
        headersList.forEach(headers => {
            headers.forEach(header => {
                if (header.name === "References" || header.name === "In-Reply-To") {
                    const ids = header.value.split(" ").filter(id => id);
                    ids.forEach(id => referencedIds.add(id));
                }
            });
        });
        const messagesToArchive = messages.filter(m => referencedIds.has(m.internetMessageId));
        log(`Found ${messagesToArchive.length} messages to archive.`);
        if (messagesToArchive.length > 0) {
            await moveMessages(token, messagesToArchive.map(m => m.id), archiveFolderId);
            log(`Archived ${messagesToArchive.length} messages.`);
        } else {
            log("No messages to archive.");
        }
    } catch (error) {
        log("Error: " + error.message);
    }
}
