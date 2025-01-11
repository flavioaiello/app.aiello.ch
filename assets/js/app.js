// assets/js/app.js

let db, orbitdb, helia;
let username = '';

// UI-Referenzen
let sendButton, messageInput, chatWindow, status;

// Helia + OrbitDB initialisieren
async function initHeliaOrbitDB() {
  try {
    // Libp2p aufsetzen, Gossipsub als Klasse (Default-Export)
    const libp2p = await window._Libp2p({
      services: {
        pubsub: new window._GossipSub({
          allowPublishToZeroTopicPeers: true
        }),
        identify: window._Identify()
      }
    });

    // Helia erstellen
    helia = await window._Helia({ libp2p });
    console.log('Helia initialisiert', helia);

    // OrbitDB erstellen
    orbitdb = await window._OrbitDB({ ipfs: helia });
    console.log('OrbitDB initialisiert', orbitdb);

    // "chat-app" Datenbank oeffnen
    db = await orbitdb.open('chat-app');
    console.log('Datenbank geöffnet:', db.address.toString());

    // Laden aller Einträge
    await db.load();
    console.log('Datenbank geladen');

    // "update"-Event => wenn neue Eintraege hinzukommen
    db.events.on('update', () => {
      console.log('Neue Einträge in der DB');
      updateChatWindow();
    });

    // Erstes UI-Update
    updateChatWindow();
    return true;
  } catch (error) {
    console.error('Fehler beim Initialisieren von Helia/OrbitDB:', error);
    return false;
  }
}

function updateChatWindow() {
  chatWindow.innerHTML = '';
  for (const entry of db.iterator()) {
    const { type, content, timestamp, username: author } = entry.payload.value;
    if (type === 'message') {
      const div = document.createElement('div');
      div.classList.add('mb-2');

      const timeStr = new Date(timestamp).toLocaleTimeString();
      const sender = author || 'Unbekannt';
      const msgClass = (sender === username) ? 'text-end' : 'text-start';
      const badgeClass = (sender === username) ? 'bg-primary' : 'bg-secondary';

      div.innerHTML = `
        <span class="badge ${badgeClass}">${sender} - ${timeStr}</span>
        <p>${content}</p>
      `;
      div.classList.add(msgClass);
      chatWindow.appendChild(div);
    }
  }
  chatWindow.scrollTop = chatWindow.scrollHeight;
}

async function sendMessage() {
  const message = messageInput.value.trim();
  if (!message || !username) return;

  try {
    await db.add({
      type: 'message',
      username,
      content: message,
      timestamp: new Date().toISOString()
    });
    messageInput.value = '';
    status.className = 'alert alert-success';
    status.textContent = 'Nachricht gesendet!';
  } catch (err) {
    console.error('Fehler beim Senden der Nachricht:', err);
    status.className = 'alert alert-danger';
    status.textContent = 'Fehler beim Senden!';
  }
}

function setupUsernameModal() {
  const saveBtn = document.getElementById('saveUsername');
  const inputEl = document.getElementById('usernameInput');
  saveBtn.addEventListener('click', () => {
    const val = inputEl.value.trim();
    if (!val) {
      alert('Bitte einen Benutzernamen eingeben!');
      return;
    }
    username = val;
    localStorage.setItem('chatUsername', username);
    const modal = bootstrap.Modal.getInstance(document.getElementById('usernameModal'));
    modal.hide();
    status.className = 'alert alert-success';
    status.textContent = `Willkommen, ${username}!`;
  });
}

document.addEventListener('DOMContentLoaded', async () => {
  // UI-Elemente greifen
  chatWindow = document.getElementById('chat-window');
  sendButton = document.getElementById('sendButton');
  messageInput = document.getElementById('messageInput');
  status = document.getElementById('status');

  sendButton.addEventListener('click', sendMessage);
  messageInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') sendMessage();
  });

  setupUsernameModal();

  const savedUsername = localStorage.getItem('chatUsername');
  if (savedUsername) {
    username = savedUsername;
    status.className = 'alert alert-success';
    status.textContent = `Willkommen zurück, ${username}!`;
  } else {
    const modal = new bootstrap.Modal(document.getElementById('usernameModal'), {
      backdrop: 'static',
      keyboard: false
    });
    modal.show();
  }

  status.textContent = 'Initialisiere Chat...';
  const ok = await initHeliaOrbitDB();
  if (ok) {
    status.className = 'alert alert-success';
    status.textContent = 'PWA bereit! Helia + OrbitDB verbunden.';
  } else {
    status.className = 'alert alert-danger';
    status.textContent = 'Fehler beim Initialisieren von Helia/OrbitDB.';
  }
});