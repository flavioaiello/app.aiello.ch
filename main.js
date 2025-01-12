import { startOrbitDB } from '@orbitdb/liftoff';

let db;

async function initOrbitDB() {
  const loader = document.getElementById('loader');
  try {
    // Start OrbitDB with @orbitdb/liftoff
    const orbitdb = await startOrbitDB();

    // Open an eventlog database for the group chat
    db = await orbitdb.open('group-chat', {
      type: 'eventlog',
      accessController: { write: ['*'] },
    });

    // Subscribe to new messages
    db.events.on('replicated', renderMessages);
    db.events.on('write', renderMessages);

    // Load existing messages
    await renderMessages();

    loader.innerText = ''; // Clear the loader message
  } catch (error) {
    console.error('Error initializing OrbitDB:', error);
    document.getElementById('output').innerText = 'Error initializing OrbitDB. Check console for details.';
    loader.innerText = 'Could not connect to OrbitDB. You might be offline.';
  }
}

async function sendMessage() {
  const message = document.getElementById('messageInput').value;
  const username = localStorage.getItem('username') || 'Anonymous';

  if (message.trim()) {
    await db.add({ username, message, timestamp: new Date().toISOString() });
    document.getElementById('messageInput').value = ''; // Clear the input field
  }
}

async function renderMessages() {
  const chatWindow = document.getElementById('chat-window');
  const messages = await db.iterator({ limit: -1 }).collect();

  chatWindow.innerHTML = messages
    .map(
      (entry) =>
        `<p><strong>${entry.payload.value.username}:</strong> ${entry.payload.value.message} <small class="text-muted">${new Date(
          entry.payload.value.timestamp
        ).toLocaleTimeString()}</small></p>`
    )
    .join('');

  chatWindow.scrollTop = chatWindow.scrollHeight; // Scroll to bottom
}

// Initialize PWA Service Worker
function initServiceWorker() {
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker
        .register('/service-worker.js')
        .then((registration) => {
          console.log('Service Worker registered:', registration);
        })
        .catch((error) => {
          console.error('Service Worker registration failed:', error);
        });
    });
  }
}

function initApp() {
  initServiceWorker();
  initOrbitDB();
}

// Handle username setup
document.getElementById('saveUsername').addEventListener('click', () => {
  const username = document.getElementById('usernameInput').value;
  if (username.trim()) {
    localStorage.setItem('username', username);
    document.getElementById('usernameModal').classList.remove('show');
  }
});

document.getElementById('sendButton').addEventListener('click', sendMessage);

initApp();
