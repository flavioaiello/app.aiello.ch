---
layout: default
title: Chat
---

<h1 class="mb-4">Dezentraler Chat mit OrbitDB und IPFS.js</h1>

<div id="status" class="alert alert-info">Initialisiere Chat...</div>

<!-- Chat Fenster -->
<div class="card">
  <div class="card-body" id="chat-window" style="height: 400px; overflow-y: scroll;">
    <!-- Nachrichten werden hier angezeigt -->
  </div>
</div>

<!-- Nachrichten Eingabe -->
<div class="input-group mt-3">
  <input type="text" id="messageInput" class="form-control" placeholder="Gib eine Nachricht ein">
  <button id="sendButton" class="btn btn-primary">Senden</button>
</div>

<!-- Benutzername Modal -->
<div class="modal fade" id="usernameModal" tabindex="-1" aria-labelledby="usernameModalLabel" aria-hidden="true">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <h1 class="modal-title fs-5" id="usernameModalLabel">Benutzername festlegen</h1>
      </div>
      <div class="modal-body">
        <input type="text" id="usernameInput" class="form-control" placeholder="Dein Benutzername">
        <p class="mt-2">
          Hinweis: Wenn die Verbindung zu IPFS nicht klappt, läuft der Chat eventuell nur lokal.
        </p>
      </div>
      <div class="modal-footer">
        <button type="button" id="saveUsername" class="btn btn-primary">Speichern</button>
      </div>
    </div>
  </div>
</div>
