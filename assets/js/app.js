// Register Service Worker
if ("serviceWorker" in navigator) {
  navigator.serviceWorker
    .register("/service-worker.js")
    .then((registration) => {
      console.log("Service Worker Scope:", registration.scope);
    })
    .catch(console.error);
}

const app = new Framework7({
  root: '#app',
  routes: [
    { 
      path: '/', 
      componentUrl: '/index.html'
    },
    { 
      path: '/:page', 
      async: ({ params }, resolve) => {
        resolve({
          componentUrl: `/${params.page}.html`
        });
      },
    },
  ],
});