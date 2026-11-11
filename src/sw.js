import { cleanupOutdatedCaches, createHandlerBoundToURL, precacheAndRoute } from "workbox-precaching";
import { NavigationRoute, registerRoute } from "workbox-routing";
import { clientsClaim } from "workbox-core";

self.skipWaiting();
clientsClaim();
cleanupOutdatedCaches();
precacheAndRoute(self.__WB_MANIFEST);

// Mismo criterio que antes con el generateSW de Workbox: la app no tiene rutas propias (todo es
// un solo path con estado en React), así que cualquier navegación cae a index.html — salvo que
// la URL termine en una extensión de archivo real (manual.html, catalogo.pdf, etc.), a la que hay
// que dejar llegar de verdad en vez de servirle la app.
registerRoute(new NavigationRoute(createHandlerBoundToURL("index.html"), {
  denylist: [/\.[^/]+$/],
}));

// Notificaciones push (Firebase Cloud Messaging) en segundo plano — con la app cerrada o la
// pestaña en otro lado. Mismo firebaseConfig que src/firebase.js (estos valores no son secretos,
// ver el comentario ahí); no se puede importar ese módulo acá porque el service worker corre en
// su propio contexto, sin acceso al bundle de la app.
importScripts("https://www.gstatic.com/firebasejs/11.1.0/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/11.1.0/firebase-messaging-compat.js");

firebase.initializeApp({
  apiKey: "AIzaSyBFLyjp_fuT-HT-UkZTtgNsxRzk-wmZbG0",
  authDomain: "aeon-stock-app.firebaseapp.com",
  projectId: "aeon-stock-app",
  storageBucket: "aeon-stock-app.firebasestorage.app",
  messagingSenderId: "939610902966",
  appId: "1:939610902966:web:b8bd4ff1e1f193eccc7117",
});

const messaging = firebase.messaging();
messaging.onBackgroundMessage((payload) => {
  const titulo = payload.notification?.title || "AEON";
  const cuerpo = payload.notification?.body || "";
  self.registration.showNotification(titulo, {
    body: cuerpo,
    icon: "icons/icon-192.png",
    badge: "icons/icon-192.png",
  });
});
