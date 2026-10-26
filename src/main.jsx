import React from "react";
import ReactDOM from "react-dom/client";
import { registerSW } from "virtual:pwa-register";
import App from "./App.jsx";
import "./index.css";

// El registro mínimo por defecto de la PWA solo llama a .register() una vez al cargar y nunca
// vuelve a chequear — en iOS Safari eso puede dejar un service worker viejo sirviendo la misma
// versión para siempre, sin detectar nunca que hubo un deploy nuevo. Acá se chequea a mano cada
// vez que la pestaña vuelve a primer plano (y cada 1h si queda abierta), y en cuanto un service
// worker nuevo toma control, se recarga sola — así no hace falta que el usuario cierre y abra la
// app a mano para ver la versión nueva.
registerSW({
  immediate: true,
  onRegisteredSW(_url, registration) {
    if (!registration) return;
    const chequear = () => registration.update().catch(() => {});
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") chequear();
    });
    setInterval(chequear, 60 * 60 * 1000);
  },
});
let refrescando = false;
navigator.serviceWorker?.addEventListener("controllerchange", () => {
  if (refrescando) return;
  refrescando = true;
  window.location.reload();
});

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
