import { initializeApp } from "firebase/app";
import { initializeFirestore, persistentLocalCache, persistentSingleTabManager } from "firebase/firestore";
import { getAuth } from "firebase/auth";
import { getStorage } from "firebase/storage";
import { getMessaging, isSupported as isMessagingSupported } from "firebase/messaging";

// These values are safe to expose in client-side code — Firebase's web
// config is not a secret. Access is controlled by Firestore security rules,
// not by hiding this object. See firestore.rules in this repo.
const firebaseConfig = {
  apiKey: "AIzaSyBFLyjp_fuT-HT-UkZTtgNsxRzk-wmZbG0",
  authDomain: "aeon-stock-app.firebaseapp.com",
  projectId: "aeon-stock-app",
  storageBucket: "aeon-stock-app.firebasestorage.app",
  messagingSenderId: "939610902966",
  appId: "1:939610902966:web:b8bd4ff1e1f193eccc7117",
};

const app = initializeApp(firebaseConfig);

// ignoreUndefinedProperties guards against accidental `undefined` fields
// (Firestore rejects writes containing them, unlike localStorage/JSON).
// localCache (IndexedDB) persiste lo ya leído entre sesiones — sin esto, cada vez que se abre la
// app (sin importar cuántas veces ya se abrió antes en el mismo dispositivo) hay que esperar a que
// TODAS las colecciones suscriptas bajen de cero por red antes de mostrar algo; con esto, lo que ya
// está en el dispositivo se muestra al toque mientras Firestore sincroniza los cambios en segundo
// plano. persistentSingleTabManager porque la app no necesita compartir el cache entre pestañas.
export const db = initializeFirestore(app, {
  ignoreUndefinedProperties: true,
  localCache: persistentLocalCache({ tabManager: persistentSingleTabManager() }),
});

export const auth = getAuth(app);

export const storage = getStorage(app);

// Notificaciones push (Firebase Cloud Messaging): no todos los navegadores lo soportan (Safari
// viejo, contexto sin service worker) — getMessaging() tira si no está soportado, por eso se
// resuelve de forma async acá y el resto de la app recibe null en vez de que explote.
export const messagingReady = isMessagingSupported().then((soportado) => (soportado ? getMessaging(app) : null));
