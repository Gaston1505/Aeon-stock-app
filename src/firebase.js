import { initializeApp } from "firebase/app";
import { initializeFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

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
export const db = initializeFirestore(app, { ignoreUndefinedProperties: true });

// Used only for ficha técnica PDFs — those can be bigger than the 1MB
// Firestore document limit, unlike the small compressed photos we keep
// as base64 in Firestore documents (equipos, remitos, etc.).
export const storage = getStorage(app);
