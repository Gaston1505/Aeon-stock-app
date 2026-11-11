import { getToken, onMessage } from "firebase/messaging";
import { doc, setDoc, serverTimestamp } from "firebase/firestore";
import { messagingReady, db } from "./firebase";

// Clave pública VAPID (Project Settings → Cloud Messaging → Web Push certificates, en Firebase
// Console) — distinta de las claves de API: esta identifica el remitente ante el navegador para
// que acepte el push, y es segura de exponer en el cliente (no es un secreto).
const VAPID_KEY = "BKEG4fV4SjEWyse99g4Ai9dC2ZsVTUUrB1r42m1ARZsghPgkktYV90PslyCEiLWfF-ZFUm5TJagbES3F17IIU1I";

export async function notificacionesSoportadas() {
  if (!("Notification" in window) || !("serviceWorker" in navigator)) return false;
  return !!(await messagingReady);
}

// Pide permiso, saca el token de este dispositivo/navegador y lo guarda en Firestore
// (pushTokens/{token}) para que la rutina de Monday/Wednesday — y a futuro cualquier otro aviso
// — le pueda mandar un push. Un dispositivo = un token = un documento, así que reactivar en el
// mismo dispositivo simplemente pisa el mismo doc en vez de duplicar.
export async function activarNotificacionesPush() {
  if (!VAPID_KEY) throw new Error("Falta configurar la clave VAPID todavía.");
  const messaging = await messagingReady;
  if (!messaging) throw new Error("Este navegador no soporta notificaciones push.");
  const permiso = await Notification.requestPermission();
  if (permiso !== "granted") throw new Error("Notificaciones no permitidas en este navegador/dispositivo.");
  const registration = await navigator.serviceWorker.ready;
  const token = await getToken(messaging, { vapidKey: VAPID_KEY, serviceWorkerRegistration: registration });
  if (!token) throw new Error("No se pudo generar el token de notificaciones.");
  await setDoc(doc(db, "pushTokens", token), {
    token, actualizadoEn: serverTimestamp(), userAgent: navigator.userAgent,
  });
  return token;
}

// Si la app está abierta y en primer plano cuando llega un push, Firebase NO dispara el
// onBackgroundMessage del service worker (eso es solo para segundo plano/cerrada) — hay que
// mostrar el aviso a mano con esto, o se pierde en silencio mientras estás mirando la pantalla.
export async function escucharNotificacionesEnPrimerPlano(onMensaje) {
  const messaging = await messagingReady;
  if (!messaging) return () => {};
  return onMessage(messaging, (payload) => onMensaje(payload));
}
