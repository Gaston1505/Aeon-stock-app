const { onSchedule } = require("firebase-functions/v2/scheduler");
const { initializeApp } = require("firebase-admin/app");
const { getFirestore } = require("firebase-admin/firestore");
const { getMessaging } = require("firebase-admin/messaging");

initializeApp();

// Mismas frases y misma rotación por número de semana que src/App.jsx (FRASES_LUNES /
// FRASES_MIERCOLES / numeroDeSemana) — así lo que llega como push coincide con lo que ya se ve
// en Recordatorios dentro de la app. Duplicado acá porque una Cloud Function no puede importar
// el bundle de la app (runtime distinto, sin Vite).
const FRASES_LUNES = [
  "Arrancá la semana: date 10 minutos para repasar los procesos clave antes de que el día te absorba.",
  "Lunes de orden: planificá las prioridades de la semana antes del mediodía — el resto se acomoda solo.",
  "Buen día para abrir la agenda y agendar esa llamada a un cliente nuevo que venís postergando.",
  "Separá lo urgente de lo importante — esta semana se gana en las primeras horas del lunes.",
  "Revisá qué quedó pendiente de la semana pasada antes de sumarle algo nuevo a la lista.",
  "Elegí una sola cosa grande para cerrar esta semana, y arrancala hoy mismo.",
  "Buen momento para repasar el embudo de cotizaciones — ¿hay algo frío que valga la pena reactivar?",
  "Arrancá ordenando lo que tenés a la vista: lo que se ve prolijo, se vende mejor.",
];
const FRASES_MIERCOLES = [
  "Mitad de semana: ¿ya hiciste seguimiento a las cotizaciones que mandaste el lunes?",
  "Miércoles: buen día para medir cómo va la semana contra lo que planificaste el lunes.",
  "A mitad de camino — revisá si el foco comercial de la semana sigue siendo el correcto.",
  "Miércoles de números: mirá las ventas comprometidas, ¿hay algo por cerrar antes del viernes?",
  "Ya pasó la mitad de la semana — es buen momento para un mensaje de seguimiento a ese cliente importante.",
  "Lo que no se mide, no mejora — dale una mirada a tus indicadores comerciales de la semana.",
  "Medio camino recorrido: ajustá el rumbo si hace falta, todavía hay tiempo de cerrar bien la semana.",
  "Miércoles comercial: ¿qué cliente necesita que lo llamés hoy, no mañana?",
];

function numeroDeSemana(fecha) {
  const inicio = new Date(fecha.getFullYear(), 0, 1);
  const dias = Math.floor((fecha - inicio) / 86400000);
  return Math.floor((dias + inicio.getDay()) / 7);
}

async function mandarMotivacion(titulo, frases) {
  const hoy = new Date();
  const cuerpo = frases[numeroDeSemana(hoy) % frases.length];

  const snap = await getFirestore().collection("pushTokens").get();
  const tokens = snap.docs.map((d) => d.data().token).filter(Boolean);
  if (tokens.length === 0) {
    console.log("Sin dispositivos con notificaciones activadas — nada que mandar.");
    return;
  }

  const resp = await getMessaging().sendEachForMulticast({
    tokens,
    notification: { title: titulo, body: cuerpo },
    webpush: { fcmOptions: { link: "https://gaston1505.github.io/Aeon-stock-app/" } },
  });

  // Un token puede quedar inválido (dispositivo desinstaló la app, permiso revocado, etc.) — lo
  // sacamos de pushTokens para no seguir intentando mandarle en vano cada semana.
  const bajas = [];
  resp.responses.forEach((r, i) => {
    if (!r.success && ["messaging/registration-token-not-registered", "messaging/invalid-argument"].includes(r.error?.code)) {
      bajas.push(snap.docs[i].id);
    }
  });
  await Promise.all(bajas.map((id) => getFirestore().collection("pushTokens").doc(id).delete()));

  console.log(`Push "${titulo}" enviado: ${resp.successCount} ok, ${resp.failureCount} fallidos, ${bajas.length} token(s) de baja.`);
}

exports.motivacionLunes = onSchedule(
  { schedule: "0 8 * * 1", timeZone: "America/Asuncion" },
  () => mandarMotivacion("Arrancá la semana", FRASES_LUNES)
);

exports.motivacionMiercoles = onSchedule(
  { schedule: "0 8 * * 3", timeZone: "America/Asuncion" },
  () => mandarMotivacion("Mitad de semana", FRASES_MIERCOLES)
);
