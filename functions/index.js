// Disparo del primer deploy a Firebase (ahora con el proyecto en Blaze y el secreto de GitHub
// Actions configurado).
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

// Manda el mismo push a todos los dispositivos activados, y de paso limpia los tokens que ya no
// sirven (dispositivo desinstaló la app, permiso revocado) para no seguir intentándoles en vano.
async function enviarPushATodos(titulo, cuerpo) {
  const snap = await getFirestore().collection("pushTokens").get();
  const tokens = snap.docs.map((d) => d.data().token).filter(Boolean);
  if (tokens.length === 0) {
    console.log("Sin dispositivos con notificaciones activadas — nada que mandar.");
    return { successCount: 0, failureCount: 0 };
  }

  const resp = await getMessaging().sendEachForMulticast({
    tokens,
    notification: { title: titulo, body: cuerpo },
    webpush: { fcmOptions: { link: "https://gaston1505.github.io/Aeon-stock-app/" } },
  });

  const bajas = [];
  resp.responses.forEach((r, i) => {
    if (!r.success && ["messaging/registration-token-not-registered", "messaging/invalid-argument"].includes(r.error?.code)) {
      bajas.push(snap.docs[i].id);
    }
  });
  await Promise.all(bajas.map((id) => getFirestore().collection("pushTokens").doc(id).delete()));

  console.log(`Push "${titulo}" enviado: ${resp.successCount} ok, ${resp.failureCount} fallidos, ${bajas.length} token(s) de baja.`);
  return resp;
}

async function mandarMotivacion(titulo, frases) {
  const hoy = new Date();
  const cuerpo = frases[numeroDeSemana(hoy) % frases.length];
  await enviarPushATodos(titulo, cuerpo);
}

exports.motivacionLunes = onSchedule(
  { schedule: "0 8 * * 1", timeZone: "America/Asuncion" },
  () => mandarMotivacion("Arrancá la semana", FRASES_LUNES)
);

exports.motivacionMiercoles = onSchedule(
  { schedule: "0 8 * * 3", timeZone: "America/Asuncion" },
  () => mandarMotivacion("Mitad de semana", FRASES_MIERCOLES)
);

// Avisos del CRM ("próximo contacto"): corre cada 15 minutos y manda el push apenas se cumple el
// momento que se eligió al cargarlo (avisoEn, calculado en el navegador — ver calcularAvisoEn en
// src/App.jsx). Filtra por proximoContactoHecho == false nomás (un solo campo, sin índice
// compuesto) y el resto (avisoEnviado, avisoEn vencido) lo resuelve acá — el volumen de esta
// colección es chico, no hace falta más que eso.
exports.avisosCrm = onSchedule(
  { schedule: "every 15 minutes", timeZone: "America/Asuncion" },
  async () => {
    const db = getFirestore();
    const snap = await db.collection("historialClientes").where("proximoContactoHecho", "==", false).get();
    const ahora = Date.now();
    const vencidos = snap.docs.filter((doc) => {
      const d = doc.data();
      return d.avisoEn && !d.avisoEnviado && d.avisoEn <= ahora;
    });
    if (vencidos.length === 0) return;
    for (const doc of vencidos) {
      const d = doc.data();
      await enviarPushATodos(`CRM — Contactar a ${d.empresaNombre}`, d.texto || "Tenés un seguimiento pendiente.");
      await doc.ref.update({ avisoEnviado: true });
    }
    console.log(`Avisos de CRM enviados: ${vencidos.length}.`);
  }
);

// Avisos de Mantenimiento postventa ("Programar visita"): dos por visita, independientes —
// avisoEn (configurable: mismo día o el día antes, a la hora elegida — ver calcularAvisoEn en
// src/App.jsx) y avisoDiaEn (fijo, el día mismo de la visita a primera hora, con el detalle de
// qué toca ese día, para que nunca dependa de haber configurado bien el primero). Mismo criterio
// de query liviana que avisosCrm: un solo campo booleano (tieneAvisoPendiente) para no necesitar
// índice compuesto, y acá adentro se resuelve cuál de los dos avisos corresponde mandar.
function resumenUnidadesVisita(unidades) {
  return (unidades || [])
    .filter((u) => !u.hecho)
    .map((u) => {
      const n = (u.equipos || []).length;
      return `Depto ${u.apartamento} (${n} equipo${n === 1 ? "" : "s"})`;
    })
    .join(", ");
}
exports.avisosMantenimiento = onSchedule(
  { schedule: "every 15 minutes", timeZone: "America/Asuncion" },
  async () => {
    const db = getFirestore();
    const snap = await db.collection("visitasPostventa").where("tieneAvisoPendiente", "==", true).get();
    const ahora = Date.now();
    let enviados = 0;
    for (const doc of snap.docs) {
      const v = doc.data();
      const pendientes = (v.unidades || []).filter((u) => !u.hecho);
      // Si ya no queda ninguna unidad pendiente (se completó la visita antes de que tocara algún
      // aviso), no tiene sentido mandar un push vacío — se da de baja directo.
      if (pendientes.length === 0) { await doc.ref.update({ tieneAvisoPendiente: false }); continue; }
      const resumen = resumenUnidadesVisita(v.unidades);
      const patch = {};
      if (v.avisoEn && !v.avisoEnviado && v.avisoEn <= ahora) {
        await enviarPushATodos(`Mantenimiento programado — ${v.edificioNombre || "Edificio"}`, `Para el ${v.fechaProgramada}: ${resumen}`);
        patch.avisoEnviado = true;
        enviados++;
      }
      if (v.avisoDiaEn && !v.avisoDiaEnviado && v.avisoDiaEn <= ahora) {
        await enviarPushATodos(`Hoy toca mantenimiento — ${v.edificioNombre || "Edificio"}`, resumen);
        patch.avisoDiaEnviado = true;
        enviados++;
      }
      const avisoEnviadoFinal = patch.avisoEnviado ?? v.avisoEnviado;
      const avisoDiaEnviadoFinal = patch.avisoDiaEnviado ?? v.avisoDiaEnviado;
      if (avisoEnviadoFinal && avisoDiaEnviadoFinal) patch.tieneAvisoPendiente = false;
      if (Object.keys(patch).length > 0) await doc.ref.update(patch);
    }
    if (enviados > 0) console.log(`Avisos de Mantenimiento enviados: ${enviados}.`);
  }
);
