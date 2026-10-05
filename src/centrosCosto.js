// Cuadro "COMERCIO EXTERIOR" de CENTROS DE UTILIDADES (los demás cuadros del archivo se descartan).
export const CENTROS_COSTO = [
  { codigo: "CO001", nombre: "COMEX", codigoLargo: "COMEX", detalle: "Importaciones / exportaciones / productos / logística" },
  { codigo: "CO002", nombre: "Morra 1 y 2", codigoLargo: "COAEM1Y2", detalle: "AEON - Morra 1 y 2 · contratos de obras" },
  { codigo: "CO003", nombre: "Home Palace", codigoLargo: "COAEHP", detalle: "AEON - Home Palace · contratos de obras" },
  { codigo: "CO004", nombre: "Laurus", codigoLargo: "COAELAU", detalle: "AEON - Laurus · contratos de obras" },
  { codigo: "CO005", nombre: "Hit", codigoLargo: "COAEHIT", detalle: "AEON - Hit · contratos de obras" },
  { codigo: "CO006", nombre: "Sun Palace", codigoLargo: "COAESUN", detalle: "AEON - Sun Palace · contratos de obras" },
  { codigo: "CO007", nombre: "Designio", codigoLargo: "COAEDES", detalle: "AEON - Designio · contratos de obras" },
];

export function centroCostoPorCodigo(codigo) {
  return CENTROS_COSTO.find((c) => c.codigo === codigo) || null;
}

// Sugerencia (siempre editable): si la obra es una de las que tienen centro propio, ese; si no, COMEX.
export function centroCostoSugerido(obra) {
  const t = String(obra || "").toLowerCase();
  if (/laurus/.test(t)) return "CO004";
  if (/home\s*palace/.test(t)) return "CO003";
  if (/morra/.test(t)) return "CO002";
  if (/\bhit\b/.test(t)) return "CO005";
  if (/sun\s*palace/.test(t)) return "CO006";
  if (/designio/.test(t)) return "CO007";
  return "CO001";
}

export function totalLineaPedido(l) {
  return Math.round((Number(l.cantidad) || 0) * (Number(l.precioUnit) || 0) * 100) / 100;
}
export function totalPedido(p) {
  return (p.lineas || []).reduce((acc, l) => acc + totalLineaPedido(l), 0);
}
