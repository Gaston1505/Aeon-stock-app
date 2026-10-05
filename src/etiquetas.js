import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
import { anchosCode128 } from "./code128.js";

// Mismo formato que la etiqueta de serie de fábrica (70×20 mm, Code 128): el código del producto
// sin guiones ni espacios + 6 dígitos de serial. Las que genera la app arrancan en 900001 para
// no pisar los seriales de fábrica, que empiezan en 000001 en cada pedido.
export const SERIAL_BASE_APP = 900000;

export function claveCodigo(codigo) {
  return String(codigo || "").replace(/[^A-Za-z0-9]/g, "").toUpperCase();
}

export function textoSerial(codigo, numero) {
  return `${claveCodigo(codigo)}${String(numero).padStart(6, "0")}`;
}

// Inversa: de un texto leído ("AEAC2T30ON000001") al producto del catálogo y el N° de unidad.
// Toma el código de catálogo más largo que sea prefijo del texto, con exactamente 6 dígitos al final.
export function parsearSerial(texto, productos) {
  const t = String(texto || "").trim().toUpperCase();
  const m = t.match(/^(.*?)(\d{6})$/);
  if (!m) return null;
  let mejor = null;
  for (const p of productos) {
    if (claveCodigo(p.nombre) === m[1] && (!mejor || p.nombre.length > mejor.nombre.length)) mejor = p;
  }
  return mejor ? { producto: mejor, numero: Number(m[2]), delaApp: Number(m[2]) > SERIAL_BASE_APP } : null;
}

const MM = 72 / 25.4;
const ETIQ_W = 70 * MM;
const ETIQ_H = 20 * MM;
const COLS = 2;
const FILAS = 14;
const PAGE_W = 595.28;
const PAGE_H = 841.89;

// `etiquetas`: [{ texto }] — una por unidad. A4 vertical, 2 columnas × 14 filas de 70×20 mm,
// con el borde marcado para recortar.
export async function generateEtiquetasPdf(etiquetas) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const negro = rgb(0, 0, 0);
  const gris = rgb(0.6, 0.6, 0.6);
  const margenX = (PAGE_W - COLS * ETIQ_W) / 2;
  const margenY = (PAGE_H - FILAS * ETIQ_H) / 2;
  const porPagina = COLS * FILAS;

  let page = null;
  etiquetas.forEach((et, idx) => {
    const pos = idx % porPagina;
    if (pos === 0) page = pdf.addPage([PAGE_W, PAGE_H]);
    const col = pos % COLS;
    const fila = Math.floor(pos / COLS);
    const x0 = margenX + col * ETIQ_W;
    const yTop = PAGE_H - margenY - fila * ETIQ_H;
    page.drawRectangle({ x: x0, y: yTop - ETIQ_H, width: ETIQ_W, height: ETIQ_H, borderColor: gris, borderWidth: 0.4 });

    const anchos = anchosCode128(et.texto);
    const modulos = anchos.reduce((a, b) => a + b, 0);
    const quiet = 10;
    const disponible = ETIQ_W - 6 * MM;
    const modulo = Math.min(0.4 * MM, disponible / (modulos + quiet * 2));
    const anchoCodigo = modulos * modulo;
    let x = x0 + (ETIQ_W - anchoCodigo) / 2;
    const altoBarras = 11 * MM;
    const yBarras = yTop - 2 * MM - altoBarras;
    anchos.forEach((m, i) => {
      if (i % 2 === 0) page.drawRectangle({ x, y: yBarras, width: m * modulo, height: altoBarras, color: negro });
      x += m * modulo;
    });
    const size = 8;
    const tw = font.widthOfTextAtSize(et.texto, size);
    page.drawText(et.texto, { x: x0 + (ETIQ_W - tw) / 2, y: yBarras - 4.2 * MM, size, font, color: negro });
  });
  if (etiquetas.length === 0) pdf.addPage([PAGE_W, PAGE_H]);
  return pdf.save();
}

export async function downloadEtiquetasPdf(etiquetas, nombre = "Etiquetas_AEON") {
  const bytes = await generateEtiquetasPdf(etiquetas);
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${nombre}.pdf`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
