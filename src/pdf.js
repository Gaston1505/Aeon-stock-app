import { PDFDocument, rgb, StandardFonts } from "pdf-lib";

// ---------- Company letterhead (matches the real quotes) ----------
const COMPANY = {
  razonSocial: "Quantum Investments S.A.",
  direccion: "San Juan XXIII esq Juan Max Boettner.",
  direccion2: "Edif. Park Plaza Piso 11 Oficina D. Asunción - Paraguay",
  telefonos: "+595 976 167335    +595 976 144599",
  emails: "ggibernau@aeon.com.py    info@aeon.com.py",
  firmante: "Gaston Gibernau",
  ruc: "RUC 80116329-3",
};

const LEGAL_TEXT =
  "TODOS LOS PRECIOS SON EN DOLARES E IVA INCLUIDO. La cotización es válida por 30 días. No incluye instalación. " +
  "Garantía 1 año por daños de fábrica desde su instalación, extendible a 36 meses, siempre y cuando se realice " +
  "servicio de mantenimiento oficial antes de los 12 y 24 meses respectivamente desde su instalacion.";

const ACCENT = rgb(0x1f / 255, 0x3a / 255, 0x5f / 255);
const ACCENT_LIGHT = rgb(0xe8 / 255, 0xee / 255, 0xf5 / 255);
const BORDER = rgb(0xe2 / 255, 0xe6 / 255, 0xeb / 255);
const INK = rgb(0x16 / 255, 0x20 / 255, 0x2a / 255);
const MUTED = rgb(0x64 / 255, 0x74 / 255, 0x8b / 255);
const WHITE = rgb(1, 1, 1);

// ---------- Número a letras (Español) ----------
const UNIDADES = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"];
const DIECIS = ["diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"];
const DECENAS = ["", "", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"];
const CENTENAS = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];

function tresDigitos(n) {
  if (n === 0) return "";
  if (n === 100) return "cien";
  const c = Math.floor(n / 100);
  const resto = n % 100;
  let out = c > 0 ? CENTENAS[c] : "";
  if (resto > 0) {
    let parte;
    if (resto < 10) parte = UNIDADES[resto];
    else if (resto < 20) parte = DIECIS[resto - 10];
    else {
      const d = Math.floor(resto / 10);
      const u = resto % 10;
      parte = DECENAS[d] + (u > 0 ? " y " + UNIDADES[u] : "");
    }
    out = out ? `${out} ${parte}` : parte;
  }
  return out;
}

function numeroALetras(num) {
  num = Math.floor(Math.max(0, Number(num) || 0));
  if (num === 0) return "cero";
  const millones = Math.floor(num / 1000000);
  const miles = Math.floor((num % 1000000) / 1000);
  const resto = num % 1000;
  const partes = [];
  if (millones > 0) partes.push(millones === 1 ? "un millón" : `${tresDigitos(millones)} millones`);
  if (miles > 0) partes.push(miles === 1 ? "mil" : `${tresDigitos(miles)} mil`);
  if (resto > 0) partes.push(tresDigitos(resto));
  return partes.join(" ");
}

export function montoEnLetras(monto) {
  const entero = Math.floor(Number(monto) || 0);
  const centavos = Math.round(((Number(monto) || 0) - entero) * 100);
  const letras = numeroALetras(entero);
  const cap = letras.charAt(0).toUpperCase() + letras.slice(1);
  return `${cap} con ${String(centavos).padStart(2, "0")}/100`;
}

// ---------- Helpers ----------
function wrapText(font, text, size, maxWidth) {
  const words = String(text || "").split(/\s+/).filter(Boolean);
  const lines = [];
  let current = "";
  for (const word of words) {
    const test = current ? `${current} ${word}` : word;
    if (current && font.widthOfTextAtSize(test, size) > maxWidth) {
      lines.push(current);
      current = word;
    } else {
      current = test;
    }
  }
  lines.push(current);
  return lines;
}

function dataUrlToBytes(dataUrl) {
  const base64 = dataUrl.split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

async function fetchBytes(url) {
  const res = await fetch(url);
  if (!res.ok) return null;
  return new Uint8Array(await res.arrayBuffer());
}

function downloadBlob(bytes, filename, type) {
  const blob = new Blob([bytes], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ---------- Cotización PDF ----------
const PAGE_W = 595.28;
const PAGE_H = 841.89;
const MARGIN = 40;
const CONTENT_W = PAGE_W - MARGIN * 2;

export async function generateCotizacionPdf(cotizacion) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  const base = import.meta.env.BASE_URL;
  const [logoBytes, firmaBytes] = await Promise.all([
    fetchBytes(`${base}aeon-logo.jpg`),
    fetchBytes(`${base}generated/firma.png`),
  ]);
  const logoImg = logoBytes ? await pdf.embedJpg(logoBytes) : null;
  const firmaImg = firmaBytes ? await pdf.embedPng(firmaBytes) : null;

  const lineas = cotizacion.lineas || [];
  const showEspec = lineas.some((l) => l.especValor);

  // Column layout
  const colCodigo = 75;
  const colFoto = 55;
  const colEspec = showEspec ? 55 : 0;
  const colCant = 32;
  const colPrecio = 62;
  const colTotal = 68;
  const colDescripcion = CONTENT_W - colCodigo - colFoto - colEspec - colCant - colPrecio - colTotal;

  let page = pdf.addPage([PAGE_W, PAGE_H]);
  let y = PAGE_H - MARGIN;

  function newPage() {
    page = pdf.addPage([PAGE_W, PAGE_H]);
    y = PAGE_H - MARGIN;
  }
  function ensureSpace(h) {
    if (y - h < MARGIN) newPage();
  }
  function text(t, x, yy, opts = {}) {
    page.drawText(String(t ?? ""), { x, y: yy, size: opts.size || 8, font: opts.bold ? bold : font, color: opts.color || INK });
  }
  function rect(x, yy, w, h, opts = {}) {
    page.drawRectangle({ x, y: yy, width: w, height: h, color: opts.fill, borderColor: opts.border, borderWidth: opts.border ? 0.5 : 0 });
  }

  // Header
  if (logoImg) {
    const w = 90;
    const h = (logoImg.height / logoImg.width) * w;
    page.drawImage(logoImg, { x: MARGIN, y: y - h, width: w, height: h });
    y -= h + 4;
  }
  text(COMPANY.razonSocial, MARGIN, y, { bold: true, size: 9 });
  y -= 11;
  text(COMPANY.direccion, MARGIN, y, { size: 7.5, color: MUTED });
  y -= 9;
  text(COMPANY.direccion2, MARGIN, y, { size: 7.5, color: MUTED });
  y -= 9;
  text(COMPANY.telefonos, MARGIN, y, { size: 7.5, color: MUTED });
  y -= 9;
  text(COMPANY.emails, MARGIN, y, { size: 7.5, color: MUTED });

  // Fecha (top right)
  const fechaLabel = "Fecha:";
  const fechaVal = fmtFecha(cotizacion.fecha);
  text(fechaLabel, PAGE_W - MARGIN - 110, PAGE_H - MARGIN, { bold: true, size: 8 });
  text(fechaVal, PAGE_W - MARGIN - 60, PAGE_H - MARGIN, { size: 8 });

  y -= 14;

  // Title bar
  rect(MARGIN, y - 18, CONTENT_W, 18, { fill: ACCENT });
  const title = `COTIZACIÓN ${(cotizacion.categoria || "").toUpperCase()}`.trim();
  const titleW = bold.widthOfTextAtSize(title, 10);
  text(title, MARGIN + CONTENT_W / 2 - titleW / 2, y - 13, { bold: true, size: 10, color: WHITE });
  y -= 18;

  // Cliente / Obra
  const rowH = 16;
  rect(MARGIN, y - rowH, CONTENT_W, rowH, { border: BORDER });
  text("Cliente:", MARGIN + 4, y - 11, { bold: true, size: 8 });
  text(cotizacion.cliente || "", MARGIN + 60, y - 11, { size: 8 });
  y -= rowH;
  rect(MARGIN, y - rowH, CONTENT_W, rowH, { border: BORDER });
  text("Obra:", MARGIN + 4, y - 11, { bold: true, size: 8 });
  text(cotizacion.obra || "", MARGIN + 60, y - 11, { size: 8 });
  y -= rowH;

  // Table header
  function drawTableHeader() {
    const startX = MARGIN;
    rect(startX, y - 20, CONTENT_W, 20, { fill: ACCENT_LIGHT });
    let cx = startX;
    const headers = [
      ["Producto", colCodigo],
      ["Foto Ref.", colFoto],
      ["Descripción", colDescripcion],
      ...(showEspec ? [[lineas.find((l) => l.especLabel)?.especLabel || "Espec.", colEspec]] : []),
      ["Cant.", colCant],
      ["Precio Unit. U$S", colPrecio],
      ["TOTAL U$S", colTotal],
    ];
    headers.forEach(([label, w]) => {
      const lw = bold.widthOfTextAtSize(label, 6.5);
      text(label, cx + w / 2 - lw / 2, y - 13, { bold: true, size: 6.5, color: ACCENT });
      cx += w;
    });
    y -= 20;
  }
  drawTableHeader();

  // Rows
  let subtotal = 0;
  for (const linea of lineas) {
    const cant = Number(linea.cantidad) || 0;
    const precio = Number(linea.precioUnit) || 0;
    const total = cant * precio;
    subtotal += total;

    const descLines = wrapText(font, linea.descripcion, 6.5, colDescripcion - 6);
    const fotoH = linea.foto ? 34 : 0;
    const rowH2 = Math.max(descLines.length * 8 + 6, fotoH + 6, 22);

    ensureSpace(rowH2 + 20);
    if (y === PAGE_H - MARGIN) drawTableHeader(); // continued on new page

    let cx = MARGIN;
    rect(cx, y - rowH2, colCodigo, rowH2, { border: BORDER });
    text(linea.codigo || "", cx + 3, y - 11, { size: 6.5 });
    cx += colCodigo;

    rect(cx, y - rowH2, colFoto, rowH2, { border: BORDER });
    if (linea.foto) {
      try {
        const imgBytes = dataUrlToBytes(linea.foto);
        const img = await pdf.embedJpg(imgBytes);
        const iw = colFoto - 8;
        const ih = (img.height / img.width) * iw;
        page.drawImage(img, { x: cx + 4, y: y - rowH2 / 2 - ih / 2, width: iw, height: Math.min(ih, rowH2 - 6) });
      } catch (e) { /* skip broken image */ }
    }
    cx += colFoto;

    rect(cx, y - rowH2, colDescripcion, rowH2, { border: BORDER });
    descLines.forEach((line, i) => text(line, cx + 3, y - 9 - i * 8, { size: 6.5 }));
    cx += colDescripcion;

    if (showEspec) {
      rect(cx, y - rowH2, colEspec, rowH2, { border: BORDER });
      const ev = linea.especValor || "";
      const ew = font.widthOfTextAtSize(ev, 6.5);
      text(ev, cx + colEspec / 2 - ew / 2, y - rowH2 / 2 - 3, { size: 6.5 });
      cx += colEspec;
    }

    rect(cx, y - rowH2, colCant, rowH2, { border: BORDER });
    const cantStr = String(cant);
    text(cantStr, cx + colCant / 2 - font.widthOfTextAtSize(cantStr, 6.5) / 2, y - rowH2 / 2 - 3, { size: 6.5 });
    cx += colCant;

    rect(cx, y - rowH2, colPrecio, rowH2, { border: BORDER });
    const precioStr = fmtNum(precio);
    text(precioStr, cx + colPrecio - 4 - font.widthOfTextAtSize(precioStr, 6.5), y - rowH2 / 2 - 3, { size: 6.5 });
    cx += colPrecio;

    rect(cx, y - rowH2, colTotal, rowH2, { border: BORDER });
    const totalStr = fmtNum(total);
    text(totalStr, cx + colTotal - 4 - font.widthOfTextAtSize(totalStr, 6.5), y - rowH2 / 2 - 3, { size: 6.5 });

    y -= rowH2;
  }

  ensureSpace(90);

  // Sub-total
  rect(MARGIN, y - 16, CONTENT_W - colTotal, 16, { fill: ACCENT_LIGHT });
  text("Sub-total", MARGIN + 4, y - 11, { bold: true, size: 8, color: ACCENT });
  rect(MARGIN + CONTENT_W - colTotal, y - 16, colTotal, 16, { fill: ACCENT_LIGHT });
  const subStr = fmtNum(subtotal);
  text(subStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(subStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
  y -= 16;

  // Comentarios
  if (cotizacion.comentarios) {
    const cLines = wrapText(font, cotizacion.comentarios, 7, CONTENT_W - 90);
    const h = Math.max(cLines.length * 9 + 6, 16);
    rect(MARGIN, y - h, 80, h, { border: BORDER });
    text("Comentarios", MARGIN + 4, y - 11, { bold: true, size: 7.5 });
    rect(MARGIN + 80, y - h, CONTENT_W - 80, h, { border: BORDER });
    cLines.forEach((line, i) => text(line, MARGIN + 84, y - 11 - i * 9, { size: 7 }));
    y -= h;
  }

  y -= 6;

  // Total IVA incluido + Dólares
  rect(MARGIN, y - 16, CONTENT_W - colTotal, 16, { border: BORDER });
  text("Total Iva Incluido", MARGIN + 4, y - 11, { bold: true, size: 8 });
  rect(MARGIN + CONTENT_W - colTotal, y - 16, colTotal, 16, { fill: ACCENT_LIGHT });
  text(subStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(subStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
  y -= 16;

  rect(MARGIN, y - 16, CONTENT_W, 16, { border: BORDER });
  text("Dólares Americanos:", MARGIN + 4, y - 11, { bold: true, size: 7.5 });
  text(montoEnLetras(subtotal), MARGIN + 110, y - 11, { size: 7.5 });
  y -= 24;

  // Fecha entrega estimada
  const entregaTxt = cotizacion.fechaEntregaEstimada || "";
  const entregaLines = wrapText(font, entregaTxt, 8, CONTENT_W - 130);
  const entregaH = Math.max(entregaLines.length * 10 + 6, 20);
  rect(MARGIN, y - entregaH, CONTENT_W, entregaH, { border: BORDER });
  text("Fecha entrega estimada:", MARGIN + 4, y - 13, { bold: true, size: 8 });
  entregaLines.forEach((line, i) => text(line, MARGIN + 130, y - 13 - i * 10, { size: 8 }));
  y -= entregaH + 10;

  // Legal terms
  ensureSpace(70);
  const legalLines = wrapText(font, LEGAL_TEXT, 6.5, CONTENT_W - 10);
  const formaPagoLine = `Forma de pago sugerida: ${cotizacion.formaPago || "A conversar"}.`;
  const obsLine = `OBS: ${cotizacion.obs || "Productos a retirar de depósito."}`;
  const allLegal = [...legalLines, formaPagoLine, obsLine];
  const legalH = allLegal.length * 8 + 8;
  rect(MARGIN, y - legalH, CONTENT_W, legalH, { border: BORDER });
  allLegal.forEach((line, i) => text(line, MARGIN + 4, y - 10 - i * 8, { size: 6.5 }));
  y -= legalH + 30;

  // Firma
  ensureSpace(70);
  const sigX = PAGE_W - MARGIN - 160;
  if (firmaImg) {
    const w = 110;
    const h = (firmaImg.height / firmaImg.width) * w;
    page.drawImage(firmaImg, { x: sigX + 25, y: y - h + 6, width: w, height: h });
    y -= h - 4;
  } else {
    y -= 30;
  }
  page.drawLine({ start: { x: sigX, y }, end: { x: sigX + 160, y }, thickness: 0.5, color: MUTED });
  y -= 11;
  const nameW = font.widthOfTextAtSize(COMPANY.firmante, 8);
  text(COMPANY.firmante, sigX + 80 - nameW / 2, y, { size: 8 });
  y -= 10;
  const razW = font.widthOfTextAtSize(COMPANY.razonSocial, 7.5);
  text(COMPANY.razonSocial, sigX + 80 - razW / 2, y, { size: 7.5 });
  y -= 9;
  const rucW = font.widthOfTextAtSize(COMPANY.ruc, 7.5);
  text(COMPANY.ruc, sigX + 80 - rucW / 2, y, { size: 7.5 });

  return pdf.save();
}

function fmtNum(n) {
  return (Number(n) || 0).toLocaleString("es-PY", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtFecha(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

export async function downloadCotizacionPdf(cotizacion) {
  const bytes = await generateCotizacionPdf(cotizacion);
  const nombre = `Cotizacion_${(cotizacion.cliente || "cliente").replace(/\s+/g, "_")}_${cotizacion.fecha || ""}.pdf`;
  downloadBlob(bytes, nombre, "application/pdf");
}

// ---------- Fichas técnicas combinadas ----------
export async function downloadFichasTecnicasPdf(cotizacion) {
  const vistos = new Set();
  const fichas = [];
  for (const linea of cotizacion.lineas || []) {
    if (linea.fichaTecnicaData && !vistos.has(linea.codigo)) {
      vistos.add(linea.codigo);
      fichas.push(linea);
    }
  }
  if (fichas.length === 0) return false;

  const out = await PDFDocument.create();
  for (const f of fichas) {
    try {
      const bytes = dataUrlToBytes(f.fichaTecnicaData);
      const src = await PDFDocument.load(bytes);
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach((p) => out.addPage(p));
    } catch (e) {
      console.error("No se pudo incorporar la ficha técnica de", f.codigo, e);
    }
  }
  if (out.getPageCount() === 0) return false;

  const bytes = await out.save();
  const nombre = `Fichas_tecnicas_${(cotizacion.cliente || "cliente").replace(/\s+/g, "_")}.pdf`;
  downloadBlob(bytes, nombre, "application/pdf");
  return true;
}
