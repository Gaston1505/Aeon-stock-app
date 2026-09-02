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

// Paleta derivada del gris del logo AEON (#686D73), igual que en la app.
const ACCENT = rgb(0x56 / 255, 0x5a / 255, 0x5f / 255);
const ACCENT_LIGHT = rgb(0xeb / 255, 0xeb / 255, 0xec / 255);
const BORDER = rgb(0xe4 / 255, 0xe5 / 255, 0xe5 / 255);
const INK = rgb(0x1c / 255, 0x1e / 255, 0x20 / 255);
const MUTED = rgb(0x68 / 255, 0x6d / 255, 0x73 / 255);
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
  // Centers a single line of text both horizontally and vertically inside a
  // table cell whose top-left-ish corner is (cellX, cellTopY) with cellTopY
  // being the current "y" cursor (top of the row) — matches how rows are drawn.
  function centerText(str, cellX, cellTopY, cellW, cellH, opts = {}) {
    const { size = 6.5, bold: isBold = false, color } = opts;
    const f = isBold ? bold : font;
    const w = f.widthOfTextAtSize(str, size);
    text(str, cellX + cellW / 2 - w / 2, cellTopY - cellH / 2 - 3, { size, bold: isBold, color });
  }
  // Scales the image to fit entirely inside the cell (preserving aspect
  // ratio, like CSS object-fit: contain) and centers it — never distorts it.
  function drawImageContained(img, cellX, cellTopY, cellW, cellH, pad = 4) {
    const maxW = cellW - pad * 2;
    const maxH = cellH - pad * 2;
    const scale = Math.min(maxW / img.width, maxH / img.height);
    const iw = img.width * scale;
    const ih = img.height * scale;
    page.drawImage(img, { x: cellX + cellW / 2 - iw / 2, y: cellTopY - cellH / 2 - ih / 2, width: iw, height: ih });
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
  const title = "COTIZACIÓN ELECTRODOMÉSTICOS";
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
    centerText(linea.codigo || "", cx, y, colCodigo, rowH2);
    cx += colCodigo;

    rect(cx, y - rowH2, colFoto, rowH2, { border: BORDER });
    if (linea.foto) {
      try {
        const imgBytes = dataUrlToBytes(linea.foto);
        const img = await pdf.embedJpg(imgBytes);
        drawImageContained(img, cx, y, colFoto, rowH2, 4);
      } catch (e) { /* skip broken image */ }
    }
    cx += colFoto;

    rect(cx, y - rowH2, colDescripcion, rowH2, { border: BORDER });
    const descBlockH = descLines.length * 8;
    const descTop = y - (rowH2 - descBlockH) / 2 - 6;
    descLines.forEach((line, i) => text(line, cx + 3, descTop - i * 8, { size: 6.5 }));
    cx += colDescripcion;

    if (showEspec) {
      rect(cx, y - rowH2, colEspec, rowH2, { border: BORDER });
      centerText(linea.especValor || "", cx, y, colEspec, rowH2);
      cx += colEspec;
    }

    rect(cx, y - rowH2, colCant, rowH2, { border: BORDER });
    centerText(String(cant), cx, y, colCant, rowH2);
    cx += colCant;

    rect(cx, y - rowH2, colPrecio, rowH2, { border: BORDER });
    centerText(fmtNum(precio), cx, y, colPrecio, rowH2);
    cx += colPrecio;

    rect(cx, y - rowH2, colTotal, rowH2, { border: BORDER });
    centerText(fmtNum(total), cx, y, colTotal, rowH2);

    y -= rowH2;
  }

  ensureSpace(170);

  const incluirDescuento = !!cotizacion.incluirDescuento;
  const incluirInstalacion = !!cotizacion.incluirInstalacion;
  const descuentoValor = incluirDescuento ? Number(cotizacion.descuento) || 0 : 0;
  const descuentoEsPorcentaje = !!cotizacion.descuentoEsPorcentaje;
  const descuentoMonto = descuentoEsPorcentaje ? subtotal * descuentoValor / 100 : descuentoValor;
  const totalConDescuento = subtotal - descuentoMonto;
  const instalacionMonto = incluirInstalacion ? Number(cotizacion.instalacionMonto) || 0 : 0;
  const totalFinal = totalConDescuento + instalacionMonto;

  // Sub-total
  rect(MARGIN, y - 16, CONTENT_W - colTotal, 16, { fill: ACCENT_LIGHT });
  text("Sub-total", MARGIN + 4, y - 11, { bold: true, size: 8, color: ACCENT });
  rect(MARGIN + CONTENT_W - colTotal, y - 16, colTotal, 16, { fill: ACCENT_LIGHT });
  const subStr = fmtNum(subtotal);
  text(subStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(subStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
  y -= 16;

  // Descuento + Total Descuento Incluido — opcionales: solo si se marcó "Incluir descuento".
  if (incluirDescuento) {
    rect(MARGIN, y - 16, CONTENT_W - colTotal, 16, { border: BORDER });
    text("Descuento", MARGIN + 4, y - 11, { bold: true, size: 8 });
    rect(MARGIN + CONTENT_W - colTotal, y - 16, colTotal, 16, { border: BORDER });
    const descStr = descuentoValor === 0 ? "-" : (descuentoEsPorcentaje ? `${fmtNum(descuentoValor)}%` : fmtNum(descuentoValor));
    text(descStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(descStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
    y -= 16;

    rect(MARGIN, y - 16, CONTENT_W - colTotal, 16, { fill: ACCENT_LIGHT });
    text("Total Descuento Incluido", MARGIN + 4, y - 11, { bold: true, size: 8, color: ACCENT });
    rect(MARGIN + CONTENT_W - colTotal, y - 16, colTotal, 16, { fill: ACCENT_LIGHT });
    const totalDescStr = fmtNum(totalConDescuento);
    text(totalDescStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(totalDescStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
    y -= 16;
  }

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

  // Instalaciones — opcional: solo si se marcó "Incluir instalación". Misma
  // estructura de 3 celdas que la fila de Comentarios (label / texto / monto),
  // con el monto alineado a la columna TOTAL U$S de la tabla de productos.
  if (incluirInstalacion) {
    const instLabelW = 80;
    const instDescW = CONTENT_W - instLabelW - colTotal;
    const instLines = wrapText(font, cotizacion.instalacionDescripcion, 7, instDescW - 8);
    const instH = Math.max(instLines.length * 9 + 6, 16);
    let cx2 = MARGIN;
    rect(cx2, y - instH, instLabelW, instH, { border: BORDER });
    text("Instalaciones", cx2 + 4, y - 11, { bold: true, size: 7.5 });
    cx2 += instLabelW;
    rect(cx2, y - instH, instDescW, instH, { border: BORDER });
    instLines.forEach((line, i) => text(line, cx2 + 4, y - 11 - i * 9, { size: 7 }));
    cx2 += instDescW;
    rect(cx2, y - instH, colTotal, instH, { border: BORDER });
    const instStr = instalacionMonto === 0 ? "-" : fmtNum(instalacionMonto);
    text(instStr, cx2 + colTotal - 4 - bold.widthOfTextAtSize(instStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
    y -= instH;
  }

  // Total IVA incluido + Dólares
  rect(MARGIN, y - 16, CONTENT_W - colTotal, 16, { border: BORDER });
  text("TOTAL IVA INCLUIDO", MARGIN + 4, y - 11, { bold: true, size: 8 });
  rect(MARGIN + CONTENT_W - colTotal, y - 16, colTotal, 16, { fill: ACCENT_LIGHT });
  const totalStr = fmtNum(totalFinal);
  text(totalStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(totalStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
  y -= 16;

  rect(MARGIN, y - 16, CONTENT_W, 16, { border: BORDER });
  text("Dólares Americanos:", MARGIN + 4, y - 11, { bold: true, size: 7.5 });
  text(montoEnLetras(totalFinal), MARGIN + 110, y - 11, { size: 7.5 });
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
function fmtGs(n) {
  return (Number(n) || 0).toLocaleString("es-PY", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
export function montoEnLetrasGuaranies(monto) {
  const letras = numeroALetras(monto);
  return letras.charAt(0).toUpperCase() + letras.slice(1);
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

// ---------- Presupuesto de Reparación / Repuestos / Mantenimiento (vía separada de Cotizaciones) ----------
const LEGAL_TEXT_REPARACION =
  "Este presupuesto corresponde a una falla NO cubierta por la garantía de fábrica Aeon (ej. golpes, uso indebido, " +
  "desgaste, u otras causas ajenas a defectos de fabricación). El trabajo se realizará únicamente una vez confirmada " +
  "la aceptación de este presupuesto por parte del cliente. Todos los precios son en dólares e IVA incluido. " +
  "Presupuesto válido por 15 días desde la fecha de emisión. La instalación es opcional y no siempre está a cargo " +
  "de Aeon; de no incluirse, el repuesto se entrega para su instalación por terceros. Forma de pago sugerida: " +
  "contado contra confirmación del trabajo realizado.";

const LEGAL_TEXT_MANTENIMIENTO =
  "Este presupuesto corresponde a un servicio de mantenimiento preventivo. Los precios de mano de obra están " +
  "expresados en guaraníes e incluyen IVA; los repuestos, si los hubiera, se cotizan por separado en dólares. " +
  "El trabajo se realizará únicamente una vez confirmada la aceptación de este presupuesto por parte del cliente. " +
  "Presupuesto válido por 15 días desde la fecha de emisión. Forma de pago sugerida: contado contra confirmación " +
  "del trabajo realizado.";

export async function generatePresupuestoReparacionPdf(presupuesto) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  const base = import.meta.env.BASE_URL;
  const logoBytes = await fetchBytes(`${base}aeon-logo.jpg`);
  const logoImg = logoBytes ? await pdf.embedJpg(logoBytes) : null;

  const tipo = presupuesto.tipo === "mantenimiento" ? "mantenimiento" : "reparacion";
  const lineas = presupuesto.lineas || [];
  const lineasMantenimiento = tipo === "mantenimiento" ? (presupuesto.lineasMantenimiento || []) : [];

  const colItem = 85;
  const colModelo = 105;
  const colCant = 35;
  const colCosto = 75;
  const colDescripcion = CONTENT_W - colItem - colModelo - colCant - colCosto;

  const colCategoria = 110;
  const colCantMtto = 40;
  const colPrecioMtto = 95;
  const colDetalleMtto = CONTENT_W - colCategoria - colCantMtto - colPrecioMtto;

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
  function centerText(str, cellX, cellTopY, cellW, cellH, opts = {}) {
    const { size = 6.5, bold: isBold = false, color } = opts;
    const f = isBold ? bold : font;
    const w = f.widthOfTextAtSize(str, size);
    text(str, cellX + cellW / 2 - w / 2, cellTopY - cellH / 2 - 3, { size, bold: isBold, color });
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

  const fechaLabel = "Fecha:";
  const fechaVal = fmtFecha(presupuesto.fecha);
  text(fechaLabel, PAGE_W - MARGIN - 110, PAGE_H - MARGIN, { bold: true, size: 8 });
  text(fechaVal, PAGE_W - MARGIN - 60, PAGE_H - MARGIN, { size: 8 });

  y -= 14;

  // Title bar
  rect(MARGIN, y - 18, CONTENT_W, 18, { fill: ACCENT });
  const title = tipo === "mantenimiento" ? "PRESUPUESTO DE MANTENIMIENTO" : "PRESUPUESTO DE REPARACIÓN / REPUESTOS";
  const titleW = bold.widthOfTextAtSize(title, 9.5);
  text(title, MARGIN + CONTENT_W / 2 - titleW / 2, y - 13, { bold: true, size: 9.5, color: WHITE });
  y -= 18;

  // Cliente / Obra / Equipo afectado
  const rowH = 16;
  rect(MARGIN, y - rowH, CONTENT_W, rowH, { border: BORDER });
  text("Cliente:", MARGIN + 4, y - 11, { bold: true, size: 8 });
  text(presupuesto.cliente || "", MARGIN + 60, y - 11, { size: 8 });
  y -= rowH;
  rect(MARGIN, y - rowH, CONTENT_W, rowH, { border: BORDER });
  text("Obra:", MARGIN + 4, y - 11, { bold: true, size: 8 });
  text(presupuesto.obra || "", MARGIN + 60, y - 11, { size: 8 });
  y -= rowH;
  rect(MARGIN, y - rowH, CONTENT_W, rowH, { border: BORDER });
  const equipoLabel = tipo === "mantenimiento" ? "Equipo(s) a mantener:" : "Equipo / Producto afectado:";
  text(equipoLabel, MARGIN + 4, y - 11, { bold: true, size: 8 });
  text(presupuesto.equipoAfectado || "", MARGIN + 155, y - 11, { size: 8 });
  y -= rowH;

  // Falla reportada / Detalle (alto dinámico según el texto)
  const detalleLabel = tipo === "mantenimiento" ? "Detalle / Observaciones:" : "Falla reportada:";
  const fallaLines = wrapText(font, presupuesto.fallaReportada || "", 7, CONTENT_W - 90);
  const fallaH = Math.max(fallaLines.length * 9 + 6, 16);
  rect(MARGIN, y - fallaH, 80, fallaH, { border: BORDER });
  text(detalleLabel, MARGIN + 4, y - 11, { bold: true, size: 7.5 });
  rect(MARGIN + 80, y - fallaH, CONTENT_W - 80, fallaH, { border: BORDER });
  fallaLines.forEach((line, i) => text(line, MARGIN + 84, y - 11 - i * 9, { size: 7 }));
  y -= fallaH;

  y -= 6;

  // Tabla de mantenimiento (₲): Categoría | Detalle | Cant. | Precio Unitario
  let subtotalMtto = 0;
  if (lineasMantenimiento.length > 0) {
    function drawMttoHeader() {
      const startX = MARGIN;
      rect(startX, y - 20, CONTENT_W, 20, { fill: ACCENT_LIGHT });
      let cx = startX;
      const headers = [
        ["Categoría", colCategoria], ["Detalle", colDetalleMtto],
        ["Cant.", colCantMtto], ["Precio Unitario Gs.", colPrecioMtto],
      ];
      headers.forEach(([label, w]) => {
        const lw = bold.widthOfTextAtSize(label, 6.5);
        text(label, cx + w / 2 - lw / 2, y - 13, { bold: true, size: 6.5, color: ACCENT });
        cx += w;
      });
      y -= 20;
    }
    drawMttoHeader();

    for (const linea of lineasMantenimiento) {
      const cant = Number(linea.cantidad) || 0;
      const precio = Number(linea.precioUnitario) || 0;
      subtotalMtto += cant * precio;

      const detLines = wrapText(font, linea.detalle || "", 6.5, colDetalleMtto - 6);
      const rowH2 = Math.max(detLines.length * 8 + 6, 22);
      ensureSpace(rowH2 + 20);
      if (y === PAGE_H - MARGIN) drawMttoHeader();

      let cx = MARGIN;
      rect(cx, y - rowH2, colCategoria, rowH2, { border: BORDER });
      centerText(linea.categoria || "", cx, y, colCategoria, rowH2);
      cx += colCategoria;

      rect(cx, y - rowH2, colDetalleMtto, rowH2, { border: BORDER });
      const detBlockH = detLines.length * 8;
      const detTop = y - (rowH2 - detBlockH) / 2 - 6;
      detLines.forEach((line, i) => text(line, cx + 3, detTop - i * 8, { size: 6.5 }));
      cx += colDetalleMtto;

      rect(cx, y - rowH2, colCantMtto, rowH2, { border: BORDER });
      centerText(String(cant), cx, y, colCantMtto, rowH2);
      cx += colCantMtto;

      rect(cx, y - rowH2, colPrecioMtto, rowH2, { border: BORDER });
      centerText(fmtGs(precio), cx, y, colPrecioMtto, rowH2);

      y -= rowH2;
    }

    ensureSpace(70);
    const colTotalMtto = colPrecioMtto;
    rect(MARGIN, y - 16, CONTENT_W - colTotalMtto, 16, { fill: ACCENT_LIGHT });
    text("Sub-total mantenimiento", MARGIN + 4, y - 11, { bold: true, size: 8, color: ACCENT });
    rect(MARGIN + CONTENT_W - colTotalMtto, y - 16, colTotalMtto, 16, { fill: ACCENT_LIGHT });
    const subMttoStr = `Gs. ${fmtGs(subtotalMtto)}`;
    text(subMttoStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(subMttoStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
    y -= 16;

    rect(MARGIN, y - 16, CONTENT_W, 16, { border: BORDER });
    text("Guaraníes:", MARGIN + 4, y - 11, { bold: true, size: 7.5 });
    text(montoEnLetrasGuaranies(subtotalMtto), MARGIN + 60, y - 11, { size: 7.5 });
    y -= 24;
  }

  // Tabla de repuestos (U$S): Ítem | Descripción | Modelo del equipo | Cant. | Costo x Unidad U$S
  if (lineas.length > 0) {
    if (lineasMantenimiento.length > 0) {
      ensureSpace(30);
      text("Repuestos", MARGIN, y, { bold: true, size: 9 });
      y -= 14;
    }

    function drawTableHeader() {
      const startX = MARGIN;
      rect(startX, y - 20, CONTENT_W, 20, { fill: ACCENT_LIGHT });
      let cx = startX;
      const headers = [
        ["Ítem", colItem], ["Descripción", colDescripcion], ["Modelo del equipo", colModelo],
        ["Cant.", colCant], ["Costo x Unidad U$S", colCosto],
      ];
      headers.forEach(([label, w]) => {
        const lw = bold.widthOfTextAtSize(label, 6.5);
        text(label, cx + w / 2 - lw / 2, y - 13, { bold: true, size: 6.5, color: ACCENT });
        cx += w;
      });
      y -= 20;
    }
    drawTableHeader();

    let subtotal = 0;
    for (const linea of lineas) {
      const cant = Number(linea.cantidad) || 0;
      const costo = Number(linea.costoUnitario) || 0;
      subtotal += cant * costo;

      const descLines = wrapText(font, linea.descripcion, 6.5, colDescripcion - 6);
      const rowH2 = Math.max(descLines.length * 8 + 6, 22);
      ensureSpace(rowH2 + 20);
      if (y === PAGE_H - MARGIN) drawTableHeader();

      let cx = MARGIN;
      rect(cx, y - rowH2, colItem, rowH2, { border: BORDER });
      centerText(linea.item || "", cx, y, colItem, rowH2);
      cx += colItem;

      rect(cx, y - rowH2, colDescripcion, rowH2, { border: BORDER });
      const descBlockH = descLines.length * 8;
      const descTop = y - (rowH2 - descBlockH) / 2 - 6;
      descLines.forEach((line, i) => text(line, cx + 3, descTop - i * 8, { size: 6.5 }));
      cx += colDescripcion;

      rect(cx, y - rowH2, colModelo, rowH2, { border: BORDER });
      centerText(linea.modeloEquipo || "", cx, y, colModelo, rowH2);
      cx += colModelo;

      rect(cx, y - rowH2, colCant, rowH2, { border: BORDER });
      centerText(String(cant), cx, y, colCant, rowH2);
      cx += colCant;

      rect(cx, y - rowH2, colCosto, rowH2, { border: BORDER });
      centerText(fmtNum(costo), cx, y, colCosto, rowH2);

      y -= rowH2;
    }

    ensureSpace(150);

    const incluirInstalacion = tipo === "reparacion" && !!presupuesto.incluirInstalacion;
    const instalacionMonto = incluirInstalacion ? Number(presupuesto.instalacionMonto) || 0 : 0;
    const totalFinal = subtotal + instalacionMonto;

    // Sub-total
    const colTotalLedger = colCosto;
    rect(MARGIN, y - 16, CONTENT_W - colTotalLedger, 16, { fill: ACCENT_LIGHT });
    text("Sub-total", MARGIN + 4, y - 11, { bold: true, size: 8, color: ACCENT });
    rect(MARGIN + CONTENT_W - colTotalLedger, y - 16, colTotalLedger, 16, { fill: ACCENT_LIGHT });
    const subStr = fmtNum(subtotal);
    text(subStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(subStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
    y -= 16;

    // Instalación (opcional)
    if (incluirInstalacion) {
      rect(MARGIN, y - 16, CONTENT_W - colTotalLedger, 16, { border: BORDER });
      text("Instalación (opcional)", MARGIN + 4, y - 11, { size: 8 });
      rect(MARGIN + CONTENT_W - colTotalLedger, y - 16, colTotalLedger, 16, { border: BORDER });
      const instStr = fmtNum(instalacionMonto);
      text(instStr, MARGIN + CONTENT_W - 4 - font.widthOfTextAtSize(instStr, 8), y - 11, { size: 8 });
      y -= 16;
    }

    // Total IVA incluido + Dólares
    rect(MARGIN, y - 16, CONTENT_W - colTotalLedger, 16, { border: BORDER });
    text("TOTAL IVA INCLUIDO", MARGIN + 4, y - 11, { bold: true, size: 8 });
    rect(MARGIN + CONTENT_W - colTotalLedger, y - 16, colTotalLedger, 16, { fill: ACCENT_LIGHT });
    const totalStr = fmtNum(totalFinal);
    text(totalStr, MARGIN + CONTENT_W - 4 - bold.widthOfTextAtSize(totalStr, 8), y - 11, { bold: true, size: 8, color: ACCENT });
    y -= 16;

    rect(MARGIN, y - 16, CONTENT_W, 16, { border: BORDER });
    text("Dólares Americanos:", MARGIN + 4, y - 11, { bold: true, size: 7.5 });
    text(montoEnLetras(totalFinal), MARGIN + 110, y - 11, { size: 7.5 });
    y -= 24;
  }

  // Plazo estimado
  const plazoLines = wrapText(font, presupuesto.plazoEstimado || "", 8, CONTENT_W - 90);
  const plazoH = Math.max(plazoLines.length * 10 + 6, 20);
  rect(MARGIN, y - plazoH, CONTENT_W, plazoH, { border: BORDER });
  text("Plazo estimado:", MARGIN + 4, y - 13, { bold: true, size: 8 });
  plazoLines.forEach((line, i) => text(line, MARGIN + 90, y - 13 - i * 10, { size: 8 }));
  y -= plazoH + 10;

  // Texto legal
  ensureSpace(60);
  const legalText = tipo === "mantenimiento" ? LEGAL_TEXT_MANTENIMIENTO : LEGAL_TEXT_REPARACION;
  const legalLines = wrapText(font, legalText, 6.5, CONTENT_W - 10);
  const legalH = legalLines.length * 8 + 8;
  rect(MARGIN, y - legalH, CONTENT_W, legalH, { border: BORDER });
  legalLines.forEach((line, i) => text(line, MARGIN + 4, y - 10 - i * 8, { size: 6.5 }));
  y -= legalH + 24;

  // Aceptación del cliente
  ensureSpace(80);
  text("ACEPTACIÓN DEL CLIENTE", MARGIN, y, { bold: true, size: 9 });
  y -= 11;
  text("(confirmar antes de iniciar el trabajo)", MARGIN, y, { size: 7, color: MUTED });
  y -= 36;

  const clienteLineX = MARGIN;
  const aeonLineX = PAGE_W - MARGIN - 160;
  page.drawLine({ start: { x: clienteLineX, y }, end: { x: clienteLineX + 200, y }, thickness: 0.5, color: MUTED });
  page.drawLine({ start: { x: aeonLineX, y }, end: { x: aeonLineX + 160, y }, thickness: 0.5, color: MUTED });
  y -= 11;
  text("Firma / Aclaración / C.I. N°", clienteLineX, y, { size: 7.5 });
  const nameW = font.widthOfTextAtSize(COMPANY.firmante, 8);
  text(COMPANY.firmante, aeonLineX + 80 - nameW / 2, y, { size: 8 });
  y -= 12;
  text("Fecha: ___/___/______", clienteLineX, y, { size: 7.5 });
  const razW = font.widthOfTextAtSize(COMPANY.razonSocial, 7.5);
  text(COMPANY.razonSocial, aeonLineX + 80 - razW / 2, y, { size: 7.5 });
  y -= 9;
  const rucW = font.widthOfTextAtSize(COMPANY.ruc, 7.5);
  text(COMPANY.ruc, aeonLineX + 80 - rucW / 2, y, { size: 7.5 });

  return pdf.save();
}

export async function downloadPresupuestoReparacionPdf(presupuesto) {
  const bytes = await generatePresupuestoReparacionPdf(presupuesto);
  const prefijo = presupuesto.tipo === "mantenimiento" ? "Presupuesto_Mantenimiento" : "Presupuesto_Reparacion";
  const nombre = `${prefijo}_${(presupuesto.cliente || "cliente").replace(/\s+/g, "_")}_${presupuesto.fecha || ""}.pdf`;
  downloadBlob(bytes, nombre, "application/pdf");
}
