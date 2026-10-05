import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
import ExcelJS from "exceljs";
import { COMPANY, fmtFecha, fmtNum } from "./pdf";
import { centroCostoPorCodigo, totalLineaPedido, totalPedido } from "./centrosCosto";

// Pedido de Facturación: sigue el esqueleto de "Pedido de Facturacion.xlsx" (hoja "Esqueleto"):
// encabezado con logo y datos de Quantum, bloque Cliente / RUC / Condición-Plazo / Término / Fecha /
// Centro de costo / Obra, tabla Cantidad · Marca · CÓDIGO · PRODUCTO · Familia · Precio UniT. ·
// Total USD, barra TOTAL y firmas SOLICITADO / AUTORIZADO.

function nombreBase(p) {
  const safe = (s) => (s || "").toString().trim().replace(/\s+/g, "_").replace(/[^a-zA-Z0-9_-]/g, "");
  return `Pedido_Facturacion_${safe(p.cliente) || "cliente"}_${safe(p.obra) || "obra"}${p.remito ? "_Remito_" + safe(p.remito) : ""}`;
}
export const nombreArchivoPedidoPdf = (p) => `${nombreBase(p)}.pdf`;
export const nombreArchivoPedidoExcel = (p) => `${nombreBase(p)}.xlsx`;

async function fetchBytes(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    return new Uint8Array(await res.arrayBuffer());
  } catch (e) {
    return null;
  }
}

function descargar(bytes, filename, type) {
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

const COLS = [30.78, 19.67, 29, 19.89, 15.67, 17.89, 15.33]; // anchos de columna del esqueleto
const COLOR_ENLACE = rgb(0x2f / 255, 0x6f / 255, 0x8f / 255);
const TELEFONOS = ["+595 976 167335", "+595 976 144599"];
const CORREOS = ["ggibernau@aeon.com.py", "info@aeon.com.py"];

// ---------- PDF ----------
export async function generatePedidoFacturacionPdf(p) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const base = import.meta.env.BASE_URL;
  const [logoBytes, firmaBytes] = await Promise.all([fetchBytes(`${base}aeon-logo.jpg`), fetchBytes(`${base}generated/firma.png`)]);
  const logo = logoBytes ? await pdf.embedJpg(logoBytes) : null;
  const firma = firmaBytes ? await pdf.embedPng(firmaBytes) : null;

  const PW = 595.28, PH = 841.89, M = 36;
  const W = PW - M * 2;
  const suma = COLS.reduce((a, b) => a + b, 0);
  const cw = COLS.map((c) => (c / suma) * W);
  const cx = [M];
  cw.forEach((w, i) => cx.push(cx[i] + w));

  const NEGRO = rgb(0, 0, 0);
  const GRIS_ENC = rgb(0xd9 / 255, 0xd9 / 255, 0xd9 / 255);
  const GRIS_TOTAL = rgb(0x80 / 255, 0x80 / 255, 0x80 / 255);
  const GRIS_TXT = rgb(0x40 / 255, 0x40 / 255, 0x40 / 255);

  let page = pdf.addPage([PW, PH]);
  let y = PH - M;

  const text = (s, x, yy, o = {}) =>
    page.drawText(String(s ?? ""), { x, y: yy, size: o.size || 9, font: o.bold ? bold : font, color: o.color || NEGRO });
  const ancho = (s, size, b) => (b ? bold : font).widthOfTextAtSize(String(s ?? ""), size);
  const fit = (s, maxW, size, b) => {
    let t = String(s ?? "");
    while (t.length > 1 && ancho(t, size, b) > maxW) t = t.slice(0, -1);
    return t;
  };
  const caja = (x, yTop, w, h, o = {}) => {
    if (o.fill) page.drawRectangle({ x, y: yTop - h, width: w, height: h, color: o.fill });
    page.drawRectangle({ x, y: yTop - h, width: w, height: h, borderColor: NEGRO, borderWidth: o.grueso ? 1.4 : 0.6 });
  };
  const celda = (x, yTop, w, h, s, o = {}) => {
    caja(x, yTop, w, h, o);
    if (s === undefined || s === null || s === "") return;
    const size = o.size || 9;
    const t = fit(s, w - 8, size, o.bold);
    const tw = ancho(t, size, o.bold);
    const tx = o.align === "right" ? x + w - 4 - tw : o.align === "left" ? x + 4 : x + (w - tw) / 2;
    text(t, tx, yTop - h / 2 - size * 0.35, { size, bold: o.bold, color: o.color });
  };

  // Encabezado
  if (logo) {
    const lw = 150, lh = (logo.height / logo.width) * lw;
    page.drawImage(logo, { x: M, y: y - lh, width: lw, height: lh });
    y -= lh + 6;
  }
  text(COMPANY.razonSocial, M, y - 9, { bold: true, size: 11 });
  y -= 22;
  text(COMPANY.direccion, M, y, { size: 8.5 }); y -= 11;
  text(COMPANY.direccion2, M, y, { size: 8.5 }); y -= 11;
  text(TELEFONOS[0], M, y, { size: 8.5 }); text(TELEFONOS[1], cx[1], y, { size: 8.5 }); y -= 11;
  text(CORREOS[0], M, y, { size: 8.5, color: COLOR_ENLACE }); text(CORREOS[1], cx[1], y, { size: 8.5, color: COLOR_ENLACE });
  y -= 22;

  const titulo = "PEDIDO DE FACTURACIÓN";
  const tw = ancho(titulo, 14, true);
  text(titulo, cx[1] + (cx[7] - cx[1] - tw) / 2, y, { bold: true, size: 14 });
  y -= 12;
  if (p.remito) {
    const ref = `Ref.: Remito N° ${p.remito}`;
    text(ref, cx[1] + (cx[7] - cx[1] - ancho(ref, 8, false)) / 2, y, { size: 8, color: GRIS_TXT });
  }
  y -= 16;

  // Bloque de datos
  const RH = 19;
  const L = (s) => ({ bold: true, size: 9.5, align: "left", grueso: true });
  const V = { bold: true, size: 10, grueso: true };
  celda(cx[0], y, cw[0], RH, "CLIENTE:", L());
  celda(cx[1], y, cw[1] + cw[2], RH, p.cliente, V);
  y -= RH;
  celda(cx[0], y, cw[0], RH, "RUC:", L());
  celda(cx[1], y, cw[1], RH, p.ruc, { ...V, align: "right" });
  celda(cx[2], y, cw[2], RH, "", V);
  y -= RH;
  celda(cx[0], y, cw[0], RH, "CONDICIÓN/PLAZO:", L());
  celda(cx[1], y, cw[1], RH, p.condicion, { ...V, align: "right" });
  celda(cx[2], y, cw[2], RH, "Término:", { bold: true, size: 9, grueso: true });
  celda(cx[3], y, cw[3], RH, p.termino, { ...V, size: 9, align: "right" });
  celda(cx[5], y, cw[5], RH, "Fecha:", { bold: true, size: 9, align: "right", grueso: true });
  celda(cx[6], y, cw[6], RH, fmtFecha(p.fecha), { bold: true, size: 9, align: "right", grueso: true });
  y -= RH;
  const cc = centroCostoPorCodigo(p.centroCostoCodigo);
  celda(cx[0], y, cw[0], RH, "CENTRO DE COSTO:", L());
  celda(cx[1], y, cw[1], RH, p.centroCostoCodigo, V);
  celda(cx[2], y, cw[2], RH, p.centroCostoNombre || cc?.nombre || "", { bold: true, size: 9, grueso: true });
  y -= RH;
  celda(cx[0], y, cw[0], RH, "OBRA:", L());
  celda(cx[1], y, cw[1] + cw[2], RH, p.obra, V);
  y -= RH + 22;

  // Tabla
  const heads = ["Cantidad", "Marca", "CÓDIGO", "PRODUCTO", "Familia", "Precio UniT.", "Total USD"];
  const drawHeader = () => {
    heads.forEach((h, i) => celda(cx[i], y, cw[i], 17, h, { bold: true, size: 8.5, fill: GRIS_ENC, grueso: false }));
    y -= 17;
  };
  drawHeader();
  // Los códigos y nombres largos se parten en varias líneas (la fila crece) en vez de cortarse.
  const partir = (texto, size, maxW) => {
    const lineas = [];
    let actual = "";
    for (const palabra of String(texto ?? "").split(/\s+/).filter(Boolean)) {
      let resto = palabra;
      while (ancho(resto, size, false) > maxW) {
        let corte = resto.length - 1;
        while (corte > 1 && ancho(resto.slice(0, corte), size, false) > maxW) corte--;
        const guion = resto.lastIndexOf("-", corte - 1);
        if (guion > 2) corte = guion + 1;
        if (actual) { lineas.push(actual); actual = ""; }
        lineas.push(resto.slice(0, corte));
        resto = resto.slice(corte);
      }
      const prueba = actual ? `${actual} ${resto}` : resto;
      if (actual && ancho(prueba, size, false) > maxW) { lineas.push(actual); actual = resto; } else actual = prueba;
    }
    if (actual) lineas.push(actual);
    return lineas.length ? lineas : [""];
  };
  const celdaLineas = (x, yTop, w, h, lineas, size, alignIzq) => {
    caja(x, yTop, w, h, {});
    const paso = size + 2;
    const alto = lineas.length * paso;
    lineas.forEach((t, i) => {
      const tw = ancho(t, size, false);
      text(t, alignIzq ? x + 4 : x + (w - tw) / 2, yTop - (h - alto) / 2 - size - i * paso + 1.5, { size });
    });
  };
  for (const l of p.lineas || []) {
    const lCodigo = partir(l.codigo, 8, cw[2] - 8);
    const lProducto = partir(l.producto, 8, cw[3] - 8);
    const lFamilia = partir(l.familia, 8, cw[4] - 8);
    const RL = Math.max(20, Math.max(lCodigo.length, lProducto.length, lFamilia.length) * 10 + 8);
    if (y - RL < M + 120) {
      page = pdf.addPage([PW, PH]);
      y = PH - M;
      drawHeader();
    }
    celda(cx[0], y, cw[0], RL, l.cantidad, { size: 9.5 });
    celda(cx[1], y, cw[1], RL, l.marca, { size: 9 });
    celdaLineas(cx[2], y, cw[2], RL, lCodigo, 8, true);
    celdaLineas(cx[3], y, cw[3], RL, lProducto, 8, false);
    celdaLineas(cx[4], y, cw[4], RL, lFamilia, 8, false);
    celda(cx[5], y, cw[5], RL, l.precioUnit === "" || l.precioUnit == null ? "" : fmtNum(l.precioUnit), { size: 9, color: GRIS_TXT });
    celda(cx[6], y, cw[6], RL, fmtNum(totalLineaPedido(l)), { size: 9, color: GRIS_TXT });
    y -= RL;
  }
  // Barra TOTAL
  const RT = 20;
  page.drawRectangle({ x: cx[0], y: y - RT, width: W, height: RT, color: GRIS_TOTAL });
  text("TOTAL", cx[5] + 6, y - RT / 2 - 4, { bold: true, size: 11, color: rgb(1, 1, 1) });
  const tot = fmtNum(totalPedido(p));
  text(tot, cx[7] - 6 - ancho(tot, 11, true), y - RT / 2 - 4, { bold: true, size: 11, color: rgb(1, 1, 1) });
  y -= RT + 36;

  // Firmas
  if (y < M + 90) {
    page = pdf.addPage([PW, PH]);
    y = PH - M - 40;
  }
  const yLinea = y - 40;
  if (firma) {
    const ratio = firma.height / firma.width;
    const fh = Math.min(ratio * cw[6] * 0.9, 36);
    const fw = fh / ratio;
    page.drawImage(firma, { x: cx[6] + (cw[6] - fw) / 2, y: yLinea + 4, width: fw, height: fh });
  }
  const dash = (x1, x2) => page.drawLine({ start: { x: x1, y: yLinea }, end: { x: x2, y: yLinea }, thickness: 0.7, color: NEGRO, dashArray: [2, 2] });
  dash(cx[1] + 6, cx[1] + cw[1] - 6);
  dash(cx[6] + 4, cx[7] - 4);
  const sol = "SOLICITADO";
  text(sol, cx[1] + (cw[1] - ancho(sol, 9, true)) / 2, yLinea - 12, { bold: true, size: 9 });
  const aut = "AUTORIZADO";
  text(aut, cx[6] + (cw[6] - ancho(aut, 9, true)) / 2, yLinea - 12, { bold: true, size: 9 });

  return pdf.save();
}

// ---------- Excel (con fórmulas, editable) ----------
export async function generatePedidoFacturacionExcel(p) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Pedido de facturación", { views: [{ showGridLines: false }], pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  COLS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const alturas = { 4: 17.4, 5: 15.6, 6: 15.6, 7: 15.6, 8: 15.6, 10: 21, 11: 21.6, 12: 18.6, 13: 18.6, 14: 18.6, 15: 18.6, 16: 18.6 };
  Object.entries(alturas).forEach(([r, h]) => { ws.getRow(Number(r)).height = h; });

  const base = import.meta.env.BASE_URL;
  const [logoBytes, firmaBytes] = await Promise.all([fetchBytes(`${base}aeon-logo.jpg`), fetchBytes(`${base}generated/firma.png`)]);
  if (logoBytes) {
    const id = wb.addImage({ buffer: logoBytes, extension: "jpeg" });
    ws.addImage(id, { tl: { col: 0.1, row: 0.2 }, ext: { width: 184, height: 33.5 } });
  }

  const grueso = { style: "medium", color: { argb: "FF000000" } };
  const fino = { style: "thin", color: { argb: "FF000000" } };
  const bordeG = { top: grueso, bottom: grueso, left: grueso, right: grueso };
  const bordeF = { top: fino, bottom: fino, left: fino, right: fino };
  const estilo = (cell, o = {}) => {
    cell.style = {
      font: { name: o.font || "Aptos Display", size: o.size || 11, bold: o.bold !== false, color: { argb: o.color || "FF000000" } },
      alignment: { vertical: "middle", horizontal: o.align || "center", wrapText: !!o.wrap },
      ...(o.fill ? { fill: { type: "pattern", pattern: "solid", fgColor: { argb: o.fill } } } : {}),
      ...(o.border ? { border: o.border } : {}),
      ...(o.numFmt ? { numFmt: o.numFmt } : {}),
    };
  };

  // Encabezado
  ws.getCell("A4").value = COMPANY.razonSocial; estilo(ws.getCell("A4"), { font: "Gotham", size: 14, align: "left" });
  [["A5", COMPANY.direccion], ["A6", COMPANY.direccion2], ["A7", TELEFONOS[0]], ["B7", TELEFONOS[1]], ["A8", CORREOS[0]], ["B8", CORREOS[1]]].forEach(([a, v]) => {
    ws.getCell(a).value = v; estilo(ws.getCell(a), { font: "Gotham", size: a.endsWith("8") ? 10 : 12, bold: false, align: "left" });
  });
  ws.mergeCells("B10:G10");
  ws.getCell("B10").value = "PEDIDO DE FACTURACIÓN"; estilo(ws.getCell("B10"), { font: "Gotham", size: 16 });
  if (p.remito) {
    ws.getCell("A11").value = `Ref.: Remito N° ${p.remito}`;
    estilo(ws.getCell("A11"), { font: "Gotham", size: 10, bold: false, align: "left", color: "FF404040" });
  }

  // Bloque de datos
  const etiqueta = (a, v) => { ws.getCell(a).value = v; estilo(ws.getCell(a), { font: "Gotham", size: 14, align: "left", wrap: true, border: bordeG }); };
  const valor = (a, v, o = {}) => { ws.getCell(a).value = v; estilo(ws.getCell(a), { size: 14, border: bordeG, ...o }); };
  etiqueta("A12", "CLIENTE:"); ws.mergeCells("B12:C12"); valor("B12", p.cliente || ""); ws.getCell("C12").border = bordeG;
  etiqueta("A13", "RUC:"); valor("B13", p.ruc || "", { align: "right" }); valor("C13", "", { size: 11 });
  etiqueta("A14", "CONDICIÓN/PLAZO:"); valor("B14", p.condicion || "", { align: "right" });
  valor("C14", "Término:", { size: 11 }); valor("D14", p.termino || "", { size: 11, align: "right" });
  valor("F14", "Fecha:", { font: "Gotham", size: 11, align: "right" });
  const [fy, fm, fd] = (p.fecha || "").split("-").map(Number);
  const f = fy ? new Date(Date.UTC(fy, fm - 1, fd)) : null; // sin hora: ExcelJS guarda las fechas en UTC
  valor("G14", f || "", { font: "Gotham", size: 11, align: "right", numFmt: "dd/mm/yyyy" });
  etiqueta("A15", "CENTRO DE COSTO:"); valor("B15", p.centroCostoCodigo || ""); valor("C15", p.centroCostoNombre || "", { size: 11 });
  etiqueta("A16", "OBRA:"); ws.mergeCells("B16:C16"); valor("B16", p.obra || ""); ws.getCell("C16").border = bordeG;

  // Tabla
  const heads = ["Cantidad", "Marca", "CÓDIGO", "PRODUCTO", "Familia", "Precio UniT.", "Total USD"];
  heads.forEach((h, i) => {
    const c = ws.getCell(19, i + 1);
    c.value = h;
    estilo(c, { font: "Aptos Narrow", size: 11, fill: "FFD9D9D9", border: bordeF });
  });
  let r = 20;
  const primera = r;
  for (const l of p.lineas || []) {
    const row = ws.getRow(r);
    row.height = 18;
    ws.getCell(r, 1).value = Number(l.cantidad) || 0; estilo(ws.getCell(r, 1), { font: "Aptos Narrow", size: 14, bold: false, border: bordeF });
    ws.getCell(r, 2).value = l.marca || ""; estilo(ws.getCell(r, 2), { font: "Aptos Narrow", size: 14, bold: false, border: bordeF });
    ws.getCell(r, 3).value = l.codigo || ""; estilo(ws.getCell(r, 3), { font: "Times New Roman", size: 12, bold: false, align: "left", border: bordeF });
    ws.getCell(r, 4).value = l.producto || ""; estilo(ws.getCell(r, 4), { size: 11, bold: false, border: bordeF });
    ws.getCell(r, 5).value = l.familia || ""; estilo(ws.getCell(r, 5), { size: 11, bold: false, border: bordeF });
    ws.getCell(r, 6).value = Number(l.precioUnit) || 0; estilo(ws.getCell(r, 6), { font: "Gotham HTF", size: 14, bold: false, color: "FF404040", border: bordeF, numFmt: "#,##0.00" });
    ws.getCell(r, 7).value = { formula: `F${r}*A${r}`, result: totalLineaPedido(l) }; estilo(ws.getCell(r, 7), { font: "Gotham HTF", size: 14, bold: false, color: "FF404040", border: bordeF, numFmt: "#,##0.00" });
    r++;
  }
  for (let c = 1; c <= 7; c++) estilo(ws.getCell(r, c), { fill: "FF808080", border: { top: fino } });
  ws.getRow(r).height = 17.4;
  ws.getCell(r, 6).value = "TOTAL"; estilo(ws.getCell(r, 6), { font: "Gotham HTF", size: 14, color: "FFFFFFFF", fill: "FF808080", align: "left", border: bordeF });
  ws.getCell(r, 7).value = { formula: `SUM(G${primera}:G${Math.max(primera, r - 1)})`, result: totalPedido(p) };
  estilo(ws.getCell(r, 7), { font: "Gotham HTF", size: 14, color: "FFFFFFFF", fill: "FF808080", align: "right", numFmt: "#,##0.00" });

  // Firmas
  const rf = r + 4;
  ws.getCell(rf, 2).value = "SOLICITADO"; estilo(ws.getCell(rf, 2), { size: 11, align: "center" });
  ws.getCell(rf, 7).value = "AUTORIZADO"; estilo(ws.getCell(rf, 7), { size: 11, align: "center", border: { top: { style: "dashed", color: { argb: "FF000000" } } } });
  ws.getCell(rf, 2).border = { top: { style: "dashed", color: { argb: "FF000000" } } };
  if (firmaBytes) {
    const id = wb.addImage({ buffer: firmaBytes, extension: "png" });
    ws.addImage(id, { tl: { col: 6.1, row: rf - 3 }, ext: { width: 88, height: 40 } });
  }

  return wb.xlsx.writeBuffer();
}

export async function downloadPedidoPdf(p) {
  descargar(await generatePedidoFacturacionPdf(p), nombreArchivoPedidoPdf(p), "application/pdf");
}
export async function downloadPedidoExcel(p) {
  descargar(await generatePedidoFacturacionExcel(p), nombreArchivoPedidoExcel(p), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

// ---------- Mail a Administración ----------
export function asuntoMailPedido(p) {
  return `Pedido de facturación — ${p.cliente || "Cliente"}${p.obra ? " — " + p.obra : ""}${p.remito ? " — Remito " + p.remito : ""}`;
}

export function textoMailPedido(p, { corto = false } = {}) {
  const cc = centroCostoPorCodigo(p.centroCostoCodigo);
  const lineas = (p.lineas || []).map((l) => {
    const pu = l.precioUnit === "" || l.precioUnit == null ? "s/precio" : `U$S ${fmtNum(l.precioUnit)}`;
    return `- ${l.cantidad} x ${l.codigo || "(sin código)"} — ${l.producto || ""} — ${pu} c/u = U$S ${fmtNum(totalLineaPedido(l))}`;
  });
  const cab = [
    "Hola, te paso el pedido de facturación:",
    "",
    p.remito ? `Remito N°: ${p.remito}` : null,
    `Fecha: ${fmtFecha(p.fecha)}`,
    `Cliente: ${p.cliente || ""}${p.ruc ? " (RUC " + p.ruc + ")" : ""}`,
    `Obra: ${p.obra || ""}`,
    `Centro de costo: ${p.centroCostoCodigo || ""}${p.centroCostoNombre || cc ? " — " + (p.centroCostoNombre || cc.nombre) : ""}`,
    `Condición: ${p.condicion || ""}${p.termino ? " — Término: " + p.termino : ""}`,
    "",
    "Detalle:",
  ].filter((x) => x !== null);
  const pie = ["", `Total a facturar: U$S ${fmtNum(totalPedido(p))}`, "", "Adjunto el pedido en PDF y Excel.", "", "Saludos,", "Gastón Gibernau", "Quantum Investments S.A."];
  const cuerpoLineas = corto && lineas.length > 6 ? [...lineas.slice(0, 6), `- ... y ${lineas.length - 6} línea(s) más (ver adjunto)`] : lineas;
  return [...cab, ...cuerpoLineas, ...pie].join("\n");
}

// Abre el redactor de mail con el texto armado. Un enlace no puede llevar adjuntos: por eso, cuando
// el navegador lo permite se comparte directo con los dos archivos (Outlook figura entre las apps),
// y si no, se descargan los archivos y se abre el mail para adjuntarlos.
export async function enviarPedidoPorMail(p, destinatario) {
  const asunto = asuntoMailPedido(p);
  const [pdfBytes, xlsxBuf] = await Promise.all([generatePedidoFacturacionPdf(p), generatePedidoFacturacionExcel(p)]);
  const archivos = [
    new File([pdfBytes], nombreArchivoPedidoPdf(p), { type: "application/pdf" }),
    new File([xlsxBuf], nombreArchivoPedidoExcel(p), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
  ];
  if (navigator.canShare && navigator.canShare({ files: archivos })) {
    try {
      await navigator.share({ files: archivos, title: asunto, text: textoMailPedido(p) });
      return "compartido";
    } catch (e) {
      if (e && e.name === "AbortError") return "cancelado";
    }
  }
  descargar(pdfBytes, archivos[0].name, "application/pdf");
  descargar(xlsxBuf, archivos[1].name, archivos[1].type);
  const cuerpo = textoMailPedido(p, { corto: true }) + "\n\n(Adjuntar los archivos descargados: PDF y Excel.)";
  window.location.href = `mailto:${encodeURIComponent(destinatario || "")}?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
  return "mailto";
}

// Mismo mail pero en Outlook web (usa la sesión de Outlook ya abierta en el navegador).
export async function abrirPedidoEnOutlookWeb(p, destinatario) {
  const asunto = asuntoMailPedido(p);
  const [pdfBytes, xlsxBuf] = await Promise.all([generatePedidoFacturacionPdf(p), generatePedidoFacturacionExcel(p)]);
  descargar(pdfBytes, nombreArchivoPedidoPdf(p), "application/pdf");
  descargar(xlsxBuf, nombreArchivoPedidoExcel(p), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const cuerpo = textoMailPedido(p, { corto: true }) + "\n\n(Adjuntar los archivos descargados: PDF y Excel.)";
  const url = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(destinatario || "")}&subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
  window.open(url, "_blank", "noopener");
}
