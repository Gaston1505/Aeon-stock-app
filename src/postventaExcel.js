import ExcelJS from "exceljs";
import { COMPANY, fmtFecha } from "./pdf";

// Planilla de Mantenimiento postventa — mismos colores y helper de estilo que cotizacionExcel.js
// (paleta derivada del gris del logo AEON), pero en formato tabla plana: una fila por unidad
// (apartamento), no el layout de invoice de una cotización.

const ACCENT = "FF565A5F";
const ACCENT_LIGHT = "FFEBEBEC";
const BORDER = "FFE4E5E5";
const MUTED = "FF686D73";
const WHITE = "FFFFFFFF";
const VERDE = "FF15803D";
const VERDE_CLARO = "FFE9F7EF";
const AMBAR = "FFB45309";
const AMBAR_CLARO = "FFFDF1E0";
const ROJO = "FFB91C1C";
const ROJO_CLARO = "FFFBEAEA";

function thinBorder() {
  const side = { style: "thin", color: { argb: BORDER } };
  return { top: side, bottom: side, left: side, right: side };
}

// Mismo patrón que cotizacionExcel.js: armar el objeto de estilo completo y asignarlo de una sola
// vez (`cell.style = ...`) en vez de ir mutando fill/font/border por separado — ExcelJS puede
// terminar compartiendo un estilo parcial entre celdas distintas si se asignan de a una propiedad.
function styleCell(cell, opts = {}) {
  const style = {};
  if (opts.fill) style.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fill } };
  style.font = { bold: !!opts.bold, size: opts.size || 10, color: { argb: opts.color || "FF1C1E20" } };
  style.border = thinBorder();
  style.alignment = { vertical: "middle", horizontal: opts.align || "left", wrapText: !!opts.wrap };
  cell.style = style;
}

async function fetchArrayBuffer(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    return new Uint8Array(await res.arrayBuffer());
  } catch (e) {
    return null;
  }
}

// Estado legible + color de un hito (12m/24m) de una unidad — mismo criterio que alertaUnidadPostventa
// en App.jsx (11/23 meses de aviso, 12/24 de vencimiento), reimplementado acá para no depender del
// componente React al generar el archivo.
function estadoHito(unidad, meses) {
  if (unidad.tipoObra !== "ventaCero" || !unidad.fechaVenta) return { texto: "No aplica", fill: null };
  const hecho = unidad[`service${meses}Hecho`];
  const fecha = unidad[`service${meses}Fecha`];
  if (hecho) return { texto: `Hecho ${fmtFecha(fecha)}`, fill: VERDE_CLARO, color: VERDE };
  const vto = new Date(unidad.fechaVenta + "T00:00:00");
  vto.setMonth(vto.getMonth() + meses);
  const hoy = new Date(new Date().toISOString().slice(0, 10) + "T00:00:00");
  const dias = Math.round((vto - hoy) / (1000 * 60 * 60 * 24));
  if (dias < 0) return { texto: `Vencido hace ${Math.abs(dias)}d`, fill: ROJO_CLARO, color: ROJO };
  if (dias <= 30) return { texto: `Por vencer (${dias}d)`, fill: AMBAR_CLARO, color: AMBAR };
  return { texto: "Pendiente", fill: null };
}

export async function generatePlanillaPostventaExcelBuffer(edificios, unidades, visitas) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Mantenimiento postventa", { views: [{ showGridLines: false }] });

  const cols = [
    { key: "edificio", width: 22 },
    { key: "apartamento", width: 12 },
    { key: "cliente", width: 24 },
    { key: "equipo", width: 26 },
    { key: "capacidad", width: 14 },
    { key: "situacion", width: 16 },
    { key: "fechaVenta", width: 13 },
    { key: "service12", width: 20 },
    { key: "service24", width: 20 },
    { key: "ultimoInforme", width: 40 },
  ];
  sheet.columns = cols;
  const lastCol = cols.length;

  const base = import.meta.env.BASE_URL;
  const logoBuf = await fetchArrayBuffer(`${base}aeon-logo.jpg`);

  let r = 1;
  if (logoBuf) {
    const imgId = workbook.addImage({ buffer: logoBuf, extension: "jpeg" });
    sheet.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: 130, height: 66 } });
  }
  sheet.getCell(r, lastCol).value = "Fecha:";
  sheet.getCell(r, lastCol).font = { bold: true, size: 9 };
  sheet.getCell(r, lastCol).alignment = { horizontal: "right" };
  r += 1;
  sheet.getCell(r, lastCol).value = fmtFecha(new Date().toISOString().slice(0, 10));
  sheet.getCell(r, lastCol).alignment = { horizontal: "right" };
  r += 4;

  [[COMPANY.razonSocial, { bold: true, size: 10 }], [COMPANY.telefonos, { size: 8.5, color: MUTED }], [COMPANY.emails, { size: 8.5, color: MUTED }]].forEach(([text, opts]) => {
    const cell = sheet.getCell(r, 1);
    cell.value = text;
    cell.font = { bold: !!opts.bold, size: opts.size, color: { argb: opts.color ? MUTED : "FF1C1E20" } };
    r += 1;
  });
  r += 1;

  sheet.mergeCells(r, 1, r, lastCol);
  const titleCell = sheet.getCell(r, 1);
  titleCell.value = "PLANILLA DE MANTENIMIENTO POSTVENTA";
  styleCell(titleCell, { fill: ACCENT, bold: true, color: WHITE, size: 11, align: "center" });
  sheet.getRow(r).height = 22;
  r += 2;

  const headers = ["Edificio", "Apartamento", "Cliente", "Equipo", "Capacidad", "Situación", "Fecha de venta", "Service 12 meses", "Service 24 meses", "Último informe técnico"];
  headers.forEach((h, i) => {
    const cell = sheet.getCell(r, i + 1);
    cell.value = h;
    styleCell(cell, { fill: ACCENT_LIGHT, bold: true, color: "FF565A5F", align: "center" });
  });
  sheet.getRow(r).height = 20;
  r += 1;

  const edificioDe = (id) => edificios.find((e) => e.id === id);
  const unidadesOrdenadas = [...unidades].sort((a, b) => {
    const ea = edificioDe(a.edificioId)?.nombre || "";
    const eb = edificioDe(b.edificioId)?.nombre || "";
    return ea.localeCompare(eb) || (a.apartamento || "").localeCompare(b.apartamento || "", undefined, { numeric: true });
  });

  for (const u of unidadesOrdenadas) {
    const edificio = edificioDe(u.edificioId);
    // El informe más reciente entre todas las visitas que ya marcaron esta unidad como hecha.
    const entregas = [];
    for (const v of visitas) {
      const entrada = (v.unidades || []).find((x) => x.unidadId === u.id);
      if (entrada?.hecho) entregas.push(entrada);
    }
    entregas.sort((a, b) => (b.fechaHecho || "").localeCompare(a.fechaHecho || ""));
    const ultimo = entregas[0];
    const s12 = estadoHito(u, 12);
    const s24 = estadoHito(u, 24);

    const row = [
      edificio?.nombre || "(edificio borrado)",
      u.apartamento || "",
      u.cliente || "",
      u.equipo || "",
      u.capacidad || "",
      u.tipoObra === "ventaCero" ? "Venta nueva" : "Obra activa",
      u.tipoObra === "ventaCero" ? fmtFecha(u.fechaVenta) : "—",
      s12.texto,
      s24.texto,
      ultimo ? `${fmtFecha(ultimo.fechaHecho)}: ${ultimo.informe || "—"}` : "",
    ];
    row.forEach((val, i) => {
      const cell = sheet.getCell(r, i + 1);
      cell.value = val;
      const opts = { wrap: true };
      if (i === 7) Object.assign(opts, { fill: s12.fill, color: s12.color, align: "center" });
      else if (i === 8) Object.assign(opts, { fill: s24.fill, color: s24.color, align: "center" });
      else if (i === 1 || i === 5 || i === 6) opts.align = "center";
      styleCell(cell, opts);
    });
    r += 1;
  }

  return workbook.xlsx.writeBuffer();
}

export function nombreArchivoPlanillaPostventa() {
  return `Mantenimiento_Postventa_${new Date().toISOString().slice(0, 10)}.xlsx`;
}

export async function downloadPlanillaPostventaExcel(edificios, unidades, visitas) {
  const buffer = await generatePlanillaPostventaExcelBuffer(edificios, unidades, visitas);
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombreArchivoPlanillaPostventa();
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
