import ExcelJS from "exceljs";
import { COMPANY, fmtFecha, fmtNum, montoEnLetras, legalTextCotizacion } from "./pdf";

// Excel de la cotización que imita el diseño del PDF (mismo logo, colores, tabla y totales),
// pero con las fotos de cada producto insertadas de verdad en la celda — la librería "xlsx"
// (SheetJS) que usa el resto de la app no soporta escribir imágenes en un .xlsx, por eso esto
// usa "exceljs" en su lugar, solo para este export puntual.

const ACCENT = "FF565A5F";
const ACCENT_LIGHT = "FFEBEBEC";
const BORDER = "FFE4E5E5";
const MUTED = "FF686D73";
const WHITE = "FFFFFFFF";

function thinBorder() {
  const side = { style: "thin", color: { argb: BORDER } };
  return { top: side, bottom: side, left: side, right: side };
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

export async function generateCotizacionExcelBuffer(cotizacion) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Cotización", { views: [{ showGridLines: false }] });

  const lineas = cotizacion.lineas || [];
  const showEspec = lineas.some((l) => l.especValor);
  const especLabel = lineas.find((l) => l.especLabel)?.especLabel || "Espec.";

  const cols = [
    { key: "codigo", width: 18 },
    { key: "foto", width: 13 },
    { key: "descripcion", width: 46 },
    ...(showEspec ? [{ key: "espec", width: 14 }] : []),
    { key: "cantidad", width: 8 },
    { key: "precio", width: 15 },
    { key: "total", width: 15 },
  ];
  sheet.columns = cols;
  const numCols = cols.length;
  const lastCol = numCols;

  function mergeRow(row, opts = {}) {
    sheet.mergeCells(row, 1, row, lastCol);
  }
  // OJO: asignar fill/font/border/alignment como propiedades sueltas (cell.fill = ..., luego
  // cell.border = ..., etc.) hace que ExcelJS vaya cacheando/reusando estilos parciales por
  // cada asignación individual — en la práctica, dos celdas con las MISMAS opciones podían
  // terminar con un alignment distinto (confirmado inspeccionando el XML real: la celda de
  // código de una línea y la barra de título perdían el horizontal/wrapText pedido). Armar el
  // objeto de estilo completo y asignarlo de una sola vez (`cell.style = ...`) evita ese bug,
  // porque ExcelJS lo resuelve como un único estilo en vez de ir mutando uno parcial.
  function styleCell(cell, opts = {}) {
    const style = {};
    if (opts.fill) style.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fill } };
    if (opts.bold || opts.color) style.font = { bold: !!opts.bold, size: opts.size || 10, color: { argb: opts.color || "FF1C1E20" } };
    else if (opts.size) style.font = { size: opts.size };
    if (opts.border !== false) style.border = thinBorder();
    style.alignment = { vertical: "middle", horizontal: opts.align || "left", wrapText: !!opts.wrap };
    cell.style = style;
  }

  const base = import.meta.env.BASE_URL;
  const [logoBuf, firmaBuf] = await Promise.all([
    fetchArrayBuffer(`${base}aeon-logo.jpg`),
    fetchArrayBuffer(`${base}generated/firma.png`),
  ]);

  let r = 1;

  // Header / letterhead
  if (logoBuf) {
    const imgId = workbook.addImage({ buffer: logoBuf, extension: "jpeg" });
    sheet.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: 130, height: 66 } });
  }
  sheet.getCell(r, lastCol).value = "Fecha:";
  sheet.getCell(r, lastCol).font = { bold: true, size: 9 };
  sheet.getCell(r, lastCol).alignment = { horizontal: "right" };
  r += 1;
  sheet.getCell(r, lastCol).value = fmtFecha(cotizacion.fecha);
  sheet.getCell(r, lastCol).alignment = { horizontal: "right" };
  r += 4; // deja lugar debajo del logo (alto ~66px)

  const infoLines = [
    [COMPANY.razonSocial, { bold: true, size: 10 }],
    [COMPANY.direccion, { size: 8.5, color: MUTED }],
    [COMPANY.direccion2, { size: 8.5, color: MUTED }],
    [COMPANY.telefonos, { size: 8.5, color: MUTED }],
    [COMPANY.emails, { size: 8.5, color: MUTED }],
  ];
  infoLines.forEach(([text, opts]) => {
    const cell = sheet.getCell(r, 1);
    cell.value = text;
    cell.font = { bold: !!opts.bold, size: opts.size, color: { argb: opts.color ? "FF686D73" : "FF1C1E20" } };
    cell.alignment = { vertical: "middle" };
    r += 1;
  });
  r += 1;

  // Title bar
  mergeRow(r);
  const titleCell = sheet.getCell(r, 1);
  titleCell.value = "COTIZACIÓN ELECTRODOMÉSTICOS";
  styleCell(titleCell, { fill: ACCENT, bold: true, color: WHITE, size: 11, align: "center" });
  sheet.getRow(r).height = 22;
  r += 1;

  // Cliente / Obra
  [["Cliente:", cotizacion.cliente || ""], ["Obra:", cotizacion.obra || ""]].forEach(([label, value]) => {
    const labelCell = sheet.getCell(r, 1);
    labelCell.value = label;
    styleCell(labelCell, { bold: true });
    sheet.getCell(r, 2).value = value;
    styleCell(sheet.getCell(r, 2));
    for (let c = 3; c <= lastCol; c++) styleCell(sheet.getCell(r, c));
    sheet.mergeCells(r, 2, r, lastCol);
    sheet.getRow(r).height = 16;
    r += 1;
  });

  // Tabla — encabezado
  const headers = [
    "Producto", "Foto Ref.", "Descripción",
    ...(showEspec ? [especLabel] : []),
    "Cant.", "Precio Unit. U$S", "TOTAL U$S",
  ];
  headers.forEach((h, i) => {
    const cell = sheet.getCell(r, i + 1);
    cell.value = h;
    styleCell(cell, { fill: ACCENT_LIGHT, bold: true, color: "FF565A5F", align: "center" });
  });
  sheet.getRow(r).height = 20;
  const headerRow = r;
  r += 1;

  let subtotal = 0;
  for (const l of lineas) {
    const cant = Number(l.cantidad) || 0;
    const precio = Number(l.precioUnit) || 0;
    const total = cant * precio;
    subtotal += total;

    const row = sheet.getRow(r);
    row.height = 48;

    let c = 1;
    const codigoCell = sheet.getCell(r, c); codigoCell.value = l.codigo || ""; styleCell(codigoCell, { align: "center", wrap: true }); c += 1;

    // wrap:true acá también aunque la celda no tenga texto largo: si otra celda alineada al
    // centro en esta fila (ej. el código) pide wrap y ésta no, ExcelJS puede terminar
    // compartiendo un mismo estilo interno entre ambas y perder el ajuste de texto en la que sí
    // lo necesitaba.
    const fotoCell = sheet.getCell(r, c); styleCell(fotoCell, { align: "center", wrap: true });
    if (l.foto) {
      try {
        // No todas las fotos de producto son JPEG (Multi Split Interior, Termocalefones son
        // PNG) — declarar "jpeg" a mano para bytes que en realidad son PNG hacía que ExcelJS
        // dejara la celda vacía en silencio. Se detecta el formato real del propio data URL.
        const extension = l.foto.startsWith("data:image/png") ? "png" : "jpeg";
        const imgId = workbook.addImage({ base64: l.foto, extension });
        sheet.addImage(imgId, { tl: { col: c - 1 + 0.15, row: r - 1 + 0.1 }, ext: { width: 50, height: 50 } });
      } catch (e) { /* foto corrupta o formato no soportado — se deja la celda vacía */ }
    }
    c += 1;

    const descCell = sheet.getCell(r, c); descCell.value = l.descripcion || ""; styleCell(descCell, { wrap: true }); c += 1;

    if (showEspec) {
      const espCell = sheet.getCell(r, c); espCell.value = l.especValor || ""; styleCell(espCell, { align: "center", wrap: true }); c += 1;
    }

    const cantCell = sheet.getCell(r, c); cantCell.value = cant; styleCell(cantCell, { align: "center", wrap: true }); c += 1;
    const precioCell = sheet.getCell(r, c); precioCell.value = precio; precioCell.numFmt = "#,##0.00"; styleCell(precioCell, { align: "right", wrap: true }); c += 1;
    const totalCell = sheet.getCell(r, c); totalCell.value = total; totalCell.numFmt = "#,##0.00"; styleCell(totalCell, { align: "right", wrap: true });

    r += 1;
  }

  const incluirDescuento = !!cotizacion.incluirDescuento;
  const incluirInstalacion = !!cotizacion.incluirInstalacion;
  const descuentoValor = incluirDescuento ? Number(cotizacion.descuento) || 0 : 0;
  const descuentoEsPorcentaje = !!cotizacion.descuentoEsPorcentaje;
  const descuentoMonto = descuentoEsPorcentaje ? (subtotal * descuentoValor) / 100 : descuentoValor;
  const totalConDescuento = subtotal - descuentoMonto;
  const instalacionMonto = incluirInstalacion ? Number(cotizacion.instalacionMonto) || 0 : 0;
  const totalFinal = totalConDescuento + instalacionMonto;

  function totalRow(label, value, opts = {}) {
    const labelCell = sheet.getCell(r, 1);
    labelCell.value = label;
    styleCell(labelCell, { bold: true, fill: opts.fill, color: opts.fill ? "FF565A5F" : undefined });
    sheet.mergeCells(r, 1, r, lastCol - 1);
    for (let c = 2; c < lastCol; c++) styleCell(sheet.getCell(r, c), { fill: opts.fill });
    const valCell = sheet.getCell(r, lastCol);
    valCell.value = value;
    if (typeof value === "number") valCell.numFmt = "#,##0.00";
    styleCell(valCell, { bold: true, fill: opts.fill, color: opts.fill ? "FF565A5F" : undefined, align: "right" });
    r += 1;
  }

  totalRow("Sub-total", subtotal, { fill: ACCENT_LIGHT });
  if (incluirDescuento) {
    totalRow("Descuento", descuentoValor === 0 ? "-" : (descuentoEsPorcentaje ? `${fmtNum(descuentoValor)}%` : descuentoValor));
    totalRow("Total Descuento Incluido", totalConDescuento, { fill: ACCENT_LIGHT });
  }

  if (cotizacion.comentarios) {
    const labelCell = sheet.getCell(r, 1);
    labelCell.value = "Comentarios";
    styleCell(labelCell, { bold: true });
    sheet.getCell(r, 2).value = cotizacion.comentarios;
    styleCell(sheet.getCell(r, 2), { wrap: true });
    sheet.mergeCells(r, 2, r, lastCol);
    // Mismas opciones (wrap incluido) que la celda de arriba: si una celda sin wrap comparte el
    // mismo estilo interno que esta, ExcelJS puede terminar aplicando ESE estilo (sin wrap) a
    // ambas y el texto se corta en vez de ajustarse.
    for (let c = 3; c <= lastCol; c++) styleCell(sheet.getCell(r, c), { wrap: true });
    sheet.getRow(r).height = 30;
    r += 1;
  }

  if (incluirInstalacion) {
    const labelCell = sheet.getCell(r, 1);
    labelCell.value = "Instalaciones";
    styleCell(labelCell, { bold: true });
    sheet.getCell(r, 2).value = cotizacion.instalacionDescripcion || "";
    styleCell(sheet.getCell(r, 2), { wrap: true });
    sheet.mergeCells(r, 2, r, lastCol - 1);
    for (let c = 3; c < lastCol; c++) styleCell(sheet.getCell(r, c), { wrap: true });
    const instCell = sheet.getCell(r, lastCol);
    instCell.value = instalacionMonto === 0 ? "-" : instalacionMonto;
    if (typeof instCell.value === "number") instCell.numFmt = "#,##0.00";
    styleCell(instCell, { bold: true, align: "right" });
    r += 1;
  }

  // Un renglón de aire entre el bloque de subtotal/descuento/instalación y el total final —
  // sin esto quedaban pegados y se leía como una sola masa de filas.
  r += 1;
  totalRow("TOTAL IVA INCLUIDO", totalFinal, { fill: ACCENT_LIGHT });

  // Etiqueta + valor en columnas separadas (igual que Cliente:/Obra:), no un solo texto largo
  // fusionado en toda la fila — más prolijo y más fácil de leer de un vistazo.
  const letrasLabelCell = sheet.getCell(r, 1);
  letrasLabelCell.value = "Dólares Americanos:";
  styleCell(letrasLabelCell, { bold: true });
  sheet.getCell(r, 2).value = montoEnLetras(totalFinal);
  styleCell(sheet.getCell(r, 2));
  for (let c = 3; c <= lastCol; c++) styleCell(sheet.getCell(r, c));
  sheet.mergeCells(r, 2, r, lastCol);
  r += 1;

  const entregaLabelCell = sheet.getCell(r, 1);
  entregaLabelCell.value = "Fecha entrega estimada:";
  styleCell(entregaLabelCell, { bold: true });
  sheet.getCell(r, 2).value = cotizacion.fechaEntregaEstimada || "";
  styleCell(sheet.getCell(r, 2), { wrap: true });
  for (let c = 3; c <= lastCol; c++) styleCell(sheet.getCell(r, c), { wrap: true });
  sheet.mergeCells(r, 2, r, lastCol);
  sheet.getRow(r).height = 18;
  r += 1;

  // Texto legal
  const legalTexto = [
    legalTextCotizacion(!!cotizacion.incluirInstalacion, cotizacion.diasValidez),
    `Forma de pago sugerida: ${cotizacion.formaPago || "A conversar"}.`,
    `OBS: ${cotizacion.obs || "Productos a retirar de depósito."}`,
  ].join("\n");
  const legalCell = sheet.getCell(r, 1);
  legalCell.value = legalTexto;
  styleCell(legalCell, { wrap: true, size: 8 });
  sheet.mergeCells(r, 1, r, lastCol);
  for (let c = 2; c <= lastCol; c++) styleCell(sheet.getCell(r, c), { wrap: true, size: 8 });
  sheet.getRow(r).height = 75;
  r += 2;

  // Firma
  if (firmaBuf) {
    const firmaId = workbook.addImage({ buffer: firmaBuf, extension: "png" });
    sheet.addImage(firmaId, { tl: { col: lastCol - 2, row: r - 1 }, ext: { width: 100, height: 45 } });
    r += 3;
  } else {
    r += 1;
  }
  // Sin merge: centrado en esta única columna, la misma donde arranca la firma arriba — si se
  // fusiona con la columna de al lado, Excel centra el texto en todo ese ancho y queda corrido
  // respecto a la imagen, que es más angosta.
  [COMPANY.firmante, COMPANY.razonSocial, COMPANY.ruc].forEach((line) => {
    const cell = sheet.getCell(r, lastCol - 1);
    cell.value = line;
    cell.alignment = { horizontal: "center" };
    cell.font = { size: 8 };
    r += 1;
  });

  sheet.getRow(headerRow).eachCell((cell) => { cell.border = thinBorder(); });

  return workbook.xlsx.writeBuffer();
}

export function nombreArchivoCotizacionExcel(cotizacion) {
  return `Cotizacion_${(cotizacion.cliente || "cliente").replace(/\s+/g, "_")}_${cotizacion.fecha || ""}.xlsx`;
}

export async function downloadCotizacionExcel(cotizacion) {
  const buffer = await generateCotizacionExcelBuffer(cotizacion);
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombreArchivoCotizacionExcel(cotizacion);
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
