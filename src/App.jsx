import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  LayoutDashboard, Package, ArrowUpFromLine, ArrowDownToLine, ShieldCheck,
  Wrench, Plus, Download, Upload, Search, X, Trash2, MessageCircle, AlertTriangle,
  CheckCircle2, Clock, ChevronRight, Boxes, Inbox, ArrowRight, Star, Lock, TrendingUp, Camera,
  Tag, FileText, FileSignature, Pencil, Menu, Hammer, PackageCheck, ScanLine,
} from "lucide-react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
import {
  collection, doc, addDoc, updateDoc, deleteDoc, onSnapshot, query, orderBy,
} from "firebase/firestore";
import { downloadCotizacionPdf, downloadFichasTecnicasPdf, downloadPresupuestoReparacionPdf } from "./pdf";

// ---------- Design tokens (paleta derivada del gris del logo AEON, #686D73) ----------
const INK = "#1C1E20";
const MUTED = "#686D73";
const ACCENT = "#565A5F";
const ACCENT_LIGHT = "#EBEBEC";
const BORDER = "#E4E5E5";
const BG = "#F5F6F8";

const STATUS = {
  "En depósito": { color: "#475569", bg: "#F1F5F9" },
  "Vendido": { color: "#15803D", bg: "#E9F7EF" },
  "En préstamo (muestra)": { color: "#1D4ED8", bg: "#EAF0FE" },
  "Muestra": { color: "#0369A1", bg: "#E0F2FE" },
  "En sustitución": { color: "#6D28D9", bg: "#F1EBFC" },
  "En reparación": { color: "#B45309", bg: "#FDF1E0" },
  "Pendiente de reparación": { color: "#C2410C", bg: "#FEEDE3" },
  "Apto para venta": { color: "#15803D", bg: "#E9F7EF" },
  "Apto para venta con descuento": { color: "#0D9488", bg: "#E3F5F3" },
  "Reservado - unidad de rescate": { color: "#BE185D", bg: "#FBE9F1" },
  "Dado de baja": { color: "#B91C1C", bg: "#FBEAEA" },
};
const ESTADOS = Object.keys(STATUS);
const RECUPERABLE_ESTADOS = [
  "Pendiente de reparación", "En reparación",
  "Apto para venta con descuento", "Reservado - unidad de rescate",
];

const TIPOS_MOVIMIENTO = ["Venta", "Muestra", "Sustitución", "Reparación"];
const TIPOS_ENTRADA = [
  "Importación inicial", "Devolución por error de fábrica",
  "Retorno de muestra", "Retorno de equipo de sustitución", "Retorno de reparación propia",
];
const ESTADOS_RESULTANTES = ["Apto para venta", "Apto para venta con descuento", "Pendiente de reparación", "Muestra", "Reservado - unidad de rescate", "Dado de baja"];

const ORIGENES_PLAYA = ["Técnico", "Gastón", "Cliente", "Otro"];
const DESTINOS_PLAYA = [
  { value: "recuperable", label: "Banco de recuperables", estado: "Pendiente de reparación" },
  { value: "socorro", label: "Equipo de socorro (unidad de rescate)", estado: "Reservado - unidad de rescate" },
  { value: "vendible", label: "Stock vendible", estado: "Apto para venta" },
  { value: "muestra", label: "Muestra (exhibición de calidad)", estado: "Muestra" },
];

// Categorías desde donde se puede dar salida a un producto — cada una filtra un pool distinto
const CATEGORIAS_ORIGEN = [
  { value: "muestras", label: "Muestras", type: "equipo", estados: ["Muestra"] },
  { value: "vendible", label: "Stock vendible", type: "equipo", estados: ["En depósito", "Apto para venta"] },
  { value: "vendible_desc", label: "Stock vendible con descuento", type: "equipo", estados: ["Apto para venta con descuento"] },
  { value: "recuperables", label: "Banco de recuperables", type: "equipo", estados: ["Pendiente de reparación", "En reparación", "Reservado - unidad de rescate"] },
  { value: "playa", label: "Zona de playa", type: "playa" },
  { value: "repuesto", label: "Repuesto", type: "producto-repuesto" },
];

// Motivo de la salida — determina en qué queda el equipo (si sigue siendo un activo a rastrear)
const MOTIVOS_SALIDA = [
  { value: "Venta", estado: "Vendido", trackea: false },
  { value: "Muestra (préstamo)", estado: "En préstamo (muestra)", trackea: true },
  { value: "Sustitución", estado: "En sustitución", trackea: true },
  { value: "Reparación", estado: "En reparación", trackea: true },
];
const MOTIVO_DEFAULT = {
  muestras: "Muestra (préstamo)",
  vendible: "Venta",
  vendible_desc: "Venta",
  recuperables: "Sustitución",
};

const MOTIVOS_BAJA = [
  "Sin arreglo técnico posible",
  "Costo de reparación no conviene",
  "Daño irreparable",
  "Obsoleto / descontinuado",
  "Extraviado / robado",
  "Otro",
];

const DECISIONES_SERVICE = ["Acepta hacer el service", "Rechaza el service", "Sin respuesta (reintentar)"];

// Firestore collection names — one collection per data type, one document per item.
const COLLECTIONS = {
  equipos: "equipos",
  movimientos: "movimientos",
  entradas: "entradas",
  ventas: "ventas",
  playa: "playa",
  ventasComprometidas: "ventasComprometidas",
  productos: "productos",
  cotizaciones: "cotizaciones",
  presupuestosReparacion: "presupuestosReparacion",
};

// ---------- Helpers ----------
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function addMonthsISO(dateStr, months) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  d.setMonth(d.getMonth() + months);
  return d.toISOString().slice(0, 10);
}
function fmtDate(d) {
  if (!d) return "—";
  const [y, m, day] = d.split("-");
  return `${day}/${m}/${y}`;
}
function calcularTotalCotizacion(c) {
  const subtotal = (c.lineas || []).reduce((acc, l) => acc + (Number(l.cantidad) || 0) * (Number(l.precioUnit) || 0), 0);
  const descuentoMonto = c.incluirDescuento
    ? (c.descuentoEsPorcentaje ? subtotal * (Number(c.descuento) || 0) / 100 : Number(c.descuento) || 0)
    : 0;
  return subtotal - descuentoMonto + (c.incluirInstalacion ? Number(c.instalacionMonto) || 0 : 0);
}

// Estado de cumplimiento de la salida de depósito para una cotización Ganada — cuánto de lo
// cotizado ya salió físicamente (via "Generar desde cotización" en Salidas), null si no aplica.
function estadoSalidaCotizacion(c) {
  if (c.estado !== "Ganada") return null;
  const lineas = c.lineas || [];
  if (lineas.length === 0) return null;
  const retiradas = c.lineasRetiradas || {};
  let completas = 0;
  let algo = 0;
  for (const l of lineas) {
    const retirado = Number(retiradas[l.codigo]) || 0;
    if (retirado <= 0) continue;
    if (retirado >= (Number(l.cantidad) || 0)) completas++;
    else algo++;
  }
  if (completas === lineas.length) return "completa";
  if (completas > 0 || algo > 0) return "parcial";
  return "pendiente";
}

// Agrupa cotizaciones por Cliente (constructora/desarrolladora) y, dentro de cada uno,
// por Obra — mismo texto de Obra = misma "cadena" que va mutando en el tiempo. La versión
// más nueva de cada obra (ya vienen ordenadas desc por createdAt) es la "activa": la que
// cuenta para los totales y cuyo estado se puede editar. Las anteriores quedan de historial.
function agruparCotizaciones(cotizaciones) {
  const porCliente = new Map();
  for (const c of cotizaciones) {
    const clienteKey = (c.cliente || "").trim() || "(Sin cliente)";
    const obraKey = (c.obra || "").trim() || "(Sin obra)";
    if (!porCliente.has(clienteKey)) porCliente.set(clienteKey, new Map());
    const porObra = porCliente.get(clienteKey);
    if (!porObra.has(obraKey)) porObra.set(obraKey, []);
    porObra.get(obraKey).push(c);
  }
  const clientes = [];
  for (const [cliente, porObra] of porCliente) {
    const obras = [];
    for (const [obra, versiones] of porObra) obras.push({ obra, versiones, activa: versiones[0] });
    obras.sort((a, b) => (b.activa.createdAt || 0) - (a.activa.createdAt || 0));
    clientes.push({ cliente, obras });
  }
  clientes.sort((a, b) => {
    const masReciente = (g) => Math.max(...g.obras.map((o) => o.activa.createdAt || 0));
    return masReciente(b) - masReciente(a);
  });
  return clientes;
}

const ESTADOS_COTIZACION = ["Pendiente", "Ganada", "Perdida"];
const ESTADO_COTIZACION_BADGE = {
  Pendiente: { color: "#B45309", bg: "#FDF1E0" },
  Ganada: { color: "#15803D", bg: "#E9F7EF" },
  Perdida: { color: "#B91C1C", bg: "#FBEAEA" },
};
// Resume Total cotizado / Ganadas / Perdidas / Pendientes tomando solo la versión activa
// de cada obra dentro de la lista de grupos de cliente que se le pase.
function resumirCotizaciones(clientes) {
  const activas = clientes.flatMap((g) => g.obras.map((o) => o.activa));
  const resumen = { total: 0, Ganada: { n: 0, total: 0 }, Perdida: { n: 0, total: 0 }, Pendiente: { n: 0, total: 0 } };
  for (const c of activas) {
    const monto = calcularTotalCotizacion(c);
    const estado = ESTADOS_COTIZACION.includes(c.estado) ? c.estado : "Pendiente";
    resumen.total += monto;
    resumen[estado].n += 1;
    resumen[estado].total += monto;
  }
  return resumen;
}

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const today = new Date(todayISO() + "T00:00:00");
  const target = new Date(dateStr + "T00:00:00");
  return Math.round((target - today) / (1000 * 60 * 60 * 24));
}
function addBusinessDaysISO(dateStr, days) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  let added = 0;
  while (added < days) {
    d.setDate(d.getDate() + 1);
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) added++;
  }
  return d.toISOString().slice(0, 10);
}
function compressImage(file, maxDim = 1280, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const reader = new FileReader();
    reader.onload = (e) => { img.src = e.target.result; };
    reader.onerror = reject;
    img.onload = () => {
      let { width, height } = img;
      if (width > height && width > maxDim) { height = Math.round(height * (maxDim / width)); width = maxDim; }
      else if (height > maxDim) { width = Math.round(width * (maxDim / height)); height = maxDim; }
      const canvas = document.createElement("canvas");
      canvas.width = width; canvas.height = height;
      canvas.getContext("2d").drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL("image/jpeg", quality));
    };
    img.onerror = reject;
    reader.readAsDataURL(file);
  });
}
// Subscribes to a Firestore collection in real time — every client sharing
// the same Firebase project sees updates from everyone else immediately.
function subscribeCollection(name, onData) {
  const q = query(collection(db, name), orderBy("createdAt", "desc"));
  return onSnapshot(q, (snap) => {
    onData(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
  }, (err) => {
    // A permission error (or any other) on one collection shouldn't leave the
    // whole app stuck on the loading screen — fall back to empty.
    console.error("Firestore subscribe error", name, err);
    onData([]);
  });
}
function addItem(name, data) {
  return addDoc(collection(db, name), { ...data, createdAt: Date.now() })
    .catch((e) => console.error("Firestore add error", name, e));
}
function updateItem(name, id, patch) {
  return updateDoc(doc(db, name, id), patch)
    .catch((e) => console.error("Firestore update error", name, id, e));
}
// El navegador nativo (window.confirm) resultó poco confiable en algunos celulares/PWA —
// a veces no aparece o queda bloqueado por el navegador sin avisar. `confirmBridge` lo
// conecta con un popup propio de la app (ver ConfirmDialog), montado una vez desde App.
let confirmBridge = null;
async function deleteItem(name, id) {
  const confirmado = confirmBridge
    ? await confirmBridge("¿Eliminar este registro? Esta acción no se puede deshacer.")
    : window.confirm("¿Eliminar este registro? Esta acción no se puede deshacer.");
  if (!confirmado) return;
  return deleteDoc(doc(db, name, id))
    .catch((e) => console.error("Firestore delete error", name, id, e));
}

// Ficha técnica PDFs are stored as base64 directly on the producto document,
// so they have to stay well under Firestore's 1MB document limit (base64
// adds ~33% overhead on top of the raw file size).
const MAX_FICHA_BYTES = 650 * 1024;
function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ---------- Importación masiva de catálogo por Excel ----------
// La foto y la ficha técnica (PDF) no vienen del Excel — se cargan después,
// a mano, editando cada producto.
function normalizarHeader(h) {
  return String(h || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim().replace(/\s+/g, " ");
}
const CAMPOS_IMPORT_CATALOGO = [
  { campo: "nombre", alias: ["nombre", "nombre / codigo", "nombre / codigo del producto", "codigo"] },
  { campo: "categoria", alias: ["categoria"] },
  { campo: "categoriaPrincipal", alias: ["categoria principal"] },
  { campo: "subcategoria", alias: ["subcategoria"] },
  { campo: "subcategoria2", alias: ["subcategoria 2"] },
  { campo: "subcategoria3", alias: ["subcategoria 3"] },
  { campo: "ordenNumerico", alias: ["orden", "orden (numero)", "orden dentro del grupo"] },
  { campo: "descripcion", alias: ["descripcion", "descripcion (aparece en la cotizacion)"] },
  { campo: "especLabel", alias: ["especificacion - etiqueta", "espec. etiqueta", "especificacion etiqueta"] },
  { campo: "especValor", alias: ["especificacion - valor", "espec. valor", "especificacion valor"] },
  { campo: "precioLista", alias: ["precio de lista u$s", "precio de lista", "precio u$s", "precio", "precio de venta u$s"] },
  { campo: "costoOrigen", alias: ["costo de origen u$s", "costo de origen", "costo origen"] },
  { campo: "costoPy", alias: ["costo puesto en py u$s", "costo puesto en py", "costo py"] },
  { campo: "contenedorTipo", alias: ["contenedor tipo", "contenedor (tipo)"] },
  { campo: "contenedorCantidad", alias: ["contenedor cantidad", "cantidad por contenedor"] },
  { campo: "stockDisponible", alias: ["stock disponible", "stock", "cantidad en stock"] },
];
function parseCatalogoExcel(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  const productos = [];
  const errores = [];
  rows.forEach((row, i) => {
    const item = {};
    for (const key in row) {
      const norm = normalizarHeader(key);
      const match = CAMPOS_IMPORT_CATALOGO.find((c) => c.alias.includes(norm));
      if (match) item[match.campo] = row[key];
    }
    const nombre = String(item.nombre || "").trim();
    if (!nombre) {
      errores.push(`Fila ${i + 2}: sin nombre/código, se omitió.`);
      return;
    }
    productos.push({
      nombre,
      categoria: String(item.categoria || "").trim(),
      categoriaPrincipal: String(item.categoriaPrincipal || "").trim(),
      subcategoria: String(item.subcategoria || "").trim(),
      subcategoria2: String(item.subcategoria2 || "").trim(),
      subcategoria3: String(item.subcategoria3 || "").trim(),
      ordenNumerico: item.ordenNumerico === "" || item.ordenNumerico == null ? null : Number(item.ordenNumerico) || 0,
      descripcion: String(item.descripcion || "").trim(),
      especLabel: String(item.especLabel || "").trim(),
      especValor: String(item.especValor || "").trim(),
      precioLista: Number(item.precioLista) || 0,
      costoOrigen: Number(item.costoOrigen) || 0,
      costoPy: Number(item.costoPy) || 0,
      contenedorTipo: String(item.contenedorTipo || "").trim(),
      contenedorCantidad: Number(item.contenedorCantidad) || 0,
      stockDisponible: item.stockDisponible === "" || item.stockDisponible == null ? null : Number(item.stockDisponible) || 0,
      foto: "", fichaTecnicaData: "", fichaTecnicaNombre: "",
    });
  });
  return { productos, errores };
}
function descargarPlantillaCatalogo() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet([{
    "Nombre / código": "AE-AC-2T-30-ON",
    "Categoría principal": "Cocina", "Subcategoría": "Anafe", "Subcategoría 2": "Vitrocerámica", "Subcategoría 3": "2",
    "Orden (número)": 2,
    "Descripción": "Anafe vitrocerámica de 2 quemadores, Voltaje 220-240V, 50~60Hz, Potencia nominal: 3000W",
    "Especificación - etiqueta": "Potencia nominal", "Especificación - valor": "3000W",
    "Precio de lista U$S": 99, "Costo de origen U$S": 55, "Costo puesto en PY U$S": 70,
    "Contenedor tipo": "40HQ", "Contenedor cantidad": 420, "Stock disponible": 5,
  }]);
  XLSX.utils.book_append_sheet(wb, ws, "Catálogo");
  XLSX.writeFile(wb, "Plantilla_Catalogo_AEON.xlsx");
}

// ---------- Small UI atoms ----------
function StatusBadge({ estado }) {
  const s = STATUS[estado] || { color: MUTED, bg: "#F1F5F9" };
  return (
    <span
      className="inline-flex items-center px-2 py-1 rounded text-xs font-medium whitespace-nowrap"
      style={{ color: s.color, backgroundColor: s.bg }}
    >
      {estado}
    </span>
  );
}

function CodeTag({ children }) {
  return (
    <span
      className="font-mono text-xs px-1.5 py-0.5 rounded tracking-wide"
      style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}
    >
      {children}
    </span>
  );
}

function EmptyState({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex flex-col items-center justify-center py-16 text-center">
      <div className="w-12 h-12 rounded-full flex items-center justify-center mb-3" style={{ backgroundColor: ACCENT_LIGHT }}>
        <Icon size={22} style={{ color: ACCENT }} />
      </div>
      <p className="text-sm font-medium" style={{ color: INK }}>{title}</p>
      <p className="text-sm mt-1" style={{ color: MUTED }}>{subtitle}</p>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block mb-3">
      <span className="block text-xs font-medium mb-1" style={{ color: MUTED }}>{label}</span>
      {children}
    </label>
  );
}

const inputClass = "w-full text-sm px-3 py-2 rounded-md border outline-none";
const inputStyle = { borderColor: BORDER, color: INK };

function TextInput(props) {
  return <input {...props} className={inputClass} style={inputStyle} />;
}
function Select({ children, ...props }) {
  return <select {...props} className={inputClass} style={inputStyle}>{children}</select>;
}
function Textarea(props) {
  return <textarea rows={3} {...props} className={inputClass} style={inputStyle} />;
}

function PrimaryButton({ children, onClick, type = "button", disabled = false }) {
  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled}
      className="inline-flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-md text-white"
      style={{ backgroundColor: ACCENT, opacity: disabled ? 0.6 : 1, cursor: disabled ? "default" : "pointer" }}
    >
      {children}
    </button>
  );
}
function SecondaryButton({ children, onClick }) {
  return (
    <button
      onClick={onClick}
      className="inline-flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-md border"
      style={{ borderColor: BORDER, color: INK }}
    >
      {children}
    </button>
  );
}

function Drawer({ open, onClose, title, children }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex justify-end" style={{ backgroundColor: "rgba(15,23,32,0.4)" }} onClick={onClose}>
      <div
        className="w-full max-w-md h-full bg-white overflow-y-auto shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b" style={{ borderColor: BORDER }}>
          <h3 className="text-sm font-semibold" style={{ color: INK }}>{title}</h3>
          <button onClick={onClose} className="p-1 rounded hover:bg-gray-100">
            <X size={18} style={{ color: MUTED }} />
          </button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function ConfirmDialog({ state, onResolve }) {
  if (!state) return null;
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-6"
      style={{ backgroundColor: "rgba(15,23,32,0.75)" }}
      onClick={() => onResolve(false)}
    >
      <div
        className="rounded-xl p-5 w-full"
        style={{ backgroundColor: "#FFFFFF", maxWidth: 340 }}
        onClick={(e) => e.stopPropagation()}
      >
        <p className="text-sm" style={{ color: INK }}>{state.message}</p>
        <div className="flex justify-end gap-2 mt-4">
          <button
            onClick={() => onResolve(false)}
            className="text-sm px-3.5 py-1.5 rounded-md font-medium"
            style={{ backgroundColor: "#FFFFFF", color: MUTED, border: `0.5px solid ${BORDER}` }}
          >
            Cancelar
          </button>
          <button
            onClick={() => onResolve(true)}
            className="text-sm px-3.5 py-1.5 rounded-md font-medium"
            style={{ backgroundColor: "#B91C1C", color: "#FFFFFF" }}
          >
            Eliminar
          </button>
        </div>
      </div>
    </div>
  );
}

function PhotoViewer({ src, onClose }) {
  if (!src) return null;
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-6"
      style={{ backgroundColor: "rgba(15,23,32,0.75)" }}
      onClick={onClose}
    >
      <button onClick={onClose} className="absolute top-4 right-4 p-2 rounded-full" style={{ backgroundColor: "rgba(255,255,255,0.15)" }}>
        <X size={20} color="#FFFFFF" />
      </button>
      <img
        src={src}
        alt="Foto del remito"
        className="rounded-lg"
        style={{ maxWidth: "90%", maxHeight: "90%", objectFit: "contain" }}
        onClick={(e) => e.stopPropagation()}
      />
    </div>
  );
}

function Table({ columns, rows, onDelete, renderCell }) {
  if (rows.length === 0) {
    return <EmptyState icon={Package} title="Todavía no hay registros" subtitle="Usá el botón de arriba para cargar el primero." />;
  }
  return (
    <div className="overflow-auto rounded-lg border" style={{ borderColor: BORDER, maxHeight: "80vh" }}>
      <table className="text-sm" style={{ minWidth: 720 }}>
        <thead>
          <tr>
            {columns.map((c) => (
              <th key={c.key} className="text-left font-medium px-3 py-2 border-b sticky top-0 z-10 whitespace-nowrap" style={{ color: MUTED, borderColor: BORDER, fontSize: 12, backgroundColor: "#FAFBFC" }}>
                {c.label}
              </th>
            ))}
            {onDelete && <th className="w-10 border-b sticky top-0 z-10" style={{ borderColor: BORDER, backgroundColor: "#FAFBFC" }}></th>}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id} className="border-b last:border-0" style={{ borderColor: BORDER }}>
              {columns.map((c) => (
                <td key={c.key} className="px-3 py-1.5 whitespace-nowrap" style={{ color: INK, fontSize: 13 }}>
                  {renderCell ? renderCell(c.key, row) : row[c.key] || "—"}
                </td>
              ))}
              {onDelete && (
                <td className="px-2 py-1.5">
                  <button onClick={() => onDelete(row.id)} className="p-1 rounded hover:bg-gray-100">
                    <Trash2 size={14} style={{ color: MUTED }} />
                  </button>
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function SearchBox({ value, onChange, placeholder }) {
  return (
    <div className="relative w-64">
      <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder || "Buscar..."}
        className="w-full text-sm pl-8 pr-3 py-2 rounded-md border outline-none"
        style={inputStyle}
      />
    </div>
  );
}

// Código interno = modelo + secuencia por unidad de ese modelo (ej. "ASAFB-12HRN1-01"), en vez
// de un contador global sin relación con el producto — así el código ya dice qué es de un vistazo,
// y una vez que empiecen a llegar series reales de fábrica, cada unidad de un mismo modelo se
// distingue con su propio número dentro de esa familia.
function nextCodigoParaModelo(equipos, modelo) {
  const base = (modelo || "").trim();
  if (!base) return "";
  const prefix = `${base}-`;
  let max = 0;
  equipos.forEach((e) => {
    if ((e.codigo || "").startsWith(prefix)) {
      const n = parseInt((e.codigo || "").slice(prefix.length), 10);
      if (!isNaN(n)) max = Math.max(max, n);
    }
  });
  return `${base}-${String(max + 1).padStart(2, "0")}`;
}

// ---------- Main App ----------
export default function App() {
  const [tab, setTab] = useState("resumen");
  const [loading, setLoading] = useState(true);
  const [equipos, setEquipos] = useState([]);
  const [movimientos, setMovimientos] = useState([]);
  const [entradas, setEntradas] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [playa, setPlaya] = useState([]);
  const [comprometidas, setComprometidas] = useState([]);
  const [productos, setProductos] = useState([]);
  const [cotizaciones, setCotizaciones] = useState([]);
  const [presupuestosReparacion, setPresupuestosReparacion] = useState([]);
  const [query, setQuery] = useState("");
  const [drawer, setDrawer] = useState(null);
  const [gestion, setGestion] = useState(null);
  const [fotoView, setFotoView] = useState(null);
  const [confirmState, setConfirmState] = useState(null);

  useEffect(() => {
    confirmBridge = (message) => new Promise((resolve) => setConfirmState({ message, resolve }));
    return () => { confirmBridge = null; };
  }, []);
  const resolveConfirm = (value) => {
    confirmState?.resolve(value);
    setConfirmState(null);
  };
  const [retiroTarget, setRetiroTarget] = useState(null);
  const [pagoTarget, setPagoTarget] = useState(null);
  const [productoEditando, setProductoEditando] = useState(null);
  const [nuevoProductoDefaults, setNuevoProductoDefaults] = useState(null);
  const [navOpen, setNavOpen] = useState(false);
  const [catalogoModoInicial, setCatalogoModoInicial] = useState(null);
  const [descargandoId, setDescargandoId] = useState(null);
  const [pdfError, setPdfError] = useState("");
  const [importandoCatalogo, setImportandoCatalogo] = useState(false);
  const [importResultado, setImportResultado] = useState("");

  useEffect(() => {
    const setters = {
      [COLLECTIONS.equipos]: setEquipos,
      [COLLECTIONS.movimientos]: setMovimientos,
      [COLLECTIONS.entradas]: setEntradas,
      [COLLECTIONS.ventas]: setVentas,
      [COLLECTIONS.playa]: setPlaya,
      [COLLECTIONS.ventasComprometidas]: setComprometidas,
      [COLLECTIONS.productos]: setProductos,
      [COLLECTIONS.cotizaciones]: setCotizaciones,
      [COLLECTIONS.presupuestosReparacion]: setPresupuestosReparacion,
    };
    const names = Object.keys(setters);
    const pending = new Set(names);
    const unsubscribers = names.map((name) =>
      subscribeCollection(name, (items) => {
        setters[name](items);
        pending.delete(name);
        if (pending.size === 0) setLoading(false);
      })
    );
    return () => unsubscribers.forEach((unsub) => unsub());
  }, []);

  const addEquipo = (data) => addItem(COLLECTIONS.equipos, data);
  const updateEquipoEstado = (id, estado) => updateItem(COLLECTIONS.equipos, id, { estado });
  const updateEquipoField = (id, field, value) => updateItem(COLLECTIONS.equipos, id, { [field]: value });
  const deleteEquipo = (id) => deleteItem(COLLECTIONS.equipos, id);

  // Carga masiva por pistola lectora: una unidad individual (cantidad 1) por cada código
  // escaneado, con secuencia propia dentro de ese modelo (modelo-01, modelo-02, ...).
  const addEquiposPorEscaneo = (data) => {
    const prefix = `${data.modelo}-`;
    let max = 0;
    equipos.forEach((e) => {
      if ((e.codigo || "").startsWith(prefix)) {
        const n = parseInt((e.codigo || "").slice(prefix.length), 10);
        if (!isNaN(n)) max = Math.max(max, n);
      }
    });
    data.series.forEach((serie) => {
      max++;
      addItem(COLLECTIONS.equipos, {
        codigo: `${data.modelo}-${String(max).padStart(2, "0")}`,
        serie, modelo: data.modelo, fechaIngreso: data.fechaIngreso,
        estado: data.estado, ubicacion: data.ubicacion, cantidad: 1, notas: "",
      });
    });
  };

  const addMovimiento = (data) => {
    // data: { fecha, categoria, sourceId, codigo, modelo, cantidad, motivo, cliente, remito, responsable, observaciones }
    addItem(COLLECTIONS.movimientos, data);

    const cat = CATEGORIAS_ORIGEN.find((c) => c.value === data.categoria);
    if (!cat) return;
    const cantidadRetirada = Number(data.cantidad) || 1;

    if (cat.type === "equipo") {
      const source = equipos.find((e) => e.id === data.sourceId);
      if (source) {
        const restante = (Number(source.cantidad) || 1) - cantidadRetirada;
        if (restante > 0) updateItem(COLLECTIONS.equipos, source.id, { cantidad: restante });
        else deleteItem(COLLECTIONS.equipos, source.id);
      }
      const motivo = MOTIVOS_SALIDA.find((m) => m.value === data.motivo);
      if (motivo && motivo.trackea) {
        // El equipo sigue siendo un activo a rastrear (préstamo, sustitución, reparación): se le crea una nueva fila con nuevo código
        addItem(COLLECTIONS.equipos, {
          codigo: nextCodigoParaModelo(equipos, data.modelo), serie: "", modelo: data.modelo,
          fechaIngreso: data.fecha, estado: motivo.estado,
          ubicacion: `Fuera de depósito (${data.cliente || "cliente"})`,
          cantidad: cantidadRetirada, notas: data.observaciones || "",
        });
      }
    } else if (cat.type === "playa") {
      const source = playa.find((p) => p.id === data.sourceId);
      if (source) {
        const restante = (Number(source.cantidad) || 1) - cantidadRetirada;
        if (restante > 0) updateItem(COLLECTIONS.playa, source.id, { cantidad: restante });
        else deleteItem(COLLECTIONS.playa, source.id);
      }
    } else if (cat.type === "producto-repuesto") {
      // A diferencia de equipo/playa, el repuesto es una ficha permanente del catálogo — nunca
      // se borra al llegar a 0, solo se descuenta su stock (piso en 0).
      const source = productos.find((p) => p.id === data.sourceId);
      if (source) {
        const restante = Math.max(0, (Number(source.stockDisponible) || 0) - cantidadRetirada);
        updateItem(COLLECTIONS.productos, source.id, { stockDisponible: restante });
      }
    }
  };
  const deleteMovimiento = (id) => deleteItem(COLLECTIONS.movimientos, id);

  // Genera una o más salidas (una por cada lote de equipo usado) a partir de una cotización
  // Ganada. `lineas` viene de SalidaDesdeCotizacionForm ya resuelto contra stock disponible:
  // [{ codigo, descripcion, precioUnit, cantidad, batches: [{ equipoId, codigo, categoria, cantidad }] }]
  const generarSalidaDesdeCotizacion = (cotizacion, lineas, remitoData) => {
    const retiros = {};
    for (const linea of lineas) {
      for (const batch of linea.batches) {
        const cat = CATEGORIAS_ORIGEN.find((c) => c.value === batch.categoria);
        addMovimiento({
          fecha: remitoData.fecha,
          categoria: batch.categoria, categoriaLabel: cat ? cat.label : "Stock vendible",
          sourceId: batch.equipoId, codigo: batch.codigo, modelo: linea.codigo,
          cantidad: batch.cantidad, motivo: "Venta",
          cliente: cotizacion.cliente, obra: cotizacion.obra,
          monto: (Number(linea.precioUnit) || 0) * batch.cantidad,
          remito: remitoData.remito, responsable: remitoData.responsable,
          observaciones: remitoData.observaciones || `Generado desde cotización — ${cotizacion.cliente}${cotizacion.obra ? " / " + cotizacion.obra : ""}`,
          lugarSalida: remitoData.lugarSalida, empresaCliente: remitoData.empresaCliente, rucCliente: remitoData.rucCliente,
          firmaNombre: remitoData.firmaNombre, firmaCedula: remitoData.firmaCedula, fotoRemito: remitoData.fotoRemito,
          cotizacionId: cotizacion.id,
        });
      }
      retiros[linea.codigo] = (retiros[linea.codigo] || 0) + linea.cantidad;
    }
    const nuevasRetiradas = { ...(cotizacion.lineasRetiradas || {}) };
    for (const [codigo, cant] of Object.entries(retiros)) {
      nuevasRetiradas[codigo] = (Number(nuevasRetiradas[codigo]) || 0) + cant;
    }
    updateItem(COLLECTIONS.cotizaciones, cotizacion.id, { lineasRetiradas: nuevasRetiradas });
  };

  const addEntrada = (data) => {
    addItem(COLLECTIONS.entradas, data);
    if (data.codigo && data.estadoResultante) updateEquipoEstadoByCodigo(data.codigo, data.estadoResultante);
  };
  const deleteEntrada = (id) => deleteItem(COLLECTIONS.entradas, id);

  function updateEquipoEstadoByCodigo(codigo, estado) {
    equipos.filter((e) => e.codigo === codigo).forEach((e) => updateItem(COLLECTIONS.equipos, e.id, { estado }));
  }

  const addVenta = (data) => addItem(COLLECTIONS.ventas, data);
  const updateVentaEstado = (id, field, value) => updateItem(COLLECTIONS.ventas, id, { [field]: value });
  const deleteVenta = (id) => deleteItem(COLLECTIONS.ventas, id);

  // Venta comprometida: reserva stock (queda físicamente en depósito, pero no disponible para otra salida)
  const addComprometida = (data) => {
    const equipo = equipos.find((e) => e.id === data.equipoId);
    if (equipo) {
      updateItem(COLLECTIONS.equipos, equipo.id, { comprometido: (Number(equipo.comprometido) || 0) + Number(data.cantidad) });
    }
    addItem(COLLECTIONS.ventasComprometidas, { estado: "Comprometida", ...data });
  };

  const cancelarComprometida = (id) => {
    const c = comprometidas.find((x) => x.id === id);
    if (!c) return;
    if (c.estado === "Comprometida") {
      const equipo = equipos.find((e) => e.id === c.equipoId);
      if (equipo) {
        updateItem(COLLECTIONS.equipos, equipo.id, { comprometido: Math.max(0, (Number(equipo.comprometido) || 0) - Number(c.cantidad)) });
      }
    }
    deleteItem(COLLECTIONS.ventasComprometidas, id);
  };

  // Retiro parcial: descuenta del saldo comprometido lo que se entrega en esta tanda (con su propia
  // ficha de remito), deja el registro en Salidas, y marca la reserva "Completada" cuando el saldo
  // llega a 0 — pero la deja abierta hasta que el usuario la cierre a mano con cerrarComprometida.
  const retirarParcial = (id, data) => {
    const c = comprometidas.find((x) => x.id === id);
    if (!c) return;
    const equipo = equipos.find((e) => e.id === c.equipoId);
    if (!equipo) return;

    const cantidadEvento = Number(data.cantidad) || 0;
    const restanteEquipo = (Number(equipo.cantidad) || 1) - cantidadEvento;
    const comprometidoRestante = Math.max(0, (Number(equipo.comprometido) || 0) - cantidadEvento);
    if (restanteEquipo > 0) updateItem(COLLECTIONS.equipos, equipo.id, { cantidad: restanteEquipo, comprometido: comprometidoRestante });
    else deleteItem(COLLECTIONS.equipos, equipo.id);

    addItem(COLLECTIONS.movimientos, {
      fecha: data.fecha || todayISO(), categoria: "vendible", categoriaLabel: "Stock vendible",
      codigo: equipo.codigo, modelo: equipo.modelo, cantidad: cantidadEvento, motivo: "Venta",
      cliente: c.razonSocial, obra: data.obra || c.obra, monto: 0,
      remito: data.remito || "", responsable: data.responsable || "",
      lugarSalida: data.lugarSalida || "", empresaCliente: data.empresaCliente || "", rucCliente: data.rucCliente || "",
      firmaNombre: data.firmaNombre || "", firmaCedula: data.firmaCedula || "", fotoRemito: data.fotoRemito || "",
      observaciones: data.observaciones ? `Retiro parcial de venta comprometida — ${data.observaciones}` : "Retiro parcial de venta comprometida",
    });

    const cantidadRetirada = (Number(c.cantidadRetirada) || 0) + cantidadEvento;
    const patch = { cantidadRetirada };
    if (cantidadRetirada >= Number(c.cantidad)) patch.estado = "Completada";
    updateItem(COLLECTIONS.ventasComprometidas, id, patch);
  };

  // Cierre manual: el usuario confirma que ya no hay más que retirar de esta reserva.
  const cerrarComprometida = (id) => updateItem(COLLECTIONS.ventasComprometidas, id, { estado: "Retirada", fechaCierre: todayISO() });

  // Cobro: independiente del retiro de mercadería — se puede cobrar sin haber entregado,
  // o entregar sin haber cobrado del todo. Se guarda como lista de pagos en la propia reserva.
  const agregarPago = (id, data) => {
    const c = comprometidas.find((x) => x.id === id);
    if (!c) return;
    const pagos = [...(c.pagos || []), { fecha: data.fecha || todayISO(), monto: Number(data.monto) || 0, formaPago: data.formaPago || "" }];
    updateItem(COLLECTIONS.ventasComprometidas, id, { pagos });
  };
  const quitarPago = (id, idx) => {
    const c = comprometidas.find((x) => x.id === id);
    if (!c) return;
    const pagos = (c.pagos || []).filter((_, i) => i !== idx);
    updateItem(COLLECTIONS.ventasComprometidas, id, { pagos });
  };

  // Catálogo de productos: precio de lista y ficha técnica de cada modelo — de acá salen
  // el precio sugerido en ventas comprometidas y el contenido de las cotizaciones.
  const addProducto = (data) => addItem(COLLECTIONS.productos, data);
  const updateProducto = (id, data) => updateItem(COLLECTIONS.productos, id, data);
  const deleteProducto = (p) => deleteItem(COLLECTIONS.productos, p.id);
  const quitarFichaTecnica = (producto) =>
    updateItem(COLLECTIONS.productos, producto.id, { fichaTecnicaData: "", fichaTecnicaNombre: "" });

  const addCotizacion = (data) => addItem(COLLECTIONS.cotizaciones, data);
  const deleteCotizacion = (id) => deleteItem(COLLECTIONS.cotizaciones, id);
  const updateCotizacion = (id, patch) => updateItem(COLLECTIONS.cotizaciones, id, patch);

  const handleDescargarPdf = async (cotizacion) => {
    setDescargandoId(cotizacion.id + ":pdf");
    setPdfError("");
    try {
      await downloadCotizacionPdf(cotizacion);
    } catch (e) {
      console.error("Error generando PDF", e);
      setPdfError("No se pudo generar el PDF de la cotización. Probá de nuevo.");
    }
    setDescargandoId(null);
  };

  const handleDescargarFichas = async (cotizacion) => {
    setDescargandoId(cotizacion.id + ":fichas");
    setPdfError("");
    try {
      const ok = await downloadFichasTecnicasPdf(cotizacion);
      if (!ok) setPdfError("Ninguno de los productos de esta cotización tiene ficha técnica cargada.");
    } catch (e) {
      console.error("Error generando fichas técnicas", e);
      setPdfError("No se pudo generar el PDF de fichas técnicas. Probá de nuevo.");
    }
    setDescargandoId(null);
  };

  const addPresupuestoReparacion = (data) => addItem(COLLECTIONS.presupuestosReparacion, data);
  const deletePresupuestoReparacion = (id) => deleteItem(COLLECTIONS.presupuestosReparacion, id);

  const handleDescargarPresupuestoPdf = async (presupuesto) => {
    setDescargandoId(presupuesto.id + ":reparacion");
    setPdfError("");
    try {
      await downloadPresupuestoReparacionPdf(presupuesto);
    } catch (e) {
      console.error("Error generando PDF de presupuesto de reparación", e);
      setPdfError("No se pudo generar el PDF del presupuesto. Probá de nuevo.");
    }
    setDescargandoId(null);
  };

  const addPlaya = (data) => addItem(COLLECTIONS.playa, { estado: "En playa", ...data });
  const deletePlaya = (id) => deleteItem(COLLECTIONS.playa, id);

  const derivarPlaya = (item, destinoValue, extra) => {
    const destino = DESTINOS_PLAYA.find((d) => d.value === destinoValue);
    if (!destino) return;
    const codigo = nextCodigoParaModelo(equipos, item.descripcion);
    addItem(COLLECTIONS.equipos, {
      codigo, serie: "", modelo: item.descripcion,
      fechaIngreso: item.fecha, estado: destino.estado,
      ubicacion: "Depósito principal",
      cantidad: (extra && extra.cantidad) || item.cantidad || 1,
      notas: (extra && extra.notas) || item.notas || "",
    });
    addItem(COLLECTIONS.entradas, {
      fecha: todayISO(), codigo,
      tipo: "Retorno de reparación propia", origen: `Zona de playa (${item.origen})`,
      motivo: `Derivado desde playa: ${item.descripcion}`,
      estadoResultante: destino.estado,
      responsable: "",
    });
    deletePlaya(item.id);
  };

  // Extraer una pieza de un equipo en playa destinado a desarmadero: suma stock al repuesto
  // del catálogo, pero el equipo de playa NO se borra ni se deriva — sigue ahí, solo con una
  // línea de historial en sus notas de qué se le sacó y cuándo (puede pasar varias veces).
  const extraerRepuestoDePlaya = (item, productoId, cantidad) => {
    const producto = productos.find((p) => p.id === productoId);
    if (!producto) return;
    const cant = Number(cantidad) || 1;
    updateItem(COLLECTIONS.productos, producto.id, { stockDisponible: (Number(producto.stockDisponible) || 0) + cant });
    const linea = `[${fmtDate(todayISO())}] Repuesto extraído: ${producto.descripcion || producto.nombre} (${producto.nombre}) x${cant}`;
    updateItem(COLLECTIONS.playa, item.id, { notas: item.notas ? `${item.notas}\n${linea}` : linea });
  };

  const handleImportarCatalogo = async (file) => {
    setImportandoCatalogo(true);
    setImportResultado("");
    try {
      const buffer = await file.arrayBuffer();
      const { productos: nuevos, errores } = parseCatalogoExcel(buffer);
      let ok = 0;
      const fallidos = [];
      for (const p of nuevos) {
        try {
          await addDoc(collection(db, COLLECTIONS.productos), { ...p, createdAt: Date.now() });
          ok++;
        } catch (e) {
          console.error("Firestore add error", COLLECTIONS.productos, e);
          fallidos.push(p.nombre);
        }
      }
      let msg = `${ok} de ${nuevos.length} producto(s) importado(s).`;
      if (fallidos.length) msg += ` Fallaron: ${fallidos.join(", ")} — probá de nuevo con esos.`;
      if (errores.length) msg += ` ${errores.length} fila(s) omitida(s) por no tener nombre/código.`;
      setImportResultado(msg);
    } catch (e) {
      console.error("Error importando catálogo", e);
      setImportResultado("No se pudo leer el archivo. Verificá que sea un .xlsx válido.");
    }
    setImportandoCatalogo(false);
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const wsEquipos = XLSX.utils.json_to_sheet(equipos.map((e) => ({
      "Código interno": e.codigo, "N° de serie": e.serie, "Modelo": e.modelo,
      "Fecha ingreso": e.fechaIngreso, "Estado actual": e.estado, "Ubicación": e.ubicacion,
      "Cantidad": e.cantidad || 1, "Comprometido": e.comprometido || 0,
      "Sin marca": e.sinMarca ? "Sí" : "No",
      "Motivo de baja": e.motivoBaja || "", "Comentario": e.notas || "",
    })));
    const wsMov = XLSX.utils.json_to_sheet(movimientos.map((m) => ({
      "Fecha": m.fecha, "Categoría de origen": m.categoriaLabel, "Código": m.codigo,
      "Modelo": m.modelo, "Cantidad": m.cantidad, "Motivo": m.motivo, "Cliente": m.cliente,
      "Obra": m.obra, "Monto U$S": m.monto,
      "N° remito": m.remito, "Responsable": m.responsable,
      "Lugar de salida": m.lugarSalida, "Empresa": m.empresaCliente, "RUC": m.rucCliente,
      "Firma - Aclaración": m.firmaNombre, "Firma - C.I.": m.firmaCedula,
      "Foto adjunta": m.fotoRemito ? "Sí" : "No",
      "Observaciones": m.observaciones,
    })));
    const wsEnt = XLSX.utils.json_to_sheet(entradas.map((e) => ({
      "Fecha": e.fecha, "Código": e.codigo, "Tipo de entrada": e.tipo, "Origen": e.origen,
      "Motivo": e.motivo, "Estado resultante": e.estadoResultante, "Responsable": e.responsable,
    })));
    const wsVentas = XLSX.utils.json_to_sheet(ventas.map((v) => ({
      "Código": v.codigo, "Cliente": v.cliente, "Obra": v.obra, "Fecha venta": v.fechaVenta,
      "Vto. service 1 (12m)": v.vtoService1, "Vto. service 2 (24m)": v.vtoService2,
      "Estado service 1": v.estadoService1, "Estado service 2": v.estadoService2,
      "Contactado service 1": v.contactadoService1 ? "Sí" : "No", "Fecha contacto service 1": v.fechaContactoService1 || "",
      "Decisión service 1": v.decisionService1 || "", "Cita service 1": v.citaService1 ? `${fmtDate(v.citaService1.fecha)} ${v.citaService1.hora || ""}` : "",
      "Contactado service 2": v.contactadoService2 ? "Sí" : "No", "Fecha contacto service 2": v.fechaContactoService2 || "",
      "Decisión service 2": v.decisionService2 || "", "Cita service 2": v.citaService2 ? `${fmtDate(v.citaService2.fecha)} ${v.citaService2.hora || ""}` : "",
    })));
    const wsRepuestos = XLSX.utils.json_to_sheet(productos.filter((p) => p.categoriaPrincipal === "Repuestos").map((p) => ({
      "Código": p.nombre, "Descripción": p.descripcion, "Tipo de equipo": p.subcategoria, "Código de equipo": p.subcategoria2,
      "Precio de lista U$S": p.precioLista, "Costo puesto en PY U$S": p.costoPy, "Stock disponible": p.stockDisponible,
    })));
    const wsPlaya = XLSX.utils.json_to_sheet(playa.map((p) => ({
      "Fecha": p.fecha, "Descripción": p.descripcion, "Origen": p.origen, "Cantidad": p.cantidad || 1, "Notas": p.notas,
    })));
    const wsComprometidas = XLSX.utils.json_to_sheet(comprometidas.map((c) => ({
      "Fecha": c.fecha, "Razón social": c.razonSocial, "Obra": c.obra, "Producto": c.modelo,
      "Cantidad": c.cantidad, "Monto U$S": c.monto, "Fecha entrega estimada": c.fechaEntrega,
      "Estado": c.estado,
    })));
    XLSX.utils.book_append_sheet(wb, wsEquipos, "Maestro de Equipos");
    XLSX.utils.book_append_sheet(wb, wsMov, "Movimientos");
    XLSX.utils.book_append_sheet(wb, wsEnt, "Entradas");
    XLSX.utils.book_append_sheet(wb, wsVentas, "Ventas y Garantias");
    XLSX.utils.book_append_sheet(wb, wsComprometidas, "Ventas Comprometidas");
    XLSX.utils.book_append_sheet(wb, wsRepuestos, "Repuestos");
    XLSX.utils.book_append_sheet(wb, wsPlaya, "Zona de Playa");
    XLSX.writeFile(wb, `Aeon_Stock_${todayISO()}.xlsx`);
  };

  const filteredEquipos = useMemo(() => {
    const q = query.toLowerCase();
    return equipos.filter((e) => !q || [e.codigo, e.serie, e.modelo, e.estado].some((v) => (v || "").toLowerCase().includes(q)));
  }, [equipos, query]);

  const filteredMovimientos = useMemo(() => {
    const q = query.toLowerCase();
    return movimientos.filter((m) => !q || [m.codigo, m.cliente, m.motivo, m.categoriaLabel, m.modelo, m.responsable].some((v) => (v || "").toLowerCase().includes(q)));
  }, [movimientos, query]);

  const filteredEntradas = useMemo(() => {
    const q = query.toLowerCase();
    return entradas.filter((e) => !q || [e.codigo, e.origen, e.tipo].some((v) => (v || "").toLowerCase().includes(q)));
  }, [entradas, query]);

  const filteredVentas = useMemo(() => {
    const q = query.toLowerCase();
    return ventas.filter((v) => !q || [v.codigo, v.cliente, v.obra].some((v2) => (v2 || "").toLowerCase().includes(q)));
  }, [ventas, query]);

  const filteredPlaya = useMemo(() => {
    const q = query.toLowerCase();
    return playa.filter((p) => !q || [p.descripcion, p.origen].some((v) => (v || "").toLowerCase().includes(q)));
  }, [playa, query]);

  const filteredProductos = useMemo(() => {
    const q = query.toLowerCase();
    return productos.filter((p) => !q || [p.nombre, p.categoria].some((v) => (v || "").toLowerCase().includes(q)));
  }, [productos, query]);

  const filteredCotizaciones = useMemo(() => {
    const q = query.toLowerCase();
    return cotizaciones.filter((c) => !q || [c.cliente, c.obra, c.categoria].some((v) => (v || "").toLowerCase().includes(q)));
  }, [cotizaciones, query]);

  const filteredPresupuestosReparacion = useMemo(() => {
    const q = query.toLowerCase();
    return presupuestosReparacion.filter((p) => !q || [p.cliente, p.obra, p.equipoAfectado].some((v) => (v || "").toLowerCase().includes(q)));
  }, [presupuestosReparacion, query]);

  const filteredComprometidas = useMemo(() => {
    const q = query.toLowerCase();
    return comprometidas.filter((c) => !q || [c.razonSocial, c.obra, c.modelo].some((v) => (v || "").toLowerCase().includes(q)));
  }, [comprometidas, query]);

  const recuperables = useMemo(() => equipos.filter((e) => RECUPERABLE_ESTADOS.includes(e.estado)), [equipos]);
  const muestras = useMemo(() => equipos.filter((e) => e.estado === "Muestra" || e.estado === "En préstamo (muestra)"), [equipos]);

  // Ventas cerradas = salidas ya retiradas del depósito con motivo Venta (fuente única de verdad, sin duplicar Equipos totales)
  const ventasCerradas = useMemo(() => movimientos.filter((m) => m.motivo === "Venta"), [movimientos]);

  const proximosServices = useMemo(() => {
    const items = [];
    ventas.forEach((v) => {
      [
        ["Service 1 (12m)", v.vtoService1, v.estadoService1, 1, v.contactadoService1, v.fechaContactoService1, v.decisionService1],
        ["Service 2 (24m)", v.vtoService2, v.estadoService2, 2, v.contactadoService2, v.fechaContactoService2, v.decisionService2],
      ].forEach(([label, date, estado, n, contactado, fechaContacto, decision]) => {
        if (!date || estado === "Cumplido") return;
        const d = daysUntil(date);
        if (d !== null && d <= 60) items.push({ ...v, label, date, dias: d, n, contactado, fechaContacto, decision });
      });
    });
    return items.sort((a, b) => a.dias - b.dias);
  }, [ventas]);

  const alertasContacto = useMemo(
    () => proximosServices.filter((s) => s.dias <= 40 && !s.contactado),
    [proximosServices]
  );

  const seguimientosPendientes = useMemo(() => {
    return proximosServices.filter((s) => {
      if (!s.contactado || s.decision) return false;
      const proximo = addBusinessDaysISO(s.fechaContacto, 5);
      return daysUntil(proximo) <= 0;
    });
  }, [proximosServices]);

  // Navegación desde un StatCard: soporta un destino compuesto "tab:extra" (hoy solo
  // "catalogo:repuestos", para que el card de Repuestos del Resumen abra el Catálogo
  // directamente en modo Repuestos en vez del tab de Repuestos suelto, ya en desuso).
  const navigateTo = (target) => {
    if (target === "catalogo:repuestos") {
      setCatalogoModoInicial("repuestos");
      setTab("catalogo");
    } else {
      setCatalogoModoInicial(null);
      setTab(target);
    }
  };

  const NAV = [
    { key: "resumen", label: "Resumen", icon: LayoutDashboard },
    { key: "panel", label: "Panel de indicadores", icon: TrendingUp },
    { key: "playa", label: "Zona de playa", icon: Inbox },
    { key: "equipos", label: "Maestro de equipos", icon: Package },
    { key: "comprometidas", label: "Ventas comprometidas", icon: Lock },
    { key: "movimientos", label: "Salidas", icon: ArrowUpFromLine },
    { key: "entradas", label: "Entradas", icon: ArrowDownToLine },
    { key: "ventas", label: "Ventas y garantías", icon: ShieldCheck },
    { key: "recuperables", label: "Banco de recuperables", icon: Wrench },
    { key: "muestras", label: "Muestras", icon: Star },
    { key: "catalogo", label: "Catálogo de productos", icon: Tag },
    { key: "cotizaciones", label: "Cotizaciones", icon: FileSignature },
    { key: "presupuestos-reparacion", label: "Presupuestos de reparación", icon: Hammer },
  ];

  if (loading) {
    return (
      <div className="flex items-center justify-center py-24" style={{ backgroundColor: BG }}>
        <p className="text-sm" style={{ color: MUTED }}>Cargando datos...</p>
      </div>
    );
  }

  return (
    <div style={{ backgroundColor: BG, minHeight: "100vh", fontFamily: "system-ui, -apple-system, sans-serif" }}>
      {/* Mobile top bar */}
      <div className="md:hidden flex items-center gap-3 px-4 py-3 border-b sticky top-0 z-20" style={{ borderColor: BORDER, backgroundColor: "#FFFFFF" }}>
        <button onClick={() => setNavOpen(true)} className="p-1 -ml-1 rounded hover:bg-gray-100">
          <Menu size={20} style={{ color: INK }} />
        </button>
        <img src={`${import.meta.env.BASE_URL}aeon-logo.jpg`} alt="AEON" className="h-8 w-auto" />
      </div>

      <div className="flex w-full">
        {/* Backdrop (mobile only, while nav is open) */}
        {navOpen && (
          <div className="fixed inset-0 z-30 md:hidden" style={{ backgroundColor: "rgba(15,23,32,0.4)" }} onClick={() => setNavOpen(false)} />
        )}

        {/* Sidebar */}
        <div
          className={`fixed md:static inset-y-0 left-0 z-40 w-56 shrink-0 border-r flex flex-col transition-transform duration-200 md:translate-x-0 ${navOpen ? "translate-x-0" : "-translate-x-full"}`}
          style={{ borderColor: BORDER, backgroundColor: "#FFFFFF" }}
        >
          <div className="px-4 py-4 border-b" style={{ borderColor: BORDER }}>
            <img src={`${import.meta.env.BASE_URL}aeon-logo.jpg`} alt="AEON" className="h-20 w-auto" />
          </div>
          <nav className="flex-1 py-2 overflow-y-auto">
            {NAV.map((n) => {
              const Icon = n.icon;
              const active = tab === n.key;
              return (
                <button
                  key={n.key}
                  onClick={() => { setTab(n.key); setQuery(""); setNavOpen(false); setCatalogoModoInicial(null); }}
                  className="w-full flex items-center gap-2.5 px-4 py-2.5 text-sm text-left"
                  style={{
                    color: active ? ACCENT : INK,
                    backgroundColor: active ? ACCENT_LIGHT : "transparent",
                    fontWeight: active ? 600 : 400,
                    borderRight: active ? `2px solid ${ACCENT}` : "2px solid transparent",
                  }}
                >
                  <Icon size={16} />
                  {n.label}
                </button>
              );
            })}
          </nav>
          <div className="px-4 py-3 border-t" style={{ borderColor: BORDER }}>
            <SecondaryButton onClick={exportExcel}><Download size={14} /> Exportar Excel</SecondaryButton>
          </div>
        </div>

      {/* Main */}
      <div className="flex-1 min-w-0 p-4 md:p-6">
        {tab === "resumen" && (
          <Resumen
            equipos={equipos} proximosServices={proximosServices} alertasContacto={alertasContacto}
            seguimientosPendientes={seguimientosPendientes} recuperables={recuperables}
            playa={playa} muestras={muestras} productos={productos} ventasCerradas={ventasCerradas}
            onNavigate={navigateTo}
          />
        )}

        {tab === "panel" && (
          <PanelView
            ventasCerradas={ventasCerradas} cotizaciones={cotizaciones} comprometidas={comprometidas}
            presupuestosReparacion={presupuestosReparacion}
          />
        )}

        {tab === "playa" && (
          <PlayaView
            playa={filteredPlaya} query={query} onQuery={setQuery}
            onNew={() => setDrawer("playa")}
            onDerivar={derivarPlaya}
            onExtraerRepuesto={extraerRepuestoDePlaya}
            productosRepuestos={productos.filter((p) => p.categoriaPrincipal === "Repuestos")}
            onDelete={deletePlaya}
          />
        )}

        {tab === "equipos" && (
          <Section
            title="Maestro de equipos"
            subtitle="Un registro por cada unidad, con su estado actual."
            query={query} onQuery={setQuery}
            onNew={() => setDrawer("equipo")}
            newLabel="Nuevo equipo"
            extraActions={
              <SecondaryButton onClick={() => setDrawer("escaneo")}>
                <ScanLine size={14} /> Cargar por escaneo
              </SecondaryButton>
            }
          >
            <Table
              columns={[
                { key: "codigo", label: "Código" }, { key: "serie", label: "N° de serie" },
                { key: "modelo", label: "Modelo" }, { key: "fechaIngreso", label: "Ingreso" },
                { key: "estado", label: "Estado" }, { key: "cantidad", label: "Cant." },
                { key: "comprometido", label: "Comprometido" }, { key: "motivoBaja", label: "Motivo de baja" },
                { key: "ubicacion", label: "Ubicación" },
              ]}
              rows={filteredEquipos}
              onDelete={deleteEquipo}
              renderCell={(key, row) => {
                if (key === "codigo") return <CodeTag>{row.codigo}</CodeTag>;
                if (key === "fechaIngreso") return fmtDate(row.fechaIngreso);
                if (key === "estado") return <StatusBadge estado={row.estado} />;
                if (key === "cantidad") return row.cantidad || 1;
                if (key === "comprometido") return row.comprometido > 0 ? (
                  <span style={{ color: "#B45309", fontWeight: 500 }}>{row.comprometido}</span>
                ) : "—";
                if (key === "motivoBaja") return row.estado === "Dado de baja" ? (row.motivoBaja || "Sin especificar") : "—";
                return row[key] || "—";
              }}
            />
          </Section>
        )}

        {tab === "comprometidas" && (
          <ComprometidasView
            comprometidas={filteredComprometidas} query={query} onQuery={setQuery}
            onNew={() => setDrawer("comprometida")}
            onCancelar={cancelarComprometida}
            onRetirar={(id) => setRetiroTarget(id)}
            onCerrar={cerrarComprometida}
            onPago={(id) => setPagoTarget(id)}
          />
        )}

        {tab === "movimientos" && (
          <Section
            title="Salidas de depósito"
            subtitle="Elegís de qué categoría sale el producto — la cantidad disponible se descuenta ahí mismo."
            query={query} onQuery={setQuery}
            onNew={() => setDrawer("movimiento")}
            newLabel="Nueva salida"
            extraActions={
              <SecondaryButton onClick={() => setDrawer("salida-cotizacion")}>
                <PackageCheck size={14} /> Generar desde cotización
              </SecondaryButton>
            }
          >
            <Table
              columns={[
                { key: "fecha", label: "Fecha" }, { key: "categoriaLabel", label: "Categoría de origen" },
                { key: "codigo", label: "Código" }, { key: "cantidad", label: "Cant." },
                { key: "motivo", label: "Motivo" }, { key: "cliente", label: "Cliente" },
                { key: "remito", label: "N° remito" }, { key: "responsable", label: "Responsable" },
              ]}
              rows={filteredMovimientos}
              onDelete={deleteMovimiento}
              renderCell={(key, row) => {
                if (key === "codigo") return <CodeTag>{row.codigo}</CodeTag>;
                if (key === "fecha") return fmtDate(row.fecha);
                if (key === "motivo") return row.motivo || "—";
                if (key === "remito") return (
                  <div className="flex items-center gap-1.5">
                    <span>{row.remito || "—"}</span>
                    {row.fotoRemito && (
                      <button onClick={() => setFotoView(row.fotoRemito)} title="Ver foto del remito">
                        <Camera size={14} style={{ color: ACCENT }} />
                      </button>
                    )}
                  </div>
                );
                return row[key] || "—";
              }}
            />
          </Section>
        )}

        {tab === "entradas" && (
          <Section
            title="Registro de entradas"
            subtitle="Todo lo que vuelve o ingresa al depósito, más allá de la importación inicial."
            query={query} onQuery={setQuery}
            onNew={() => setDrawer("entrada")}
            newLabel="Nueva entrada"
          >
            <Table
              columns={[
                { key: "fecha", label: "Fecha" }, { key: "codigo", label: "Código" },
                { key: "tipo", label: "Tipo de entrada" }, { key: "origen", label: "Origen" },
                { key: "estadoResultante", label: "Estado resultante" }, { key: "responsable", label: "Responsable" },
              ]}
              rows={filteredEntradas}
              onDelete={deleteEntrada}
              renderCell={(key, row) => {
                if (key === "codigo") return <CodeTag>{row.codigo}</CodeTag>;
                if (key === "fecha") return fmtDate(row.fecha);
                if (key === "estadoResultante") return <StatusBadge estado={row.estadoResultante} />;
                return row[key] || "—";
              }}
            />
          </Section>
        )}

        {tab === "ventas" && (
          <Section
            title="Ventas y garantías"
            subtitle="Vencimientos de service para mantener vigente la garantía extendida a 36 meses."
            query={query} onQuery={setQuery}
            onNew={() => setDrawer("venta")}
            newLabel="Nueva venta"
          >
            <Table
              columns={[
                { key: "codigo", label: "Código" }, { key: "cliente", label: "Cliente" }, { key: "obra", label: "Obra" },
                { key: "fechaVenta", label: "Fecha venta" }, { key: "vtoService1", label: "Service 1 (12m)" },
                { key: "vtoService2", label: "Service 2 (24m)" },
              ]}
              rows={filteredVentas}
              onDelete={deleteVenta}
              renderCell={(key, row) => {
                if (key === "codigo") return <CodeTag>{row.codigo}</CodeTag>;
                if (key === "fechaVenta") return fmtDate(row.fechaVenta);
                if (key === "vtoService1") return <ServiceCell venta={row} field="Service1" label={fmtDate(row.vtoService1)} onUpdate={updateVentaEstado} onGestionar={(v, f) => setGestion({ venta: v, field: f })} />;
                if (key === "vtoService2") return <ServiceCell venta={row} field="Service2" label={fmtDate(row.vtoService2)} onUpdate={updateVentaEstado} onGestionar={(v, f) => setGestion({ venta: v, field: f })} />;
                return row[key] || "—";
              }}
            />
          </Section>
        )}

        {tab === "recuperables" && (
          <RecuperablesView recuperables={recuperables} query={query} onQuery={setQuery} onUpdateEstado={updateEquipoEstado} onUpdateField={updateEquipoField} />
        )}

        {tab === "muestras" && (
          <MuestrasView muestras={muestras} query={query} onQuery={setQuery} onUpdateField={updateEquipoField} />
        )}

        {tab === "catalogo" && (
          <CatalogoView
            productos={filteredProductos} query={query} onQuery={setQuery}
            modoInicial={catalogoModoInicial}
            onNew={(modo) => { setProductoEditando(null); setNuevoProductoDefaults(modo === "repuestos" ? { categoriaPrincipal: "Repuestos" } : null); setDrawer("producto"); }}
            onEdit={(p) => { setProductoEditando(p); setDrawer("producto"); }}
            onDelete={deleteProducto}
            onQuitarFicha={quitarFichaTecnica}
            onImportar={handleImportarCatalogo}
            importando={importandoCatalogo}
            importResultado={importResultado}
          />
        )}

        {tab === "cotizaciones" && (
          <CotizacionesView
            cotizaciones={filteredCotizaciones} query={query} onQuery={setQuery}
            onNew={() => setDrawer("cotizacion")}
            onDelete={deleteCotizacion}
            onUpdate={updateCotizacion}
            onDescargarPdf={handleDescargarPdf}
            onDescargarFichas={handleDescargarFichas}
            descargandoId={descargandoId}
            pdfError={pdfError}
          />
        )}

        {tab === "presupuestos-reparacion" && (
          <PresupuestosReparacionView
            presupuestos={filteredPresupuestosReparacion} query={query} onQuery={setQuery}
            onNew={() => setDrawer("presupuesto-reparacion")}
            onDelete={deletePresupuestoReparacion}
            onDescargarPdf={handleDescargarPresupuestoPdf}
            descargandoId={descargandoId}
            pdfError={pdfError}
          />
        )}
        </div>
      </div>

      {/* Drawers */}
      <Drawer open={drawer === "equipo"} onClose={() => setDrawer(null)} title="Nuevo equipo">
        <EquipoForm equipos={equipos} onSave={(d) => { addEquipo(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "escaneo"} onClose={() => setDrawer(null)} title="Cargar por escaneo">
        <EscaneoUnidadesForm onSave={(d) => { addEquiposPorEscaneo(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "movimiento"} onClose={() => setDrawer(null)} title="Nueva salida">
        <MovimientoForm equipos={equipos} playa={playa} productos={productos} onSave={(d) => { addMovimiento(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "entrada"} onClose={() => setDrawer(null)} title="Nueva entrada">
        <EntradaForm equipos={equipos} onSave={(d) => { addEntrada(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "venta"} onClose={() => setDrawer(null)} title="Nueva venta">
        <VentaForm equipos={equipos} onSave={(d) => { addVenta(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "comprometida"} onClose={() => setDrawer(null)} title="Nueva venta comprometida">
        <ComprometidaForm equipos={equipos} productos={productos} onSave={(d) => { addComprometida(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "playa"} onClose={() => setDrawer(null)} title="Nuevo ingreso a playa">
        <PlayaForm onSave={(d) => { addPlaya(d); setDrawer(null); }} />
      </Drawer>
      <Drawer
        open={drawer === "producto"} onClose={() => { setDrawer(null); setNuevoProductoDefaults(null); }}
        title={productoEditando ? "Editar producto" : nuevoProductoDefaults?.categoriaPrincipal === "Repuestos" ? "Nuevo repuesto" : "Nuevo producto"}
      >
        <ProductoForm
          producto={productoEditando}
          defaults={nuevoProductoDefaults}
          onSave={(d) => {
            if (productoEditando) updateProducto(productoEditando.id, d);
            else addProducto(d);
            setDrawer(null);
            setProductoEditando(null);
            setNuevoProductoDefaults(null);
          }}
        />
      </Drawer>
      <Drawer open={drawer === "cotizacion"} onClose={() => setDrawer(null)} title="Nueva cotización">
        <CotizacionForm productos={productos} onSave={(d) => { addCotizacion(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "presupuesto-reparacion"} onClose={() => setDrawer(null)} title="Nuevo presupuesto de reparación">
        <PresupuestoReparacionForm
          productos={productos.filter((p) => p.categoriaPrincipal === "Repuestos")}
          onSave={(d) => { addPresupuestoReparacion(d); setDrawer(null); }}
        />
      </Drawer>
      <Drawer open={drawer === "salida-cotizacion"} onClose={() => setDrawer(null)} title="Generar salida desde cotización">
        <SalidaDesdeCotizacionForm
          cotizaciones={cotizaciones}
          equipos={equipos}
          onGenerar={(cotizacion, lineas, remitoData) => { generarSalidaDesdeCotizacion(cotizacion, lineas, remitoData); setDrawer(null); }}
        />
      </Drawer>
      <Drawer open={!!gestion} onClose={() => setGestion(null)} title={gestion ? `Seguimiento — ${gestion.field === "Service1" ? "Service 1 (12m)" : "Service 2 (24m)"}` : ""}>
        {gestion && (
          <GestionServiceForm
            venta={gestion.venta} field={gestion.field}
            onUpdate={updateVentaEstado}
            onClose={() => setGestion(null)}
          />
        )}
      </Drawer>
      <Drawer open={!!retiroTarget} onClose={() => setRetiroTarget(null)} title="Registrar retiro">
        {retiroTarget && (
          <RetiroParcialForm
            comprometida={comprometidas.find((c) => c.id === retiroTarget)}
            onSave={(d) => { retirarParcial(retiroTarget, d); setRetiroTarget(null); }}
          />
        )}
      </Drawer>
      <Drawer open={!!pagoTarget} onClose={() => setPagoTarget(null)} title="Pagos">
        {pagoTarget && (
          <PagosForm
            comprometida={comprometidas.find((c) => c.id === pagoTarget)}
            onAgregar={(d) => agregarPago(pagoTarget, d)}
            onQuitar={(idx) => quitarPago(pagoTarget, idx)}
          />
        )}
      </Drawer>
      <PhotoViewer src={fotoView} onClose={() => setFotoView(null)} />
      <ConfirmDialog state={confirmState} onResolve={resolveConfirm} />
    </div>
  );
}

function ServiceCell({ venta, field, label, onUpdate, onGestionar }) {
  const estadoKey = `estado${field}`;
  const estado = venta[estadoKey] || "Pendiente";
  const d = daysUntil(field === "Service1" ? venta.vtoService1 : venta.vtoService2);
  const soon = estado !== "Cumplido" && d !== null && d <= 30;
  const contactado = venta[`contactado${field}`];
  return (
    <div className="flex items-center gap-1.5 flex-wrap">
      <span style={{ color: soon ? "#B45309" : INK }}>{label}</span>
      {soon && <AlertTriangle size={13} style={{ color: "#B45309" }} />}
      <select
        value={estado}
        onChange={(e) => onUpdate(venta.id, estadoKey, e.target.value)}
        className="text-xs rounded px-1 py-0.5 border"
        style={{ borderColor: BORDER, color: MUTED }}
      >
        <option>Pendiente</option>
        <option>Cumplido</option>
        <option>Vencido</option>
      </select>
      {estado !== "Cumplido" && (
        <button
          onClick={() => onGestionar(venta, field)}
          className="text-xs px-1.5 py-0.5 rounded"
          style={{ backgroundColor: contactado ? "#E9F7EF" : ACCENT_LIGHT, color: contactado ? "#15803D" : ACCENT }}
        >
          {contactado ? "Seguimiento" : "Gestionar"}
        </button>
      )}
    </div>
  );
}

function Section({ title, subtitle, query, onQuery, onNew, newLabel, extraActions, children }) {
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>{title}</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>{subtitle}</p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
          {extraActions}
          <PrimaryButton onClick={onNew}><Plus size={15} /> {newLabel}</PrimaryButton>
        </div>
      </div>
      {children}
    </div>
  );
}

function StatCard({ label, value, icon: Icon, tint, onClick }) {
  const Wrapper = onClick ? "button" : "div";
  return (
    <Wrapper
      onClick={onClick}
      className="rounded-xl p-4 flex items-center gap-3 text-left w-full"
      style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}`, cursor: onClick ? "pointer" : "default" }}
    >
      <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ backgroundColor: tint || ACCENT_LIGHT }}>
        <Icon size={17} style={{ color: ACCENT }} />
      </div>
      <div>
        <p className="text-xl font-semibold" style={{ color: INK }}>{value}</p>
        <p className="text-xs" style={{ color: MUTED }}>{label}</p>
      </div>
    </Wrapper>
  );
}

function sumCantidad(list) {
  return list.reduce((acc, e) => acc + (Number(e.cantidad) || 1), 0);
}

function VentasCerradasPanel({ ventasCerradas }) {
  const [q, setQ] = useState("");
  const filtered = ventasCerradas.filter((v) => {
    const s = q.toLowerCase();
    return !s || [v.cliente, v.obra, v.modelo].some((x) => (x || "").toLowerCase().includes(s));
  });
  const totalCantidad = filtered.reduce((acc, v) => acc + (Number(v.cantidad) || 0), 0);
  const totalMonto = filtered.reduce((acc, v) => acc + (Number(v.monto) || 0), 0);

  const porProducto = useMemo(() => {
    const map = {};
    filtered.forEach((v) => {
      const key = v.modelo || "—";
      if (!map[key]) map[key] = { modelo: key, cantidad: 0, monto: 0 };
      map[key].cantidad += Number(v.cantidad) || 0;
      map[key].monto += Number(v.monto) || 0;
    });
    return Object.values(map).sort((a, b) => b.monto - a.monto);
  }, [filtered]);

  return (
    <div className="rounded-xl p-4 mt-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <div className="flex items-center justify-between mb-3 gap-3 flex-wrap">
        <div className="flex items-center gap-2">
          <TrendingUp size={15} style={{ color: ACCENT }} />
          <h3 className="text-sm font-semibold" style={{ color: INK }}>Ventas cerradas</h3>
        </div>
        <SearchBox value={q} onChange={setQ} placeholder="Buscar cliente, obra o producto..." />
      </div>

      {filtered.length === 0 ? (
        <p className="text-sm" style={{ color: MUTED }}>Todavía no hay ventas con retiro registrado.</p>
      ) : (
        <>
          <div className="overflow-x-auto rounded-lg border mb-3" style={{ borderColor: BORDER }}>
            <table className="w-full text-sm">
              <thead>
                <tr style={{ backgroundColor: "#FAFBFC" }}>
                  {["Fecha", "Razón social", "Obra", "Producto", "Cantidad", "Monto U$S"].map((h) => (
                    <th key={h} className="text-left font-medium px-3 py-2 border-b" style={{ color: MUTED, borderColor: BORDER, fontSize: 12 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.slice(0, 10).map((v, i) => (
                  <tr key={i} className="border-b last:border-0" style={{ borderColor: BORDER }}>
                    <td className="px-3 py-2" style={{ color: INK }}>{fmtDate(v.fecha)}</td>
                    <td className="px-3 py-2" style={{ color: INK }}>{v.cliente || "—"}</td>
                    <td className="px-3 py-2" style={{ color: INK }}>{v.obra || "—"}</td>
                    <td className="px-3 py-2" style={{ color: INK }}>{v.modelo}</td>
                    <td className="px-3 py-2" style={{ color: INK }}>{v.cantidad}</td>
                    <td className="px-3 py-2" style={{ color: INK }}>{Number(v.monto || 0).toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{ backgroundColor: ACCENT_LIGHT }}>
                  <td colSpan={4} className="px-3 py-2 text-sm font-medium" style={{ color: ACCENT }}>Total</td>
                  <td className="px-3 py-2 text-sm font-medium" style={{ color: ACCENT }}>{totalCantidad}</td>
                  <td className="px-3 py-2 text-sm font-medium" style={{ color: ACCENT }}>{totalMonto.toLocaleString()}</td>
                </tr>
              </tfoot>
            </table>
          </div>

          <p className="text-xs font-medium mb-1.5" style={{ color: MUTED }}>Cerrado por producto (todas las obras)</p>
          <div className="overflow-x-auto rounded-lg border" style={{ borderColor: BORDER }}>
            <table className="w-full text-sm">
              <thead>
                <tr style={{ backgroundColor: "#FAFBFC" }}>
                  {["Producto", "Cantidad total", "Monto total U$S"].map((h) => (
                    <th key={h} className="text-left font-medium px-3 py-2 border-b" style={{ color: MUTED, borderColor: BORDER, fontSize: 12 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {porProducto.map((p) => (
                  <tr key={p.modelo} className="border-b last:border-0" style={{ borderColor: BORDER }}>
                    <td className="px-3 py-2" style={{ color: INK }}>{p.modelo}</td>
                    <td className="px-3 py-2" style={{ color: INK }}>{p.cantidad}</td>
                    <td className="px-3 py-2" style={{ color: INK }}>{p.monto.toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}

function Resumen({ equipos, proximosServices, alertasContacto, seguimientosPendientes, recuperables, playa, muestras, productos, ventasCerradas, onNavigate }) {
  const vendible = equipos.filter((e) => e.estado === "En depósito" || e.estado === "Apto para venta");
  const bajas = equipos.filter((e) => e.estado === "Dado de baja");
  const totalUnidades = sumCantidad(equipos.filter((e) => e.estado !== "Dado de baja"));
  // El stock de repuestos vive en el Catálogo de productos (categoriaPrincipal="Repuestos"),
  // no en la colección "repuestos" suelta — ese origen quedó en desuso.
  const totalRepuestos = productos
    .filter((p) => p.categoriaPrincipal === "Repuestos")
    .reduce((acc, p) => acc + (Number(p.stockDisponible) || 0), 0);
  const totalPlaya = sumCantidad(playa);
  const totalVendido = ventasCerradas.reduce((acc, v) => acc + (Number(v.cantidad) || 0), 0);

  const cardsActivos = [
    { label: "Equipos totales activos", value: totalUnidades, icon: Package, tab: "equipos" },
    { label: "Zona de playa (sin clasificar)", value: totalPlaya, icon: Inbox, tab: "playa" },
    { label: "Stock vendible", value: sumCantidad(vendible), icon: ArrowDownToLine, tab: "equipos" },
    { label: "Banco de recuperables", value: sumCantidad(recuperables), icon: Wrench, tab: "recuperables" },
    { label: "Muestras", value: sumCantidad(muestras), icon: Star, tab: "muestras" },
    { label: "Repuestos (unidades)", value: totalRepuestos, icon: Boxes, tab: "catalogo:repuestos" },
  ];
  const cardsHistorico = [
    { label: "Vendidos y retirados", value: totalVendido, icon: CheckCircle2, tab: "movimientos" },
    { label: "Dados de baja", value: sumCantidad(bajas), icon: AlertTriangle, tab: "equipos" },
  ];

  return (
    <div>
      <h2 className="text-lg font-semibold mb-1" style={{ color: INK }}>Resumen</h2>
      <p className="text-sm mb-4" style={{ color: MUTED }}>
        "Equipos totales activos" es el stock que todavía cuenta como inventario. Vendidos y retirados, y Dados de baja, ya salieron del circuito — quedan abajo como historial, aparte.
      </p>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
        {cardsActivos.map((c) => (
          <StatCard key={c.label} label={c.label} value={c.value} icon={c.icon} onClick={() => onNavigate(c.tab)} />
        ))}
      </div>

      <p className="text-xs font-medium mb-2" style={{ color: MUTED }}>Historial (ya no cuenta como stock)</p>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-6">
        {cardsHistorico.map((c) => (
          <div key={c.label} style={{ opacity: 0.75 }}>
            <StatCard label={c.label} value={c.value} icon={c.icon} onClick={() => onNavigate(c.tab)} tint="#F1F5F9" />
          </div>
        ))}
        <div />
      </div>

      {(alertasContacto.length > 0 || seguimientosPendientes.length > 0) && (
        <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#FEF3E2", border: "0.5px solid #F5D9A8" }}>
          <div className="flex items-center gap-2 mb-2">
            <AlertTriangle size={15} style={{ color: "#B45309" }} />
            <h3 className="text-sm font-semibold" style={{ color: "#92400E" }}>Garantías que necesitan tu atención</h3>
          </div>
          {alertasContacto.length > 0 && (
            <div className="mb-2">
              <p className="text-xs font-medium mb-1" style={{ color: "#92400E" }}>A 40 días o menos — contactar al cliente:</p>
              {alertasContacto.map((s, i) => (
                <p key={i} className="text-sm" style={{ color: "#92400E" }}>
                  · {s.cliente} — {s.obra} — {s.label} ({s.dias < 0 ? `vencido hace ${Math.abs(s.dias)}d` : `en ${s.dias}d`})
                </p>
              ))}
            </div>
          )}
          {seguimientosPendientes.length > 0 && (
            <div>
              <p className="text-xs font-medium mb-1" style={{ color: "#92400E" }}>Sin respuesta — reenviar recordatorio:</p>
              {seguimientosPendientes.map((s, i) => (
                <p key={i} className="text-sm" style={{ color: "#92400E" }}>
                  · {s.cliente} — {s.obra} — {s.label} (contactado el {fmtDate(s.fechaContacto)})
                </p>
              ))}
            </div>
          )}
          <p className="text-xs mt-2" style={{ color: "#92400E" }}>Gestioná el contacto y la decisión del cliente desde la pestaña "Ventas y garantías".</p>
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="rounded-xl p-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
          <div className="flex items-center gap-2 mb-3">
            <Clock size={15} style={{ color: ACCENT }} />
            <h3 className="text-sm font-semibold" style={{ color: INK }}>Services próximos a vencer</h3>
          </div>
          {proximosServices.length === 0 ? (
            <p className="text-sm" style={{ color: MUTED }}>No hay vencimientos en los próximos 60 días.</p>
          ) : (
            <div className="space-y-2">
              {proximosServices.slice(0, 6).map((s, i) => (
                <div key={i} className="flex items-center justify-between text-sm">
                  <div>
                    <span className="font-medium" style={{ color: INK }}>{s.cliente}</span>
                    <span style={{ color: MUTED }}> · {s.obra} · {s.label}</span>
                  </div>
                  <span style={{ color: s.dias <= 15 ? "#B91C1C" : "#B45309" }}>
                    {s.dias < 0 ? `vencido hace ${Math.abs(s.dias)}d` : `en ${s.dias}d`}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="rounded-xl p-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
          <div className="flex items-center gap-2 mb-3">
            <Star size={15} style={{ color: ACCENT }} />
            <h3 className="text-sm font-semibold" style={{ color: INK }}>Muestras</h3>
          </div>
          {muestras.length === 0 ? (
            <p className="text-sm" style={{ color: MUTED }}>No hay equipos clasificados como muestra.</p>
          ) : (
            <div className="space-y-2.5">
              {muestras.slice(0, 6).map((e) => (
                <div key={e.id} className="text-sm">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <CodeTag>{e.codigo}</CodeTag>
                      <span style={{ color: INK }}>{e.modelo}</span>
                      <span style={{ color: MUTED }}>· cant. {e.cantidad || 1}</span>
                    </div>
                    <StatusBadge estado={e.estado} />
                  </div>
                  {e.notas && <p className="text-xs mt-0.5" style={{ color: MUTED }}>{e.notas}</p>}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <VentasCerradasPanel ventasCerradas={ventasCerradas} />
    </div>
  );
}

// ---------- Panel de indicadores ----------
const MESES_LABEL = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

function mesPrefijo(fecha) {
  return fecha ? fecha.slice(0, 7) : ""; // "YYYY-MM"
}
function sumarMonto(fechaDesdeIncl, fechaHastaExcl, lista) {
  return lista.reduce((acc, x) => {
    if (!x.fecha || x.fecha < fechaDesdeIncl || x.fecha >= fechaHastaExcl) return acc;
    return acc + (Number(x.monto) || 0);
  }, 0);
}

function IndicadorCard({ label, value, sub, subColor }) {
  return (
    <div className="rounded-xl p-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <p className="text-xs" style={{ color: MUTED }}>{label}</p>
      <p className="text-xl font-semibold mt-1" style={{ color: INK }}>{value}</p>
      {sub && <p className="text-xs mt-0.5" style={{ color: subColor || MUTED }}>{sub}</p>}
    </div>
  );
}

function GraficoEstacionalidad({ ventasCerradas }) {
  const anios = useMemo(() => {
    const set = new Set(ventasCerradas.map((v) => Number(mesPrefijo(v.fecha).slice(0, 4))).filter(Boolean));
    if (set.size === 0) set.add(new Date().getFullYear());
    return [...set].sort((a, b) => b - a);
  }, [ventasCerradas]);
  const [anio, setAnio] = useState(anios[0]);
  const [codigo, setCodigo] = useState("");

  const codigos = useMemo(() => {
    const set = new Set(ventasCerradas.map((v) => v.codigo).filter(Boolean));
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [ventasCerradas]);

  const datos = useMemo(() => {
    const meses = Array(12).fill(0);
    for (const v of ventasCerradas) {
      if (!v.fecha) continue;
      const [y, m] = v.fecha.split("-");
      if (Number(y) !== anio) continue;
      if (codigo && v.codigo !== codigo) continue;
      meses[Number(m) - 1] += Number(v.cantidad) || 0;
    }
    return meses;
  }, [ventasCerradas, anio, codigo]);

  const max = Math.max(1, ...datos);
  const mesPico = datos.indexOf(Math.max(...datos));
  const hayVentas = datos.some((v) => v > 0);

  return (
    <div className="rounded-xl p-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <div className="flex items-start justify-between mb-3 gap-3 flex-wrap">
        <div>
          <p className="text-sm font-semibold" style={{ color: INK }}>Estacionalidad de ventas</p>
          <p className="text-xs mt-0.5" style={{ color: MUTED }}>Unidades vendidas por mes — para detectar la época fuerte de cada producto.</p>
        </div>
        <div className="flex items-center gap-2">
          <div style={{ width: 90 }}>
            <Select value={anio} onChange={(e) => setAnio(Number(e.target.value))}>
              {anios.map((a) => <option key={a} value={a}>{a}</option>)}
            </Select>
          </div>
          <div style={{ width: 180 }}>
            <Select value={codigo} onChange={(e) => setCodigo(e.target.value)}>
              <option value="">Todos los productos</option>
              {codigos.map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </div>
        </div>
      </div>

      {!hayVentas ? (
        <p className="text-sm py-6 text-center" style={{ color: MUTED }}>Sin ventas registradas para {anio}{codigo ? ` de ${codigo}` : ""}.</p>
      ) : (
        <>
          <div className="flex items-end gap-1.5" style={{ height: 140 }}>
            {datos.map((v, i) => (
              <div key={i} className="flex-1 flex flex-col items-center justify-end h-full">
                <div
                  className="w-full rounded-t"
                  style={{ height: `${Math.max(2, (v / max) * 100)}%`, backgroundColor: i === mesPico && v > 0 ? ACCENT : ACCENT_LIGHT }}
                  title={`${MESES_LABEL[i]}: ${v} unidad(es)`}
                />
              </div>
            ))}
          </div>
          <div className="flex gap-1.5 mt-1.5">
            {MESES_LABEL.map((m, i) => (
              <p key={m} className="flex-1 text-center text-[10px]" style={{ color: i === mesPico && datos[i] > 0 ? ACCENT : MUTED, fontWeight: i === mesPico ? 600 : 400 }}>{m}</p>
            ))}
          </div>
          {datos[mesPico] > 0 && (
            <p className="text-xs mt-3" style={{ color: ACCENT }}>
              Mes con más ventas: <strong>{MESES_LABEL[mesPico]}</strong> ({datos[mesPico]} unidad(es)) — buen momento para reforzar stock y campañas con anticipación.
            </p>
          )}
        </>
      )}
    </div>
  );
}

function RankingCard({ title, subtitle, rows, valueLabel }) {
  if (rows.length === 0) return null;
  const max = rows[0][1];
  return (
    <div className="rounded-xl p-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <p className="text-sm font-semibold" style={{ color: INK }}>{title}</p>
      {subtitle && <p className="text-xs mt-0.5" style={{ color: MUTED }}>{subtitle}</p>}
      <div className="space-y-2 mt-3">
        {rows.map(([label, value], i) => (
          <div key={label} className="flex items-center gap-3">
            <span className="text-xs w-4" style={{ color: MUTED }}>{i + 1}</span>
            <span className="text-xs flex-1 truncate" style={{ color: INK }}>{label}</span>
            <div className="flex-1 rounded-full overflow-hidden" style={{ backgroundColor: ACCENT_LIGHT, height: 6 }}>
              <div style={{ width: `${max > 0 ? (value / max) * 100 : 0}%`, backgroundColor: ACCENT, height: 6 }} />
            </div>
            <span className="text-xs w-20 text-right shrink-0" style={{ color: MUTED }}>{valueLabel(value)}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function GraficoTendenciaVentas({ ventasDirectas, comprometidas }) {
  const meses = useMemo(() => {
    const hoy = todayISO();
    const [y0, m0] = hoy.split("-").map(Number);
    const arr = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(y0, m0 - 1 - i, 1);
      const inicio = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-01`;
      const dSig = new Date(d.getFullYear(), d.getMonth() + 1, 1);
      const fin = `${dSig.getFullYear()}-${String(dSig.getMonth() + 1).padStart(2, "0")}-01`;
      const total = sumarMonto(inicio, fin, ventasDirectas) + sumarMonto(inicio, fin, comprometidas);
      arr.push({ label: MESES_LABEL[d.getMonth()], total });
    }
    return arr;
  }, [ventasDirectas, comprometidas]);

  const max = Math.max(1, ...meses.map((m) => m.total));
  const hayDatos = meses.some((m) => m.total > 0);

  return (
    <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <p className="text-sm font-semibold" style={{ color: INK }}>Tendencia de ventas — últimos 6 meses</p>
      <p className="text-xs mt-0.5 mb-3" style={{ color: MUTED }}>Ventas directas + ventas comprometidas, en U$S por mes.</p>
      {!hayDatos ? (
        <p className="text-sm py-6 text-center" style={{ color: MUTED }}>Todavía no hay ventas registradas.</p>
      ) : (
        <>
          <div className="flex items-end gap-2" style={{ height: 110 }}>
            {meses.map((m, i) => (
              <div key={i} className="flex-1 flex flex-col items-center justify-end h-full">
                <div
                  className="w-full rounded-t"
                  style={{ height: `${Math.max(2, (m.total / max) * 100)}%`, backgroundColor: i === meses.length - 1 ? ACCENT : ACCENT_LIGHT }}
                  title={`${m.label}: U$S ${m.total.toLocaleString()}`}
                />
              </div>
            ))}
          </div>
          <div className="flex gap-2 mt-1.5">
            {meses.map((m, i) => (
              <p key={i} className="flex-1 text-center text-[10px]" style={{ color: i === meses.length - 1 ? ACCENT : MUTED, fontWeight: i === meses.length - 1 ? 600 : 400 }}>{m.label}</p>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function ReparacionMantenimientoCard({ presupuestos }) {
  const nReparacion = presupuestos.filter((p) => p.tipo !== "mantenimiento").length;
  const nMantenimiento = presupuestos.filter((p) => p.tipo === "mantenimiento").length;
  const total = nReparacion + nMantenimiento;
  if (total === 0) return null;
  const pctRep = (nReparacion / total) * 100;
  const pctMtto = (nMantenimiento / total) * 100;
  const VERDE = "#2F7A4A";

  return (
    <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <p className="text-sm font-semibold mb-3" style={{ color: INK }}>Presupuestos de reparación — mezcla</p>
      <div className="flex rounded-full overflow-hidden" style={{ height: 10, backgroundColor: ACCENT_LIGHT }}>
        {pctRep > 0 && <div style={{ width: `${pctRep}%`, backgroundColor: ACCENT }} title={`Reparación / repuestos: ${nReparacion}`} />}
        {pctMtto > 0 && <div style={{ width: `${pctMtto}%`, backgroundColor: VERDE }} title={`Mantenimiento: ${nMantenimiento}`} />}
      </div>
      <div className="flex items-center gap-4 mt-3 flex-wrap">
        <div className="flex items-center gap-1.5">
          <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: ACCENT }} />
          <span className="text-xs" style={{ color: MUTED }}>Reparación / repuestos ({nReparacion})</span>
        </div>
        <div className="flex items-center gap-1.5">
          <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: VERDE }} />
          <span className="text-xs" style={{ color: MUTED }}>Mantenimiento ({nMantenimiento})</span>
        </div>
      </div>
    </div>
  );
}

function PanelView({ ventasCerradas, cotizaciones, comprometidas, presupuestosReparacion }) {
  const hoy = todayISO();
  const mesActual = mesPrefijo(hoy);
  const [yActual, mActual] = mesActual.split("-").map(Number);
  const mesAnteriorDate = new Date(yActual, mActual - 2, 1);
  const mesAnterior = `${mesAnteriorDate.getFullYear()}-${String(mesAnteriorDate.getMonth() + 1).padStart(2, "0")}`;
  const inicioMesActual = `${mesActual}-01`;
  const inicioMesSiguiente = new Date(yActual, mActual, 1);
  const finMesActual = `${inicioMesSiguiente.getFullYear()}-${String(inicioMesSiguiente.getMonth() + 1).padStart(2, "0")}-01`;
  const inicioMesAnterior = `${mesAnterior}-01`;

  const ventasDirectas = ventasCerradas.filter((v) => Number(v.monto) > 0);
  const ventasMes = sumarMonto(inicioMesActual, finMesActual, ventasDirectas) + sumarMonto(inicioMesActual, finMesActual, comprometidas);
  const ventasMesAnt = sumarMonto(inicioMesAnterior, inicioMesActual, ventasDirectas) + sumarMonto(inicioMesAnterior, inicioMesActual, comprometidas);
  const variacion = ventasMesAnt > 0 ? ((ventasMes - ventasMesAnt) / ventasMesAnt) * 100 : null;

  const grupos = useMemo(() => agruparCotizaciones(cotizaciones), [cotizaciones]);
  const resumenCot = useMemo(() => resumirCotizaciones(grupos), [grupos]);
  const totalGanadaPerdida = resumenCot.Ganada.n + resumenCot.Perdida.n;
  const tasaConversion = totalGanadaPerdida > 0 ? (resumenCot.Ganada.n / totalGanadaPerdida) * 100 : null;

  const hace90 = new Date(Date.parse(hoy) - 90 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
  const topProductos = useMemo(() => {
    const map = new Map();
    for (const v of ventasCerradas) {
      if (!v.fecha || v.fecha < hace90) continue;
      const key = v.codigo || v.modelo || "—";
      map.set(key, (map.get(key) || 0) + (Number(v.cantidad) || 0));
    }
    return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, 5);
  }, [ventasCerradas, hace90]);

  const topClientes = useMemo(() => {
    const map = new Map();
    for (const c of cotizaciones) {
      if (c.estado !== "Ganada") continue;
      const key = (c.cliente || "").trim() || "(Sin cliente)";
      map.set(key, (map.get(key) || 0) + calcularTotalCotizacion(c));
    }
    return [...map.entries()].filter(([, total]) => total > 0).sort((a, b) => b[1] - a[1]).slice(0, 5);
  }, [cotizaciones]);

  return (
    <div>
      <h2 className="text-lg font-semibold mb-1" style={{ color: INK }}>Panel de indicadores</h2>
      <p className="text-sm mb-4" style={{ color: MUTED }}>
        Un vistazo rápido del negocio. "Ventas del mes" suma salidas directas con monto cargado + el valor total de las ventas comprometidas del mes.
      </p>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
        <IndicadorCard
          label="Ventas del mes"
          value={`U$S ${ventasMes.toLocaleString()}`}
          sub={variacion === null ? "Sin datos del mes anterior" : `${variacion >= 0 ? "▲" : "▼"} ${Math.abs(variacion).toFixed(0)}% vs. mes anterior`}
          subColor={variacion === null ? MUTED : variacion >= 0 ? "#15803D" : "#B91C1C"}
        />
        <IndicadorCard
          label="Conversión de cotizaciones"
          value={tasaConversion === null ? "—" : `${tasaConversion.toFixed(0)}%`}
          sub={`${resumenCot.Ganada.n} ganada(s) · ${resumenCot.Perdida.n} perdida(s)`}
        />
        <IndicadorCard
          label="Cotizaciones en juego"
          value={resumenCot.Pendiente.n}
          sub={`U$S ${resumenCot.Pendiente.total.toLocaleString()} pendientes de definir`}
        />
        <IndicadorCard
          label="Producto más vendido (90 días)"
          value={topProductos[0] ? topProductos[0][0] : "—"}
          sub={topProductos[0] ? `${topProductos[0][1]} unidad(es)` : "Sin ventas registradas"}
        />
      </div>

      <GraficoTendenciaVentas ventasDirectas={ventasDirectas} comprometidas={comprometidas} />

      {(topProductos.length > 0 || topClientes.length > 0) && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-3 mb-4">
          <RankingCard
            title="Top 5 productos — últimos 90 días"
            rows={topProductos}
            valueLabel={(v) => `${v} un.`}
          />
          <RankingCard
            title="Top clientes — cotizaciones ganadas"
            rows={topClientes}
            valueLabel={(v) => `U$S ${v.toLocaleString()}`}
          />
        </div>
      )}

      <ReparacionMantenimientoCard presupuestos={presupuestosReparacion} />

      <GraficoEstacionalidad ventasCerradas={ventasCerradas} />
    </div>
  );
}

function PlayaCard({ item, productosRepuestos, onDerivar, onExtraerRepuesto, onDelete }) {
  const [destino, setDestino] = useState("");
  const [cantidad, setCantidad] = useState(item.cantidad || 1);
  const [error, setError] = useState("");

  const [extrayendo, setExtrayendo] = useState(false);
  const [repuestoId, setRepuestoId] = useState("");
  const [cantidadExtraida, setCantidadExtraida] = useState(1);

  const confirmar = () => {
    if (!destino) {
      setError("Elegí un destino primero.");
      return;
    }
    onDerivar(item, destino, { cantidad: Number(cantidad) || 1 });
  };

  const confirmarExtraccion = () => {
    if (!repuestoId) {
      setError("Elegí qué repuesto extraer.");
      return;
    }
    onExtraerRepuesto(item, repuestoId, Number(cantidadExtraida) || 1);
    setExtrayendo(false);
    setRepuestoId("");
    setCantidadExtraida(1);
    setError("");
  };

  return (
    <div className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <div className="flex items-start justify-between mb-1">
        <div>
          <p className="text-sm font-medium" style={{ color: INK }}>{item.descripcion}</p>
          <p className="text-xs" style={{ color: MUTED }}>{fmtDate(item.fecha)} · dejado por {item.origen} · cant. {item.cantidad || 1}</p>
        </div>
        <button onClick={() => onDelete(item.id)} className="p-1 rounded hover:bg-gray-100">
          <Trash2 size={14} style={{ color: MUTED }} />
        </button>
      </div>
      {item.notas && <p className="text-xs mt-1 mb-2 whitespace-pre-line" style={{ color: MUTED }}>{item.notas}</p>}

      <div className="flex items-center gap-2 mt-2">
        <Select value={destino} onChange={(e) => { setDestino(e.target.value); setError(""); }} style={{ ...inputStyle, padding: "4px 8px", fontSize: 12 }}>
          <option value="">Derivar a...</option>
          {DESTINOS_PLAYA.map((d) => <option key={d.value} value={d.value}>{d.label}</option>)}
        </Select>
        <button onClick={confirmar} className="p-1.5 rounded shrink-0" style={{ backgroundColor: ACCENT }}>
          <ArrowRight size={14} color="#FFFFFF" />
        </button>
      </div>

      {!extrayendo ? (
        <button onClick={() => setExtrayendo(true)} className="text-xs font-medium mt-2.5" style={{ color: ACCENT }}>
          + Extraer repuesto (queda en playa)
        </button>
      ) : (
        <div className="mt-2.5 pt-3 border-t space-y-2" style={{ borderColor: BORDER }}>
          <p className="text-xs" style={{ color: MUTED }}>Suma stock a un repuesto del catálogo sin sacar este equipo de playa — para desarmaderos de los que se van rescatando piezas de a poco.</p>
          <Select value={repuestoId} onChange={(e) => setRepuestoId(e.target.value)} style={{ ...inputStyle, padding: "4px 8px", fontSize: 12 }}>
            <option value="">Repuesto del catálogo...</option>
            {productosRepuestos.map((p) => (
              <option key={p.id} value={p.id}>{p.nombre}{p.descripcion ? " — " + p.descripcion : ""}</option>
            ))}
          </Select>
          <div className="flex items-center gap-2">
            <TextInput type="number" min="1" value={cantidadExtraida} onChange={(e) => setCantidadExtraida(e.target.value)} placeholder="Cantidad" />
            <button onClick={confirmarExtraccion} className="text-xs px-3 py-2 rounded shrink-0" style={{ backgroundColor: ACCENT, color: "#FFFFFF" }}>
              Confirmar
            </button>
            <button onClick={() => { setExtrayendo(false); setError(""); }} className="text-xs px-2 py-2 rounded shrink-0" style={{ color: MUTED }}>
              Cancelar
            </button>
          </div>
          {productosRepuestos.length === 0 && (
            <p className="text-xs" style={{ color: "#B45309" }}>No hay repuestos cargados todavía — creá uno en Catálogo de productos → Repuestos.</p>
          )}
        </div>
      )}

      {error && <p className="text-xs mt-2" style={{ color: "#B91C1C" }}>{error}</p>}
    </div>
  );
}

function PlayaView({ playa, productosRepuestos, query, onQuery, onNew, onDerivar, onExtraerRepuesto, onDelete }) {
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Zona de playa</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>
            Productos que llegaron al depósito (dejados por el técnico o por vos) sin clasificar todavía.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
          <PrimaryButton onClick={onNew}><Plus size={15} /> Nuevo ingreso</PrimaryButton>
        </div>
      </div>
      {playa.length === 0 ? (
        <EmptyState icon={Inbox} title="Playa vacía" subtitle="Todo lo que llegue sin clasificar va a aparecer acá para decidir su destino." />
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {playa.map((item) => (
            <PlayaCard
              key={item.id} item={item} productosRepuestos={productosRepuestos}
              onDerivar={onDerivar} onExtraerRepuesto={onExtraerRepuesto} onDelete={onDelete}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function RecuperablesView({ recuperables, query, onQuery, onUpdateEstado, onUpdateField }) {
  const filtered = recuperables.filter((e) => {
    const q = query.toLowerCase();
    return !q || [e.codigo, e.modelo].some((v) => (v || "").toLowerCase().includes(q));
  });
  const shareOutlet = (equipo) => {
    const text = `${equipo.modelo} — ${equipo.codigo}. Disponible con descuento por imperfección/uso. Consultar precio.`;
    window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, "_blank");
  };
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Banco de recuperables</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>Equipos en reparación, pendientes o reservados como unidad de rescate.</p>
        </div>
        <SearchBox value={query} onChange={onQuery} />
      </div>
      {filtered.length === 0 ? (
        <EmptyState icon={Wrench} title="Banco vacío" subtitle="Los equipos que entren con desperfecto aparecerán acá." />
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {filtered.map((e) => (
            <div key={e.id} className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
              <div className="flex items-center justify-between mb-2">
                <CodeTag>{e.codigo}</CodeTag>
                <StatusBadge estado={e.estado} />
              </div>
              <p className="text-sm font-medium mb-2" style={{ color: INK }}>{e.modelo}</p>
              <div className="flex items-center gap-2">
                <Select value={e.estado} onChange={(ev) => onUpdateEstado(e.id, ev.target.value)} style={{ ...inputStyle, padding: "4px 8px", fontSize: 12 }}>
                  {RECUPERABLE_ESTADOS.concat(["Apto para venta", "Muestra", "Dado de baja"]).map((s) => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </Select>
                {e.estado === "Apto para venta con descuento" && (
                  <button onClick={() => shareOutlet(e)} className="p-1.5 rounded border shrink-0" style={{ borderColor: BORDER }} title="Compartir por WhatsApp">
                    <MessageCircle size={14} style={{ color: "#15803D" }} />
                  </button>
                )}
              </div>
              {e.estado === "Dado de baja" && (
                <div className="mt-2 pt-2 border-t" style={{ borderColor: BORDER }}>
                  <Select
                    value={e.motivoBaja || ""}
                    onChange={(ev) => onUpdateField(e.id, "motivoBaja", ev.target.value)}
                    style={{ ...inputStyle, padding: "4px 8px", fontSize: 12 }}
                  >
                    <option value="">Motivo de la baja...</option>
                    {MOTIVOS_BAJA.map((m) => <option key={m} value={m}>{m}</option>)}
                  </Select>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Ventas comprometidas ----------
const COMPROMETIDA_BADGE = {
  Comprometida: { color: "#B45309", bg: "#FDF1E0" },
  Completada: { color: "#0D9488", bg: "#E3F5F3" },
  Retirada: { color: "#15803D", bg: "#E9F7EF" },
};

function ComprometidasView({ comprometidas, query, onQuery, onNew, onCancelar, onRetirar, onCerrar, onPago }) {
  const pendientes = comprometidas.filter((c) => c.estado === "Comprometida");
  const totalMonto = pendientes.reduce((acc, c) => acc + (Number(c.monto) || 0), 0);

  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Ventas comprometidas</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>
            Mercadería vendida pero todavía en depósito — reserva el stock para que no se use en otra salida. Se puede retirar en varias tandas.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
          <PrimaryButton onClick={onNew}><Plus size={15} /> Nueva venta comprometida</PrimaryButton>
        </div>
      </div>

      {pendientes.length > 0 && (
        <div className="mb-4 px-3 py-2 rounded-lg text-sm" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>
          {pendientes.length} venta(s) comprometida(s) con saldo pendiente — total U$S {totalMonto.toLocaleString()}
        </div>
      )}

      {comprometidas.length === 0 ? (
        <EmptyState icon={Lock} title="No hay ventas comprometidas" subtitle="Cuando reservás mercadería vendida antes del retiro, va a aparecer acá." />
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {comprometidas.map((c) => {
            const retirado = Number(c.cantidadRetirada) || 0;
            const saldo = Math.max(0, (Number(c.cantidad) || 0) - retirado);
            const badge = COMPROMETIDA_BADGE[c.estado] || COMPROMETIDA_BADGE.Comprometida;
            const pagado = (c.pagos || []).reduce((acc, p) => acc + (Number(p.monto) || 0), 0);
            const saldoPago = Math.max(0, (Number(c.monto) || 0) - pagado);
            return (
              <div key={c.id} className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
                <div className="flex items-start justify-between mb-1">
                  <div>
                    <p className="text-sm font-medium" style={{ color: INK }}>{c.razonSocial}</p>
                    <p className="text-xs" style={{ color: MUTED }}>{c.obra}</p>
                  </div>
                  <span className="text-xs px-2 py-0.5 rounded" style={{ color: badge.color, backgroundColor: badge.bg }}>
                    {c.estado}
                  </span>
                </div>
                <p className="text-sm mt-2" style={{ color: INK }}>
                  {c.modelo} · {retirado} de {c.cantidad} retirado{retirado > 0 && c.estado !== "Retirada" ? ` · saldo ${saldo}` : ""}
                </p>
                <p className="text-xs mt-0.5" style={{ color: MUTED }}>
                  Monto: U$S {Number(c.monto || 0).toLocaleString()} · Entrega estimada: {fmtDate(c.fechaEntrega)}
                </p>
                <p className="text-xs mt-0.5" style={{ color: saldoPago > 0 ? "#B45309" : "#15803D" }}>
                  Pagado: U$S {pagado.toLocaleString()} de U$S {Number(c.monto || 0).toLocaleString()}
                  {saldoPago > 0 ? ` · saldo U$S ${saldoPago.toLocaleString()}` : " · pagado por completo"}
                </p>
                <div className="mt-2">
                  <button onClick={() => onPago(c.id)} className="text-xs px-2.5 py-1.5 rounded border" style={{ borderColor: BORDER, color: ACCENT }}>
                    Pagos
                  </button>
                </div>
                {c.estado === "Comprometida" && (
                  <div className="flex gap-2 mt-3">
                    <button onClick={() => onRetirar(c.id)} className="text-xs px-2.5 py-1.5 rounded" style={{ backgroundColor: ACCENT, color: "#FFFFFF" }}>
                      Registrar retiro
                    </button>
                    {retirado === 0 && (
                      <button onClick={() => onCancelar(c.id)} className="text-xs px-2.5 py-1.5 rounded border" style={{ borderColor: BORDER, color: MUTED }}>
                        Cancelar
                      </button>
                    )}
                  </div>
                )}
                {c.estado === "Completada" && (
                  <div className="mt-3">
                    <p className="text-xs mb-2" style={{ color: "#0D9488" }}>Todo lo comprometido ya se retiró.</p>
                    <button onClick={() => onCerrar(c.id)} className="text-xs px-2.5 py-1.5 rounded" style={{ backgroundColor: "#E9F7EF", color: "#15803D" }}>
                      Cerrar
                    </button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------- Muestras ----------
function ComentarioEditor({ value, onSave, placeholder = "Ej: en muestra por defecto de pintura en la puerta" }) {
  const [draft, setDraft] = useState(value || "");
  const [editing, setEditing] = useState(false);

  useEffect(() => setDraft(value || ""), [value]);

  const commit = () => {
    setEditing(false);
    if (draft !== value) onSave(draft);
  };

  return (
    <input
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onFocus={() => setEditing(true)}
      onBlur={commit}
      placeholder={placeholder}
      className="w-full text-xs px-2 py-1.5 rounded border outline-none"
      style={{ borderColor: editing ? ACCENT : BORDER, color: INK }}
    />
  );
}

const MUESTRAS_TABS = [
  { key: "todas", label: "Todas" },
  { key: "con-marca", label: "Con marca" },
  { key: "sin-marca", label: "Sin marca" },
];

function MuestrasView({ muestras, query, onQuery, onUpdateField }) {
  const [tabMarca, setTabMarca] = useState("todas");

  const buscadas = muestras.filter((e) => {
    const q = query.toLowerCase();
    return !q || [e.codigo, e.modelo].some((v) => (v || "").toLowerCase().includes(q));
  });
  const filtered = buscadas.filter((e) => {
    if (tabMarca === "con-marca") return !e.sinMarca;
    if (tabMarca === "sin-marca") return !!e.sinMarca;
    return true;
  });

  const totalTodas = sumCantidad(buscadas);
  const totalSinMarca = sumCantidad(buscadas.filter((e) => e.sinMarca));
  const totalConMarca = totalTodas - totalSinMarca;

  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Muestras</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>
            Equipos usados como pieza de exhibición — en depósito o prestados a un cliente. Incluye tanto muestras con
            marca como muestras sin marca de fábrica (China) — el total de acá es el que se reporta al seguro.
          </p>
        </div>
        <SearchBox value={query} onChange={onQuery} />
      </div>

      <div className="flex items-center gap-2 mb-4 flex-wrap">
        {MUESTRAS_TABS.map((t) => {
          const n = t.key === "todas" ? totalTodas : t.key === "con-marca" ? totalConMarca : totalSinMarca;
          return (
            <button
              key={t.key}
              onClick={() => setTabMarca(t.key)}
              className="text-xs px-3 py-1.5 rounded-full font-medium"
              style={tabMarca === t.key
                ? { backgroundColor: ACCENT, color: "#FFFFFF" }
                : { backgroundColor: "#F2F3F4", color: MUTED }}
            >
              {t.label} ({n})
            </button>
          );
        })}
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={Star} title="No hay muestras cargadas" subtitle="Los equipos clasificados como Muestra van a aparecer acá." />
      ) : (
        <div className="rounded-lg border overflow-hidden" style={{ borderColor: BORDER }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ backgroundColor: "#FAFBFC" }}>
                {["Producto", "Código", "Cantidad", "Estado", "Sin marca", "Comentario"].map((h) => (
                  <th key={h} className="text-left font-medium px-3 py-2.5 border-b" style={{ color: MUTED, borderColor: BORDER, fontSize: 12 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((e) => (
                <tr key={e.id} className="border-b last:border-0" style={{ borderColor: BORDER }}>
                  <td className="px-3 py-2.5" style={{ color: INK }}>{e.modelo}</td>
                  <td className="px-3 py-2.5"><CodeTag>{e.codigo}</CodeTag></td>
                  <td className="px-3 py-2.5" style={{ color: INK }}>{e.cantidad || 1}</td>
                  <td className="px-3 py-2.5"><StatusBadge estado={e.estado} /></td>
                  <td className="px-3 py-2.5 text-center">
                    <input
                      type="checkbox" checked={!!e.sinMarca}
                      onChange={(ev) => onUpdateField(e.id, "sinMarca", ev.target.checked)}
                      title="Muestra de fábrica sin logo de la marca"
                    />
                  </td>
                  <td className="px-3 py-2.5" style={{ minWidth: 260 }}>
                    <ComentarioEditor value={e.notas} onSave={(v) => onUpdateField(e.id, "notas", v)} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ---------- Forms ----------
function EquipoForm({ equipos, onSave }) {
  const [codigo, setCodigo] = useState("");
  const [codigoAuto, setCodigoAuto] = useState(true);
  const [serie, setSerie] = useState("");
  const [modelo, setModelo] = useState("");
  const [fechaIngreso, setFechaIngreso] = useState(todayISO());
  const [estado, setEstado] = useState("En depósito");
  const [ubicacion, setUbicacion] = useState("");
  const [cantidad, setCantidad] = useState(1);
  const [notas, setNotas] = useState("");
  const [motivoBaja, setMotivoBaja] = useState("");
  const [error, setError] = useState("");

  // El código interno se arma solo a partir del modelo (modelo-01, modelo-02...) mientras el
  // usuario no lo edite a mano — si lo toca, dejamos de pisarlo aunque siga cambiando el modelo.
  const handleModelo = (v) => {
    setModelo(v);
    if (codigoAuto) setCodigo(nextCodigoParaModelo(equipos, v));
  };
  const handleCodigo = (v) => {
    setCodigo(v);
    setCodigoAuto(false);
  };

  const submit = () => {
    if (!codigo.trim() || !modelo.trim()) {
      setError("Completá al menos código y modelo.");
      return;
    }
    onSave({ codigo, serie, modelo, fechaIngreso, estado, ubicacion, cantidad: Number(cantidad) || 1, notas, motivoBaja: estado === "Dado de baja" ? motivoBaja : "" });
  };

  return (
    <div>
      <Field label="Modelo"><TextInput value={modelo} onChange={(e) => handleModelo(e.target.value)} placeholder="Ej: AE-AK630-9M-3G-CS-ON" /></Field>
      <Field label="Código interno"><TextInput value={codigo} onChange={(e) => handleCodigo(e.target.value)} placeholder="Se arma solo a partir del modelo" /></Field>
      <Field label="N° de serie"><TextInput value={serie} onChange={(e) => setSerie(e.target.value)} placeholder="N° de serie de fábrica" /></Field>
      <Field label="Cantidad"><TextInput type="number" value={cantidad} onChange={(e) => setCantidad(e.target.value)} /></Field>
      <Field label="Fecha de ingreso"><TextInput type="date" value={fechaIngreso} onChange={(e) => setFechaIngreso(e.target.value)} /></Field>
      <Field label="Estado"><Select value={estado} onChange={(e) => setEstado(e.target.value)}>{ESTADOS.map((s) => <option key={s}>{s}</option>)}</Select></Field>
      {estado === "Dado de baja" && (
        <Field label="Motivo de la baja">
          <Select value={motivoBaja} onChange={(e) => setMotivoBaja(e.target.value)}>
            <option value="">Elegir motivo...</option>
            {MOTIVOS_BAJA.map((m) => <option key={m} value={m}>{m}</option>)}
          </Select>
        </Field>
      )}
      <Field label="Ubicación"><TextInput value={ubicacion} onChange={(e) => setUbicacion(e.target.value)} placeholder="Ej: Depósito principal, estante 3" /></Field>
      <Field label="Comentario"><TextInput value={notas} onChange={(e) => setNotas(e.target.value)} placeholder="Opcional" /></Field>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar equipo</PrimaryButton>
    </div>
  );
}

const FORMATOS_BARCODE = ["code_128", "code_39", "ean_13", "ean_8", "upc_a", "upc_e", "qr_code", "itf", "codabar"];

function EscaneoUnidadesForm({ onSave }) {
  const [modelo, setModelo] = useState("");
  const [estado, setEstado] = useState("En depósito");
  const [ubicacion, setUbicacion] = useState("Depósito principal");
  const [fechaIngreso, setFechaIngreso] = useState(todayISO());
  const [scanInput, setScanInput] = useState("");
  const [series, setSeries] = useState([]);
  const [error, setError] = useState("");

  const [camaraActiva, setCamaraActiva] = useState(false);
  const [camaraError, setCamaraError] = useState("");
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const detectorRef = useRef(null);
  const intervalRef = useRef(null);
  const lastScanRef = useRef({ valor: "", ts: 0 });

  const agregarCodigo = (valor) => {
    const v = (valor || "").trim();
    if (!v) return;
    setSeries((prev) => {
      if (prev.includes(v)) {
        setError(`"${v}" ya fue escaneado en esta carga.`);
        return prev;
      }
      setError("");
      return [...prev, v];
    });
  };

  const registrarScan = () => {
    agregarCodigo(scanInput);
    setScanInput("");
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      registrarScan();
    }
  };

  const quitarSerie = (s) => setSeries(series.filter((x) => x !== s));

  const detenerCamara = () => {
    if (intervalRef.current) { clearInterval(intervalRef.current); intervalRef.current = null; }
    if (streamRef.current) { streamRef.current.getTracks().forEach((t) => t.stop()); streamRef.current = null; }
    setCamaraActiva(false);
  };

  const iniciarCamara = async () => {
    setCamaraError("");
    if (!("BarcodeDetector" in window)) {
      setCamaraError("Este navegador no soporta lectura de códigos por cámara (probá con Chrome en Android). Usá la pistola o escribí el código a mano.");
      return;
    }
    try {
      detectorRef.current = new window.BarcodeDetector({ formats: FORMATOS_BARCODE });
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      streamRef.current = stream;
      setCamaraActiva(true);
    } catch (err) {
      setCamaraError("No se pudo acceder a la cámara. Revisá que le hayas dado permiso al navegador.");
    }
  };

  useEffect(() => {
    if (!camaraActiva || !streamRef.current || !videoRef.current || !detectorRef.current) return;
    const video = videoRef.current;
    video.srcObject = streamRef.current;
    video.play().catch(() => {});
    intervalRef.current = setInterval(async () => {
      if (video.readyState < 2) return;
      try {
        const codigos = await detectorRef.current.detect(video);
        if (codigos.length > 0) {
          const valor = codigos[0].rawValue;
          const ahora = Date.now();
          if (valor === lastScanRef.current.valor && ahora - lastScanRef.current.ts < 2000) return;
          lastScanRef.current = { valor, ts: ahora };
          agregarCodigo(valor);
        }
      } catch (e) {
        // Cuadro no decodificable — se sigue intentando en el próximo ciclo.
      }
    }, 350);
    return () => {
      if (intervalRef.current) { clearInterval(intervalRef.current); intervalRef.current = null; }
    };
  }, [camaraActiva]);

  // Apaga la cámara si se cierra el formulario sin tocar "Cerrar cámara" (ej. se cierra el drawer).
  useEffect(() => () => {
    if (streamRef.current) streamRef.current.getTracks().forEach((t) => t.stop());
  }, []);

  const submit = () => {
    if (!modelo.trim()) {
      setError("Ingresá el modelo.");
      return;
    }
    if (series.length === 0) {
      setError("Escaneá al menos una unidad antes de guardar.");
      return;
    }
    detenerCamara();
    onSave({ modelo, estado, ubicacion, fechaIngreso, series });
  };

  return (
    <div>
      <p className="text-xs mb-3" style={{ color: MUTED }}>
        Cada código escaneado carga una unidad individual (cantidad 1) con ese número de serie —
        para cuando lleguen productos con serie propia por caja.
      </p>
      <Field label="Modelo"><TextInput value={modelo} onChange={(e) => setModelo(e.target.value)} placeholder="Ej: AE-AK630-9M-3G-CS-ON" /></Field>
      <Field label="Estado inicial">
        <Select value={estado} onChange={(e) => setEstado(e.target.value)}>
          {ESTADOS.filter((s) => s !== "Dado de baja").map((s) => <option key={s}>{s}</option>)}
        </Select>
      </Field>
      <Field label="Ubicación"><TextInput value={ubicacion} onChange={(e) => setUbicacion(e.target.value)} /></Field>
      <Field label="Fecha de ingreso"><TextInput type="date" value={fechaIngreso} onChange={(e) => setFechaIngreso(e.target.value)} /></Field>

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Escaneo de unidades</p>

      <div className="mb-2">
        {!camaraActiva ? (
          <SecondaryButton onClick={iniciarCamara}><Camera size={14} /> Escanear con la cámara</SecondaryButton>
        ) : (
          <SecondaryButton onClick={detenerCamara}><X size={14} /> Cerrar cámara</SecondaryButton>
        )}
      </div>
      {camaraError && <p className="text-xs mb-2" style={{ color: "#B45309" }}>{camaraError}</p>}
      {camaraActiva && (
        <div className="mb-3 rounded-lg overflow-hidden border" style={{ borderColor: BORDER }}>
          <video ref={videoRef} muted playsInline className="w-full" style={{ maxHeight: 260, backgroundColor: "#000" }} />
        </div>
      )}

      <Field label={`Código / N° de serie — ${series.length} escaneada${series.length === 1 ? "" : "s"}`}>
        <input
          autoFocus
          value={scanInput}
          onChange={(e) => setScanInput(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder="Apuntá la pistola acá y escaneá — o tipeá el código y Enter"
          className="w-full text-sm px-3 py-2 rounded-md border outline-none"
          style={{ borderColor: BORDER, color: INK }}
        />
      </Field>

      {series.length > 0 && (
        <div className="mb-3 rounded border overflow-y-auto" style={{ borderColor: BORDER, maxHeight: 220 }}>
          {series.map((s, i) => (
            <div key={s} className="flex items-center justify-between px-2.5 py-1.5 text-xs border-b last:border-0" style={{ borderColor: BORDER }}>
              <span style={{ color: INK }}>{i + 1}. {s}</span>
              <button onClick={() => quitarSerie(s)}><X size={13} style={{ color: MUTED }} /></button>
            </div>
          ))}
        </div>
      )}

      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>
        Guardar{series.length > 0 ? ` ${series.length} unidad${series.length === 1 ? "" : "es"}` : ""}
      </PrimaryButton>
    </div>
  );
}

function MovimientoForm({ equipos, playa, productos, onSave }) {
  const [fecha, setFecha] = useState(todayISO());
  const [categoria, setCategoria] = useState("");
  const [sourceId, setSourceId] = useState("");
  const [cantidad, setCantidad] = useState(1);
  const [motivo, setMotivo] = useState("");
  const [cliente, setCliente] = useState("");
  const [monto, setMonto] = useState("");
  const [remito, setRemito] = useState("");
  const [responsable, setResponsable] = useState("");
  const [observaciones, setObservaciones] = useState("");
  const [lugarSalida, setLugarSalida] = useState("Depósito principal");
  const [empresaCliente, setEmpresaCliente] = useState("");
  const [rucCliente, setRucCliente] = useState("");
  const [obra, setObra] = useState("");
  const [firmaNombre, setFirmaNombre] = useState("");
  const [firmaCedula, setFirmaCedula] = useState("");
  const [fotoRemito, setFotoRemito] = useState("");
  const [subiendoFoto, setSubiendoFoto] = useState(false);
  const [error, setError] = useState("");

  const cat = CATEGORIAS_ORIGEN.find((c) => c.value === categoria);

  const opciones = useMemo(() => {
    if (!cat) return [];
    if (cat.type === "equipo") return equipos.filter((e) => cat.estados.includes(e.estado));
    if (cat.type === "playa") return playa;
    if (cat.type === "producto-repuesto") return productos.filter((p) => p.categoriaPrincipal === "Repuestos");
    return [];
  }, [cat, equipos, playa, productos]);

  const source = opciones.find((o) => o.id === sourceId);
  const comprometido = source && cat.type === "equipo" ? (Number(source.comprometido) || 0) : 0;
  const disponible = !source ? 0
    : cat.type === "producto-repuesto" ? Math.max(0, Number(source.stockDisponible) || 0)
    : Math.max(0, (Number(source.cantidad) || 1) - comprometido);

  const handleCategoria = (v) => {
    setCategoria(v);
    setSourceId("");
    setCantidad(1);
    setMotivo(MOTIVO_DEFAULT[v] || "");
  };

  const handleFoto = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setSubiendoFoto(true);
    try {
      const dataUrl = await compressImage(file);
      setFotoRemito(dataUrl);
    } catch (err) {
      setError("No se pudo procesar la foto, probá con otra imagen.");
    }
    setSubiendoFoto(false);
  };

  const submit = () => {
    if (!cat || !source) {
      setError("Elegí una categoría y un producto.");
      return;
    }
    if (!remito.trim()) {
      setError("Ingresá el N° de remito.");
      return;
    }
    if (!firmaNombre.trim()) {
      setError("Falta la aclaración de firma de quien retira/recibe.");
      return;
    }
    const cant = Number(cantidad) || 0;
    if (cant <= 0 || cant > disponible) {
      setError(
        comprometido > 0
          ? `Solo hay ${disponible} libres — ${comprometido} de este producto ya están comprometidas a otra venta y no se pueden retirar para esto.`
          : `La cantidad debe ser mayor a 0 y no puede superar lo disponible (${disponible}).`
      );
      return;
    }
    if (cat.type === "equipo" && !motivo) {
      setError("Elegí el motivo de la salida.");
      return;
    }
    const modelo = cat.type === "equipo" ? source.modelo
      : cat.type === "playa" ? source.descripcion
      : cat.type === "producto-repuesto" ? (source.descripcion || source.nombre)
      : source.nombre;
    const codigo = cat.type === "equipo" ? source.codigo
      : cat.type === "producto-repuesto" ? source.nombre
      : (source.codigo || modelo);
    onSave({
      fecha, categoria: cat.value, categoriaLabel: cat.label, sourceId, codigo, modelo,
      cantidad: cant, motivo: cat.type === "equipo" ? motivo : "",
      cliente, obra, monto: motivo === "Venta" ? Number(monto) || 0 : 0,
      remito, responsable, observaciones,
      lugarSalida, empresaCliente, rucCliente, firmaNombre, firmaCedula, fotoRemito,
    });
  };

  return (
    <div>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Categoría de origen">
        <Select value={categoria} onChange={(e) => handleCategoria(e.target.value)}>
          <option value="">De dónde sale el producto...</option>
          {CATEGORIAS_ORIGEN.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
        </Select>
      </Field>
      {cat && (
        <Field label="Producto">
          <Select value={sourceId} onChange={(e) => { setSourceId(e.target.value); setCantidad(1); }}>
            <option value="">Seleccionar...</option>
            {opciones.map((o) => {
              const label = cat.type === "equipo" ? `${o.codigo} — ${o.modelo}`
                : cat.type === "playa" ? o.descripcion
                : cat.type === "producto-repuesto" ? `${o.nombre}${o.descripcion ? " — " + o.descripcion : ""}`
                : o.nombre;
              const libres = cat.type === "equipo" ? Math.max(0, (Number(o.cantidad) || 1) - (Number(o.comprometido) || 0))
                : cat.type === "producto-repuesto" ? (Number(o.stockDisponible) || 0)
                : (o.cantidad || 1);
              return <option key={o.id} value={o.id}>{label} (disponible: {libres})</option>;
            })}
          </Select>
          {opciones.length === 0 && (
            <p className="text-xs mt-1" style={{ color: "#B45309" }}>No hay stock disponible en esta categoría.</p>
          )}
        </Field>
      )}
      {comprometido > 0 && (
        <div className="flex items-start gap-1.5 mb-3 p-2 rounded" style={{ backgroundColor: "#FEF3E2" }}>
          <AlertTriangle size={14} style={{ color: "#B45309", marginTop: 2 }} />
          <p className="text-xs" style={{ color: "#92400E" }}>
            {comprometido} unidad(es) de este producto ya están comprometidas a otra venta (ver pestaña "Ventas comprometidas") — solo quedan {disponible} libres.
          </p>
        </div>
      )}
      {source && (
        <Field label={`Cantidad a retirar (disponible: ${disponible})`}>
          <TextInput type="number" min="1" max={disponible} value={cantidad} onChange={(e) => setCantidad(e.target.value)} />
        </Field>
      )}
      {cat && cat.type === "equipo" && (
        <Field label="Motivo de la salida">
          <Select value={motivo} onChange={(e) => setMotivo(e.target.value)}>
            {MOTIVOS_SALIDA.map((m) => <option key={m.value} value={m.value}>{m.value}</option>)}
          </Select>
        </Field>
      )}
      <Field label="Cliente / destino"><TextInput value={cliente} onChange={(e) => setCliente(e.target.value)} /></Field>
      {motivo === "Venta" && (
        <Field label="Monto U$S"><TextInput type="number" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="Opcional" /></Field>
      )}

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Ficha de remito</p>
      <Field label="Lugar de salida"><TextInput value={lugarSalida} onChange={(e) => setLugarSalida(e.target.value)} /></Field>
      <Field label="Empresa (razón social)"><TextInput value={empresaCliente} onChange={(e) => setEmpresaCliente(e.target.value)} /></Field>
      <Field label="RUC"><TextInput value={rucCliente} onChange={(e) => setRucCliente(e.target.value)} /></Field>
      <Field label="Obra"><TextInput value={obra} onChange={(e) => setObra(e.target.value)} placeholder="Opcional" /></Field>
      <Field label="N° de remito"><TextInput value={remito} onChange={(e) => setRemito(e.target.value)} /></Field>
      <Field label="Responsable"><TextInput value={responsable} onChange={(e) => setResponsable(e.target.value)} placeholder="Ej: Gastón" /></Field>
      <Field label="Firma — Aclaración (quien retira/recibe)"><TextInput value={firmaNombre} onChange={(e) => setFirmaNombre(e.target.value)} /></Field>
      <Field label="Firma — C.I. N°"><TextInput value={firmaCedula} onChange={(e) => setFirmaCedula(e.target.value)} /></Field>

      <Field label="Foto del remito en papel">
        <input type="file" accept="image/*" capture="environment" onChange={handleFoto} className="text-xs" />
      </Field>
      {subiendoFoto && <p className="text-xs mb-2" style={{ color: MUTED }}>Procesando imagen...</p>}
      {fotoRemito && (
        <div className="mb-3 relative inline-block">
          <img src={fotoRemito} alt="Remito" className="rounded border" style={{ maxWidth: 160, borderColor: BORDER }} />
          <button onClick={() => setFotoRemito("")} className="absolute -top-2 -right-2 rounded-full p-0.5" style={{ backgroundColor: "#B91C1C" }}>
            <X size={12} color="#FFFFFF" />
          </button>
        </div>
      )}

      <Field label="Observaciones"><TextInput value={observaciones} onChange={(e) => setObservaciones(e.target.value)} /></Field>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar salida</PrimaryButton>
    </div>
  );
}

// Estados de equipo desde los que se puede despachar una venta (coincide con las categorías
// "vendible" y "vendible_desc" de CATEGORIAS_ORIGEN).
const ESTADOS_EQUIPO_VENDIBLE = ["En depósito", "Apto para venta", "Apto para venta con descuento"];

function categoriaOrigenPorEstadoEquipo(estado) {
  return CATEGORIAS_ORIGEN.find((c) => c.type === "equipo" && c.estados.includes(estado))?.value || "vendible";
}

// Para una línea de cotización, busca en `equipos` los lotes (por modelo === código de la
// línea) con stock realmente libre (cantidad - comprometido) en un estado apto para venta.
function candidatosParaLinea(equipos, codigoLinea) {
  return equipos
    .filter((e) => e.modelo === codigoLinea && ESTADOS_EQUIPO_VENDIBLE.includes(e.estado))
    .map((e) => ({ id: e.id, codigo: e.codigo, estado: e.estado, disponible: Math.max(0, (Number(e.cantidad) || 1) - (Number(e.comprometido) || 0)) }))
    .filter((e) => e.disponible > 0)
    .sort((a, b) => b.disponible - a.disponible);
}

function SalidaDesdeCotizacionForm({ cotizaciones, equipos, onGenerar }) {
  const [cotizacionId, setCotizacionId] = useState("");
  const [seleccion, setSeleccion] = useState({});
  const [cantidades, setCantidades] = useState({});

  const [fecha, setFecha] = useState(todayISO());
  const [remito, setRemito] = useState("");
  const [responsable, setResponsable] = useState("");
  const [lugarSalida, setLugarSalida] = useState("Depósito principal");
  const [empresaCliente, setEmpresaCliente] = useState("");
  const [rucCliente, setRucCliente] = useState("");
  const [firmaNombre, setFirmaNombre] = useState("");
  const [firmaCedula, setFirmaCedula] = useState("");
  const [observaciones, setObservaciones] = useState("");
  const [fotoRemito, setFotoRemito] = useState("");
  const [subiendoFoto, setSubiendoFoto] = useState(false);
  const [error, setError] = useState("");

  const cotizacionesGanadas = useMemo(() => {
    return cotizaciones
      .filter((c) => c.estado === "Ganada")
      .filter((c) => (c.lineas || []).some((l) => (Number(l.cantidad) || 0) - (Number((c.lineasRetiradas || {})[l.codigo]) || 0) > 0))
      .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
  }, [cotizaciones]);

  const cotizacion = cotizaciones.find((c) => c.id === cotizacionId);

  const filas = useMemo(() => {
    if (!cotizacion) return [];
    const retiradas = cotizacion.lineasRetiradas || {};
    return (cotizacion.lineas || []).map((l) => {
      const pendiente = Math.max(0, (Number(l.cantidad) || 0) - (Number(retiradas[l.codigo]) || 0));
      const candidatos = candidatosParaLinea(equipos, l.codigo);
      const stockDisponible = candidatos.reduce((acc, e) => acc + e.disponible, 0);
      return { ...l, pendiente, candidatos, stockDisponible };
    });
  }, [cotizacion, equipos]);

  const handleCotizacion = (id) => {
    setCotizacionId(id);
    const c = cotizaciones.find((x) => x.id === id);
    if (!c) { setSeleccion({}); setCantidades({}); return; }
    const retiradas = c.lineasRetiradas || {};
    const nuevaSel = {};
    const nuevaCant = {};
    for (const l of c.lineas || []) {
      const pendiente = Math.max(0, (Number(l.cantidad) || 0) - (Number(retiradas[l.codigo]) || 0));
      const stockDisponible = candidatosParaLinea(equipos, l.codigo).reduce((acc, e) => acc + e.disponible, 0);
      const aRetirar = Math.min(pendiente, stockDisponible);
      nuevaSel[l.codigo] = aRetirar > 0;
      nuevaCant[l.codigo] = aRetirar;
    }
    setSeleccion(nuevaSel);
    setCantidades(nuevaCant);
    setError("");
  };

  const toggleLinea = (codigo) => setSeleccion((s) => ({ ...s, [codigo]: !s[codigo] }));
  const setCantidadLinea = (codigo, v) => setCantidades((c) => ({ ...c, [codigo]: v }));

  const handleFoto = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setSubiendoFoto(true);
    try {
      const dataUrl = await compressImage(file);
      setFotoRemito(dataUrl);
    } catch (err) {
      setError("No se pudo procesar la foto, probá con otra imagen.");
    }
    setSubiendoFoto(false);
  };

  const submit = () => {
    if (!cotizacion) {
      setError("Elegí una cotización.");
      return;
    }
    if (!remito.trim()) {
      setError("Ingresá el N° de remito.");
      return;
    }
    if (!firmaNombre.trim()) {
      setError("Falta la aclaración de firma de quien retira/recibe.");
      return;
    }

    const lineasParaGenerar = [];
    for (const f of filas) {
      if (!seleccion[f.codigo]) continue;
      const cantidad = Math.min(Number(cantidades[f.codigo]) || 0, f.pendiente, f.stockDisponible);
      if (cantidad <= 0) continue;
      let restante = cantidad;
      const batches = [];
      for (const cand of f.candidatos) {
        if (restante <= 0) break;
        const tomar = Math.min(cand.disponible, restante);
        if (tomar <= 0) continue;
        batches.push({ equipoId: cand.id, codigo: cand.codigo, categoria: categoriaOrigenPorEstadoEquipo(cand.estado), cantidad: tomar });
        restante -= tomar;
      }
      if (batches.length > 0) {
        lineasParaGenerar.push({ codigo: f.codigo, descripcion: f.descripcion, precioUnit: f.precioUnit, cantidad: cantidad - restante, batches });
      }
    }
    if (lineasParaGenerar.length === 0) {
      setError("No hay productos seleccionados con stock disponible para retirar.");
      return;
    }
    onGenerar(cotizacion, lineasParaGenerar, {
      fecha, remito, responsable, lugarSalida, empresaCliente, rucCliente,
      firmaNombre, firmaCedula, observaciones, fotoRemito,
    });
  };

  return (
    <div>
      <Field label="Cotización ganada">
        <Select value={cotizacionId} onChange={(e) => handleCotizacion(e.target.value)}>
          <option value="">Seleccionar...</option>
          {cotizacionesGanadas.map((c) => (
            <option key={c.id} value={c.id}>{c.cliente || "(Sin cliente)"} — {c.obra || "(Sin obra)"} ({fmtDate(c.fecha)})</option>
          ))}
        </Select>
        {cotizacionesGanadas.length === 0 && (
          <p className="text-xs mt-1" style={{ color: "#B45309" }}>No hay cotizaciones ganadas con productos pendientes de salida.</p>
        )}
      </Field>

      {cotizacion && (
        <>
          <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Productos a retirar</p>
          <div className="mb-3 rounded border overflow-hidden" style={{ borderColor: BORDER }}>
            {filas.map((f) => {
              const sinStock = f.stockDisponible === 0;
              const stockInsuficiente = f.stockDisponible > 0 && f.stockDisponible < f.pendiente;
              const yaCompleto = f.pendiente === 0;
              return (
                <div key={f.codigo} className="px-2.5 py-2 text-xs border-b last:border-0" style={{ borderColor: BORDER, opacity: sinStock || yaCompleto ? 0.55 : 1 }}>
                  <div className="flex items-center gap-2">
                    <input
                      type="checkbox" checked={!!seleccion[f.codigo]} disabled={sinStock || yaCompleto}
                      onChange={() => toggleLinea(f.codigo)}
                    />
                    <div className="flex-1 min-w-0">
                      <span className="font-medium" style={{ color: INK }}>{f.descripcion || f.codigo}</span>
                      <span style={{ color: MUTED }}> · {f.codigo}</span>
                    </div>
                    {seleccion[f.codigo] && !sinStock && !yaCompleto && (
                      <input
                        type="number" min="1" max={Math.min(f.pendiente, f.stockDisponible)}
                        value={cantidades[f.codigo] || 0}
                        onChange={(e) => setCantidadLinea(f.codigo, e.target.value)}
                        className="w-16 text-xs px-1.5 py-1 rounded border"
                        style={{ borderColor: BORDER }}
                      />
                    )}
                  </div>
                  <div className="mt-1 pl-6" style={{ color: MUTED }}>
                    {yaCompleto && "Ya se generó la salida completa de este producto."}
                    {!yaCompleto && sinStock && "Sin stock disponible en depósito — no se puede incluir en esta salida."}
                    {!yaCompleto && !sinStock && stockInsuficiente && `Pendiente: ${f.pendiente} · solo hay ${f.stockDisponible} disponible(s).`}
                    {!yaCompleto && !sinStock && !stockInsuficiente && `Pendiente: ${f.pendiente} · disponible: ${f.stockDisponible}.`}
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Ficha de remito</p>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Lugar de salida"><TextInput value={lugarSalida} onChange={(e) => setLugarSalida(e.target.value)} /></Field>
      <Field label="Empresa (razón social)"><TextInput value={empresaCliente} onChange={(e) => setEmpresaCliente(e.target.value)} /></Field>
      <Field label="RUC"><TextInput value={rucCliente} onChange={(e) => setRucCliente(e.target.value)} /></Field>
      <Field label="N° de remito"><TextInput value={remito} onChange={(e) => setRemito(e.target.value)} /></Field>
      <Field label="Responsable"><TextInput value={responsable} onChange={(e) => setResponsable(e.target.value)} placeholder="Ej: Gastón" /></Field>
      <Field label="Firma — Aclaración (quien retira/recibe)"><TextInput value={firmaNombre} onChange={(e) => setFirmaNombre(e.target.value)} /></Field>
      <Field label="Firma — C.I. N°"><TextInput value={firmaCedula} onChange={(e) => setFirmaCedula(e.target.value)} /></Field>

      <Field label="Foto del remito en papel">
        <input type="file" accept="image/*" capture="environment" onChange={handleFoto} className="text-xs" />
      </Field>
      {subiendoFoto && <p className="text-xs mb-2" style={{ color: MUTED }}>Procesando imagen...</p>}
      {fotoRemito && (
        <div className="mb-3 relative inline-block">
          <img src={fotoRemito} alt="Remito" className="rounded border" style={{ maxWidth: 160, borderColor: BORDER }} />
          <button onClick={() => setFotoRemito("")} className="absolute -top-2 -right-2 rounded-full p-0.5" style={{ backgroundColor: "#B91C1C" }}>
            <X size={12} color="#FFFFFF" />
          </button>
        </div>
      )}

      <Field label="Observaciones"><TextInput value={observaciones} onChange={(e) => setObservaciones(e.target.value)} placeholder="Opcional" /></Field>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Generar salida</PrimaryButton>
    </div>
  );
}

function EntradaForm({ equipos, onSave }) {
  const [fecha, setFecha] = useState(todayISO());
  const [codigo, setCodigo] = useState("");
  const [tipo, setTipo] = useState(TIPOS_ENTRADA[0]);
  const [origen, setOrigen] = useState("");
  const [motivo, setMotivo] = useState("");
  const [estadoResultante, setEstadoResultante] = useState(ESTADOS_RESULTANTES[0]);
  const [responsable, setResponsable] = useState("");
  const [error, setError] = useState("");

  const submit = () => {
    if (!codigo) {
      setError("Elegí el equipo que está ingresando.");
      return;
    }
    onSave({ fecha, codigo, tipo, origen, motivo, estadoResultante, responsable });
  };

  return (
    <div>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Equipo">
        <Select value={codigo} onChange={(e) => setCodigo(e.target.value)}>
          <option value="">Seleccionar equipo...</option>
          {equipos.map((eq) => <option key={eq.id} value={eq.codigo}>{eq.codigo} — {eq.modelo}</option>)}
        </Select>
      </Field>
      <Field label="Tipo de entrada"><Select value={tipo} onChange={(e) => setTipo(e.target.value)}>{TIPOS_ENTRADA.map((t) => <option key={t}>{t}</option>)}</Select></Field>
      <Field label="Origen"><TextInput value={origen} onChange={(e) => setOrigen(e.target.value)} placeholder="Ej: Cliente, Fábrica, Técnico" /></Field>
      <Field label="Motivo"><TextInput value={motivo} onChange={(e) => setMotivo(e.target.value)} /></Field>
      <Field label="Estado resultante"><Select value={estadoResultante} onChange={(e) => setEstadoResultante(e.target.value)}>{ESTADOS_RESULTANTES.map((s) => <option key={s}>{s}</option>)}</Select></Field>
      <Field label="Responsable"><TextInput value={responsable} onChange={(e) => setResponsable(e.target.value)} /></Field>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar entrada</PrimaryButton>
    </div>
  );
}

function VentaForm({ equipos, onSave }) {
  const [codigo, setCodigo] = useState("");
  const [cliente, setCliente] = useState("");
  const [obra, setObra] = useState("");
  const [fechaVenta, setFechaVenta] = useState(todayISO());
  const [error, setError] = useState("");

  const submit = () => {
    if (!codigo || !cliente.trim()) {
      setError("Elegí el equipo e ingresá el cliente.");
      return;
    }
    onSave({
      codigo, cliente, obra, fechaVenta,
      vtoService1: addMonthsISO(fechaVenta, 12),
      vtoService2: addMonthsISO(fechaVenta, 24),
      estadoService1: "Pendiente", estadoService2: "Pendiente",
    });
  };

  return (
    <div>
      <Field label="Equipo">
        <Select value={codigo} onChange={(e) => setCodigo(e.target.value)}>
          <option value="">Seleccionar equipo...</option>
          {equipos.map((eq) => <option key={eq.id} value={eq.codigo}>{eq.codigo} — {eq.modelo}</option>)}
        </Select>
      </Field>
      <Field label="Cliente"><TextInput value={cliente} onChange={(e) => setCliente(e.target.value)} /></Field>
      <Field label="Obra"><TextInput value={obra} onChange={(e) => setObra(e.target.value)} /></Field>
      <Field label="Fecha de venta"><TextInput type="date" value={fechaVenta} onChange={(e) => setFechaVenta(e.target.value)} /></Field>
      <p className="text-xs mb-3" style={{ color: MUTED }}>
        Los vencimientos de service (12 y 24 meses) se calculan automáticamente a partir de la fecha de venta.
      </p>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar venta</PrimaryButton>
    </div>
  );
}

function GestionServiceForm({ venta, field, onUpdate, onClose }) {
  const [contactado, setContactado] = useState(!!venta[`contactado${field}`]);
  const [fechaContacto, setFechaContacto] = useState(venta[`fechaContacto${field}`] || "");
  const [decision, setDecision] = useState(venta[`decision${field}`] || "");
  const [cita, setCita] = useState(venta[`cita${field}`] || { fecha: "", hora: "", equipos: [] });
  const [nuevoProducto, setNuevoProducto] = useState("");
  const [nuevaCantidad, setNuevaCantidad] = useState(1);

  const vto = field === "Service1" ? venta.vtoService1 : venta.vtoService2;
  const dias = daysUntil(vto);
  const proximoSeguimiento = fechaContacto ? addBusinessDaysISO(fechaContacto, 5) : "";
  const seguimientoVencido = fechaContacto && !decision && daysUntil(proximoSeguimiento) <= 0;

  const registrarContacto = () => {
    const hoy = todayISO();
    setContactado(true);
    setFechaContacto(hoy);
    onUpdate(venta.id, `contactado${field}`, true);
    onUpdate(venta.id, `fechaContacto${field}`, hoy);
  };

  const reintentar = () => {
    const hoy = todayISO();
    setFechaContacto(hoy);
    onUpdate(venta.id, `fechaContacto${field}`, hoy);
  };

  const elegirDecision = (d) => {
    setDecision(d);
    onUpdate(venta.id, `decision${field}`, d);
    if (d === "Rechaza el service") onUpdate(venta.id, `estado${field}`, "Vencido");
  };

  const agregarEquipoCita = () => {
    if (!nuevoProducto.trim()) return;
    const next = { ...cita, equipos: [...(cita.equipos || []), { producto: nuevoProducto, cantidad: Number(nuevaCantidad) || 1 }] };
    setCita(next);
    onUpdate(venta.id, `cita${field}`, next);
    setNuevoProducto("");
    setNuevaCantidad(1);
  };

  const quitarEquipoCita = (idx) => {
    const next = { ...cita, equipos: cita.equipos.filter((_, i) => i !== idx) };
    setCita(next);
    onUpdate(venta.id, `cita${field}`, next);
  };

  const actualizarCitaCampo = (campo, valor) => {
    const next = { ...cita, [campo]: valor };
    setCita(next);
    onUpdate(venta.id, `cita${field}`, next);
  };

  return (
    <div>
      <div className="mb-4 p-3 rounded" style={{ backgroundColor: "#F7F8FA" }}>
        <p className="text-sm font-medium" style={{ color: INK }}>{venta.cliente} — {venta.obra}</p>
        <p className="text-xs mt-0.5" style={{ color: MUTED }}>
          Vence {fmtDate(vto)} ({dias < 0 ? `vencido hace ${Math.abs(dias)}d` : `en ${dias}d`})
        </p>
      </div>

      {!contactado ? (
        <>
          <p className="text-sm mb-3" style={{ color: MUTED }}>Todavía no se registró contacto con el cliente por este service.</p>
          <PrimaryButton onClick={registrarContacto}>Registrar contacto con cliente (hoy)</PrimaryButton>
        </>
      ) : (
        <>
          <p className="text-sm mb-1" style={{ color: INK }}>Contactado el <strong>{fmtDate(fechaContacto)}</strong>.</p>
          {!decision && (
            <p className="text-xs mb-3" style={{ color: seguimientoVencido ? "#B91C1C" : MUTED }}>
              Seguimiento sugerido: {fmtDate(proximoSeguimiento)}
              {seguimientoVencido ? " — ya venció, reenviá el recordatorio." : ""}
            </p>
          )}

          {!decision ? (
            <>
              <p className="text-xs font-medium mb-2" style={{ color: MUTED }}>Decisión final del cliente</p>
              <div className="flex flex-col gap-2 mb-3">
                {DECISIONES_SERVICE.map((d) => (
                  <SecondaryButton key={d} onClick={() => elegirDecision(d)}>{d}</SecondaryButton>
                ))}
              </div>
              {seguimientoVencido && (
                <button onClick={reintentar} className="text-xs px-2.5 py-1.5 rounded" style={{ backgroundColor: "#FDF1E0", color: "#B45309" }}>
                  Marcar que se reenvió el recordatorio hoy
                </button>
              )}
            </>
          ) : (
            <div className="mb-3 p-2 rounded" style={{ backgroundColor: decision.startsWith("Acepta") ? "#E9F7EF" : "#FBEAEA" }}>
              <p className="text-sm" style={{ color: decision.startsWith("Acepta") ? "#15803D" : "#B91C1C" }}>{decision}</p>
            </div>
          )}

          {decision === "Acepta hacer el service" && (
            <div className="mt-2 pt-3 border-t" style={{ borderColor: BORDER }}>
              <p className="text-xs font-medium mb-2" style={{ color: MUTED }}>Cita de service</p>
              <Field label="Fecha"><TextInput type="date" value={cita.fecha} onChange={(e) => actualizarCitaCampo("fecha", e.target.value)} /></Field>
              <Field label="Hora aproximada"><TextInput type="time" value={cita.hora} onChange={(e) => actualizarCitaCampo("hora", e.target.value)} /></Field>
              <p className="text-xs font-medium mb-1.5" style={{ color: MUTED }}>Equipos a intervenir</p>
              {(cita.equipos || []).map((eq, i) => (
                <div key={i} className="flex items-center justify-between text-sm mb-1 px-2 py-1 rounded" style={{ backgroundColor: "#F7F8FA" }}>
                  <span style={{ color: INK }}>{eq.producto} · cant. {eq.cantidad}</span>
                  <button onClick={() => quitarEquipoCita(i)}><X size={13} style={{ color: MUTED }} /></button>
                </div>
              ))}
              <div className="flex gap-2 mt-2">
                <TextInput value={nuevoProducto} onChange={(e) => setNuevoProducto(e.target.value)} placeholder="Producto / modelo" />
                <TextInput type="number" value={nuevaCantidad} onChange={(e) => setNuevaCantidad(e.target.value)} style={{ width: 70 }} />
                <SecondaryButton onClick={agregarEquipoCita}><Plus size={14} /></SecondaryButton>
              </div>
            </div>
          )}
        </>
      )}

      <div className="mt-5">
        <SecondaryButton onClick={onClose}>Cerrar</SecondaryButton>
      </div>
    </div>
  );
}

function ComprometidaForm({ equipos, productos, onSave }) {
  const [fecha, setFecha] = useState(todayISO());
  const [razonSocial, setRazonSocial] = useState("");
  const [obra, setObra] = useState("");
  const [equipoId, setEquipoId] = useState("");
  const [cantidad, setCantidad] = useState(1);
  const [monto, setMonto] = useState("");
  const [montoTocado, setMontoTocado] = useState(false);
  const [fechaEntrega, setFechaEntrega] = useState("");
  const [error, setError] = useState("");

  const vendibles = equipos.filter((e) => e.estado === "En depósito" || e.estado === "Apto para venta" || e.estado === "Apto para venta con descuento");
  const equipo = vendibles.find((e) => e.id === equipoId);
  const disponible = equipo ? Math.max(0, (Number(equipo.cantidad) || 1) - (Number(equipo.comprometido) || 0)) : 0;
  const productoCatalogo = equipo ? (productos || []).find((p) => p.nombre === equipo.modelo) : null;

  const sugerirMonto = (cant, prod) => {
    if (montoTocado || !prod) return;
    setMonto(String((Number(prod.precioLista) || 0) * cant));
  };

  const handleEquipo = (id) => {
    setEquipoId(id);
    setCantidad(1);
    const eq = vendibles.find((e) => e.id === id);
    const prod = eq ? (productos || []).find((p) => p.nombre === eq.modelo) : null;
    sugerirMonto(1, prod);
  };

  const handleCantidad = (v) => {
    setCantidad(v);
    sugerirMonto(Number(v) || 0, productoCatalogo);
  };

  const submit = () => {
    if (!razonSocial.trim() || !equipo) {
      setError("Ingresá la razón social y elegí el producto.");
      return;
    }
    const cant = Number(cantidad) || 0;
    if (cant <= 0 || cant > disponible) {
      setError(`La cantidad no puede superar lo disponible sin comprometer (${disponible}).`);
      return;
    }
    onSave({
      fecha, razonSocial, obra, equipoId, modelo: equipo.modelo, codigo: equipo.codigo,
      cantidad: cant, monto: Number(monto) || 0, fechaEntrega,
    });
  };

  return (
    <div>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Razón social"><TextInput value={razonSocial} onChange={(e) => setRazonSocial(e.target.value)} /></Field>
      <Field label="Obra"><TextInput value={obra} onChange={(e) => setObra(e.target.value)} /></Field>
      <Field label="Producto">
        <Select value={equipoId} onChange={(e) => handleEquipo(e.target.value)}>
          <option value="">Seleccionar equipo...</option>
          {vendibles.map((eq) => {
            const libres = Math.max(0, (Number(eq.cantidad) || 1) - (Number(eq.comprometido) || 0));
            return <option key={eq.id} value={eq.id}>{eq.codigo} — {eq.modelo} (disponible: {libres})</option>;
          })}
        </Select>
      </Field>
      {equipo && productoCatalogo && (
        <p className="text-xs mb-3" style={{ color: MUTED }}>
          Precio de lista: U$S {Number(productoCatalogo.precioLista || 0).toLocaleString()} — el monto se sugiere solo, pero se puede editar libremente si se negoció otro precio.
        </p>
      )}
      {equipo && (
        <Field label={`Cantidad a comprometer (disponible: ${disponible})`}>
          <TextInput type="number" min="1" max={disponible} value={cantidad} onChange={(e) => handleCantidad(e.target.value)} />
        </Field>
      )}
      <Field label="Monto U$S"><TextInput type="number" value={monto} onChange={(e) => { setMonto(e.target.value); setMontoTocado(true); }} /></Field>
      <Field label="Fecha estimada de entrega"><TextInput type="date" value={fechaEntrega} onChange={(e) => setFechaEntrega(e.target.value)} /></Field>
      <p className="text-xs mb-3" style={{ color: MUTED }}>
        Esta cantidad queda reservada: no se va a poder retirar del depósito para otra salida hasta que la marques como retirada.
      </p>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar venta comprometida</PrimaryButton>
    </div>
  );
}

function RetiroParcialForm({ comprometida, onSave }) {
  const retirado = Number(comprometida.cantidadRetirada) || 0;
  const saldo = Math.max(0, (Number(comprometida.cantidad) || 0) - retirado);

  const [cantidad, setCantidad] = useState(saldo);
  const [fecha, setFecha] = useState(todayISO());
  const [lugarSalida, setLugarSalida] = useState("Depósito principal");
  const [empresaCliente, setEmpresaCliente] = useState(comprometida.razonSocial || "");
  const [rucCliente, setRucCliente] = useState("");
  const [obra, setObra] = useState(comprometida.obra || "");
  const [remito, setRemito] = useState("");
  const [responsable, setResponsable] = useState("");
  const [firmaNombre, setFirmaNombre] = useState("");
  const [firmaCedula, setFirmaCedula] = useState("");
  const [fotoRemito, setFotoRemito] = useState("");
  const [subiendoFoto, setSubiendoFoto] = useState(false);
  const [observaciones, setObservaciones] = useState("");
  const [error, setError] = useState("");

  const handleFoto = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setSubiendoFoto(true);
    try {
      const dataUrl = await compressImage(file);
      setFotoRemito(dataUrl);
    } catch (err) {
      setError("No se pudo procesar la foto, probá con otra imagen.");
    }
    setSubiendoFoto(false);
  };

  const submit = () => {
    const cant = Number(cantidad) || 0;
    if (cant <= 0 || cant > saldo) {
      setError(`La cantidad debe ser mayor a 0 y no puede superar el saldo pendiente (${saldo}).`);
      return;
    }
    if (!remito.trim()) {
      setError("Ingresá el N° de remito.");
      return;
    }
    if (!firmaNombre.trim()) {
      setError("Falta la aclaración de firma de quien retira/recibe.");
      return;
    }
    onSave({ cantidad: cant, fecha, lugarSalida, empresaCliente, rucCliente, obra, remito, responsable, firmaNombre, firmaCedula, fotoRemito, observaciones });
  };

  return (
    <div>
      <div className="mb-4 p-3 rounded" style={{ backgroundColor: "#F7F8FA" }}>
        <p className="text-sm font-medium" style={{ color: INK }}>{comprometida.razonSocial} — {comprometida.obra}</p>
        <p className="text-xs mt-0.5" style={{ color: MUTED }}>{comprometida.modelo} · saldo pendiente: {saldo} de {comprometida.cantidad}</p>
      </div>

      <Field label={`Cantidad a retirar (saldo: ${saldo})`}>
        <TextInput type="number" min="1" max={saldo} value={cantidad} onChange={(e) => setCantidad(e.target.value)} />
      </Field>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Ficha de remito</p>
      <Field label="Lugar de salida"><TextInput value={lugarSalida} onChange={(e) => setLugarSalida(e.target.value)} /></Field>
      <Field label="Empresa (razón social)"><TextInput value={empresaCliente} onChange={(e) => setEmpresaCliente(e.target.value)} /></Field>
      <Field label="RUC"><TextInput value={rucCliente} onChange={(e) => setRucCliente(e.target.value)} /></Field>
      <Field label="Obra"><TextInput value={obra} onChange={(e) => setObra(e.target.value)} /></Field>
      <Field label="N° de remito"><TextInput value={remito} onChange={(e) => setRemito(e.target.value)} /></Field>
      <Field label="Responsable"><TextInput value={responsable} onChange={(e) => setResponsable(e.target.value)} placeholder="Ej: Gastón" /></Field>
      <Field label="Firma — Aclaración (quien retira/recibe)"><TextInput value={firmaNombre} onChange={(e) => setFirmaNombre(e.target.value)} /></Field>
      <Field label="Firma — C.I. N°"><TextInput value={firmaCedula} onChange={(e) => setFirmaCedula(e.target.value)} /></Field>

      <Field label="Foto del remito en papel">
        <input type="file" accept="image/*" capture="environment" onChange={handleFoto} className="text-xs" />
      </Field>
      {subiendoFoto && <p className="text-xs mb-2" style={{ color: MUTED }}>Procesando imagen...</p>}
      {fotoRemito && (
        <div className="mb-3 relative inline-block">
          <img src={fotoRemito} alt="Remito" className="rounded border" style={{ maxWidth: 160, borderColor: BORDER }} />
          <button onClick={() => setFotoRemito("")} className="absolute -top-2 -right-2 rounded-full p-0.5" style={{ backgroundColor: "#B91C1C" }}>
            <X size={12} color="#FFFFFF" />
          </button>
        </div>
      )}

      <Field label="Observaciones"><TextInput value={observaciones} onChange={(e) => setObservaciones(e.target.value)} placeholder="Opcional" /></Field>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Registrar retiro</PrimaryButton>
    </div>
  );
}

const FORMAS_PAGO = ["Efectivo", "Transferencia", "Cheque", "Tarjeta", "Otro"];

function PagosForm({ comprometida, onAgregar, onQuitar }) {
  const pagos = comprometida.pagos || [];
  const pagado = pagos.reduce((acc, p) => acc + (Number(p.monto) || 0), 0);
  const total = Number(comprometida.monto) || 0;
  const saldo = Math.max(0, total - pagado);

  const [fecha, setFecha] = useState(todayISO());
  const [monto, setMonto] = useState("");
  const [formaPago, setFormaPago] = useState(FORMAS_PAGO[0]);
  const [error, setError] = useState("");

  const submit = () => {
    const m = Number(monto) || 0;
    if (m <= 0) {
      setError("Ingresá un monto mayor a 0.");
      return;
    }
    onAgregar({ fecha, monto: m, formaPago });
    setMonto("");
    setError("");
  };

  return (
    <div>
      <div className="mb-4 p-3 rounded" style={{ backgroundColor: "#F7F8FA" }}>
        <p className="text-sm font-medium" style={{ color: INK }}>{comprometida.razonSocial} — {comprometida.obra}</p>
        <p className="text-xs mt-0.5" style={{ color: MUTED }}>
          Pagado U$S {pagado.toLocaleString()} de U$S {total.toLocaleString()} · saldo U$S {saldo.toLocaleString()}
        </p>
      </div>

      {pagos.length > 0 && (
        <div className="mb-4 rounded border overflow-hidden" style={{ borderColor: BORDER }}>
          {pagos.map((p, i) => (
            <div key={i} className="flex items-center justify-between px-2.5 py-2 text-xs border-b last:border-0" style={{ borderColor: BORDER }}>
              <div>
                <span className="font-medium" style={{ color: INK }}>U$S {Number(p.monto).toLocaleString()}</span>
                <span style={{ color: MUTED }}> · {p.formaPago} · {fmtDate(p.fecha)}</span>
              </div>
              <button onClick={() => onQuitar(i)}><X size={13} style={{ color: MUTED }} /></button>
            </div>
          ))}
        </div>
      )}

      <p className="text-xs font-semibold mb-2" style={{ color: ACCENT }}>Registrar pago</p>
      <Field label="Monto U$S"><TextInput type="number" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="0" /></Field>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Forma de pago">
        <Select value={formaPago} onChange={(e) => setFormaPago(e.target.value)}>
          {FORMAS_PAGO.map((f) => <option key={f} value={f}>{f}</option>)}
        </Select>
      </Field>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Agregar pago</PrimaryButton>
    </div>
  );
}

function PlayaForm({ onSave }) {
  const [fecha, setFecha] = useState(todayISO());
  const [descripcion, setDescripcion] = useState("");
  const [origen, setOrigen] = useState(ORIGENES_PLAYA[0]);
  const [cantidad, setCantidad] = useState(1);
  const [notas, setNotas] = useState("");
  const [error, setError] = useState("");

  const submit = () => {
    if (!descripcion.trim()) {
      setError("Describí qué producto llegó.");
      return;
    }
    onSave({ fecha, descripcion, origen, cantidad: Number(cantidad) || 1, notas });
  };

  return (
    <div>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Descripción">
        <TextInput value={descripcion} onChange={(e) => setDescripcion(e.target.value)} placeholder="Ej: Horno con display roto, modelo AE-AK630" />
      </Field>
      <Field label="Cantidad"><TextInput type="number" value={cantidad} onChange={(e) => setCantidad(e.target.value)} /></Field>
      <Field label="Quién lo dejó"><Select value={origen} onChange={(e) => setOrigen(e.target.value)}>{ORIGENES_PLAYA.map((o) => <option key={o}>{o}</option>)}</Select></Field>
      <Field label="Notas"><TextInput value={notas} onChange={(e) => setNotas(e.target.value)} placeholder="Opcional" /></Field>
      <p className="text-xs mb-3" style={{ color: MUTED }}>
        Una vez cargado, vas a poder derivarlo a Banco de recuperables, Equipo de socorro, Stock vendible o Muestra — o extraerle repuestos sin sacarlo de playa.
      </p>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar en playa</PrimaryButton>
    </div>
  );
}

// ---------- Catálogo de productos ----------
function ProductoForm({ producto, defaults, onSave }) {
  const [nombre, setNombre] = useState(producto?.nombre || "");
  const [categoria, setCategoria] = useState(producto?.categoria || "");
  const [categoriaPrincipal, setCategoriaPrincipal] = useState(producto?.categoriaPrincipal || defaults?.categoriaPrincipal || "");
  const [subcategoria, setSubcategoria] = useState(producto?.subcategoria || "");
  const [subcategoria2, setSubcategoria2] = useState(producto?.subcategoria2 || "");
  const [subcategoria3, setSubcategoria3] = useState(producto?.subcategoria3 || "");
  const [ordenNumerico, setOrdenNumerico] = useState(producto ? String(producto.ordenNumerico ?? "") : "");
  const [descripcion, setDescripcion] = useState(producto?.descripcion || "");
  const [especLabel, setEspecLabel] = useState(producto?.especLabel || "");
  const [especValor, setEspecValor] = useState(producto?.especValor || "");
  const [precioLista, setPrecioLista] = useState(producto ? String(producto.precioLista ?? "") : "");
  const [costoOrigen, setCostoOrigen] = useState(producto ? String(producto.costoOrigen ?? "") : "");
  const [costoPy, setCostoPy] = useState(producto ? String(producto.costoPy ?? "") : "");
  const [contenedorTipo, setContenedorTipo] = useState(producto?.contenedorTipo || "");
  const [contenedorCantidad, setContenedorCantidad] = useState(producto ? String(producto.contenedorCantidad ?? "") : "");
  const [stockDisponible, setStockDisponible] = useState(producto ? String(producto.stockDisponible ?? "") : "");
  const [foto, setFoto] = useState(producto?.foto || "");
  const [subiendoFoto, setSubiendoFoto] = useState(false);
  const [fichaFile, setFichaFile] = useState(null);
  const [error, setError] = useState("");
  const [guardando, setGuardando] = useState(false);

  const handleFoto = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setSubiendoFoto(true);
    try {
      const dataUrl = await compressImage(file, 480, 0.75);
      setFoto(dataUrl);
    } catch (err) {
      setError("No se pudo procesar la foto, probá con otra imagen.");
    }
    setSubiendoFoto(false);
  };

  const handleFicha = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (file.size > MAX_FICHA_BYTES) {
      setError(`Ese PDF pesa ${(file.size / 1024).toFixed(0)}KB — el máximo es ${(MAX_FICHA_BYTES / 1024).toFixed(0)}KB. Probá comprimirlo o recortar imágenes pesadas.`);
      e.target.value = "";
      return;
    }
    setError("");
    setFichaFile(file);
  };

  const submit = async () => {
    if (!nombre.trim()) {
      setError("Ingresá el nombre/código del producto.");
      return;
    }
    setGuardando(true);
    setError("");
    try {
      let ficha = {};
      if (fichaFile) {
        const fichaTecnicaData = await readFileAsDataUrl(fichaFile);
        ficha = { fichaTecnicaData, fichaTecnicaNombre: fichaFile.name };
      } else if (producto) {
        // conservar la ficha técnica existente si no se subió una nueva
        ficha = { fichaTecnicaData: producto.fichaTecnicaData || "", fichaTecnicaNombre: producto.fichaTecnicaNombre || "" };
      }
      onSave({
        nombre, categoria, descripcion, especLabel, especValor,
        categoriaPrincipal, subcategoria, subcategoria2, subcategoria3,
        ordenNumerico: ordenNumerico === "" ? null : Number(ordenNumerico) || 0,
        precioLista: Number(precioLista) || 0,
        costoOrigen: Number(costoOrigen) || 0, costoPy: Number(costoPy) || 0,
        contenedorTipo, contenedorCantidad: Number(contenedorCantidad) || 0,
        stockDisponible: stockDisponible === "" ? null : Number(stockDisponible) || 0,
        foto, ...ficha,
      });
    } catch (err) {
      setError("No se pudo leer la ficha técnica. Probá de nuevo.");
      setGuardando(false);
    }
  };

  const esRepuesto = categoriaPrincipal === "Repuestos";

  return (
    <div>
      <Field label={esRepuesto ? "Código del repuesto" : "Nombre / código del producto"}>
        <TextInput value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder={esRepuesto ? "Ej: AE-AWG-2T-30-ON-Rep-1" : "Ej: AE-AK630-9M-3G-CS-ON"} />
      </Field>
      <Field label="Categoría"><TextInput value={categoria} onChange={(e) => setCategoria(e.target.value)} placeholder="Ej: Electrodomésticos, Aires acondicionados..." /></Field>
      <Field label="Descripción (aparece en la cotización)"><Textarea value={descripcion} onChange={(e) => setDescripcion(e.target.value)} placeholder={esRepuesto ? "Ej: Motor" : "Descripción técnica completa del producto"} /></Field>
      <div className="flex gap-2">
        <Field label="Especificación — etiqueta"><TextInput value={especLabel} onChange={(e) => setEspecLabel(e.target.value)} placeholder="Ej: Capacidad BTU" /></Field>
        <Field label="Especificación — valor"><TextInput value={especValor} onChange={(e) => setEspecValor(e.target.value)} placeholder="Ej: 12.000" /></Field>
      </div>
      <Field label="Precio de lista (venta) U$S"><TextInput type="number" value={precioLista} onChange={(e) => setPrecioLista(e.target.value)} /></Field>

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Categorización (para agrupar y ordenar el catálogo)</p>
      <div className="flex gap-2">
        <Field label="Categoría principal"><TextInput value={categoriaPrincipal} onChange={(e) => setCategoriaPrincipal(e.target.value)} placeholder="Ej: Cocina" /></Field>
        <Field label={esRepuesto ? "Tipo de equipo" : "Subcategoría"}>
          <TextInput value={subcategoria} onChange={(e) => setSubcategoria(e.target.value)} placeholder={esRepuesto ? "Ej: Horno" : "Ej: Campana"} />
        </Field>
      </div>
      <div className="flex gap-2">
        <Field label={esRepuesto ? "Código de equipo" : "Subcategoría 2"}>
          <TextInput value={subcategoria2} onChange={(e) => setSubcategoria2(e.target.value)} placeholder={esRepuesto ? "Ej: AE-AK630-9M-3G-CS-ON" : "Ej: Pared"} />
        </Field>
        <Field label="Subcategoría 3"><TextInput value={subcategoria3} onChange={(e) => setSubcategoria3(e.target.value)} placeholder="Ej: Telescópica" /></Field>
      </div>
      <Field label="Orden dentro del grupo (número — ej. potencia, litraje, hornallas)">
        <TextInput type="number" value={ordenNumerico} onChange={(e) => setOrdenNumerico(e.target.value)} placeholder="Ej: 700" />
      </Field>

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Datos internos de costo (no aparecen en la cotización)</p>
      <div className="flex gap-2">
        <Field label="Costo de origen U$S"><TextInput type="number" value={costoOrigen} onChange={(e) => setCostoOrigen(e.target.value)} /></Field>
        <Field label="Costo puesto en PY U$S"><TextInput type="number" value={costoPy} onChange={(e) => setCostoPy(e.target.value)} /></Field>
      </div>
      <div className="flex gap-2">
        <Field label="Contenedor (tipo)"><TextInput value={contenedorTipo} onChange={(e) => setContenedorTipo(e.target.value)} placeholder="Ej: 40HQ" /></Field>
        <Field label="Cantidad por contenedor"><TextInput type="number" value={contenedorCantidad} onChange={(e) => setContenedorCantidad(e.target.value)} /></Field>
      </div>
      <Field label="Stock disponible (unidades en depósito)"><TextInput type="number" value={stockDisponible} onChange={(e) => setStockDisponible(e.target.value)} /></Field>

      <Field label="Foto de referencia">
        <input type="file" accept="image/*" onChange={handleFoto} className="text-xs" />
      </Field>
      {subiendoFoto && <p className="text-xs mb-2" style={{ color: MUTED }}>Procesando imagen...</p>}
      {foto && (
        <div className="mb-3 relative inline-block">
          <img src={foto} alt="Producto" className="rounded border" style={{ maxWidth: 140, borderColor: BORDER }} />
          <button onClick={() => setFoto("")} className="absolute -top-2 -right-2 rounded-full p-0.5" style={{ backgroundColor: "#B91C1C" }}>
            <X size={12} color="#FFFFFF" />
          </button>
        </div>
      )}

      <Field label={`Ficha técnica (PDF, máx. ${(MAX_FICHA_BYTES / 1024).toFixed(0)}KB)`}>
        <input type="file" accept="application/pdf" onChange={handleFicha} className="text-xs" />
      </Field>
      {fichaFile && <p className="text-xs mb-2" style={{ color: MUTED }}>{fichaFile.name} ({(fichaFile.size / 1024).toFixed(0)}KB)</p>}

      {producto?.fichaTecnicaNombre && !fichaFile && (
        <p className="text-xs mb-2" style={{ color: MUTED }}>Ficha técnica actual: {producto.fichaTecnicaNombre} (se conserva si no subís una nueva)</p>
      )}

      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit} disabled={guardando}>
        {guardando ? "Guardando..." : producto ? "Guardar cambios" : "Guardar producto"}
      </PrimaryButton>
    </div>
  );
}

function ProductoCard({ p, onEdit, onDelete, onQuitarFicha }) {
  const tieneCosto = p.costoOrigen || p.costoPy;
  return (
    <div className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <div className="flex items-start gap-3 mb-2">
        {p.foto ? (
          <img src={p.foto} alt={p.nombre} className="rounded border shrink-0" style={{ width: 56, height: 56, objectFit: "contain", borderColor: BORDER, backgroundColor: "#FAFBFC" }} />
        ) : (
          <div className="rounded border shrink-0 flex items-center justify-center" style={{ width: 56, height: 56, borderColor: BORDER, backgroundColor: "#FAFBFC" }}>
            <Tag size={20} style={{ color: MUTED }} />
          </div>
        )}
        <div className="min-w-0 flex-1">
          <div className="flex items-start justify-between gap-2">
            <CodeTag>{p.nombre}</CodeTag>
            <div className="flex items-center gap-1 shrink-0">
              <button onClick={() => onEdit(p)} className="p-1 rounded hover:bg-gray-100" title="Editar producto">
                <Pencil size={13} style={{ color: MUTED }} />
              </button>
              <button onClick={() => onDelete(p)} className="p-1 rounded hover:bg-gray-100" title="Eliminar producto">
                <Trash2 size={13} style={{ color: MUTED }} />
              </button>
            </div>
          </div>
          {p.categoria && <p className="text-xs mt-1" style={{ color: MUTED }}>{p.categoria}</p>}
        </div>
      </div>
      <div className="flex items-center gap-2">
        <p className="text-sm font-medium" style={{ color: INK }}>U$S {Number(p.precioLista || 0).toLocaleString()}</p>
        {p.stockDisponible != null && (
          <span className="text-xs px-1.5 py-0.5 rounded" style={{ backgroundColor: p.stockDisponible > 0 ? "#E9F7EF" : "#FBEAEA", color: p.stockDisponible > 0 ? "#15803D" : "#B91C1C" }}>
            Stock: {p.stockDisponible}
          </span>
        )}
      </div>
      {p.especLabel && <p className="text-xs mt-0.5" style={{ color: MUTED }}>{p.especLabel}: {p.especValor}</p>}
      {p.descripcion && <p className="text-xs mt-1.5 line-clamp-2" style={{ color: MUTED }}>{p.descripcion}</p>}

      {(tieneCosto || p.contenedorTipo) && (
        <div className="mt-2 p-2 rounded" style={{ backgroundColor: "#F7F8FA" }}>
          <p className="text-[10px] font-semibold uppercase tracking-wide mb-1" style={{ color: MUTED }}>Info interna — no aparece en la cotización</p>
          {tieneCosto && (
            <p className="text-xs" style={{ color: MUTED }}>
              Origen U$S {Number(p.costoOrigen || 0).toLocaleString()} · Puesto en PY U$S {Number(p.costoPy || 0).toLocaleString()}
            </p>
          )}
          {p.contenedorTipo && (
            <p className="text-xs" style={{ color: MUTED }}>
              Contenedor {p.contenedorTipo}{p.contenedorCantidad ? ` · ${p.contenedorCantidad} un.` : ""}
            </p>
          )}
        </div>
      )}

      <div className="mt-2.5 pt-2.5 border-t flex items-center justify-between" style={{ borderColor: BORDER }}>
        {p.fichaTecnicaData ? (
          <div className="flex items-center gap-1.5">
            <a href={p.fichaTecnicaData} target="_blank" rel="noreferrer" className="text-xs flex items-center gap-1" style={{ color: ACCENT }}>
              <FileText size={13} /> Ficha técnica
            </a>
            <button onClick={() => onQuitarFicha(p)} title="Quitar ficha técnica">
              <X size={12} style={{ color: MUTED }} />
            </button>
          </div>
        ) : (
          <span className="text-xs" style={{ color: MUTED }}>Sin ficha técnica</span>
        )}
      </div>
    </div>
  );
}

function ordenarProductos(productos) {
  return [...productos].sort((a, b) => {
    const oa = a.ordenNumerico, ob = b.ordenNumerico;
    if (oa != null && ob != null && oa !== ob) return oa - ob;
    if (oa != null && ob == null) return -1;
    if (oa == null && ob != null) return 1;
    return (a.nombre || "").localeCompare(b.nombre || "");
  });
}

const NIVELES_CATEGORIA_PRODUCTO = [
  (p) => p.categoriaPrincipal,
  (p) => p.subcategoria,
  (p) => p.subcategoria2,
  (p) => p.subcategoria3,
];

// Orden manual pedido para cada pestaña del catálogo — lo que no figure acá cae al final,
// ordenado alfabéticamente, así un valor nuevo nunca desaparece del catálogo.
function ordenarClaves(claves, ordenPersonalizado) {
  if (!ordenPersonalizado) return [...claves].sort((a, b) => a.localeCompare(b));
  const idx = (v) => {
    const i = ordenPersonalizado.indexOf(v);
    return i === -1 ? ordenPersonalizado.length : i;
  };
  return [...claves].sort((a, b) => {
    const diff = idx(a) - idx(b);
    return diff !== 0 ? diff : a.localeCompare(b);
  });
}

// Árbol de categorías: cada nodo tiene `productos` (hoja, ya ordenados) o `hijos` (subgrupos).
// Productos sin valor en un nivel quedan agrupados como "Otros" en ese mismo nivel, en vez de
// perderse — así un catálogo cargado a medias sigue siendo navegable. `ordenesPorNivel` es un
// objeto { [nivel]: ["valor1", "valor2", ...] } opcional para forzar un orden manual en vez del
// alfabético por default.
function construirArbolCategorias(productos, nivel = 0, ordenesPorNivel = null) {
  if (nivel >= NIVELES_CATEGORIA_PRODUCTO.length) return { productos: ordenarProductos(productos), hijos: null };
  const get = NIVELES_CATEGORIA_PRODUCTO[nivel];
  const grupos = new Map();
  const sinValor = [];
  for (const p of productos) {
    const key = (get(p) || "").trim();
    if (!key) { sinValor.push(p); continue; }
    if (!grupos.has(key)) grupos.set(key, []);
    grupos.get(key).push(p);
  }
  if (grupos.size === 0) return { productos: ordenarProductos(productos), hijos: null };
  const claves = ordenarClaves([...grupos.keys()], ordenesPorNivel && ordenesPorNivel[nivel]);
  const hijos = claves.map((valor) => ({ valor, ...construirArbolCategorias(grupos.get(valor), nivel + 1, ordenesPorNivel) }));
  if (sinValor.length > 0) hijos.push({ valor: "Otros", ...construirArbolCategorias(sinValor, NIVELES_CATEGORIA_PRODUCTO.length) });
  return { productos: null, hijos };
}

// Pestañas del catálogo de productos (no incluye Repuestos, que tiene su propio modo aparte).
// `nivelInicio` salta los niveles ya implícitos en la pestaña (categoriaPrincipal, y para las
// de Cocina también subcategoria) para no repetir un título redundante con el nombre de la pestaña.
const CATALOGO_TABS = [
  {
    key: "aires", label: "Aire Acondicionado", nivelInicio: 1,
    filtro: (p) => p.categoriaPrincipal === "Aire Acondicionado",
    ordenesPorNivel: { 1: ["Split Pared", "Cassette", "Piso-Techo", "Ducto", "Multi Split Interior", "Multi Split Exterior"] },
  },
  {
    key: "anafes", label: "Anafes", nivelInicio: 2,
    filtro: (p) => p.categoriaPrincipal === "Cocina" && p.subcategoria === "Anafe",
    ordenesPorNivel: { 2: ["Vitrocerámica", "Inducción", "Combinado"] },
  },
  {
    key: "campanas", label: "Campanas", nivelInicio: 2,
    filtro: (p) => p.categoriaPrincipal === "Cocina" && p.subcategoria === "Campana",
    ordenesPorNivel: { 2: ["Pared", "Isla"] },
  },
  {
    key: "hornos", label: "Hornos", nivelInicio: 2,
    filtro: (p) => p.categoriaPrincipal === "Cocina" && p.subcategoria === "Horno",
    ordenesPorNivel: { 2: ["Mecánico", "Semi-Digital", "Digital"] },
  },
  {
    key: "termocalefones", label: "Termocalefones", nivelInicio: 1,
    filtro: (p) => p.categoriaPrincipal === "Termocalefones",
    ordenesPorNivel: null,
  },
];

const CATEGORIA_TITULO_CLASE = ["text-base font-semibold", "text-sm font-semibold", "text-xs font-semibold", "text-xs font-medium"];

function CategoriaNodo({ nodo, nivel, onEdit, onDelete, onQuitarFicha }) {
  if (nodo.productos) {
    return (
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
        {nodo.productos.map((p) => (
          <ProductoCard key={p.id} p={p} onEdit={onEdit} onDelete={onDelete} onQuitarFicha={onQuitarFicha} />
        ))}
      </div>
    );
  }
  return (
    <div className={nivel === 0 ? "space-y-5" : "space-y-3 pl-3 border-l-2"} style={nivel === 0 ? {} : { borderColor: BORDER }}>
      {nodo.hijos.map((hijo) => (
        <div key={hijo.valor}>
          <p className={CATEGORIA_TITULO_CLASE[Math.min(nivel, 3)]} style={{ color: nivel === 0 ? INK : MUTED }}>{hijo.valor}</p>
          <div className="mt-2">
            <CategoriaNodo nodo={hijo} nivel={nivel + 1} onEdit={onEdit} onDelete={onDelete} onQuitarFicha={onQuitarFicha} />
          </div>
        </div>
      ))}
    </div>
  );
}

function CatalogoView({ productos, query, onQuery, onNew, onEdit, onDelete, onQuitarFicha, onImportar, importando, importResultado, modoInicial }) {
  const fileInputRef = useRef(null);
  const [modo, setModo] = useState(modoInicial === "repuestos" ? "repuestos" : "productos"); // "productos" | "repuestos" — carpetas totalmente separadas
  const [catTab, setCatTab] = useState(CATALOGO_TABS[0].key);

  const productosFiltrados = useMemo(
    () => productos.filter((p) => (p.categoriaPrincipal === "Repuestos") === (modo === "repuestos")),
    [productos, modo]
  );

  const buscando = query.trim().length > 0;

  // Todo lo que no cae en ninguna pestaña conocida (ej. Enfriador de vinos, o una categoría
  // nueva que todavía no tiene su propia pestaña) — para que nunca desaparezca del catálogo.
  const otrosProductos = useMemo(() => {
    if (modo !== "productos") return [];
    return productosFiltrados.filter((p) => !CATALOGO_TABS.some((t) => t.filtro(p)));
  }, [productosFiltrados, modo]);

  const tabs = useMemo(() => {
    if (modo !== "productos") return [];
    const tabs = CATALOGO_TABS.map((t) => ({ key: t.key, label: t.label, count: productosFiltrados.filter(t.filtro).length }));
    if (otrosProductos.length > 0) tabs.push({ key: "otros", label: "Otros", count: otrosProductos.length });
    return tabs;
  }, [productosFiltrados, otrosProductos, modo]);

  const tabActivo = CATALOGO_TABS.find((t) => t.key === catTab);
  const hayProductosEnTab = catTab === "otros" ? otrosProductos.length > 0 : (tabActivo ? productosFiltrados.some(tabActivo.filtro) : false);

  const arbol = useMemo(() => {
    if (modo === "repuestos") {
      const a = construirArbolCategorias(productosFiltrados);
      // En modo repuestos todos comparten categoriaPrincipal="Repuestos" — ese nivel es redundante
      // con el selector de arriba, así que se muestra directamente su contenido.
      if (a.hijos && a.hijos.length === 1 && a.hijos[0].valor === "Repuestos") return a.hijos[0];
      return a;
    }
    // Buscando: se muestra todo el catálogo sin recortar por pestaña, para no esconder resultados.
    if (buscando) return construirArbolCategorias(productosFiltrados);
    if (catTab === "otros") return construirArbolCategorias(otrosProductos);
    if (!tabActivo) return { productos: [], hijos: null };
    return construirArbolCategorias(productosFiltrados.filter(tabActivo.filtro), tabActivo.nivelInicio, tabActivo.ordenesPorNivel);
  }, [modo, productosFiltrados, buscando, catTab, otrosProductos, tabActivo]);

  const handleFileChange = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (file) onImportar(file);
  };

  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Catálogo de productos</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>
            Precio de lista y ficha técnica de cada modelo — de acá sale el precio sugerido en ventas comprometidas y el contenido de las cotizaciones.
          </p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <SearchBox value={query} onChange={onQuery} />
          <SecondaryButton onClick={descargarPlantillaCatalogo}><Download size={14} /> Plantilla</SecondaryButton>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileChange} className="hidden" />
          <SecondaryButton onClick={() => fileInputRef.current?.click()}>
            <Upload size={14} /> {importando ? "Importando..." : "Importar Excel"}
          </SecondaryButton>
          <PrimaryButton onClick={() => onNew(modo)}><Plus size={15} /> {modo === "repuestos" ? "Nuevo repuesto" : "Nuevo producto"}</PrimaryButton>
        </div>
      </div>

      <div className="flex gap-2 mb-4">
        {[{ key: "productos", label: "Productos" }, { key: "repuestos", label: "Repuestos" }].map((op) => (
          <button
            key={op.key}
            onClick={() => setModo(op.key)}
            className="text-sm px-3.5 py-1.5 rounded-md font-medium"
            style={modo === op.key
              ? { backgroundColor: ACCENT, color: "#FFFFFF" }
              : { backgroundColor: "#FFFFFF", color: MUTED, border: `0.5px solid ${BORDER}` }}
          >
            {op.label}
          </button>
        ))}
      </div>

      {importResultado && (
        <div className="mb-4 px-3 py-2 rounded-lg text-sm" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>{importResultado}</div>
      )}

      {modo === "productos" && !buscando && tabs.length > 0 && (
        <div className="flex gap-1.5 mb-4 flex-wrap">
          {tabs.map((t) => (
            <button
              key={t.key}
              onClick={() => setCatTab(t.key)}
              className="text-xs px-3 py-1.5 rounded-full font-medium"
              style={catTab === t.key
                ? { backgroundColor: ACCENT, color: "#FFFFFF" }
                : { backgroundColor: "#F2F3F4", color: MUTED }}
            >
              {t.label} ({t.count})
            </button>
          ))}
        </div>
      )}

      {modo === "productos" && buscando && (
        <p className="text-xs mb-3" style={{ color: MUTED }}>Resultados de búsqueda en todo el catálogo de productos.</p>
      )}

      {productosFiltrados.length === 0 ? (
        <EmptyState
          icon={Tag}
          title={modo === "repuestos" ? "Todavía no hay repuestos cargados" : "Todavía no hay productos cargados"}
          subtitle="Usá el botón de arriba para cargar el primero."
        />
      ) : modo === "productos" && !buscando && !hayProductosEnTab ? (
        <EmptyState icon={Tag} title="Sin productos en esta categoría" subtitle="Elegí otra pestaña o cargá un producto nuevo acá." />
      ) : (
        <CategoriaNodo nodo={arbol} nivel={0} onEdit={onEdit} onDelete={onDelete} onQuitarFicha={onQuitarFicha} />
      )}
    </div>
  );
}

// ---------- Cotizaciones ----------
const FECHA_ENTREGA_DEFAULT = "Una vez aprobado el presupuesto la entrega se concreta de 150 a 200 dias";
const OBS_DEFAULT = "Productos a retirar de depósito.";

function CotizacionForm({ productos, onSave }) {
  const [fecha, setFecha] = useState(todayISO());
  const [cliente, setCliente] = useState("");
  const [obra, setObra] = useState("");
  const [clienteReal, setClienteReal] = useState("");
  const [categoria, setCategoria] = useState("");
  const [comentarios, setComentarios] = useState("");
  const [incluirDescuento, setIncluirDescuento] = useState(false);
  const [descuento, setDescuento] = useState("");
  const [incluirInstalacion, setIncluirInstalacion] = useState(false);
  const [instalacionDescripcion, setInstalacionDescripcion] = useState("Instalación de equipos");
  const [instalacionMonto, setInstalacionMonto] = useState("");
  const [fechaEntregaEstimada, setFechaEntregaEstimada] = useState(FECHA_ENTREGA_DEFAULT);
  const [formaPago, setFormaPago] = useState("A conversar");
  const [obs, setObs] = useState(OBS_DEFAULT);
  const [lineas, setLineas] = useState([]);
  const [productoId, setProductoId] = useState("");
  const [cantidadNueva, setCantidadNueva] = useState(1);
  const [precioNuevo, setPrecioNuevo] = useState("");
  const [error, setError] = useState("");

  const productoSel = productos.find((p) => p.id === productoId);
  const subtotal = lineas.reduce((acc, l) => acc + (Number(l.cantidad) || 0) * (Number(l.precioUnit) || 0), 0);

  // Agrupa el selector de productos por categoría (categoriaPrincipal > subcategoria > ...)
  // para que un catálogo grande siga siendo navegable en vez de una lista plana larguísima.
  const productosPorGrupo = useMemo(() => {
    const grupos = new Map();
    for (const p of productos) {
      const path = [p.categoriaPrincipal, p.subcategoria, p.subcategoria2, p.subcategoria3].filter((v) => (v || "").trim()).join(" — ");
      const key = path || "Otros";
      if (!grupos.has(key)) grupos.set(key, []);
      grupos.get(key).push(p);
    }
    return [...grupos.entries()].sort(([a], [b]) => a.localeCompare(b));
  }, [productos]);

  const handleProducto = (id) => {
    setProductoId(id);
    const p = productos.find((x) => x.id === id);
    setCantidadNueva(1);
    setPrecioNuevo(p ? String(Number(p.precioLista) || 0) : "");
  };

  const agregarLinea = () => {
    if (!productoSel) {
      setError("Elegí un producto del catálogo.");
      return;
    }
    setLineas([...lineas, {
      codigo: productoSel.nombre, descripcion: productoSel.descripcion, foto: productoSel.foto,
      especLabel: productoSel.especLabel, especValor: productoSel.especValor,
      fichaTecnicaData: productoSel.fichaTecnicaData || "",
      cantidad: Number(cantidadNueva) || 1, precioUnit: Number(precioNuevo) || 0,
    }]);
    setProductoId("");
    setCantidadNueva(1);
    setPrecioNuevo("");
    setError("");
  };

  const quitarLinea = (idx) => setLineas(lineas.filter((_, i) => i !== idx));

  const submit = () => {
    if (!cliente.trim()) {
      setError("Ingresá el cliente.");
      return;
    }
    if (lineas.length === 0) {
      setError("Agregá al menos un producto.");
      return;
    }
    onSave({
      fecha, cliente, obra, categoria, comentarios, lineas,
      incluirDescuento, descuento: Number(descuento) || 0, descuentoEsPorcentaje: true,
      incluirInstalacion, instalacionDescripcion, instalacionMonto: Number(instalacionMonto) || 0,
      fechaEntregaEstimada, formaPago, obs,
      clienteReal, estado: "Pendiente",
    });
  };

  return (
    <div>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Cliente (constructora/desarrolladora — sale en el PDF)"><TextInput value={cliente} onChange={(e) => setCliente(e.target.value)} /></Field>
      <Field label="Obra"><TextInput value={obra} onChange={(e) => setObra(e.target.value)} placeholder="Usá el mismo nombre si es una obra que ya cotizaste" /></Field>
      <Field label="Categoría (título de la cotización)"><TextInput value={categoria} onChange={(e) => setCategoria(e.target.value)} placeholder="Ej: AIRES ACONDICIONADOS" /></Field>

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Datos internos (no aparecen en el PDF)</p>
      <Field label="Cliente real / inversor"><TextInput value={clienteReal} onChange={(e) => setClienteReal(e.target.value)} placeholder="Ej: Pepe Gómez" /></Field>

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Productos</p>
      <div className="p-2.5 rounded mb-3" style={{ backgroundColor: "#F7F8FA" }}>
        <Field label="Producto del catálogo">
          <Select value={productoId} onChange={(e) => handleProducto(e.target.value)}>
            <option value="">Seleccionar...</option>
            {productosPorGrupo.map(([grupo, items]) => (
              <optgroup key={grupo} label={grupo}>
                {items.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
              </optgroup>
            ))}
          </Select>
        </Field>
        {productoSel && (
          <>
            <div className="flex gap-2">
              <Field label="Cantidad"><TextInput type="number" min="1" value={cantidadNueva} onChange={(e) => setCantidadNueva(e.target.value)} /></Field>
              <Field label="Precio Unit. U$S"><TextInput type="number" value={precioNuevo} onChange={(e) => setPrecioNuevo(e.target.value)} /></Field>
            </div>
            <SecondaryButton onClick={agregarLinea}><Plus size={14} /> Agregar a la cotización</SecondaryButton>
          </>
        )}
      </div>

      {lineas.length > 0 && (
        <div className="mb-3 rounded border overflow-hidden" style={{ borderColor: BORDER }}>
          {lineas.map((l, i) => (
            <div key={i} className="flex items-center justify-between px-2.5 py-2 text-xs border-b last:border-0" style={{ borderColor: BORDER }}>
              <div className="min-w-0">
                <span className="font-medium" style={{ color: INK }}>{l.codigo}</span>
                <span style={{ color: MUTED }}> · cant. {l.cantidad} × U$S {Number(l.precioUnit).toLocaleString()} = U$S {(l.cantidad * l.precioUnit).toLocaleString()}</span>
              </div>
              <button onClick={() => quitarLinea(i)}><X size={13} style={{ color: MUTED }} /></button>
            </div>
          ))}
          <div className="px-2.5 py-2 text-xs font-semibold flex justify-between" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>
            <span>Subtotal</span><span>U$S {subtotal.toLocaleString()}</span>
          </div>
        </div>
      )}

      <label className="flex items-center gap-2 mb-3 text-sm" style={{ color: INK }}>
        <input type="checkbox" checked={incluirDescuento} onChange={(e) => setIncluirDescuento(e.target.checked)} />
        Incluir descuento
      </label>
      {incluirDescuento && (
        <Field label="Descuento %"><TextInput type="number" value={descuento} onChange={(e) => setDescuento(e.target.value)} placeholder="Ej: 10" /></Field>
      )}

      <label className="flex items-center gap-2 mb-3 text-sm" style={{ color: INK }}>
        <input type="checkbox" checked={incluirInstalacion} onChange={(e) => setIncluirInstalacion(e.target.checked)} />
        Incluir instalación
      </label>
      {incluirInstalacion && (
        <div className="flex gap-2">
          <Field label="Instalación — descripción"><TextInput value={instalacionDescripcion} onChange={(e) => setInstalacionDescripcion(e.target.value)} placeholder="Ej: Instalación de equipos" /></Field>
          <Field label="Instalación — monto U$S"><TextInput type="number" value={instalacionMonto} onChange={(e) => setInstalacionMonto(e.target.value)} placeholder="0" /></Field>
        </div>
      )}
      {lineas.length > 0 && (
        <div className="px-2.5 py-2 mb-3 text-sm font-semibold flex justify-between rounded" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>
          <span>Total final</span>
          <span>U$S {(subtotal - (incluirDescuento ? subtotal * (Number(descuento) || 0) / 100 : 0) + (incluirInstalacion ? Number(instalacionMonto) || 0 : 0)).toLocaleString()}</span>
        </div>
      )}

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Datos del PDF</p>
      <Field label="Comentarios (opcional)"><TextInput value={comentarios} onChange={(e) => setComentarios(e.target.value)} /></Field>
      <Field label="Fecha de entrega estimada"><TextInput value={fechaEntregaEstimada} onChange={(e) => setFechaEntregaEstimada(e.target.value)} /></Field>
      <Field label="Forma de pago"><TextInput value={formaPago} onChange={(e) => setFormaPago(e.target.value)} /></Field>
      <Field label="Observaciones"><TextInput value={obs} onChange={(e) => setObs(e.target.value)} /></Field>

      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar cotización</PrimaryButton>
    </div>
  );
}

function ResumenCotizaciones({ resumen }) {
  return (
    <div className="mb-4 grid grid-cols-2 sm:grid-cols-4 gap-2">
      <div className="px-3 py-2 rounded-lg" style={{ backgroundColor: ACCENT_LIGHT }}>
        <p className="text-xs" style={{ color: ACCENT }}>Total cotizado</p>
        <p className="text-sm font-semibold" style={{ color: ACCENT }}>U$S {resumen.total.toLocaleString()}</p>
      </div>
      {ESTADOS_COTIZACION.map((estado) => {
        const badge = ESTADO_COTIZACION_BADGE[estado];
        return (
          <div key={estado} className="px-3 py-2 rounded-lg" style={{ backgroundColor: badge.bg }}>
            <p className="text-xs" style={{ color: badge.color }}>{estado} ({resumen[estado].n})</p>
            <p className="text-sm font-semibold" style={{ color: badge.color }}>U$S {resumen[estado].total.toLocaleString()}</p>
          </div>
        );
      })}
    </div>
  );
}

const SALIDA_COTIZACION_BADGE = {
  completa: { label: "Salida completa", bg: "#E8F3EC", color: "#1F7A44" },
  parcial: { label: "Salida parcial", bg: "#FEF3E2", color: "#B45309" },
  pendiente: { label: "Sin salida", bg: "#F2F3F4", color: "#686D73" },
};

function CotizacionCard({ c, esActiva, onDelete, onUpdate, onDescargarPdf, onDescargarFichas, descargandoId }) {
  const total = calcularTotalCotizacion(c);
  const tieneFichas = (c.lineas || []).some((l) => l.fichaTecnicaData);
  const estado = ESTADOS_COTIZACION.includes(c.estado) ? c.estado : "Pendiente";
  const estadoSalida = estadoSalidaCotizacion(c);
  const badgeSalida = estadoSalida ? SALIDA_COTIZACION_BADGE[estadoSalida] : null;
  return (
    <div className="rounded-lg p-3.5" style={{ backgroundColor: esActiva ? "#FFFFFF" : "#FAFBFC", border: `0.5px solid ${BORDER}` }}>
      <div className="flex items-start justify-between mb-1">
        <p className="text-xs" style={{ color: MUTED }}>{fmtDate(c.fecha)}</p>
        <button onClick={() => onDelete(c.id)} className="p-1 rounded hover:bg-gray-100">
          <Trash2 size={13} style={{ color: MUTED }} />
        </button>
      </div>
      {c.categoria && <CodeTag>{c.categoria}</CodeTag>}
      <div className="flex items-center gap-1.5 mt-2 flex-wrap">
        <p className="text-sm" style={{ color: INK }}>{(c.lineas || []).length} producto(s) · U$S {total.toLocaleString()}</p>
        {badgeSalida && (
          <span className="text-[10px] px-1.5 py-0.5 rounded-full font-medium" style={{ backgroundColor: badgeSalida.bg, color: badgeSalida.color }}>
            {badgeSalida.label}
          </span>
        )}
      </div>

      {esActiva && (
        <div className="mt-2 p-2 rounded" style={{ backgroundColor: "#F7F8FA" }}>
          <p className="text-[10px] font-semibold uppercase tracking-wide mb-1" style={{ color: MUTED }}>Info interna — no aparece en el PDF</p>
          <div className="flex items-center gap-2 flex-wrap">
            <div style={{ width: 110 }}>
              <Select value={estado} onChange={(e) => onUpdate(c.id, { estado: e.target.value })}>
                {ESTADOS_COTIZACION.map((op) => <option key={op} value={op}>{op}</option>)}
              </Select>
            </div>
            <div className="flex-1 min-w-[140px]">
              <ComentarioEditor value={c.clienteReal} onSave={(v) => onUpdate(c.id, { clienteReal: v })} placeholder="Cliente real / inversor" />
            </div>
          </div>
        </div>
      )}

      <div className="flex gap-2 mt-3 flex-wrap">
        <button
          onClick={() => onDescargarPdf(c)}
          disabled={descargandoId === `${c.id}:pdf`}
          className="text-xs px-2.5 py-1.5 rounded flex items-center gap-1"
          style={{ backgroundColor: ACCENT, color: "#FFFFFF", opacity: descargandoId === `${c.id}:pdf` ? 0.6 : 1 }}
        >
          <Download size={13} /> {descargandoId === `${c.id}:pdf` ? "Generando..." : "Descargar PDF"}
        </button>
        {tieneFichas && (
          <button
            onClick={() => onDescargarFichas(c)}
            disabled={descargandoId === `${c.id}:fichas`}
            className="text-xs px-2.5 py-1.5 rounded border flex items-center gap-1"
            style={{ borderColor: BORDER, color: INK, opacity: descargandoId === `${c.id}:fichas` ? 0.6 : 1 }}
          >
            <FileText size={13} /> {descargandoId === `${c.id}:fichas` ? "Generando..." : "Fichas técnicas"}
          </button>
        )}
      </div>
    </div>
  );
}

function ObraGrupo({ grupo, onDelete, onUpdate, onDescargarPdf, onDescargarFichas, descargandoId }) {
  const [expandido, setExpandido] = useState(false);
  const historial = grupo.versiones.slice(1);
  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <p className="text-xs font-medium" style={{ color: INK }}>{grupo.obra}</p>
        {historial.length > 0 && (
          <button onClick={() => setExpandido(!expandido)} className="text-xs" style={{ color: ACCENT }}>
            {expandido ? "Ocultar historial" : `Ver historial (${historial.length})`}
          </button>
        )}
      </div>
      <CotizacionCard
        c={grupo.activa} esActiva
        onDelete={onDelete} onUpdate={onUpdate}
        onDescargarPdf={onDescargarPdf} onDescargarFichas={onDescargarFichas} descargandoId={descargandoId}
      />
      {expandido && (
        <div className="mt-2 pl-3 border-l-2 space-y-2" style={{ borderColor: BORDER }}>
          {historial.map((v) => (
            <CotizacionCard
              key={v.id} c={v} esActiva={false}
              onDelete={onDelete} onUpdate={onUpdate}
              onDescargarPdf={onDescargarPdf} onDescargarFichas={onDescargarFichas} descargandoId={descargandoId}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function ClienteGrupo({ grupo, onDelete, onUpdate, onDescargarPdf, onDescargarFichas, descargandoId }) {
  const resumen = useMemo(() => resumirCotizaciones([grupo]), [grupo]);
  return (
    <div className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
      <div className="flex items-start justify-between mb-3 gap-3 flex-wrap">
        <p className="text-sm font-semibold" style={{ color: INK }}>{grupo.cliente}</p>
        <div className="flex items-center gap-3 text-xs flex-wrap">
          <span style={{ color: ACCENT }}>Cotizado U$S {resumen.total.toLocaleString()}</span>
          {ESTADOS_COTIZACION.map((estado) => (
            <span key={estado} style={{ color: ESTADO_COTIZACION_BADGE[estado].color }}>
              {estado}: {resumen[estado].n}
            </span>
          ))}
        </div>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        {grupo.obras.map((o) => (
          <ObraGrupo
            key={o.obra} grupo={o}
            onDelete={onDelete} onUpdate={onUpdate}
            onDescargarPdf={onDescargarPdf} onDescargarFichas={onDescargarFichas} descargandoId={descargandoId}
          />
        ))}
      </div>
    </div>
  );
}

function CotizacionesView({ cotizaciones, query, onQuery, onNew, onDelete, onUpdate, onDescargarPdf, onDescargarFichas, descargandoId, pdfError }) {
  const grupos = useMemo(() => agruparCotizaciones(cotizaciones), [cotizaciones]);
  const resumen = useMemo(() => resumirCotizaciones(grupos), [grupos]);

  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Cotizaciones</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>Armá una cotización con productos del catálogo y descargala en PDF, con tu firma. Agrupadas por cliente y obra — una obra que volvés a cotizar queda como nueva versión de la misma, con historial.</p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
          <PrimaryButton onClick={onNew}><Plus size={15} /> Nueva cotización</PrimaryButton>
        </div>
      </div>

      {pdfError && (
        <div className="mb-4 px-3 py-2 rounded-lg text-sm" style={{ backgroundColor: "#FBEAEA", color: "#B91C1C" }}>{pdfError}</div>
      )}

      {cotizaciones.length === 0 ? (
        <EmptyState icon={FileSignature} title="Todavía no hay cotizaciones" subtitle="Usá el botón de arriba para armar la primera." />
      ) : (
        <>
          <ResumenCotizaciones resumen={resumen} />
          <div className="space-y-3">
            {grupos.map((g) => (
              <ClienteGrupo
                key={g.cliente} grupo={g}
                onDelete={onDelete} onUpdate={onUpdate}
                onDescargarPdf={onDescargarPdf} onDescargarFichas={onDescargarFichas} descargandoId={descargandoId}
              />
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// ---------- Presupuestos de reparación / repuestos ----------
// Vía totalmente separada de Cotizaciones: se arma con productos de la carpeta
// "Repuestos" del catálogo y genera un PDF con el modelo "Presupuesto de Reparación".
const MANTENIMIENTO_CATEGORIAS = ["Aire Acondicionado", "Termocalefón", "Campana"];

const MANTENIMIENTO_PRECIOS = {
  "Aire Acondicionado": {
    conBtu: true,
    opciones: [
      { label: "9.000 - 12.000 BTU", precio: 300000 },
      { label: "18.000 - 36.000 BTU", precio: 420000 },
      { label: "48.000 - 64.000 BTU", precio: 576000 },
    ],
    descripcion: "Mantenimiento preventivo: limpieza de filtros, serpentinas y chequeo general.",
  },
  "Termocalefón": {
    conBtu: false,
    precio: 360000,
    descripcion: "Desarme, limpieza con hidrolavadora, cambio de vela de magnesio (incluido), chequeo de resistencia y montaje.",
  },
  "Campana": {
    conBtu: false,
    precio: 300000,
    descripcion: "Limpieza de filtro, turbina, campana y ducto. De ser necesario, cambio de filtro con costo adicional.",
  },
};

const PLAZO_ESTIMADO_DEFAULT = {
  reparacion: "A confirmar según disponibilidad de repuesto y agenda del técnico.",
  mantenimiento: "A coordinar según disponibilidad de agenda del técnico.",
};

function PresupuestoReparacionForm({ productos, onSave }) {
  const [tipo, setTipo] = useState("reparacion");
  const [fecha, setFecha] = useState(todayISO());
  const [cliente, setCliente] = useState("");
  const [obra, setObra] = useState("");
  const [equipoAfectado, setEquipoAfectado] = useState("");
  const [fallaReportada, setFallaReportada] = useState("");
  const [lineas, setLineas] = useState([]);
  const [productoId, setProductoId] = useState("");
  const [cantidadNueva, setCantidadNueva] = useState(1);
  const [costoNuevo, setCostoNuevo] = useState("");
  const [incluirInstalacion, setIncluirInstalacion] = useState(false);
  const [instalacionMonto, setInstalacionMonto] = useState("");
  const [plazoEstimado, setPlazoEstimado] = useState(PLAZO_ESTIMADO_DEFAULT.reparacion);
  const [error, setError] = useState("");

  const [lineasMantenimiento, setLineasMantenimiento] = useState([]);
  const [categoriaActiva, setCategoriaActiva] = useState(MANTENIMIENTO_CATEGORIAS[0]);
  const [btuSel, setBtuSel] = useState(0);
  const [cantidadMtto, setCantidadMtto] = useState(1);
  const [precioMtto, setPrecioMtto] = useState(String(MANTENIMIENTO_PRECIOS[MANTENIMIENTO_CATEGORIAS[0]].opciones[0].precio));

  const productoSel = productos.find((p) => p.id === productoId);
  const subtotal = lineas.reduce((acc, l) => acc + (Number(l.cantidad) || 0) * (Number(l.costoUnitario) || 0), 0);
  const subtotalMtto = lineasMantenimiento.reduce((acc, l) => acc + (Number(l.cantidad) || 0) * (Number(l.precioUnitario) || 0), 0);
  const mttoInfo = MANTENIMIENTO_PRECIOS[categoriaActiva];

  const productosPorGrupo = useMemo(() => {
    const grupos = new Map();
    for (const p of productos) {
      const key = p.subcategoria || "Otros";
      if (!grupos.has(key)) grupos.set(key, []);
      grupos.get(key).push(p);
    }
    return [...grupos.entries()].sort(([a], [b]) => a.localeCompare(b));
  }, [productos]);

  const handleTipo = (t) => {
    setTipo(t);
    if (Object.values(PLAZO_ESTIMADO_DEFAULT).includes(plazoEstimado)) setPlazoEstimado(PLAZO_ESTIMADO_DEFAULT[t]);
  };

  const handleProducto = (id) => {
    setProductoId(id);
    const p = productos.find((x) => x.id === id);
    setCantidadNueva(1);
    setCostoNuevo(p ? String(Number(p.precioLista) || 0) : "");
  };

  const agregarLinea = () => {
    if (!productoSel) {
      setError("Elegí un repuesto del catálogo.");
      return;
    }
    setLineas([...lineas, {
      item: productoSel.descripcion || productoSel.nombre,
      descripcion: productoSel.descripcion || "",
      modeloEquipo: productoSel.subcategoria2 || "",
      cantidad: Number(cantidadNueva) || 1, costoUnitario: Number(costoNuevo) || 0,
    }]);
    setProductoId("");
    setCantidadNueva(1);
    setCostoNuevo("");
    setError("");
  };

  const quitarLinea = (idx) => setLineas(lineas.filter((_, i) => i !== idx));

  const handleCategoriaActiva = (cat) => {
    setCategoriaActiva(cat);
    const info = MANTENIMIENTO_PRECIOS[cat];
    setCantidadMtto(1);
    setBtuSel(0);
    setPrecioMtto(String(info.conBtu ? info.opciones[0].precio : info.precio));
  };

  const handleBtuSel = (idx) => {
    setBtuSel(idx);
    setPrecioMtto(String(mttoInfo.opciones[idx].precio));
  };

  const agregarLineaMtto = () => {
    const detalle = mttoInfo.conBtu ? mttoInfo.opciones[btuSel].label : mttoInfo.descripcion;
    setLineasMantenimiento([...lineasMantenimiento, {
      categoria: categoriaActiva, detalle,
      cantidad: Number(cantidadMtto) || 1, precioUnitario: Number(precioMtto) || 0,
    }]);
    setCantidadMtto(1);
  };

  const quitarLineaMtto = (idx) => setLineasMantenimiento(lineasMantenimiento.filter((_, i) => i !== idx));

  const submit = () => {
    if (!cliente.trim()) {
      setError("Ingresá el cliente.");
      return;
    }
    if (tipo === "reparacion" && lineas.length === 0) {
      setError("Agregá al menos un repuesto.");
      return;
    }
    if (tipo === "mantenimiento" && lineasMantenimiento.length === 0 && lineas.length === 0) {
      setError("Agregá al menos un servicio de mantenimiento o un repuesto.");
      return;
    }
    onSave({
      tipo, fecha, cliente, obra, equipoAfectado, fallaReportada, lineas,
      lineasMantenimiento: tipo === "mantenimiento" ? lineasMantenimiento : [],
      incluirInstalacion: tipo === "reparacion" && incluirInstalacion,
      instalacionMonto: Number(instalacionMonto) || 0,
      plazoEstimado,
    });
  };

  return (
    <div>
      <div className="flex gap-2 mb-4">
        <button
          type="button" onClick={() => handleTipo("reparacion")}
          className="flex-1 text-xs font-medium px-3 py-2 rounded-lg"
          style={tipo === "reparacion" ? { backgroundColor: ACCENT, color: "#FFFFFF" } : { backgroundColor: "#F2F3F4", color: MUTED }}
        >
          Reparación / Repuestos
        </button>
        <button
          type="button" onClick={() => handleTipo("mantenimiento")}
          className="flex-1 text-xs font-medium px-3 py-2 rounded-lg"
          style={tipo === "mantenimiento" ? { backgroundColor: ACCENT, color: "#FFFFFF" } : { backgroundColor: "#F2F3F4", color: MUTED }}
        >
          Mantenimiento
        </button>
      </div>

      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Cliente"><TextInput value={cliente} onChange={(e) => setCliente(e.target.value)} /></Field>
      <Field label="Obra"><TextInput value={obra} onChange={(e) => setObra(e.target.value)} /></Field>
      <Field label={tipo === "mantenimiento" ? "Equipo(s) a mantener" : "Equipo / Producto afectado"}>
        <TextInput
          value={equipoAfectado} onChange={(e) => setEquipoAfectado(e.target.value)}
          placeholder={tipo === "mantenimiento" ? "Ej: 3 splits, 1 termocalefón" : "Ej: Horno AE-AK630-9M-3G-CS-ON"}
        />
      </Field>
      <Field label={tipo === "mantenimiento" ? "Detalle / Observaciones" : "Falla reportada"}>
        <Textarea
          value={fallaReportada} onChange={(e) => setFallaReportada(e.target.value)}
          placeholder={tipo === "mantenimiento" ? "Observaciones adicionales (opcional)" : "Descripción de la falla que reportó el cliente"}
        />
      </Field>

      {tipo === "mantenimiento" && (
        <>
          <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Servicios de mantenimiento</p>
          <div className="flex gap-1.5 mb-2 flex-wrap">
            {MANTENIMIENTO_CATEGORIAS.map((cat) => (
              <button
                key={cat} type="button" onClick={() => handleCategoriaActiva(cat)}
                className="text-xs px-2.5 py-1.5 rounded-full"
                style={categoriaActiva === cat ? { backgroundColor: ACCENT, color: "#FFFFFF" } : { backgroundColor: "#F2F3F4", color: MUTED }}
              >
                {cat}
              </button>
            ))}
          </div>
          <div className="p-2.5 rounded mb-3" style={{ backgroundColor: "#F7F8FA" }}>
            {mttoInfo.conBtu && (
              <Field label="Capacidad (BTU)">
                <Select value={btuSel} onChange={(e) => handleBtuSel(Number(e.target.value))}>
                  {mttoInfo.opciones.map((o, i) => <option key={i} value={i}>{o.label}</option>)}
                </Select>
              </Field>
            )}
            <p className="text-xs mb-2" style={{ color: MUTED }}>{mttoInfo.descripcion}</p>
            <div className="flex gap-2">
              <Field label="Cantidad"><TextInput type="number" min="1" value={cantidadMtto} onChange={(e) => setCantidadMtto(e.target.value)} /></Field>
              <Field label="Precio unitario ₲"><TextInput type="number" value={precioMtto} onChange={(e) => setPrecioMtto(e.target.value)} /></Field>
            </div>
            <SecondaryButton onClick={agregarLineaMtto}><Plus size={14} /> Agregar mantenimiento</SecondaryButton>
          </div>

          {lineasMantenimiento.length > 0 && (
            <div className="mb-3 rounded border overflow-hidden" style={{ borderColor: BORDER }}>
              {lineasMantenimiento.map((l, i) => (
                <div key={i} className="flex items-center justify-between px-2.5 py-2 text-xs border-b last:border-0" style={{ borderColor: BORDER }}>
                  <div className="min-w-0">
                    <span className="font-medium" style={{ color: INK }}>{l.categoria}</span>
                    <span style={{ color: MUTED }}> · {l.detalle} · cant. {l.cantidad} × ₲ {Number(l.precioUnitario).toLocaleString()} = ₲ {(l.cantidad * l.precioUnitario).toLocaleString()}</span>
                  </div>
                  <button onClick={() => quitarLineaMtto(i)}><X size={13} style={{ color: MUTED }} /></button>
                </div>
              ))}
              <div className="px-2.5 py-2 text-xs font-semibold flex justify-between" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>
                <span>Subtotal mantenimiento</span><span>₲ {subtotalMtto.toLocaleString()}</span>
              </div>
            </div>
          )}
        </>
      )}

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Repuestos{tipo === "mantenimiento" ? " (opcional)" : ""}</p>
      <div className="p-2.5 rounded mb-3" style={{ backgroundColor: "#F7F8FA" }}>
        <Field label="Repuesto del catálogo">
          <Select value={productoId} onChange={(e) => handleProducto(e.target.value)}>
            <option value="">Seleccionar...</option>
            {productosPorGrupo.map(([grupo, items]) => (
              <optgroup key={grupo} label={grupo}>
                {items.map((p) => <option key={p.id} value={p.id}>{p.descripcion || p.nombre} — {p.subcategoria2}</option>)}
              </optgroup>
            ))}
          </Select>
        </Field>
        {productoSel && (
          <>
            <div className="flex gap-2">
              <Field label="Cantidad"><TextInput type="number" min="1" value={cantidadNueva} onChange={(e) => setCantidadNueva(e.target.value)} /></Field>
              <Field label="Costo x unidad U$S"><TextInput type="number" value={costoNuevo} onChange={(e) => setCostoNuevo(e.target.value)} /></Field>
            </div>
            <SecondaryButton onClick={agregarLinea}><Plus size={14} /> Agregar al presupuesto</SecondaryButton>
          </>
        )}
      </div>

      {lineas.length > 0 && (
        <div className="mb-3 rounded border overflow-hidden" style={{ borderColor: BORDER }}>
          {lineas.map((l, i) => (
            <div key={i} className="flex items-center justify-between px-2.5 py-2 text-xs border-b last:border-0" style={{ borderColor: BORDER }}>
              <div className="min-w-0">
                <span className="font-medium" style={{ color: INK }}>{l.item}</span>
                <span style={{ color: MUTED }}> · cant. {l.cantidad} × U$S {Number(l.costoUnitario).toLocaleString()} = U$S {(l.cantidad * l.costoUnitario).toLocaleString()}</span>
              </div>
              <button onClick={() => quitarLinea(i)}><X size={13} style={{ color: MUTED }} /></button>
            </div>
          ))}
          <div className="px-2.5 py-2 text-xs font-semibold flex justify-between" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>
            <span>Subtotal repuestos</span><span>U$S {subtotal.toLocaleString()}</span>
          </div>
        </div>
      )}

      {tipo === "reparacion" && (
        <>
          <label className="flex items-center gap-2 mb-3 text-sm" style={{ color: INK }}>
            <input type="checkbox" checked={incluirInstalacion} onChange={(e) => setIncluirInstalacion(e.target.checked)} />
            Incluir instalación
          </label>
          {incluirInstalacion && (
            <Field label="Instalación — monto U$S"><TextInput type="number" value={instalacionMonto} onChange={(e) => setInstalacionMonto(e.target.value)} placeholder="0" /></Field>
          )}
          {lineas.length > 0 && (
            <div className="px-2.5 py-2 mb-3 text-sm font-semibold flex justify-between rounded" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>
              <span>Total final</span>
              <span>U$S {(subtotal + (incluirInstalacion ? Number(instalacionMonto) || 0 : 0)).toLocaleString()}</span>
            </div>
          )}
        </>
      )}

      <Field label="Plazo estimado"><TextInput value={plazoEstimado} onChange={(e) => setPlazoEstimado(e.target.value)} /></Field>

      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar presupuesto</PrimaryButton>
    </div>
  );
}

function PresupuestosReparacionView({ presupuestos, query, onQuery, onNew, onDelete, onDescargarPdf, descargandoId, pdfError }) {
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Presupuestos de reparación</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>Reparación / repuestos y mantenimiento (AA, termocalefón, campana) — vía separada de las cotizaciones de venta.</p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
          <PrimaryButton onClick={onNew}><Plus size={15} /> Nuevo presupuesto</PrimaryButton>
        </div>
      </div>

      {pdfError && (
        <div className="mb-4 px-3 py-2 rounded-lg text-sm" style={{ backgroundColor: "#FBEAEA", color: "#B91C1C" }}>{pdfError}</div>
      )}

      {presupuestos.length === 0 ? (
        <EmptyState icon={Hammer} title="Todavía no hay presupuestos de reparación" subtitle="Usá el botón de arriba para armar el primero." />
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {presupuestos.map((p) => {
            const esMtto = p.tipo === "mantenimiento";
            const subtotal = (p.lineas || []).reduce((acc, l) => acc + (Number(l.cantidad) || 0) * (Number(l.costoUnitario) || 0), 0);
            const total = subtotal + (p.incluirInstalacion ? Number(p.instalacionMonto) || 0 : 0);
            const subtotalMtto = (p.lineasMantenimiento || []).reduce((acc, l) => acc + (Number(l.cantidad) || 0) * (Number(l.precioUnitario) || 0), 0);
            return (
              <div key={p.id} className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
                <div className="flex items-start justify-between mb-1">
                  <div>
                    <div className="flex items-center gap-1.5">
                      <p className="text-sm font-medium" style={{ color: INK }}>{p.cliente}</p>
                      <span
                        className="text-[10px] px-1.5 py-0.5 rounded-full font-medium"
                        style={esMtto ? { backgroundColor: "#EAF4EE", color: "#2F7A4A" } : { backgroundColor: ACCENT_LIGHT, color: ACCENT }}
                      >
                        {esMtto ? "Mantenimiento" : "Reparación"}
                      </span>
                    </div>
                    <p className="text-xs" style={{ color: MUTED }}>{p.obra} · {fmtDate(p.fecha)}</p>
                  </div>
                  <button onClick={() => onDelete(p.id)} className="p-1 rounded hover:bg-gray-100">
                    <Trash2 size={13} style={{ color: MUTED }} />
                  </button>
                </div>
                {p.equipoAfectado && <p className="text-xs mt-1" style={{ color: MUTED }}>{p.equipoAfectado}</p>}
                {esMtto ? (
                  <p className="text-sm mt-2" style={{ color: INK }}>
                    {(p.lineasMantenimiento || []).length} servicio(s) · ₲ {subtotalMtto.toLocaleString()}
                    {(p.lineas || []).length > 0 && ` · +${(p.lineas || []).length} repuesto(s) U$S ${subtotal.toLocaleString()}`}
                  </p>
                ) : (
                  <p className="text-sm mt-2" style={{ color: INK }}>{(p.lineas || []).length} repuesto(s) · U$S {total.toLocaleString()}</p>
                )}
                <button
                  onClick={() => onDescargarPdf(p)}
                  disabled={descargandoId === `${p.id}:reparacion`}
                  className="text-xs px-2.5 py-1.5 rounded flex items-center gap-1 mt-3"
                  style={{ backgroundColor: ACCENT, color: "#FFFFFF", opacity: descargandoId === `${p.id}:reparacion` ? 0.6 : 1 }}
                >
                  <Download size={13} /> {descargandoId === `${p.id}:reparacion` ? "Generando..." : "Descargar PDF"}
                </button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
