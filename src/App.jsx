import React, { useState, useEffect, useMemo } from "react";
import {
  LayoutDashboard, Package, ArrowUpFromLine, ArrowDownToLine, ShieldCheck,
  Wrench, Plus, Download, Search, X, Trash2, MessageCircle, AlertTriangle,
  CheckCircle2, Clock, ChevronRight, Boxes, Inbox, ArrowRight, Star, Lock, TrendingUp, Camera,
  Tag, FileText, FileSignature,
} from "lucide-react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
import {
  collection, doc, addDoc, updateDoc, deleteDoc, onSnapshot, query, orderBy,
} from "firebase/firestore";
import { downloadCotizacionPdf, downloadFichasTecnicasPdf } from "./pdf";

// ---------- Design tokens ----------
const INK = "#16202A";
const MUTED = "#64748B";
const ACCENT = "#1F3A5F";
const ACCENT_LIGHT = "#E8EEF5";
const BORDER = "#E2E6EB";
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
  { value: "repuesto", label: "Repuesto", estado: null },
];

// Categorías desde donde se puede dar salida a un producto — cada una filtra un pool distinto
const CATEGORIAS_ORIGEN = [
  { value: "muestras", label: "Muestras", type: "equipo", estados: ["Muestra"] },
  { value: "vendible", label: "Stock vendible", type: "equipo", estados: ["En depósito", "Apto para venta"] },
  { value: "vendible_desc", label: "Stock vendible con descuento", type: "equipo", estados: ["Apto para venta con descuento"] },
  { value: "recuperables", label: "Banco de recuperables", type: "equipo", estados: ["Pendiente de reparación", "En reparación", "Reservado - unidad de rescate"] },
  { value: "playa", label: "Zona de playa", type: "playa" },
  { value: "repuesto", label: "Repuesto", type: "repuesto" },
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
  repuestos: "repuestos",
  playa: "playa",
  ventasComprometidas: "ventasComprometidas",
  productos: "productos",
  cotizaciones: "cotizaciones",
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
function deleteItem(name, id) {
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
    <div className="overflow-x-auto rounded-lg border" style={{ borderColor: BORDER }}>
      <table className="w-full text-sm" style={{ minWidth: 720 }}>
        <thead>
          <tr style={{ backgroundColor: "#FAFBFC" }}>
            {columns.map((c) => (
              <th key={c.key} className="text-left font-medium px-3 py-2.5 border-b" style={{ color: MUTED, borderColor: BORDER, fontSize: 12 }}>
                {c.label}
              </th>
            ))}
            {onDelete && <th className="w-10 border-b" style={{ borderColor: BORDER }}></th>}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id} className="border-b last:border-0" style={{ borderColor: BORDER }}>
              {columns.map((c) => (
                <td key={c.key} className="px-3 py-2.5 align-top" style={{ color: INK }}>
                  {renderCell ? renderCell(c.key, row) : row[c.key] || "—"}
                </td>
              ))}
              {onDelete && (
                <td className="px-2 py-2.5 align-top">
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

function nextCodigo(equipos) {
  let max = 0;
  equipos.forEach((e) => {
    const m = /AEO-(\d+)/.exec(e.codigo || "");
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return `AEO-${String(max + 1).padStart(4, "0")}`;
}

// ---------- Main App ----------
export default function App() {
  const [tab, setTab] = useState("resumen");
  const [loading, setLoading] = useState(true);
  const [equipos, setEquipos] = useState([]);
  const [movimientos, setMovimientos] = useState([]);
  const [entradas, setEntradas] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [repuestos, setRepuestos] = useState([]);
  const [playa, setPlaya] = useState([]);
  const [comprometidas, setComprometidas] = useState([]);
  const [productos, setProductos] = useState([]);
  const [cotizaciones, setCotizaciones] = useState([]);
  const [query, setQuery] = useState("");
  const [drawer, setDrawer] = useState(null);
  const [gestion, setGestion] = useState(null);
  const [fotoView, setFotoView] = useState(null);
  const [retiroTarget, setRetiroTarget] = useState(null);
  const [descargandoId, setDescargandoId] = useState(null);
  const [pdfError, setPdfError] = useState("");

  useEffect(() => {
    const setters = {
      [COLLECTIONS.equipos]: setEquipos,
      [COLLECTIONS.movimientos]: setMovimientos,
      [COLLECTIONS.entradas]: setEntradas,
      [COLLECTIONS.ventas]: setVentas,
      [COLLECTIONS.repuestos]: setRepuestos,
      [COLLECTIONS.playa]: setPlaya,
      [COLLECTIONS.ventasComprometidas]: setComprometidas,
      [COLLECTIONS.productos]: setProductos,
      [COLLECTIONS.cotizaciones]: setCotizaciones,
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
          codigo: nextCodigo(equipos), serie: "", modelo: data.modelo,
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
    } else if (cat.type === "repuesto") {
      const source = repuestos.find((r) => r.id === data.sourceId);
      if (source) {
        const restante = (Number(source.cantidad) || 1) - cantidadRetirada;
        if (restante > 0) updateItem(COLLECTIONS.repuestos, source.id, { cantidad: restante });
        else deleteItem(COLLECTIONS.repuestos, source.id);
      }
    }
  };
  const deleteMovimiento = (id) => deleteItem(COLLECTIONS.movimientos, id);

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

  const addRepuesto = (data) => addItem(COLLECTIONS.repuestos, data);
  const deleteRepuesto = (id) => deleteItem(COLLECTIONS.repuestos, id);

  // Catálogo de productos: precio de lista y ficha técnica de cada modelo — de acá salen
  // el precio sugerido en ventas comprometidas y el contenido de las cotizaciones.
  const addProducto = (data) => addItem(COLLECTIONS.productos, data);
  const deleteProducto = (p) => deleteItem(COLLECTIONS.productos, p.id);
  const quitarFichaTecnica = (producto) =>
    updateItem(COLLECTIONS.productos, producto.id, { fichaTecnicaData: "", fichaTecnicaNombre: "" });

  const addCotizacion = (data) => addItem(COLLECTIONS.cotizaciones, data);
  const deleteCotizacion = (id) => deleteItem(COLLECTIONS.cotizaciones, id);

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

  const addPlaya = (data) => addItem(COLLECTIONS.playa, { estado: "En playa", ...data });
  const deletePlaya = (id) => deleteItem(COLLECTIONS.playa, id);

  const derivarPlaya = (item, destinoValue, extra) => {
    const destino = DESTINOS_PLAYA.find((d) => d.value === destinoValue);
    if (!destino) return;
    if (destino.value === "repuesto") {
      addRepuesto({
        nombre: extra.nombre || item.descripcion,
        modeloAsociado: extra.modeloAsociado || "",
        cantidad: extra.cantidad || 1,
        precio: extra.precio || 0,
      });
    } else {
      const codigo = nextCodigo(equipos);
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
    }
    deletePlaya(item.id);
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const wsEquipos = XLSX.utils.json_to_sheet(equipos.map((e) => ({
      "Código interno": e.codigo, "N° de serie": e.serie, "Modelo": e.modelo,
      "Fecha ingreso": e.fechaIngreso, "Estado actual": e.estado, "Ubicación": e.ubicacion,
      "Cantidad": e.cantidad || 1, "Comprometido": e.comprometido || 0,
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
    const wsRepuestos = XLSX.utils.json_to_sheet(repuestos.map((r) => ({
      "Nombre del repuesto": r.nombre, "Modelo asociado": r.modeloAsociado,
      "Cantidad": r.cantidad, "Precio U$S": r.precio,
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

  const filteredRepuestos = useMemo(() => {
    const q = query.toLowerCase();
    return repuestos.filter((r) => !q || [r.nombre, r.modeloAsociado].some((v) => (v || "").toLowerCase().includes(q)));
  }, [repuestos, query]);

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

  const NAV = [
    { key: "resumen", label: "Resumen", icon: LayoutDashboard },
    { key: "playa", label: "Zona de playa", icon: Inbox },
    { key: "equipos", label: "Maestro de equipos", icon: Package },
    { key: "comprometidas", label: "Ventas comprometidas", icon: Lock },
    { key: "movimientos", label: "Salidas", icon: ArrowUpFromLine },
    { key: "entradas", label: "Entradas", icon: ArrowDownToLine },
    { key: "ventas", label: "Ventas y garantías", icon: ShieldCheck },
    { key: "recuperables", label: "Banco de recuperables", icon: Wrench },
    { key: "muestras", label: "Muestras", icon: Star },
    { key: "repuestos", label: "Repuestos", icon: Boxes },
    { key: "catalogo", label: "Catálogo de productos", icon: Tag },
    { key: "cotizaciones", label: "Cotizaciones", icon: FileSignature },
  ];

  if (loading) {
    return (
      <div className="flex items-center justify-center py-24" style={{ backgroundColor: BG }}>
        <p className="text-sm" style={{ color: MUTED }}>Cargando datos...</p>
      </div>
    );
  }

  return (
    <div className="flex w-full" style={{ backgroundColor: BG, minHeight: 560, fontFamily: "system-ui, -apple-system, sans-serif" }}>
      {/* Sidebar */}
      <div className="w-56 shrink-0 border-r flex flex-col" style={{ borderColor: BORDER, backgroundColor: "#FFFFFF" }}>
        <div className="px-4 py-4 border-b" style={{ borderColor: BORDER }}>
          <p className="text-sm font-bold tracking-wide" style={{ color: ACCENT }}>AEON</p>
          <p className="text-xs" style={{ color: MUTED }}>Control de stock</p>
        </div>
        <nav className="flex-1 py-2">
          {NAV.map((n) => {
            const Icon = n.icon;
            const active = tab === n.key;
            return (
              <button
                key={n.key}
                onClick={() => { setTab(n.key); setQuery(""); }}
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
      <div className="flex-1 min-w-0 p-6">
        {tab === "resumen" && (
          <Resumen
            equipos={equipos} proximosServices={proximosServices} alertasContacto={alertasContacto}
            seguimientosPendientes={seguimientosPendientes} recuperables={recuperables}
            playa={playa} muestras={muestras} repuestos={repuestos} ventasCerradas={ventasCerradas}
            onNavigate={setTab}
          />
        )}

        {tab === "playa" && (
          <PlayaView
            playa={filteredPlaya} query={query} onQuery={setQuery}
            onNew={() => setDrawer("playa")}
            onDerivar={derivarPlaya}
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
          />
        )}

        {tab === "movimientos" && (
          <Section
            title="Salidas de depósito"
            subtitle="Elegís de qué categoría sale el producto — la cantidad disponible se descuenta ahí mismo."
            query={query} onQuery={setQuery}
            onNew={() => setDrawer("movimiento")}
            newLabel="Nueva salida"
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

        {tab === "repuestos" && (
          <Section
            title="Repuestos"
            subtitle="Piezas sueltas en stock, asociadas al modelo de equipo que corresponden."
            query={query} onQuery={setQuery}
            onNew={() => setDrawer("repuesto")}
            newLabel="Nuevo repuesto"
          >
            <Table
              columns={[
                { key: "nombre", label: "Repuesto" }, { key: "modeloAsociado", label: "Modelo asociado" },
                { key: "cantidad", label: "Cantidad" }, { key: "precio", label: "Precio U$S" },
              ]}
              rows={filteredRepuestos}
              onDelete={deleteRepuesto}
              renderCell={(key, row) => {
                if (key === "precio") return `$${Number(row.precio || 0).toFixed(2)}`;
                return row[key] || "—";
              }}
            />
          </Section>
        )}

        {tab === "catalogo" && (
          <CatalogoView
            productos={filteredProductos} query={query} onQuery={setQuery}
            onNew={() => setDrawer("producto")}
            onDelete={deleteProducto}
            onQuitarFicha={quitarFichaTecnica}
          />
        )}

        {tab === "cotizaciones" && (
          <CotizacionesView
            cotizaciones={filteredCotizaciones} query={query} onQuery={setQuery}
            onNew={() => setDrawer("cotizacion")}
            onDelete={deleteCotizacion}
            onDescargarPdf={handleDescargarPdf}
            onDescargarFichas={handleDescargarFichas}
            descargandoId={descargandoId}
            pdfError={pdfError}
          />
        )}
      </div>

      {/* Drawers */}
      <Drawer open={drawer === "equipo"} onClose={() => setDrawer(null)} title="Nuevo equipo">
        <EquipoForm sugerido={nextCodigo(equipos)} onSave={(d) => { addEquipo(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "movimiento"} onClose={() => setDrawer(null)} title="Nueva salida">
        <MovimientoForm equipos={equipos} playa={playa} repuestos={repuestos} onSave={(d) => { addMovimiento(d); setDrawer(null); }} />
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
      <Drawer open={drawer === "repuesto"} onClose={() => setDrawer(null)} title="Nuevo repuesto">
        <RepuestoForm onSave={(d) => { addRepuesto(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "producto"} onClose={() => setDrawer(null)} title="Nuevo producto">
        <ProductoForm onSave={(d) => { addProducto(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "cotizacion"} onClose={() => setDrawer(null)} title="Nueva cotización">
        <CotizacionForm productos={productos} onSave={(d) => { addCotizacion(d); setDrawer(null); }} />
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
      <PhotoViewer src={fotoView} onClose={() => setFotoView(null)} />
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

function Section({ title, subtitle, query, onQuery, onNew, newLabel, children }) {
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>{title}</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>{subtitle}</p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
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

function Resumen({ equipos, proximosServices, alertasContacto, seguimientosPendientes, recuperables, playa, muestras, repuestos, ventasCerradas, onNavigate }) {
  const vendible = equipos.filter((e) => e.estado === "En depósito" || e.estado === "Apto para venta");
  const bajas = equipos.filter((e) => e.estado === "Dado de baja");
  const totalUnidades = sumCantidad(equipos.filter((e) => e.estado !== "Dado de baja"));
  const totalRepuestos = sumCantidad(repuestos);
  const totalPlaya = sumCantidad(playa);
  const totalVendido = ventasCerradas.reduce((acc, v) => acc + (Number(v.cantidad) || 0), 0);

  const cardsActivos = [
    { label: "Equipos totales activos", value: totalUnidades, icon: Package, tab: "equipos" },
    { label: "Zona de playa (sin clasificar)", value: totalPlaya, icon: Inbox, tab: "playa" },
    { label: "Stock vendible", value: sumCantidad(vendible), icon: ArrowDownToLine, tab: "equipos" },
    { label: "Banco de recuperables", value: sumCantidad(recuperables), icon: Wrench, tab: "recuperables" },
    { label: "Muestras", value: sumCantidad(muestras), icon: Star, tab: "muestras" },
    { label: "Repuestos (unidades)", value: totalRepuestos, icon: Boxes, tab: "repuestos" },
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
      <div className="grid grid-cols-3 gap-3 mb-4">
        {cardsActivos.map((c) => (
          <StatCard key={c.label} label={c.label} value={c.value} icon={c.icon} onClick={() => onNavigate(c.tab)} />
        ))}
      </div>

      <p className="text-xs font-medium mb-2" style={{ color: MUTED }}>Historial (ya no cuenta como stock)</p>
      <div className="grid grid-cols-3 gap-3 mb-6">
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

      <div className="grid grid-cols-2 gap-4">
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

function PlayaCard({ item, onDerivar, onDelete }) {
  const [destino, setDestino] = useState("");
  const [nombre, setNombre] = useState(item.descripcion);
  const [modeloAsociado, setModeloAsociado] = useState("");
  const [cantidad, setCantidad] = useState(item.cantidad || 1);
  const [precio, setPrecio] = useState("");
  const [error, setError] = useState("");

  const confirmar = () => {
    if (!destino) {
      setError("Elegí un destino primero.");
      return;
    }
    onDerivar(item, destino, { nombre, modeloAsociado, cantidad: Number(cantidad) || 1, precio: Number(precio) || 0 });
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
      {item.notas && <p className="text-xs mt-1 mb-2" style={{ color: MUTED }}>{item.notas}</p>}
      <div className="flex items-center gap-2 mt-2">
        <Select value={destino} onChange={(e) => { setDestino(e.target.value); setError(""); }} style={{ ...inputStyle, padding: "4px 8px", fontSize: 12 }}>
          <option value="">Derivar a...</option>
          {DESTINOS_PLAYA.map((d) => <option key={d.value} value={d.value}>{d.label}</option>)}
        </Select>
        <button onClick={confirmar} className="p-1.5 rounded shrink-0" style={{ backgroundColor: ACCENT }}>
          <ArrowRight size={14} color="#FFFFFF" />
        </button>
      </div>
      {destino === "repuesto" && (
        <div className="mt-3 pt-3 border-t space-y-2" style={{ borderColor: BORDER }}>
          <TextInput value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Nombre del repuesto" />
          <TextInput value={modeloAsociado} onChange={(e) => setModeloAsociado(e.target.value)} placeholder="Modelo asociado" />
          <div className="flex gap-2">
            <TextInput type="number" value={cantidad} onChange={(e) => setCantidad(e.target.value)} placeholder="Cantidad" />
            <TextInput type="number" value={precio} onChange={(e) => setPrecio(e.target.value)} placeholder="Precio U$S" />
          </div>
        </div>
      )}
      {error && <p className="text-xs mt-2" style={{ color: "#B91C1C" }}>{error}</p>}
    </div>
  );
}

function PlayaView({ playa, query, onQuery, onNew, onDerivar, onDelete }) {
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
        <div className="grid grid-cols-2 gap-3">
          {playa.map((item) => (
            <PlayaCard key={item.id} item={item} onDerivar={onDerivar} onDelete={onDelete} />
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
        <div className="grid grid-cols-2 gap-3">
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

function ComprometidasView({ comprometidas, query, onQuery, onNew, onCancelar, onRetirar, onCerrar }) {
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
        <div className="grid grid-cols-2 gap-3">
          {comprometidas.map((c) => {
            const retirado = Number(c.cantidadRetirada) || 0;
            const saldo = Math.max(0, (Number(c.cantidad) || 0) - retirado);
            const badge = COMPROMETIDA_BADGE[c.estado] || COMPROMETIDA_BADGE.Comprometida;
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
function ComentarioEditor({ value, onSave }) {
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
      placeholder="Ej: en muestra por defecto de pintura en la puerta"
      className="w-full text-xs px-2 py-1.5 rounded border outline-none"
      style={{ borderColor: editing ? ACCENT : BORDER, color: INK }}
    />
  );
}

function MuestrasView({ muestras, query, onQuery, onUpdateField }) {
  const filtered = muestras.filter((e) => {
    const q = query.toLowerCase();
    return !q || [e.codigo, e.modelo].some((v) => (v || "").toLowerCase().includes(q));
  });
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Muestras</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>
            Equipos usados como pieza de exhibición — en depósito o prestados a un cliente.
          </p>
        </div>
        <SearchBox value={query} onChange={onQuery} />
      </div>
      {filtered.length === 0 ? (
        <EmptyState icon={Star} title="No hay muestras cargadas" subtitle="Los equipos clasificados como Muestra van a aparecer acá." />
      ) : (
        <div className="rounded-lg border overflow-hidden" style={{ borderColor: BORDER }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ backgroundColor: "#FAFBFC" }}>
                {["Producto", "Código", "Cantidad", "Estado", "Comentario"].map((h) => (
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
function EquipoForm({ sugerido, onSave }) {
  const [codigo, setCodigo] = useState(sugerido);
  const [serie, setSerie] = useState("");
  const [modelo, setModelo] = useState("");
  const [fechaIngreso, setFechaIngreso] = useState(todayISO());
  const [estado, setEstado] = useState("En depósito");
  const [ubicacion, setUbicacion] = useState("");
  const [cantidad, setCantidad] = useState(1);
  const [notas, setNotas] = useState("");
  const [motivoBaja, setMotivoBaja] = useState("");
  const [error, setError] = useState("");

  const submit = () => {
    if (!codigo.trim() || !modelo.trim()) {
      setError("Completá al menos código y modelo.");
      return;
    }
    onSave({ codigo, serie, modelo, fechaIngreso, estado, ubicacion, cantidad: Number(cantidad) || 1, notas, motivoBaja: estado === "Dado de baja" ? motivoBaja : "" });
  };

  return (
    <div>
      <Field label="Código interno"><TextInput value={codigo} onChange={(e) => setCodigo(e.target.value)} /></Field>
      <Field label="N° de serie"><TextInput value={serie} onChange={(e) => setSerie(e.target.value)} placeholder="N° de serie de fábrica" /></Field>
      <Field label="Modelo"><TextInput value={modelo} onChange={(e) => setModelo(e.target.value)} placeholder="Ej: AE-AK630-9M-3G-CS-ON" /></Field>
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

function MovimientoForm({ equipos, playa, repuestos, onSave }) {
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
    if (cat.type === "repuesto") return repuestos;
    return [];
  }, [cat, equipos, playa, repuestos]);

  const source = opciones.find((o) => o.id === sourceId);
  const comprometido = source && cat.type === "equipo" ? (Number(source.comprometido) || 0) : 0;
  const disponible = source ? Math.max(0, (Number(source.cantidad) || 1) - comprometido) : 0;

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
    const modelo = cat.type === "equipo" ? source.modelo : cat.type === "playa" ? source.descripcion : source.nombre;
    const codigo = cat.type === "equipo" ? source.codigo : (source.codigo || modelo);
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
              const label = cat.type === "equipo" ? `${o.codigo} — ${o.modelo}` : cat.type === "playa" ? o.descripcion : `${o.nombre}${o.modeloAsociado ? " — " + o.modeloAsociado : ""}`;
              const libres = cat.type === "equipo" ? Math.max(0, (Number(o.cantidad) || 1) - (Number(o.comprometido) || 0)) : (o.cantidad || 1);
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
        Una vez cargado, vas a poder derivarlo a Banco de recuperables, Equipo de socorro, Stock vendible, Muestra o Repuestos.
      </p>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar en playa</PrimaryButton>
    </div>
  );
}

function RepuestoForm({ onSave }) {
  const [nombre, setNombre] = useState("");
  const [modeloAsociado, setModeloAsociado] = useState("");
  const [cantidad, setCantidad] = useState(1);
  const [precio, setPrecio] = useState("");
  const [error, setError] = useState("");

  const submit = () => {
    if (!nombre.trim()) {
      setError("Ingresá el nombre del repuesto.");
      return;
    }
    onSave({ nombre, modeloAsociado, cantidad: Number(cantidad) || 1, precio: Number(precio) || 0 });
  };

  return (
    <div>
      <Field label="Nombre del repuesto"><TextInput value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej: Placa electrónica" /></Field>
      <Field label="Modelo asociado"><TextInput value={modeloAsociado} onChange={(e) => setModeloAsociado(e.target.value)} placeholder="Ej: Horno AE-AK630-9M-3G-CS-ON" /></Field>
      <Field label="Cantidad"><TextInput type="number" value={cantidad} onChange={(e) => setCantidad(e.target.value)} /></Field>
      <Field label="Precio U$S"><TextInput type="number" value={precio} onChange={(e) => setPrecio(e.target.value)} /></Field>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar repuesto</PrimaryButton>
    </div>
  );
}

// ---------- Catálogo de productos ----------
function ProductoForm({ onSave }) {
  const [nombre, setNombre] = useState("");
  const [categoria, setCategoria] = useState("");
  const [descripcion, setDescripcion] = useState("");
  const [especLabel, setEspecLabel] = useState("");
  const [especValor, setEspecValor] = useState("");
  const [precioLista, setPrecioLista] = useState("");
  const [foto, setFoto] = useState("");
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
      }
      onSave({
        nombre, categoria, descripcion, especLabel, especValor,
        precioLista: Number(precioLista) || 0, foto, ...ficha,
      });
    } catch (err) {
      setError("No se pudo leer la ficha técnica. Probá de nuevo.");
      setGuardando(false);
    }
  };

  return (
    <div>
      <Field label="Nombre / código del producto"><TextInput value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej: AE-AK630-9M-3G-CS-ON" /></Field>
      <Field label="Categoría"><TextInput value={categoria} onChange={(e) => setCategoria(e.target.value)} placeholder="Ej: Electrodomésticos, Aires acondicionados..." /></Field>
      <Field label="Descripción (aparece en la cotización)"><Textarea value={descripcion} onChange={(e) => setDescripcion(e.target.value)} placeholder="Descripción técnica completa del producto" /></Field>
      <div className="flex gap-2">
        <Field label="Especificación — etiqueta"><TextInput value={especLabel} onChange={(e) => setEspecLabel(e.target.value)} placeholder="Ej: Capacidad BTU" /></Field>
        <Field label="Especificación — valor"><TextInput value={especValor} onChange={(e) => setEspecValor(e.target.value)} placeholder="Ej: 12.000" /></Field>
      </div>
      <Field label="Precio de lista U$S"><TextInput type="number" value={precioLista} onChange={(e) => setPrecioLista(e.target.value)} /></Field>

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

      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit} disabled={guardando}>{guardando ? "Guardando..." : "Guardar producto"}</PrimaryButton>
    </div>
  );
}

function CatalogoView({ productos, query, onQuery, onNew, onDelete, onQuitarFicha }) {
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Catálogo de productos</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>
            Precio de lista y ficha técnica de cada modelo — de acá sale el precio sugerido en ventas comprometidas y el contenido de las cotizaciones.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
          <PrimaryButton onClick={onNew}><Plus size={15} /> Nuevo producto</PrimaryButton>
        </div>
      </div>

      {productos.length === 0 ? (
        <EmptyState icon={Tag} title="Todavía no hay productos cargados" subtitle="Usá el botón de arriba para cargar el primero." />
      ) : (
        <div className="grid grid-cols-3 gap-3">
          {productos.map((p) => (
            <div key={p.id} className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
              <div className="flex items-start gap-3 mb-2">
                {p.foto ? (
                  <img src={p.foto} alt={p.nombre} className="rounded border shrink-0" style={{ width: 56, height: 56, objectFit: "cover", borderColor: BORDER }} />
                ) : (
                  <div className="rounded border shrink-0 flex items-center justify-center" style={{ width: 56, height: 56, borderColor: BORDER, backgroundColor: "#FAFBFC" }}>
                    <Tag size={20} style={{ color: MUTED }} />
                  </div>
                )}
                <div className="min-w-0 flex-1">
                  <div className="flex items-start justify-between gap-2">
                    <CodeTag>{p.nombre}</CodeTag>
                    <button onClick={() => onDelete(p)} className="p-1 rounded hover:bg-gray-100 shrink-0">
                      <Trash2 size={13} style={{ color: MUTED }} />
                    </button>
                  </div>
                  {p.categoria && <p className="text-xs mt-1" style={{ color: MUTED }}>{p.categoria}</p>}
                </div>
              </div>
              <p className="text-sm font-medium" style={{ color: INK }}>U$S {Number(p.precioLista || 0).toLocaleString()}</p>
              {p.especLabel && <p className="text-xs mt-0.5" style={{ color: MUTED }}>{p.especLabel}: {p.especValor}</p>}
              {p.descripcion && <p className="text-xs mt-1.5 line-clamp-2" style={{ color: MUTED }}>{p.descripcion}</p>}
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
          ))}
        </div>
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
  const [categoria, setCategoria] = useState("");
  const [comentarios, setComentarios] = useState("");
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
    onSave({ fecha, cliente, obra, categoria, comentarios, fechaEntregaEstimada, formaPago, obs, lineas });
  };

  return (
    <div>
      <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Cliente"><TextInput value={cliente} onChange={(e) => setCliente(e.target.value)} /></Field>
      <Field label="Obra"><TextInput value={obra} onChange={(e) => setObra(e.target.value)} /></Field>
      <Field label="Categoría (título de la cotización)"><TextInput value={categoria} onChange={(e) => setCategoria(e.target.value)} placeholder="Ej: AIRES ACONDICIONADOS" /></Field>

      <p className="text-xs font-semibold mt-4 mb-2" style={{ color: ACCENT }}>Productos</p>
      <div className="p-2.5 rounded mb-3" style={{ backgroundColor: "#F7F8FA" }}>
        <Field label="Producto del catálogo">
          <Select value={productoId} onChange={(e) => handleProducto(e.target.value)}>
            <option value="">Seleccionar...</option>
            {productos.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
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

function CotizacionesView({ cotizaciones, query, onQuery, onNew, onDelete, onDescargarPdf, onDescargarFichas, descargandoId, pdfError }) {
  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Cotizaciones</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>Armá una cotización con productos del catálogo y descargala en PDF, con tu firma.</p>
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
        <div className="grid grid-cols-2 gap-3">
          {cotizaciones.map((c) => {
            const subtotal = (c.lineas || []).reduce((acc, l) => acc + (Number(l.cantidad) || 0) * (Number(l.precioUnit) || 0), 0);
            const tieneFichas = (c.lineas || []).some((l) => l.fichaTecnicaData);
            return (
              <div key={c.id} className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
                <div className="flex items-start justify-between mb-1">
                  <div>
                    <p className="text-sm font-medium" style={{ color: INK }}>{c.cliente}</p>
                    <p className="text-xs" style={{ color: MUTED }}>{c.obra} · {fmtDate(c.fecha)}</p>
                  </div>
                  <button onClick={() => onDelete(c.id)} className="p-1 rounded hover:bg-gray-100">
                    <Trash2 size={13} style={{ color: MUTED }} />
                  </button>
                </div>
                {c.categoria && <CodeTag>{c.categoria}</CodeTag>}
                <p className="text-sm mt-2" style={{ color: INK }}>{(c.lineas || []).length} producto(s) · U$S {subtotal.toLocaleString()}</p>
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
          })}
        </div>
      )}
    </div>
  );
}
