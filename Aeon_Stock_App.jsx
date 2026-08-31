import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard, Package, ArrowUpFromLine, ArrowDownToLine, ShieldCheck,
  Wrench, Plus, Download, Search, X, Trash2, MessageCircle, AlertTriangle,
  CheckCircle2, Clock, ChevronRight, Boxes, Inbox, ArrowRight, Star, Lock, TrendingUp, Camera
} from "lucide-react";
import * as XLSX from "xlsx";

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

const STORAGE_KEYS = {
  equipos: "aeon-stock:equipos",
  movimientos: "aeon-stock:movimientos",
  entradas: "aeon-stock:entradas",
  ventas: "aeon-stock:ventas",
  repuestos: "aeon-stock:repuestos",
  playa: "aeon-stock:playa",
  ventasComprometidas: "aeon-stock:ventasComprometidas",
};

// ---------- Helpers ----------
function uid() {
  return Math.random().toString(36).slice(2, 10);
}
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
async function loadCollection(key) {
  try {
    const res = await window.storage.get(key, true);
    return res && res.value ? JSON.parse(res.value) : [];
  } catch (e) {
    return [];
  }
}
async function saveCollection(key, data) {
  try {
    await window.storage.set(key, JSON.stringify(data), true);
  } catch (e) {
    console.error("Storage error", key, e);
  }
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

function PrimaryButton({ children, onClick, type = "button" }) {
  return (
    <button
      type={type}
      onClick={onClick}
      className="inline-flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-md text-white"
      style={{ backgroundColor: ACCENT }}
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
  const [query, setQuery] = useState("");
  const [drawer, setDrawer] = useState(null);
  const [gestion, setGestion] = useState(null);
  const [fotoView, setFotoView] = useState(null);

  useEffect(() => {
    (async () => {
      const [eq, mv, en, vt, rp, pl, cm] = await Promise.all([
        loadCollection(STORAGE_KEYS.equipos),
        loadCollection(STORAGE_KEYS.movimientos),
        loadCollection(STORAGE_KEYS.entradas),
        loadCollection(STORAGE_KEYS.ventas),
        loadCollection(STORAGE_KEYS.repuestos),
        loadCollection(STORAGE_KEYS.playa),
        loadCollection(STORAGE_KEYS.ventasComprometidas),
      ]);
      setEquipos(eq);
      setMovimientos(mv);
      setEntradas(en);
      setVentas(vt);
      setRepuestos(rp);
      setPlaya(pl);
      setComprometidas(cm);
      setLoading(false);
    })();
  }, []);

  const persist = useCallback((key, setter, updater) => {
    setter((prev) => {
      const next = updater(prev);
      saveCollection(key, next);
      return next;
    });
  }, []);

  const addEquipo = (data) => persist(STORAGE_KEYS.equipos, setEquipos, (prev) => [{ id: uid(), ...data }, ...prev]);
  const updateEquipoEstado = (id, estado) => persist(STORAGE_KEYS.equipos, setEquipos, (prev) => prev.map((e) => (e.id === id ? { ...e, estado } : e)));
  const updateEquipoField = (id, field, value) => persist(STORAGE_KEYS.equipos, setEquipos, (prev) => prev.map((e) => (e.id === id ? { ...e, [field]: value } : e)));
  const deleteEquipo = (id) => persist(STORAGE_KEYS.equipos, setEquipos, (prev) => prev.filter((e) => e.id !== id));

  const addMovimiento = (data) => {
    // data: { fecha, categoria, sourceId, codigo, modelo, cantidad, motivo, cliente, remito, responsable, observaciones }
    persist(STORAGE_KEYS.movimientos, setMovimientos, (prev) => [{ id: uid(), ...data }, ...prev]);

    const cat = CATEGORIAS_ORIGEN.find((c) => c.value === data.categoria);
    if (!cat) return;
    const cantidadRetirada = Number(data.cantidad) || 1;

    if (cat.type === "equipo") {
      persist(STORAGE_KEYS.equipos, setEquipos, (prev) => {
        const next = [];
        for (const e of prev) {
          if (e.id === data.sourceId) {
            const restante = (Number(e.cantidad) || 1) - cantidadRetirada;
            if (restante > 0) next.push({ ...e, cantidad: restante });
          } else {
            next.push(e);
          }
        }
        return next;
      });
      const motivo = MOTIVOS_SALIDA.find((m) => m.value === data.motivo);
      if (motivo && motivo.trackea) {
        // El equipo sigue siendo un activo a rastrear (préstamo, sustitución, reparación): se le crea una nueva fila con nuevo código
        persist(STORAGE_KEYS.equipos, setEquipos, (prev) => [{
          id: uid(), codigo: nextCodigo(prev), serie: "", modelo: data.modelo,
          fechaIngreso: data.fecha, estado: motivo.estado,
          ubicacion: `Fuera de depósito (${data.cliente || "cliente"})`,
          cantidad: cantidadRetirada, notas: data.observaciones || "",
        }, ...prev]);
      }
    } else if (cat.type === "playa") {
      persist(STORAGE_KEYS.playa, setPlaya, (prev) => {
        const next = [];
        for (const p of prev) {
          if (p.id === data.sourceId) {
            const restante = (Number(p.cantidad) || 1) - cantidadRetirada;
            if (restante > 0) next.push({ ...p, cantidad: restante });
          } else next.push(p);
        }
        return next;
      });
    } else if (cat.type === "repuesto") {
      persist(STORAGE_KEYS.repuestos, setRepuestos, (prev) => {
        const next = [];
        for (const r of prev) {
          if (r.id === data.sourceId) {
            const restante = (Number(r.cantidad) || 1) - cantidadRetirada;
            if (restante > 0) next.push({ ...r, cantidad: restante });
          } else next.push(r);
        }
        return next;
      });
    }
  };
  const deleteMovimiento = (id) => persist(STORAGE_KEYS.movimientos, setMovimientos, (prev) => prev.filter((m) => m.id !== id));

  const addEntrada = (data) => {
    persist(STORAGE_KEYS.entradas, setEntradas, (prev) => [{ id: uid(), ...data }, ...prev]);
    if (data.codigo && data.estadoResultante) updateEquipoEstadoByCodigo(data.codigo, data.estadoResultante);
  };
  const deleteEntrada = (id) => persist(STORAGE_KEYS.entradas, setEntradas, (prev) => prev.filter((e) => e.id !== id));

  function updateEquipoEstadoByCodigo(codigo, estado) {
    persist(STORAGE_KEYS.equipos, setEquipos, (prev) => prev.map((e) => (e.codigo === codigo ? { ...e, estado } : e)));
  }

  const addVenta = (data) => persist(STORAGE_KEYS.ventas, setVentas, (prev) => [{ id: uid(), ...data }, ...prev]);
  const updateVentaEstado = (id, field, value) => persist(STORAGE_KEYS.ventas, setVentas, (prev) => prev.map((v) => (v.id === id ? { ...v, [field]: value } : v)));
  const deleteVenta = (id) => persist(STORAGE_KEYS.ventas, setVentas, (prev) => prev.filter((v) => v.id !== id));

  // Venta comprometida: reserva stock (queda físicamente en depósito, pero no disponible para otra salida)
  const addComprometida = (data) => {
    persist(STORAGE_KEYS.equipos, setEquipos, (prev) => prev.map((e) => (
      e.id === data.equipoId ? { ...e, comprometido: (Number(e.comprometido) || 0) + Number(data.cantidad) } : e
    )));
    persist(STORAGE_KEYS.ventasComprometidas, setComprometidas, (prev) => [{ id: uid(), estado: "Comprometida", ...data }, ...prev]);
  };

  const cancelarComprometida = (id) => {
    const c = comprometidas.find((x) => x.id === id);
    if (!c) return;
    if (c.estado === "Comprometida") {
      persist(STORAGE_KEYS.equipos, setEquipos, (prev) => prev.map((e) => (
        e.id === c.equipoId ? { ...e, comprometido: Math.max(0, (Number(e.comprometido) || 0) - Number(c.cantidad)) } : e
      )));
    }
    persist(STORAGE_KEYS.ventasComprometidas, setComprometidas, (prev) => prev.filter((x) => x.id !== id));
  };

  // Retirar: libera la reserva, descuenta stock real, deja el registro de Movimiento y marca la comprometida como cumplida
  const retirarComprometida = (id, remito, responsable) => {
    const c = comprometidas.find((x) => x.id === id);
    if (!c) return;
    const equipo = equipos.find((e) => e.id === c.equipoId);
    if (!equipo) return;

    persist(STORAGE_KEYS.equipos, setEquipos, (prev) => {
      const next = [];
      for (const e of prev) {
        if (e.id === c.equipoId) {
          const restante = (Number(e.cantidad) || 1) - Number(c.cantidad);
          const comprometidoRestante = Math.max(0, (Number(e.comprometido) || 0) - Number(c.cantidad));
          if (restante > 0) next.push({ ...e, cantidad: restante, comprometido: comprometidoRestante });
        } else next.push(e);
      }
      return next;
    });

    persist(STORAGE_KEYS.movimientos, setMovimientos, (prev) => [{
      id: uid(), fecha: todayISO(), categoria: "vendible", categoriaLabel: "Stock vendible",
      codigo: equipo.codigo, modelo: equipo.modelo, cantidad: c.cantidad, motivo: "Venta",
      cliente: c.razonSocial, obra: c.obra, monto: c.monto, remito: remito || "", responsable: responsable || "",
      observaciones: "Retiro de venta comprometida",
    }, ...prev]);

    persist(STORAGE_KEYS.ventasComprometidas, setComprometidas, (prev) => prev.map((x) => (
      x.id === id ? { ...x, estado: "Retirada", fechaRetiro: todayISO() } : x
    )));
  };

  const addRepuesto = (data) => persist(STORAGE_KEYS.repuestos, setRepuestos, (prev) => [{ id: uid(), ...data }, ...prev]);
  const deleteRepuesto = (id) => persist(STORAGE_KEYS.repuestos, setRepuestos, (prev) => prev.filter((r) => r.id !== id));

  const addPlaya = (data) => persist(STORAGE_KEYS.playa, setPlaya, (prev) => [{ id: uid(), estado: "En playa", ...data }, ...prev]);
  const deletePlaya = (id) => persist(STORAGE_KEYS.playa, setPlaya, (prev) => prev.filter((p) => p.id !== id));

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
      persist(STORAGE_KEYS.equipos, setEquipos, (prev) => [{
        id: uid(), codigo, serie: "", modelo: item.descripcion,
        fechaIngreso: item.fecha, estado: destino.estado,
        ubicacion: "Depósito principal",
        cantidad: (extra && extra.cantidad) || item.cantidad || 1,
        notas: (extra && extra.notas) || item.notas || "",
      }, ...prev]);
      persist(STORAGE_KEYS.entradas, setEntradas, (prev) => [{
        id: uid(), fecha: todayISO(), codigo,
        tipo: "Retorno de reparación propia", origen: `Zona de playa (${item.origen})`,
        motivo: `Derivado desde playa: ${item.descripcion}`,
        estadoResultante: destino.estado,
        responsable: "",
      }, ...prev]);
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
            onRetirar={retirarComprometida}
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
        <ComprometidaForm equipos={equipos} onSave={(d) => { addComprometida(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "playa"} onClose={() => setDrawer(null)} title="Nuevo ingreso a playa">
        <PlayaForm onSave={(d) => { addPlaya(d); setDrawer(null); }} />
      </Drawer>
      <Drawer open={drawer === "repuesto"} onClose={() => setDrawer(null)} title="Nuevo repuesto">
        <RepuestoForm onSave={(d) => { addRepuesto(d); setDrawer(null); }} />
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
function RetirarPopover({ item, onRetirar, onClose }) {
  const [remito, setRemito] = useState("");
  const [responsable, setResponsable] = useState("");
  return (
    <div className="mt-2 pt-2 border-t space-y-2" style={{ borderColor: BORDER }}>
      <TextInput value={remito} onChange={(e) => setRemito(e.target.value)} placeholder="N° de remito" />
      <TextInput value={responsable} onChange={(e) => setResponsable(e.target.value)} placeholder="Responsable" />
      <div className="flex gap-2">
        <PrimaryButton onClick={() => onRetirar(item.id, remito, responsable)}>Confirmar retiro</PrimaryButton>
        <SecondaryButton onClick={onClose}>Cancelar</SecondaryButton>
      </div>
    </div>
  );
}

function ComprometidasView({ comprometidas, query, onQuery, onNew, onCancelar, onRetirar }) {
  const [retirando, setRetirando] = useState(null);
  const pendientes = comprometidas.filter((c) => c.estado === "Comprometida");
  const totalMonto = pendientes.reduce((acc, c) => acc + (Number(c.monto) || 0), 0);

  return (
    <div>
      <div className="flex items-start justify-between mb-4 gap-4 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold" style={{ color: INK }}>Ventas comprometidas</h2>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>
            Mercadería vendida pero todavía en depósito — reserva el stock para que no se use en otra salida.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <SearchBox value={query} onChange={onQuery} />
          <PrimaryButton onClick={onNew}><Plus size={15} /> Nueva venta comprometida</PrimaryButton>
        </div>
      </div>

      {pendientes.length > 0 && (
        <div className="mb-4 px-3 py-2 rounded-lg text-sm" style={{ backgroundColor: ACCENT_LIGHT, color: ACCENT }}>
          {pendientes.length} venta(s) comprometida(s) pendiente(s) de retiro — total U$S {totalMonto.toLocaleString()}
        </div>
      )}

      {comprometidas.length === 0 ? (
        <EmptyState icon={Lock} title="No hay ventas comprometidas" subtitle="Cuando reservás mercadería vendida antes del retiro, va a aparecer acá." />
      ) : (
        <div className="grid grid-cols-2 gap-3">
          {comprometidas.map((c) => (
            <div key={c.id} className="rounded-lg p-3.5" style={{ backgroundColor: "#FFFFFF", border: `0.5px solid ${BORDER}` }}>
              <div className="flex items-start justify-between mb-1">
                <div>
                  <p className="text-sm font-medium" style={{ color: INK }}>{c.razonSocial}</p>
                  <p className="text-xs" style={{ color: MUTED }}>{c.obra}</p>
                </div>
                <span
                  className="text-xs px-2 py-0.5 rounded"
                  style={{
                    color: c.estado === "Retirada" ? "#15803D" : "#B45309",
                    backgroundColor: c.estado === "Retirada" ? "#E9F7EF" : "#FDF1E0",
                  }}
                >
                  {c.estado}
                </span>
              </div>
              <p className="text-sm mt-2" style={{ color: INK }}>{c.modelo} · cant. {c.cantidad}</p>
              <p className="text-xs mt-0.5" style={{ color: MUTED }}>
                Monto: U$S {Number(c.monto || 0).toLocaleString()} · Entrega estimada: {fmtDate(c.fechaEntrega)}
              </p>
              {c.estado === "Comprometida" && (
                retirando === c.id ? (
                  <RetirarPopover item={c} onRetirar={(id, r, resp) => { onRetirar(id, r, resp); setRetirando(null); }} onClose={() => setRetirando(null)} />
                ) : (
                  <div className="flex gap-2 mt-3">
                    <button onClick={() => setRetirando(c.id)} className="text-xs px-2.5 py-1.5 rounded" style={{ backgroundColor: ACCENT, color: "#FFFFFF" }}>
                      Marcar retirado
                    </button>
                    <button onClick={() => onCancelar(c.id)} className="text-xs px-2.5 py-1.5 rounded border" style={{ borderColor: BORDER, color: MUTED }}>
                      Cancelar
                    </button>
                  </div>
                )
              )}
            </div>
          ))}
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

function ComprometidaForm({ equipos, onSave }) {
  const [fecha, setFecha] = useState(todayISO());
  const [razonSocial, setRazonSocial] = useState("");
  const [obra, setObra] = useState("");
  const [equipoId, setEquipoId] = useState("");
  const [cantidad, setCantidad] = useState(1);
  const [monto, setMonto] = useState("");
  const [fechaEntrega, setFechaEntrega] = useState("");
  const [error, setError] = useState("");

  const vendibles = equipos.filter((e) => e.estado === "En depósito" || e.estado === "Apto para venta" || e.estado === "Apto para venta con descuento");
  const equipo = vendibles.find((e) => e.id === equipoId);
  const disponible = equipo ? Math.max(0, (Number(equipo.cantidad) || 1) - (Number(equipo.comprometido) || 0)) : 0;

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
        <Select value={equipoId} onChange={(e) => { setEquipoId(e.target.value); setCantidad(1); }}>
          <option value="">Seleccionar equipo...</option>
          {vendibles.map((eq) => {
            const libres = Math.max(0, (Number(eq.cantidad) || 1) - (Number(eq.comprometido) || 0));
            return <option key={eq.id} value={eq.id}>{eq.codigo} — {eq.modelo} (disponible: {libres})</option>;
          })}
        </Select>
      </Field>
      {equipo && (
        <Field label={`Cantidad a comprometer (disponible: ${disponible})`}>
          <TextInput type="number" min="1" max={disponible} value={cantidad} onChange={(e) => setCantidad(e.target.value)} />
        </Field>
      )}
      <Field label="Monto U$S"><TextInput type="number" value={monto} onChange={(e) => setMonto(e.target.value)} /></Field>
      <Field label="Fecha estimada de entrega"><TextInput type="date" value={fechaEntrega} onChange={(e) => setFechaEntrega(e.target.value)} /></Field>
      <p className="text-xs mb-3" style={{ color: MUTED }}>
        Esta cantidad queda reservada: no se va a poder retirar del depósito para otra salida hasta que la marques como retirada.
      </p>
      {error && <p className="text-xs mb-2" style={{ color: "#B91C1C" }}>{error}</p>}
      <PrimaryButton onClick={submit}>Guardar venta comprometida</PrimaryButton>
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
