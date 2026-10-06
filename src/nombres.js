// Nombres de empresa/cliente: mayúscula inicial en cada palabra, formas societarias (SA, SRL...)
// siempre en mayúsculas, y "CCI" siempre como "CCI SA". No baja a minúscula el resto de cada
// palabra, así las siglas ya escritas en mayúsculas no se rompen.
const FORMAS_SOCIETARIAS = new Set([
  "sa", "s.a", "s.a.", "srl", "s.r.l", "s.r.l.", "sae", "s.a.e", "s.a.e.", "saeca", "s.a.e.c.a.",
  "sas", "s.a.s", "s.a.s.", "sacif", "saci", "ltda", "ltda.", "eirl", "spa", "llc", "inc", "ltd", "gmbh",
]);

// Alias: nombres que en realidad son la misma empresa y se quieren siempre escritos igual.
const ALIAS_EMPRESA = [
  { patron: /^cci(\s+s\.?\s?a\.?)?$/i, nombre: "CCI SA" },
];

export function normalizarEmpresa(s) {
  if (typeof s !== "string") return s;
  const limpio = s.trim().replace(/\s+/g, " ");
  if (!limpio) return limpio;
  const palabras = limpio.split(" ").map((p) => {
    if (FORMAS_SOCIETARIAS.has(p.toLowerCase())) return p.toUpperCase();
    return p.charAt(0).toUpperCase() + p.slice(1);
  });
  const resultado = palabras.join(" ");
  const alias = ALIAS_EMPRESA.find((a) => a.patron.test(resultado));
  return alias ? alias.nombre : resultado;
}

// Campos que guardan el nombre de una empresa/cliente en las distintas colecciones.
export const CAMPOS_EMPRESA = ["cliente", "empresaCliente", "razonSocial", "clienteReal", "empresa"];

export function normalizarCamposEmpresa(obj) {
  if (!obj || typeof obj !== "object") return obj;
  let copia = null;
  for (const k of CAMPOS_EMPRESA) {
    if (typeof obj[k] === "string") {
      const n = normalizarEmpresa(obj[k]);
      if (n !== obj[k]) {
        if (!copia) copia = { ...obj };
        copia[k] = n;
      }
    }
  }
  return copia || obj;
}
