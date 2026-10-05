// Codificador Code 128 (sets B y C) — devuelve los anchos de barras/espacios en módulos, listo
// para dibujar. Mismo tipo de código lineal que usa la etiqueta de serie de fábrica.
const PATRONES = [
  "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213",
  "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132",
  "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211",
  "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
  "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331",
  "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111",
  "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214",
  "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
  "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141",
  "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141",
  "114131", "311141", "411131", "211412", "211214", "211232", "2331112",
];
const START_B = 104;
const START_C = 105;
const CODE_B = 100;
const CODE_C = 99;
const STOP = 106;

// Valores de símbolo para `texto` (ASCII 32–126), eligiendo set C para tramos largos de dígitos.
export function valoresCode128(texto) {
  const n = texto.length;
  for (let k = 0; k < n; k++) {
    const c = texto.charCodeAt(k);
    if (c < 32 || c > 126) throw new Error(`Carácter no soportado en Code 128: "${texto[k]}"`);
  }
  const digitos = (desde) => {
    let c = 0;
    while (desde + c < n && texto[desde + c] >= "0" && texto[desde + c] <= "9") c++;
    return c;
  };
  const vals = [];
  let modo = null;
  let i = 0;
  while (i < n) {
    let run = digitos(i);
    if (modo === "C") {
      if (run >= 2) {
        vals.push(parseInt(texto.substr(i, 2), 10));
        i += 2;
        continue;
      }
      vals.push(CODE_B);
      modo = "B";
      continue;
    }
    const usarC = run >= 4 && (modo === null || run >= 6 || i + run === n);
    if (usarC) {
      if (run % 2 === 1) {
        if (modo === null) { vals.push(START_B); modo = "B"; }
        vals.push(texto.charCodeAt(i) - 32);
        i++;
        run--;
        vals.push(CODE_C);
      } else {
        vals.push(modo === null ? START_C : CODE_C);
      }
      modo = "C";
      continue;
    }
    if (modo === null) { vals.push(START_B); modo = "B"; }
    vals.push(texto.charCodeAt(i) - 32);
    i++;
  }
  if (modo === null) vals.push(START_B);
  let suma = vals[0];
  for (let k = 1; k < vals.length; k++) suma += vals[k] * k;
  vals.push(suma % 103, STOP);
  return vals;
}

// Anchos alternados barra/espacio (empieza con barra), en módulos.
export function anchosCode128(texto) {
  const anchos = [];
  for (const v of valoresCode128(texto)) for (const ch of PATRONES[v]) anchos.push(Number(ch));
  return anchos;
}

export function modulosCode128(texto) {
  return anchosCode128(texto).reduce((a, b) => a + b, 0);
}
