import { PDFDocument } from "pdf-lib";

// Firestore guarda el documento entero (todos los campos del producto/venta juntos) en un límite
// de 1 MB — un PDF en base64 pesa ~37% más que el archivo original, así que cualquier adjunto de
// más de ~650 KB corre riesgo real de hacer fallar el guardado (en silencio, sin aviso claro).
// Esta misma lógica ya se usó (y se probó) en la herramienta de carga masiva de fichas técnicas
// (public/herramientas/fichas-tecnicas.html) — acá queda disponible para cualquier adjunto de la
// app en uso normal: ficha técnica de un producto, contrato firmado, factura.
const TARGET_BYTES = 665600;
const PDFJS_VERSION = "3.11.174";

let pdfjsPromise = null;
function cargarPdfjs() {
  if (typeof window !== "undefined" && window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (!pdfjsPromise) {
    pdfjsPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.min.js`;
      script.onload = () => {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.js`;
        resolve(window.pdfjsLib);
      };
      script.onerror = () => reject(new Error("No se pudo cargar el compresor de PDF (sin conexión a internet)."));
      document.head.appendChild(script);
    });
  }
  return pdfjsPromise;
}

function arrayBufferABase64(buf) {
  let binary = "";
  const bytes = new Uint8Array(buf);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

// Rasteriza cada página (vía pdf.js) y arma un PDF nuevo con esas páginas como imagen JPEG (vía
// pdf-lib) — pierde texto seleccionable y algo de nitidez, pero reduce drásticamente el tamaño.
// Reintenta bajando calidad/resolución hasta 4 veces si todavía no entra en el objetivo.
async function rasterizarYComprimir(arrayBuffer) {
  const pdfjsLib = await cargarPdfjs();
  let quality = 0.72, scale = 1.6, out = null;
  for (let intento = 0; intento < 4; intento++) {
    const pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer.slice(0) }).promise;
    const outDoc = await PDFDocument.create();
    for (let i = 1; i <= pdfDoc.numPages; i++) {
      const page = await pdfDoc.getPage(i);
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      const ctx = canvas.getContext("2d");
      await page.render({ canvasContext: ctx, viewport }).promise;
      const jpegUrl = canvas.toDataURL("image/jpeg", quality);
      const jpegBytes = Uint8Array.from(atob(jpegUrl.split(",")[1]), (c) => c.charCodeAt(0));
      const img = await outDoc.embedJpg(jpegBytes);
      const outPage = outDoc.addPage([viewport.width, viewport.height]);
      outPage.drawImage(img, { x: 0, y: 0, width: viewport.width, height: viewport.height });
    }
    out = await outDoc.save();
    if (out.byteLength <= TARGET_BYTES) break;
    quality = Math.max(0.35, quality - 0.15);
    scale = Math.max(0.9, scale - 0.3);
  }
  return out;
}

// Punto de entrada único: siempre devuelve un data URL de PDF listo para guardar en Firestore.
// Si el archivo ya entra en el límite, lo deja intacto (texto seleccionable, calidad original).
export async function prepararPdfParaFirestore(file) {
  const buf = await file.arrayBuffer();
  if (buf.byteLength <= TARGET_BYTES) {
    return { dataUrl: `data:application/pdf;base64,${arrayBufferABase64(buf)}`, comprimido: false };
  }
  const comprimido = await rasterizarYComprimir(buf);
  return {
    dataUrl: `data:application/pdf;base64,${arrayBufferABase64(comprimido.buffer || comprimido)}`,
    comprimido: true,
    bytesOriginal: buf.byteLength,
    bytesFinal: comprimido.byteLength,
  };
}
