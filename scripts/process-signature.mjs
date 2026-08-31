// Turns the photographed signature into a transparent-background PNG for
// use on quotation PDFs. Samples the paper color from the four corners,
// then fades every pixel's alpha by its color distance from that
// background — background pixels vanish, ink stays opaque, edges stay
// anti-aliased instead of a hard cutout. Then trims to the ink's bounds.
import sharp from "sharp";
import { mkdirSync, existsSync } from "fs";

const SRC = "public/firma-gaston.jpg";
const OUT_DIR = "public/generated";
const OUT = `${OUT_DIR}/firma.png`;

const LOW = 15;  // distance below this -> fully transparent (background)
const HIGH = 70; // distance above this -> fully opaque (ink)

if (!existsSync(SRC)) {
  console.log("no signature file found at", SRC, "- skipping");
  process.exit(0);
}

mkdirSync(OUT_DIR, { recursive: true });

const { data, info } = await sharp(SRC).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
const { width, height, channels } = info;

function pixelAt(x, y) {
  const i = (y * width + x) * channels;
  return [data[i], data[i + 1], data[i + 2]];
}
const corners = [pixelAt(2, 2), pixelAt(width - 3, 2), pixelAt(2, height - 3), pixelAt(width - 3, height - 3)];
const bg = [0, 1, 2].map((c) => Math.round(corners.reduce((s, p) => s + p[c], 0) / corners.length));

const out = Buffer.from(data);
let minX = width, maxX = 0, minY = height, maxY = 0;
for (let y = 0; y < height; y++) {
  for (let x = 0; x < width; x++) {
    const i = (y * width + x) * channels;
    const r = data[i], g = data[i + 1], b = data[i + 2];
    const dist = Math.sqrt((r - bg[0]) ** 2 + (g - bg[1]) ** 2 + (b - bg[2]) ** 2);
    const alpha = Math.max(0, Math.min(255, Math.round(((dist - LOW) / (HIGH - LOW)) * 255)));
    out[i + 3] = alpha;
    if (alpha > 20) {
      if (x < minX) minX = x;
      if (x > maxX) maxX = x;
      if (y < minY) minY = y;
      if (y > maxY) maxY = y;
    }
  }
}

const pad = 6;
const left = Math.max(0, minX - pad);
const top = Math.max(0, minY - pad);
const cropW = Math.min(width - left, maxX - minX + pad * 2);
const cropH = Math.min(height - top, maxY - minY + pad * 2);

await sharp(out, { raw: { width, height, channels } })
  .extract({ left, top, width: cropW, height: cropH })
  .png()
  .toFile(OUT);

console.log("signature processed ->", OUT, `(${cropW}x${cropH})`);
